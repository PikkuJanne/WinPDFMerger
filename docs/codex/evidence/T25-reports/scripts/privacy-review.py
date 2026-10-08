from __future__ import annotations
import collections, hashlib, json, re, subprocess, sys
from pathlib import Path

repo = Path(sys.argv[1]).resolve()
out = Path(sys.argv[2]).resolve()
def git(*args):
    return subprocess.check_output(['git','-C',str(repo),*args], stderr=subprocess.PIPE)
def textgit(*args):
    return git(*args).decode('utf-8','replace').strip()
head = textgit('rev-parse','HEAD')
commits = textgit('rev-list','--all').splitlines()
objectlines = git('rev-list','--objects','--all').splitlines()
objects = [line.split(b' ',1)[0].decode() for line in objectlines]
history_paths = collections.defaultdict(set)
modes = collections.Counter()
for commit in commits:
    for entry in git('ls-tree','-r','-z',commit).split(b'\0'):
        if not entry: continue
        desc, rawpath = entry.split(b'\t',1)
        mode, kind, oid = desc.decode().split()
        path = rawpath.decode('utf-8','replace')
        history_paths[oid].add(path)
        modes[mode] += 1
index_entries=[]
for entry in git('ls-files','--stage','-z').split(b'\0'):
    if not entry: continue
    desc, rawpath = entry.split(b'\t',1)
    mode, oid, stage = desc.decode().split()
    index_entries.append((mode,oid,stage,rawpath.decode('utf-8','replace')))

patterns = {
 'private_key_material': re.compile(rb'-----BEGIN (?:RSA |EC |DSA |OPENSSH |ENCRYPTED )?PRIVATE KEY-----'),
 'github_token': re.compile(rb'(?:gh[pousr]_[A-Za-z0-9]{30,255}|github_pat_[A-Za-z0-9_]{40,255})'),
 'aws_access_key': re.compile(rb'\b(?:AKIA|ASIA)[A-Z0-9]{16}\b'),
 'openai_api_key': re.compile(rb'\bsk-(?:proj-|svcacct-)?[A-Za-z0-9_-]{35,255}\b'),
 'slack_token': re.compile(rb'\bxox[baprs]-[A-Za-z0-9-]{20,255}\b'),
 'azure_sas_secret': re.compile(rb'(?i)[?&](?:sig)=[A-Za-z0-9%+/]{20,300}'),
 'credentialed_url': re.compile(rb'(?i)https?://[^\s<>"\x27/@:]{1,200}:[^\s<>"\x27/@]{1,200}@'),
 'literal_password_or_token_assignment': re.compile(rb'(?i)(?:password|passwd|secret|api[_-]?key|access[_-]?token|auth[_-]?token)\s*[=:]\s*["\x27][^"\x27\r\n]{4,200}["\x27]'),
 'local_user_profile_path': re.compile(rb'(?i)\b[A-Z]:[\\/]{1,2}(?:Users|Documents and Settings)[\\/]{1,2}[^\\/\s"\x27<>]+'),
}
def scan_matches(data):
    lower=data.lower()
    gates={
      'private_key_material': b'PRIVATE KEY-----' in data,
      'github_token': b'gh' in data or b'github_pat_' in data,
      'aws_access_key': b'AKIA' in data or b'ASIA' in data,
      'openai_api_key': b'sk-' in data,
      'slack_token': b'xox' in data,
      'azure_sas_secret': b'sig=' in lower,
      'credentialed_url': b'http' in lower,
      'literal_password_or_token_assignment': any(marker in lower for marker in (b'password',b'passwd',b'secret',b'api',b'token')),
      'local_user_profile_path': any(marker in lower for marker in (b':\\users\\',b':\\\\users\\\\',b':/users/',b':\\documents and settings\\',b':\\\\documents and settings\\\\',b':/documents and settings/')),
    }
    return {name:len(pattern.findall(data)) for name,pattern in patterns.items() if gates[name] and pattern.search(data)}

blob_reports=[]
blob_data={}
p = subprocess.Popen(['git','-C',str(repo),'cat-file','--batch'],stdin=subprocess.PIPE,stdout=subprocess.PIPE,stderr=subprocess.PIPE)
assert p.stdin and p.stdout
for oid in objects:
    p.stdin.write((oid+'\n').encode()); p.stdin.flush()
    header=p.stdout.readline().decode().split()
    if len(header)!=3: raise RuntimeError('Unexpected object reader response')
    _,kind,size=header; data=p.stdout.read(int(size)); terminator=p.stdout.read(1)
    if terminator!=b'\n': raise RuntimeError('Object reader framing mismatch')
    if kind!='blob': continue
    blob_data[oid]=data
    matches=scan_matches(data)
    matches={name:count for name,count in matches.items() if count}
    if matches:
        # Known source names are safe; unclassified names are represented by path hashes only.
        paths=sorted(history_paths[oid])
        safe_paths=[path if path.startswith(('docs/codex/','tests/','tools/','.github/','src/')) or path in ('README.md','LICENSE','WinPDFMerge.ps1','WinPDFMerge.bat','.gitignore','AGENTS.md','SECURITY.md') else 'unclassified-sha256:'+hashlib.sha256(path.encode()).hexdigest() for path in paths]
        blob_reports.append({'blob':oid,'path_classes':safe_paths,'categories':matches})
p.stdin.close(); p.wait(timeout=30)
if p.returncode: raise RuntimeError('Object reader failed')
paths=set(path for paths in history_paths.values() for path in paths)
exts=collections.Counter(Path(path).suffix.lower() or '[no-extension]' for path in paths)
blocked_suffixes={'.exe','.dll','.msi','.pfx','.p12','.pem','.key','.doc','.docx','.xls','.xlsx','.ppt','.pptx','.msg','.eml','.zip','.7z','.rar'}
blocked=[path for path in paths if Path(path).suffix.lower() in blocked_suffixes]
pdf_paths=sorted(path for path in paths if Path(path).suffix.lower()=='.pdf')
pdf_results=[]
for path in pdf_paths:
    variants=[{'blob':oid,'sha256':hashlib.sha256(blob_data[oid]).hexdigest(),'bytes':len(blob_data[oid])} for oid, names in history_paths.items() if path in names and oid in blob_data]
    pdf_results.append({'path_class':path if path.startswith('tests/fixtures/numbered/') else 'unclassified-sha256:'+hashlib.sha256(path.encode()).hexdigest(),'variants':variants})
magic=collections.Counter()
for oid,data in blob_data.items():
    if data.startswith(b'MZ'): magic['PE_or_DOS_binary']+=1
    if data.startswith(b'PK\x03\x04'): magic['zip_container']+=1
    if data.startswith(b'%PDF'): magic['PDF']+=1
    if data.startswith(b'\x89PNG\r\n\x1a\n'): magic['PNG']+=1
    if data.startswith(b'\x00\x00\x01\x00'): magic['ICO']+=1
    if b'\x00' in data and not data.startswith(b'%PDF'): magic['other_NUL_containing']+=1
current_worktree=[]
for _,_,_,path in index_entries:
    file=repo/path
    if not file.is_file(): continue
    data=file.read_bytes()
    findings=scan_matches(data)
    if findings: current_worktree.append({'path':path,'categories':findings})
report={
 'schema_version':1,'review_class':'read-only repository/CI/privacy review, heuristic not comprehensive secret certification',
 'commit_under_review':head,'reachable_commits':len(commits),'reachable_objects':len(objects),'unique_reachable_blobs':len(blob_data),
 'total_reachable_blob_bytes':sum(map(len,blob_data.values())),'unique_historical_paths':len(paths),'index_entries':len(index_entries),'staged_changed_paths':len(git('diff','--cached','--name-only','-z').split(b'\0'))-1,
 'historical_path_suffix_counts':dict(sorted(exts.items())),'historical_tree_entry_mode_counts':dict(modes),
 'prohibited_suffix_paths_count':len(blocked),'pdf_history':pdf_results,'magic_counts':dict(magic),
 'history_candidate_blob_count':len(blob_reports),'history_candidates':blob_reports,'current_worktree_candidates':current_worktree,
 'pattern_descriptions':list(patterns),
 'scan_scope':['all local refs returned by git rev-list --all; unique reachable Git blobs scanned as raw bytes','every historical tree scanned for file paths/modes','current tracked files read from worktree','all index entries, with staged diff count separately'],
 'limitations':['heuristic patterns can miss unrecognized credentials or encoded/binary content; every candidate requires contextual review','unreachable Git objects, reflogs, remote refs not present locally, fork content and third-party service logs/artifacts are outside this content scan','no secret scanner installed; no history mutation; no runtime tests or package built'],
 'commands':['git rev-parse HEAD','git rev-list --all','git rev-list --objects --all','git ls-tree -r -z <each reachable commit>','git ls-files --stage -z','git cat-file --batch','git diff --cached --name-only -z']
}
out.mkdir(parents=True,exist_ok=True)
(out/'repository-scan.json').write_text(json.dumps(report,indent=2)+'\n',encoding='utf-8')
print(json.dumps({key:report[key] for key in ('commit_under_review','reachable_commits','reachable_objects','unique_reachable_blobs','unique_historical_paths','index_entries','staged_changed_paths','historical_path_suffix_counts','prohibited_suffix_paths_count','pdf_history','magic_counts','history_candidate_blob_count')},indent=2))
print('Sanitized candidate inventory retained in external review report; no candidate values printed.')