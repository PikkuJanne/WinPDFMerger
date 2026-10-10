"""Preserve the reviewed initial writer and derive prose/count-only V2."""
from pathlib import Path
import difflib, hashlib, json, re

p = Path('tests/.work/T32-WriteCompletion.py')
old = p.read_text(encoding='utf-8')
assert hashlib.sha256(p.read_bytes()).hexdigest() == '1d9760b565a347ae9a857513bc08024a835630b91441a1bb182f6b3dd671dacc'
def prose(match):
    content = match.group(1)
    content = re.sub(r'(?<=[0-9])(?=[A-Za-z])(?![a-fA-F0-9]+\b)', ' ', content)
    for a,b in [('tree5014f5b','tree `5014f5b`'),('Clean detached-R','Clean detached checkout at R'),
                ('pinnedPS','pinned PS'),('actualPS','actual PS'),('Both pre-tag','Both pre-tag'),
                ('exactR','exact R'),('PS5.1','PS 5.1'),('PS7.6.6','PS 7.6.6'),
                ('PDFtk2.02','PDFtk 2.02'),('Ghostscript10.08.0','Ghostscript 10.08.0'),
                ('GS10.08.0','GS 10.08.0'),('T33publication','T33 publication'),
                ('T34synchronized','T34 synchronized'),('T33 finalpublication','T33 final publication'),
                ('required Windows shells/native','required Windows shells and native'),
                ('GateE','Gate E'),('GateF','Gate F'),('BOTH','both'),('notpublished','not published'),
                ('remainnot_run','remain not_run'),('published_atnull','published_at null'),
                ('exacttwoassets','exactly two assets'),('silent retry','silent retry'),
                ('silentretry','silent retry'),('intermediateversion','intermediate version'),
                ('falsecompletion','false completion'),('neverpass','never pass'),
                ('nonrequired/\nunperformed','nonrequired/\nunperformed'),('full regression1072','full regression 1072')]:
        content = content.replace(a,b)
    return "f'''" + content + "'''"
new = re.sub(r"f'''(.*?)'''", prose, old, flags=re.S)
needle = "'decoded_image_checks':images['checks'],'PDFs':21,'PDF_pages':106,'issues':[]},"
assert needle in new
new = new.replace(needle, "'decoded_image_checks':images['checks'],'PDFs':21,'PDF_pages':106,'issues':[],\n"
    "                              'independent_draft_checks':draft['checks_total'],\n"
    "                              'independent_downloaded_package_checks':draft['fresh_downloaded_package_audit']['checks']},")
new = new.replace('Authenticated producer download and separate independent download/byte review match',
    'Authenticated producer download and separate independent download review (30 checks)\nplus downloaded-package byte inspection (266 checks) match')
q = p.with_name('T32-WriteCompletionV2.py')
assert not q.exists() and new != old
q.write_text(new,encoding='utf-8')
q.with_suffix('.diff.txt').write_text(''.join(difflib.unified_diff(old.splitlines(True),new.splitlines(True),fromfile=p.name,tofile=q.name)),encoding='utf-8')
print(json.dumps({'original_sha256':hashlib.sha256(p.read_bytes()).hexdigest(),'V2_sha256':hashlib.sha256(q.read_bytes()).hexdigest(),'scope':'Human prose spacing and actual draft/download audit count summaries only; execution/writes unperformed.'}))
