from pathlib import Path
from collections import Counter
import hashlib,json,re,datetime
from PIL import Image
repo=Path.cwd().resolve();work=repo/'tests/.work';rows=[];coverage=[];groups={}
roots=[work/'T19-dirty-ps51-4fa6813564b540e38022b969e512e000',work/'T19-dirty-ps7-75b7d014511d4f928f8dc7e25acf6f89']
sha=lambda b:hashlib.sha256(b).hexdigest()
for root in roots:
    text=(root/'PreservationNative.stdout.txt').read_text(encoding='utf-8-sig')
    p=Path(re.findall(r'^Preservation native observations: (.+)$',text,re.M)[0].strip());d=json.loads(p.read_text(encoding='utf-8-sig'))
    for row in d['Observations'][:4]:
        records=row.get('Originals',[]) if row['Label']=='original-corpus' else [row[k] for k in ['Master','Email'] if row[k]]
        for record in records:
            s=record['Snapshot'];label=root.name+'/'+row['Label']+'/'+Path(record['Pdf']).name
            rows.append({'label':label,'pdf_sha256':s['file']['sha256'],'bytes':s['file']['bytes'],'pages':s['page_identifiers'],'fields':s['forms']['canonical_fields'],'widgets':len(s['forms']['widgets']),'relations':s['forms']['relations'],'forms_issues':s['forms']['issues'],'named_destinations':s['named_destinations'],'bookmarks':s['bookmarks'],'annotation_subtypes':dict(Counter(a['subtype'] for page in s['pages'] for a in page['annotations'])),'embedded_files':s['embedded_files'],'tagging':s['tagging'],'rotations':[page['structural']['rotation_degrees'] for page in s['pages']],'geometry':[page['pdfium']['size_points'] for page in s['pages']],'parser':s['parser'],'snapshot':record['SnapshotPath']})
            for render in s['renders']:
                f=Path(render['path']);im=Image.open(f).convert('RGB');pixelsha=sha(im.tobytes());group=groups.setdefault(pixelsha,{'representative':str(f),'pixel_sha256':pixelsha,'width':im.width,'height':im.height,'members':[]});group['members'].append(label+'/'+render['identifier']);coverage.append({'pdf':label,'page':render['page_index'],'identifier':render['identifier'],'png_path':str(f),'png_sha256':sha(f.read_bytes()),'pixel_sha256':pixelsha})
out={'task':'T19','scope':'dirty native outputs only; visual inspection pending','created_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'rows':rows,'coverage':coverage,'pixel_groups':list(groups.values())}
(work/'T19-dirty-output-summary.json').write_text(json.dumps(out,indent=2)+'\n')
for row in rows:
    print(row['label'],row['bytes'],'fields',[(f.get('name'),f.get('value')) for f in row['fields']],'widgets',row['widgets'],'named',[(x['name'],x['target']['page_index']) for x in row['named_destinations']],'bmk',[(x['title'],x['page_index']) for x in row['bookmarks']],'annots',row['annotation_subtypes'],'embedded',len(row['embedded_files']),'tagroot',row['tagging']['root_present'],'tagassoc',len(row['tagging']['mcid_associations']),'rotation',row['rotations'])
print('coverage',len(coverage),'unique pixel groups',len(groups));print(json.dumps([{'path':g['representative'],'members':len(g['members'])} for g in groups.values()],indent=2))
