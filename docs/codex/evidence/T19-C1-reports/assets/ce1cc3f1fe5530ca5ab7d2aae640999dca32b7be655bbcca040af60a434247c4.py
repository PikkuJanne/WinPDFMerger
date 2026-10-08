"""Read-only original feature oracle validation; does not invoke the application."""
from pathlib import Path
import argparse, datetime, hashlib, json, os, shutil, subprocess, sys, uuid

def sha(path): return hashlib.sha256(path.read_bytes()).hexdigest()

def main():
    parser=argparse.ArgumentParser();parser.add_argument('--receipt',type=Path,required=True);args=parser.parse_args()
    repo=Path.cwd().resolve();work=repo/'tests/.work';receipt=json.loads(args.receipt.read_text(encoding='utf-8-sig'))
    model=receipt['model'];corpus=Path(receipt['corpus']);generator=repo/'tests/fixtures/features/generate_features.py'
    oracle=repo/'tools/test/feature_oracle.py';manifest=repo/'tests/fixtures/features/manifest.json'
    assert model['generator_sha256']==sha(generator)==receipt['generator_source_sha256']
    assert model['versions']=={'python':'3.12.14','reportlab':'4.4.9','pypdf':'6.10.0'}
    assert not manifest.exists(), 'Bind the actual initial manifest once; preserve any existing bytes.'
    manifest.write_text(json.dumps(model,indent=2,ensure_ascii=False)+'\n',encoding='utf-8')
    root=work/('T19-original-oracle-'+uuid.uuid4().hex);root.mkdir();sources=root/'sources';sources.mkdir()
    for source in [generator,oracle,manifest,Path(__file__),args.receipt]:shutil.copyfile(source,sources/source.name)
    results=[];checks=[];failure=None
    def require(condition,message):
        checks.append({'check':message,'pass':bool(condition)})
        if not condition:raise AssertionError(message)
    try:
        for fixture in model['fixtures']:
            pdf=corpus/fixture['file'];before=sha(pdf);name=pdf.stem;output=root/(name+'.json');renders=root/(name+'-renders')
            argv=[sys.executable,'-B',str(oracle),'--pdf',str(pdf),'--output',str(output),'--render-dir',str(renders),'--dpi','144']
            stdout=root/(name+'.stdout.txt');stderr=root/(name+'.stderr.txt');started=datetime.datetime.now(datetime.timezone.utc).isoformat()
            with stdout.open('xb') as out,stderr.open('xb') as err:
                code=subprocess.run(argv,cwd=repo,env={k:v for k,v in os.environ.items() if k.casefold()!='psmodulepath'},stdin=subprocess.DEVNULL,stdout=out,stderr=err,timeout=120).returncode
            row={'fixture':fixture['file'],'argv':argv,'exit_code':code,'started_at_utc':started,'finished_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),
                 'stdout':str(stdout),'stderr':str(stderr),'stdout_sha256':sha(stdout),'stderr_sha256':sha(stderr),'source_before_sha256':before,'source_after_sha256':sha(pdf),'snapshot':str(output) if output.exists() else None}
            results.append(row);require(code==0,'Actual read-only oracle process exits zero: '+name)
            snapshot=json.loads(output.read_text(encoding='utf-8'));row['snapshot_sha256']=sha(output)
            require(before==fixture['sha256']==sha(pdf)==snapshot['file']['sha256'],'Original source hash preserved and bound: '+name)
            require(pdf.stat().st_size==fixture['bytes']==snapshot['file']['bytes'],'Actual source bytes match manifest: '+name)
            require(snapshot['parser']=={'strict':True,'warnings':[]},'Strict pypdf original has no observed repair warning: '+name)
            require(snapshot['page_count']==2 and snapshot['page_identifiers']==fixture['page_identifiers'],'Two actual PDFium page IDs in manifest order: '+name)
            require([p['structural']['rotation_degrees'] for p in snapshot['pages']]==fixture['structural_rotations'],'Structural original rotations: '+name)
            require([p['pdfium']['rotation_degrees'] for p in snapshot['pages']]==fixture['structural_rotations'],'Independent PDFium rotations: '+name)
            require([p['structural']['media_box'] for p in snapshot['pages']]==fixture['media_boxes'],'Original media boxes: '+name)
            require([p['pdfium']['size_points'] for p in snapshot['pages']]==[[612.0,792.0],[792.0,612.0]],'Independent PDFium rotated dimensions: '+name)
            form=snapshot['forms'];field=fixture['canonical_field'];require(form['acroform_present'] and form['need_appearances'] is False,'Canonical AcroForm with explicit appearances: '+name)
            require(len(form['canonical_fields'])==1 and form['canonical_fields'][0]['full_name']==field['name'] and form['canonical_fields'][0]['value']==field['value'],'One canonical original field and distinct value: '+name)
            require(len(form['widgets'])==field['widget_count'] and len(form['relations'])==2 and all(all(v for k,v in r.items() if k not in ['widget_ref','canonical_field_ref']) for r in form['relations']),'Widget Parent/Kids/P/effective-value associations: '+name)
            require(not form['issues'] and not form['duplicate_canonical_names'],'No original canonical tree issue or duplicate local name: '+name)
            require(all(w['appearance']['present'] and w['appearance']['normal']['text_literals']==[field['value']] for w in form['widgets']),'Widget normal AP stream contains expected original value: '+name)
            named=snapshot['named_destinations'];dest=fixture['named_destination'];require(len(named)==1 and named[0]['name']==dest['name'] and named[0]['target']['page_index']==dest['page_index'] and named[0]['target']['page_identifier']==fixture['page_identifiers'][1],'Named destination resolves actual second page: '+name)
            require([b['title'] for b in snapshot['bookmarks']]==fixture['bookmark_titles'] and [b['page_index'] for b in snapshot['bookmarks']]==[0,1],'Original outline hierarchy resolves actual pages: '+name)
            from collections import Counter
            annots=[a for page in snapshot['pages'] for a in page['annotations']];require(dict(Counter(a['subtype'] for a in annots))==fixture['annotation_counts'],'Original annotation subtype counts: '+name)
            require(sum(a['uri']==fixture['uri'] for a in annots)==2 and sum(a['destination'] is not None and a['destination']['page_index']==1 for a in annots)==2,'URI records and internal link page targets remain separate: '+name)
            attachment=fixture['attachment'];embedded=snapshot['embedded_files'];require(len(embedded)==1 and embedded[0]['name']==attachment['name'] and not snapshot['embedded_file_issues'],'Canonical original EmbeddedFiles entry: '+name)
            require(all(s['decoded_sha256']==attachment['sha256'] and s['decoded_bytes']==attachment['bytes'] for s in embedded[0]['streams'].values()) and len(embedded[0]['streams'])==2,'Decoded attachment bytes and distinct legacy/Unicode references: '+name)
            icon=[a['file_attachment'] for a in annots if a['subtype']=='/FileAttachment'];require(len(icon)==1 and icon[0]['file_spec_ref']==embedded[0]['file_spec_ref'],'Page FileAttachment references canonical embedded filespec: '+name)
            tags=snapshot['tagging'];require(tags['root_present'] and tags['marked'] is True and not tags['issues'],'Minimal original structure tree present: '+name)
            require(len(tags['mcid_associations'])==2 and all(a['element_in_structure_tree'] and a['element_content_reference_matches'] and a['element_page_index']==a['page_index'] for a in tags['mcid_associations']),'MCID slot/ParentTree/structure element page relationships: '+name)
            require([p['content_mcids'] for p in snapshot['pages']]==[[{'tag':'/P','mcid':0}],[{'tag':'/P','mcid':0}]],'Original marked content matches minimal manifest: '+name)
            require(snapshot['signature_observations']=={'signature_field_count':0,'xfa_present':False,'doc_mdp_present':False,'validation_performed':False},'No fabricated signature, XFA or validity evidence: '+name)
            require(snapshot['active_content_observations']=={'annotation_action_types':['/URI'],'catalog_open_action_present':False,'catalog_additional_actions_present':False,'javascript_name_tree_present':False,'uris_followed':False,'javascript_platform_provided':False,'xfa_render_disabled':True},'No JavaScript/open action/XFA execution and no URI follow: '+name)
            require(len(snapshot['renders'])==2 and all(Path(r['path']).is_file() and sha(Path(r['path']))==r['sha256'] for r in snapshot['renders']),'Actual PDFium render files hash-bound; no visual-pass claim: '+name)
            row['renders']=snapshot['renders']
    except Exception as exc:failure=type(exc).__name__+': '+str(exc)
    summary={'schema_version':1,'task':'T19','result':'fail' if failure else 'pass','partial':len(results)!=2,'source_receipt':str(args.receipt.resolve()),'source_receipt_sha256':sha(args.receipt),
             'generator_sha256':sha(generator),'oracle_sha256':sha(oracle),'manifest_sha256':sha(manifest),'command_producer_sha256':sha(Path(__file__)),
             'checks':checks,'check_count':len(checks),'results':results,'failure':failure,'scope':'Actual original PDFs only; read-only strict pypdf/PDFium characterization and render preparation. Not application, engine preservation, manual visual or accessibility/signature acceptance.'}
    (root/'receipt.json').write_text(json.dumps(summary,indent=2,ensure_ascii=False)+'\n',encoding='utf-8')
    print(json.dumps({'receipt':str(root/'receipt.json'),'sha256':sha(root/'receipt.json'),**summary},ensure_ascii=False))
    return 1 if failure else 0

if __name__=='__main__':raise SystemExit(main())
