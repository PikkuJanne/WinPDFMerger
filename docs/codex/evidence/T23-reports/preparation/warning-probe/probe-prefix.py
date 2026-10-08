import pathlib,re
p=pathlib.Path(__file__).with_name('probe.py')
s=p.read_text(encoding='utf-8-sig')
s=s.replace("variants={", """prefix=b'T23 original synthetic warning prefix\\n'
def prepend_adjusted(blob,prefix):
    xref=blob.index(b'xref\\n')
    before,after=blob[:xref],blob[xref:]
    after=re.sub(rb'(\\d{10})( 00000 n)',lambda m:('%010d'%(int(m[1])+len(prefix))).encode()+m[2],after)
    after=re.sub(rb'startxref\\s+(\\d+)',lambda m:b'startxref\\n'+str(int(m[1])+len(prefix)).encode(),after)
    return prefix+before+after
variants={\n 'leading_garbage_adjusted':prepend_adjusted(base,prefix),\n 'leading_garbage_original_offsets':prefix+base,""")
exec(compile(s,str(p),'exec'))
