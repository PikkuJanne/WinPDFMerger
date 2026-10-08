import hashlib, io, json, random, re, sys, zlib
from pathlib import Path
from contextlib import closing

def make_pdf(path, identifier, raster=False):
    content=(f'BT /F1 16 Tf 24 260 Td ({identifier}) Tj ET\n').encode()
    objects=[b'<< /Type /Catalog /Pages 2 0 R >>',b'<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
             b'<< /Type /Page /Parent 2 0 R /MediaBox [0 0 432 288] /Resources << /Font << /F1 4 0 R >>'+(b' /XObject << /Im0 6 0 R >>' if raster else b'')+b' >> /Contents 5 0 R >>',
             b'<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>']
    if raster: content+=b'q 384 0 0 220 24 24 cm /Im0 Do Q\n'
    objects.append(b'<< /Length '+str(len(content)).encode()+b' >>\nstream\n'+content+b'endstream')
    if raster:
        pixels=random.Random(230053+int(identifier.split('-')[-1])).randbytes(1200*800*3)
        image=zlib.compress(pixels,9)
        objects.append(b'<< /Type /XObject /Subtype /Image /Width 1200 /Height 800 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /FlateDecode /Length '+str(len(image)).encode()+b' >>\nstream\n'+image+b'\nendstream')
    buffer=io.BytesIO();buffer.write(b'%PDF-1.4\n%\xe2\xe3\xcf\xd3\n');offsets=[0]
    for number,obj in enumerate(objects,1):
        offsets.append(buffer.tell());buffer.write(str(number).encode()+b' 0 obj\n'+obj+b'\nendobj\n')
    xref=buffer.tell();buffer.write(f'xref\n0 {len(objects)+1}\n0000000000 65535 f \n'.encode())
    for offset in offsets[1:]:buffer.write(f'{offset:010} 00000 n \n'.encode())
    buffer.write(f'trailer\n<< /Size {len(objects)+1} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n'.encode())
    raw=buffer.getvalue();path.write_bytes(raw)
    return {'file':path.name,'identifier':identifier,'raster':raster,'bytes':len(raw),'sha256':hashlib.sha256(raw).hexdigest()}

if sys.argv[1]=='generate':
    root=Path(sys.argv[2]);root.mkdir()
    near=root/'many-inputs';mixed=root/'mixed-inputs';near.mkdir();mixed.mkdir()
    many=[make_pdf(near/(f'document{n:04}-'+('x'*70)+'.pdf'),f'T23-LIMIT-{n:04}') for n in range(1,181)]
    mix=[make_pdf(mixed/f'{n}.pdf',f'T23-MIXED-{n:04}',n%4==0) for n in range(1,25)]
    base=Path(sys.argv[3]).read_bytes()
    if hashlib.sha256(base).hexdigest()!='ed0457c1d675cfc502f96ee5f9a9b4fd0dddc4010a21af13c891528089be2109':raise ValueError('Original numbered fixture pin differs.')
    prefix=b'T23 original synthetic warning prefix\n';xref=base.index(b'xref\n');before,after=base[:xref],base[xref:]
    after=re.sub(rb'(\d{10})( 00000 n)',lambda m:f'{int(m[1])+len(prefix):010}'.encode()+m[2],after)
    after=re.sub(rb'startxref\s+(\d+)',lambda m:b'startxref\n'+str(int(m[1])+len(prefix)).encode(),after)
    warning=prefix+before+after;(root/'warning-input.pdf').write_bytes(warning)
    print(json.dumps({'provenance':'Original stdlib-only vector text and deterministic noise rasters; no external or private documents','seed_base':230053,'near':many,'mixed':mix,'warning':{'file':'warning-input.pdf','bytes':len(warning),'sha256':hashlib.sha256(warning).hexdigest(),'prefix_hex':prefix.hex(),'prefix_bytes':len(prefix),'scope':'Native helper only: deliberately nonconformant header prefix, every live xref and startxref offset adjusted; actual application envelope guard refuses this input.'}}));raise SystemExit(0)
if sys.argv[1]=='inspect':
    import pypdfium2 as pdfium
    if sys.version.split()[0]!='3.12.14' or str(pdfium.PYPDFIUM_INFO)!='5.13.0' or str(pdfium.PDFIUM_INFO)!='153.0.7999.0':raise RuntimeError('Use approved independent PDFium versions.')
    path=Path(sys.argv[2]);expected=json.loads(Path(sys.argv[3]).read_text(encoding='utf-8-sig'));before=path.read_bytes();pages=[]
    with pdfium.PdfDocument(io.BytesIO(before)) as document:
        if len(document)!=len(expected):raise ValueError('Independent page count differs.')
        for n in range(len(document)):
            with closing(document[n]) as page, closing(page.get_textpage()) as text:
                found=re.findall(r'T23-(?:LIMIT|MIXED)-[0-9]{4}|T03-[0-9]{2}-P[0-9]{2}',text.get_text_range())
                if found!=[expected[n]]:raise ValueError(f'Visible page identity/order differs at {n+1}: {found}, expected {expected[n]}')
                pages.append({'identifier':found[0],'rotation':page.get_rotation(),'size_points':list(page.get_size())})
    if path.read_bytes()!=before:raise ValueError('Independent read changed source bytes.')
    print(json.dumps({'page_count':len(pages),'pages':pages,'sha256':hashlib.sha256(before).hexdigest(),'python':sys.version.split()[0],'pypdfium2':str(pdfium.PYPDFIUM_INFO),'pdfium':str(pdfium.PDFIUM_INFO)}));raise SystemExit(0)
raise ValueError('Unknown synthetic acceptance operation.')