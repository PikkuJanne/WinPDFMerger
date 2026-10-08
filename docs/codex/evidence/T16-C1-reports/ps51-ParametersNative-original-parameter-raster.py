from pathlib import Path
import hashlib, io, json, random, sys
path=Path(sys.argv[1])
pixels=random.Random(160038).randbytes(1200*800*3)
content=b'q 432 0 0 260 0 28 cm /Im0 Do Q\nBT /F1 12 Tf 24 8 Td (T03-16-P01) Tj ET\n'
objects=[b'<< /Type /Catalog /Pages 2 0 R >>',b'<< /Type /Pages /Kids [3 0 R] /Count 1 >>',b'<< /Type /Page /Parent 2 0 R /MediaBox [0 0 432 288] /Resources << /Font << /F1 4 0 R >> /XObject << /Im0 5 0 R >> >> /Contents 6 0 R >>',b'<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',b'<< /Type /XObject /Subtype /Image /Width 1200 /Height 800 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length '+str(len(pixels)).encode()+b' >>\nstream\n'+pixels+b'\nendstream',b'<< /Length '+str(len(content)).encode()+b' >>\nstream\n'+content+b'endstream']
buffer=io.BytesIO();buffer.write(b'%PDF-1.4\n%\xe2\xe3\xcf\xd3\n');offsets=[0]
for number,obj in enumerate(objects,1):
    offsets.append(buffer.tell());buffer.write(str(number).encode()+b' 0 obj\n'+obj+b'\nendobj\n')
xref=buffer.tell();buffer.write(b'xref\n0 7\n0000000000 65535 f \n')
for offset in offsets[1:]:buffer.write(f'{offset:010} 00000 n \n'.encode())
buffer.write(b'trailer\n<< /Size 7 /Root 1 0 R >>\nstartxref\n'+str(xref).encode()+b'\n%%EOF\n')
raw=buffer.getvalue();path.write_bytes(raw)
print(json.dumps({'provenance':'Original deterministic stdlib-only RGB noise raster and synthetic text; no external content','seed':160038,'pixel_dimensions':[1200,800],'page_size_points':[432,288],'visible_id':'T03-16-P01','pages':1,'bytes':len(raw),'sha256':hashlib.sha256(raw).hexdigest()}))