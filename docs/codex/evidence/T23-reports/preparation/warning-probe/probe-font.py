import runpy,pathlib
p=pathlib.Path(__file__).with_name('probe.py')
s=p.read_text(encoding='utf-8-sig')
s=s.replace(" 'baseline':base,", " 'baseline':base,\n 'unknown_font':base.replace(b'/Helvetica',b'/ZetZZetZZ'),\n 'unknown_font_normal':base.replace(b'/Helvetica ',b'/ZetZZetZZ '),\n 'bad_font_encoding':base.replace(b'/WinAnsiEncoding',b'/BadTestEncoding'),\n 'bad_pdf_version':base.replace(b'%PDF-1.4',b'%PDF-1.9'),\n 'negative_mediabox':base.replace(b'[ 0 0 432 288 ]',b'[ 0 0 000 288 ]'),")
exec(compile(s,str(p),'exec'))
