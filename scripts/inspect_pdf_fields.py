from pathlib import Path
from pypdf import PdfReader

pdf_path = Path('/home/ubuntu/upload/pasted_file_CWc60c_س1-2026.pdf')
reader = PdfReader(str(pdf_path))
print('pages', len(reader.pages))
fields = reader.get_fields() or {}
print('fields', len(fields))
for name, field in fields.items():
    value = field.get('/V', '')
    kids = field.get('/Kids') or []
    rect = field.get('/Rect', '')
    print(name, field.get('/FT'), value, rect, 'kids=', len(kids))
    for kid in kids:
        obj = kid.get_object()
        print('  kid', obj.get('/Subtype'), obj.get('/Rect'), obj.get('/AS'))
