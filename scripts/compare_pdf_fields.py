from pypdf import PdfReader
from pathlib import Path

paths = {
    'old': Path('/home/ubuntu/taamin-form-filler/client/public/assets/taamin-template.pdf'),
    'new': Path('/home/ubuntu/upload/pasted_file_CWc60c_س1-2026.pdf'),
}
for label, path in paths.items():
    reader = PdfReader(str(path))
    fields = reader.get_fields() or {}
    print(f'[{label}] pages={len(reader.pages)} fields={len(fields)}')
    for name, field in fields.items():
        rect = field.get('/Rect', '')
        kids = field.get('/Kids') or []
        if kids:
            rect = kids[0].get_object().get('/Rect', '')
        print(name, field.get('/FT'), rect)
