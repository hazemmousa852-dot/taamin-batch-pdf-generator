from pathlib import Path
from pypdf import PdfReader

for label, path in [('new', Path('/home/ubuntu/upload/pasted_file_CWc60c_س1-2026.pdf')), ('old', Path('/home/ubuntu/taamin-form-filler/client/public/assets/taamin-template.pdf'))]:
    reader = PdfReader(str(path))
    print(f'[{label}]')
    for page_num, page in enumerate(reader.pages, start=1):
        print('page', page_num)
        for annot_ref in page.get('/Annots', []):
            annot = annot_ref.get_object()
            if annot.get('/Subtype') != '/Widget':
                continue
            parent = annot.get('/Parent')
            field = parent.get_object() if parent else annot
            name = field.get('/T') or annot.get('/T') or ''
            print(name, field.get('/FT') or annot.get('/FT'), list(annot.get('/Rect') or []))
