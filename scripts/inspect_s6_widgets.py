from pypdf import PdfReader
from pathlib import Path

path = Path('client/public/assets/taamin-s6-template.pdf')
reader = PdfReader(str(path))
for page_index, page in enumerate(reader.pages, start=1):
    print(f'PAGE {page_index}')
    annots = page.get('/Annots') or []
    for ref in annots:
        obj = ref.get_object()
        if obj.get('/Subtype') != '/Widget':
            continue
        parent = obj.get('/Parent')
        field = parent.get_object() if parent else obj
        name = field.get('/T') or obj.get('/T')
        rect = obj.get('/Rect')
        flags = field.get('/Ff')
        if name:
            print(repr(str(name)), 'rect=', [round(float(v), 2) for v in rect], 'flags=', flags)
