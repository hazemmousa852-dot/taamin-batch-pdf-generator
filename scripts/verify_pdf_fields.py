"""Fail fast when an official PDF template no longer matches the verified maps."""

from pathlib import Path
from pypdf import PdfReader

ROOT = Path(__file__).resolve().parents[1]
ASSETS = ROOT / "client" / "public" / "assets"

EXPECTED = {
    "taamin-template.pdf": {
        "بتكم", ": ةفصبلطلا مدقم_1", "Text Field6", "يلتلا مقر:نوف", "مقر: ىموق",
        "Text Field3", "fill_3", ":يــــنيمأتلا مـــقرلا", "ةيــــــــــــسنلجا_1",
        "ةيــــــــــــسنلجا", "ةنهلما", "ةنهلما_1", "fill_7", "fill_6", "Text Field4",
        "Text Field5", "Text Field10", "Text Field9", "Text Field8", "Text Field2",
        "Text Field1", "Text Field0", "%", "أشنلما مقرة", "fill_4", "ةيرق",
        "ةيرق_1", ": ةظفامح", ": ةظفامح_1", "fill_2",
        ":نييمأتلا اهمقر", ":نييمأتلا اهمقر_1", ":ناوــــــــــــنعلا", "انأ رقأ",
    },
    "taamin-s6-template.pdf": {
        "بتكم", ":بلطلا مدقم ةفص", ":بلطلا مدقم ةفص_1", ":نوفيلتلا مقر", "موقلا مقري",
        "Text Field0", "Text Field1", ": يـموقلا مقرلا", ": نييمأتلا مقرلا", ":هأشنلما مسا",
        ":هأشنلما مسا_1", "fill_1", "Text Field2", "Text Field3", "Text Field4",
        "Text Field5", "Text Field6", "Text Field7",
        ":نييمأتلا اهمقر",
        "ا: هيلع نمؤلما مــــــــس---------------------------------------:نييمأتلا همقر",
    },
}

for filename, expected in EXPECTED.items():
    reader = PdfReader(ASSETS / filename)
    fields = set((reader.get_fields() or {}).keys())
    missing = expected - fields
    if missing:
        raise SystemExit(f"{filename}: missing fields: {sorted(missing)}")
    orphaned = []
    widget_count = 0
    for page_number, page in enumerate(reader.pages, 1):
        for reference in page.get("/Annots", []):
            widget = reference.get_object()
            if widget.get("/Subtype") != "/Widget":
                continue
            widget_count += 1
            parent = widget.get("/Parent")
            effective = parent.get_object() if parent else widget
            name = effective.get("/T")
            if name and name not in fields:
                orphaned.append((page_number, name))
    if orphaned:
        raise SystemExit(f"{filename}: orphaned widgets: {orphaned}")
    print(f"{filename}: {len(expected)} verified fields, {widget_count} canonical widgets")

