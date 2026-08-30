"""Attach orphaned widgets in the official S1/S6 PDFs to their AcroForm trees."""

from pathlib import Path
from pypdf import PdfReader, PdfWriter

ROOT = Path(__file__).resolve().parents[1]
ASSETS = ROOT / "client" / "public" / "assets"

for filename in ("taamin-template.pdf", "taamin-s6-template.pdf"):
    source = ASSETS / filename
    temporary = source.with_suffix(".repaired.pdf")
    reader = PdfReader(source)
    writer = PdfWriter()
    writer.clone_document_from_reader(reader)
    reattached = writer.reattach_fields()
    with temporary.open("wb") as stream:
        writer.write(stream)
    temporary.replace(source)
    print(f"{filename}: reattached {len(reattached)} widgets")

