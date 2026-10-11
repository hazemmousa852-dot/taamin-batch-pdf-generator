"""Fail fast when an official PDF template no longer matches the verified maps.

- قوالب س1 وس6: تُفحص أسماء حقول AcroForm كما كان.
- قالب نموذج 31: لا يحتوي حقول AcroForm، لذلك يُفحص انحدار المعايرة: يُعاد
  تشغيل سكربت المعايرة على القالب ويُقارَن ناتجه بإحداثيات client/src/lib/form31-map.ts
  بترتيب الخانات. أي اختلاف يعني أن القالب تغيّر وأن الخريطة تحتاج معايرة جديدة.
"""

from __future__ import annotations

import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from pypdf import PdfReader

ROOT = Path(__file__).resolve().parents[1]
ASSETS = ROOT / "client" / "public" / "assets"
FORM31_MAP = ROOT / "client" / "src" / "lib" / "form31-map.ts"
FORM31_TEMPLATE = "taamin-form31-template.pdf"
TOLERANCE = 1.2

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

ANCHOR_PATTERN = re.compile(
    r"x0:\s*(-?\d+(?:\.\d+)?),\s*y0:\s*(-?\d+(?:\.\d+)?),\s*"
    r"x1:\s*(-?\d+(?:\.\d+)?),\s*y1:\s*(-?\d+(?:\.\d+)?)"
)


def numbers_from_text(text: str) -> list[tuple[float, float, float, float]]:
    return [tuple(float(group) for group in match) for match in ANCHOR_PATTERN.findall(text)]


def calibration_section(output: str) -> str:
    start = output.index("========== انسخ هذا القسم")
    end = output.index("========== نهاية القسم ==========")
    return output[start:end]


def check_acrofom_template(filename: str, expected: set[str]) -> None:
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


CALIBRATION_LINE = re.compile(r"^\s*([A-Za-z0-9_]+):\s*\{(.*?)\}\s*,", re.M)


def numbers_of(body: str) -> tuple[float, float, float, float]:
    """يستخرج قيم x0/y0/x1/y1 بالاسم — لا بالترتيب — حتى لا تختلط بأرقام أسماء الحقول."""
    values = dict(re.findall(r"\b(x0|y0|x1|y1)\s*:\s*(-?\d+(?:\.\d+)?)", body))
    return (float(values["x0"]), float(values["y0"]), float(values["x1"]), float(values["y1"]))


def calibrated_anchors(output: str) -> dict[str, tuple[float, float, float, float]]:
    return {name: numbers_of(body) for name, body in CALIBRATION_LINE.findall(calibration_section(output))}


def mapped_anchors() -> dict[str, tuple[float, float, float, float]]:
    """يقرأ الخريطة المستخدمة في التطبيق ويُرجع الخانات بمفاتيح مطابقة لناتج المعايرة."""
    source = FORM31_MAP.read_text(encoding="utf-8")
    text_block = source.split("FORM31_TEXT_ANCHORS", 1)[1].split("export const", 1)[0]
    box_block = source.split("FORM31_BOX_ANCHORS", 1)[1].split("export const", 1)[0]
    date_block = source.split("FORM31_DATE_ANCHORS", 1)[1]
    anchors: dict[str, tuple[float, float, float, float]] = {}
    for name, body in re.findall(r"(\w+):\s*\{([^}]*)\}", text_block):
        anchors[name] = numbers_of(body)
    for entry in re.finditer(r"id:\s*\"(\w+)\"[^}]*?anchor:\s*\{([^}]*)\}", box_block, re.S):
        anchors[entry.group(1)] = numbers_of(entry.group(2))
    for chunk in date_block.split("source:")[1:]:
        name = re.search(r'"?(\w+)"?', chunk).group(1)
        parts = re.findall(r"\{([^}]*)\}", chunk)[:3]
        for index, part in enumerate(parts):
            anchors[f"{name}_{('year', 'month', 'day')[index]}"] = numbers_of(part)
    return anchors


def check_form31_regression() -> None:
    """يقارن ناتج المعايرة الحالي بإحداثيات الخريطة المستخدمة في التطبيق."""
    script = ROOT / "scripts" / "form31-calibrate.py"
    with tempfile.TemporaryDirectory() as workdir:
        shutil.copy(ASSETS / FORM31_TEMPLATE, Path(workdir) / FORM31_TEMPLATE)
        result = subprocess.run(
            [sys.executable, str(script), FORM31_TEMPLATE],
            cwd=workdir, capture_output=True, text=True,
        )
    if result.returncode != 0:
        raise SystemExit(f"form31: تعذّر تشغيل المعايرة:\n{result.stdout}\n{result.stderr}")
    calibrated = calibrated_anchors(result.stdout)
    mapped = mapped_anchors()
    missing = sorted(set(calibrated) - set(mapped))
    extra = sorted(set(mapped) - set(calibrated))
    if missing or extra:
        raise SystemExit(f"{FORM31_TEMPLATE}: اختلاف في أسماء الخانات — ناقص من الخريطة: {missing}، "
                         f"زائد في الخريطة: {extra}")
    drifted = sorted(
        (name, calibrated[name], mapped[name])
        for name in calibrated
        if max(abs(a - b) for a, b in zip(calibrated[name], mapped[name])) > TOLERANCE
    )
    if drifted:
        details = "\n".join(f"  {name}: معايرة={expected} خريطة={actual}" for name, expected, actual in drifted[:8])
        raise SystemExit(f"{FORM31_TEMPLATE}: انحراف في {len(drifted)} خانة عن الخريطة:\n{details}\n"
                         "شغّل python scripts/form31-calibrate.py ثم حدّث client/src/lib/form31-map.ts")
    reader = PdfReader(ASSETS / FORM31_TEMPLATE)
    print(f"{FORM31_TEMPLATE}: {len(mapped)} خانة مطابقة للخريطة داخل {len(reader.pages)} صفحة "
          f"(بدون حقول AcroForm)")


for filename, expected_fields in EXPECTED.items():
    check_acrofom_template(filename, expected_fields)
check_form31_regression()
