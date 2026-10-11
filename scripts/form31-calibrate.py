#!/usr/bin/env python3
"""معايرة خريطة حقول نموذج رقم (31) — نسخة تعمل مع أي قالب، بحقول AcroForm أو بدونها.

الاستخدام:
    python scripts/form31-calibrate.py "client/public/assets/taamin-form31-template.pdf"

طريقة العمل (قياس هندسي من الملف نفسه):
    1) الشبكات الرقمية: تُكتشف من المستطيلات المتجهة المرسومة في الملف (مجموعة مربعات
       متساوية ومتلاصقة)، وإن لم توجد تُكتشف من ملامح الصورة.
    2) خانات النصوص: تُكتشف من أسطر النقاط المنقّطة المطبوعة (كل سطر نقاط = منطقة قيمة)،
       ويُؤخذ ارتفاع السطر من حبر التسمية العربية على نفس السطر.
    3) شرحات التواريخ: تُستخرج من مواضع الشرطات «/» والنقطتين «:» وكلمة «20» المطبوعة.

المخرج:
    * form31-calibration.png — صورة الفحص البصري (كل خانة عليها مستطيل).
    * قسم ANCHORS بصيغة TypeScript جاهز للنسخ داخل client/src/lib/form31-map.ts.

المتطلبات: pymupdf (أو fitz) و numpy.
"""

from __future__ import annotations

import re
import sys
from dataclasses import dataclass

try:
    import numpy as np
    import pymupdf  # type: ignore
except ImportError:  # pragma: no cover
    import fitz as pymupdf  # type: ignore

REFERENCE_WIDTH = 612.0
REFERENCE_HEIGHT = 792.0

# ارتفاع سطر الكتابة فوق السطر المنقّط (بلونطات) — مشتق من القالب الرسمي الحالي
LINE_HEIGHT = 15.1
# إزاحة قمة خانة الكتابة أعلى نقطة أساس (baseline) السطر المنقّط. القيمة 18.5
# تجعل قاعدة النص تجلس أعلى النقاط بمقدار ~5 نقاط (مقاسة من القالب الرسمي)
# فلا يقطع النص السطر المنقّط كما يحدث في النماذج المطبوعة اليدوية.
BASELINE_TO_BAND = 18.5
# ارتفاع سطر خانة التاريخ (مطابق لخانة تاريخ الإصابة في القالب الرسمي)
DATE_LINE_HEIGHT = 20.64


@dataclass
class GridRole:
    key: str
    label: str
    box_count: int
    region: tuple[float, float, float, float]


@dataclass
class TextRole:
    key: str
    label: str
    region: tuple[float, float, float, float]


GRID_ROLES: list[GridRole] = [
    GridRole("headerInsurance", "الرقم التأميني (أعلى الصفحة)", 9, (90, 42, 215, 68)),
    GridRole("headerNationalId", "الرقم القومي (أعلى الصفحة)", 14, (45, 62, 215, 88)),
    GridRole("applicantInsuranceGrid", "خانات الرقم التأميني لمقدم الطلب", 9, (315, 213, 515, 243)),
    GridRole("applicantNationalIdGrid", "خانات الرقم القومي لمقدم الطلب", 14, (215, 246, 515, 276)),
    GridRole("establishmentNumberGrid", "خانات رقم المنشأة", 8, (85, 332, 275, 364)),
    GridRole("insuredInsuranceGrid", "خانات الرقم التأميني للمؤمن عليه", 9, (270, 388, 470, 418)),
]

TEXT_ROLES: list[TextRole] = [
    TextRole("office", "المكتب", (440, 38, 600, 70)),
    TextRole("applicantName", "اسم مقدم الطلب", (340, 178, 540, 205)),
    TextRole("applicantRole", "صفة مقدم الطلب", (40, 178, 250, 205)),
    TextRole("applicantPhone", "رقم تليفون مقدم الطلب", (40, 208, 280, 232)),
    TextRole("establishmentName", "اسم المنشأة", (340, 330, 510, 356)),
    TextRole("insuredName", "اسم المؤمن عليه", (95, 358, 480, 384)),
    TextRole("companion", "هل يوجد مرافق", (110, 414, 255, 440)),
    TextRole("travelExpenses", "هل توجد مصاريف انتقال", (350, 474, 470, 500)),
    TextRole("treatmentPlace", "العلاج داخل أم خارج البلاد", (60, 474, 180, 500)),
    TextRole("monthlyWage", "أجر الاشتراك الشهري للمصاب", (345, 506, 445, 533)),
    TextRole("wagePaymentMethod", "طريقة أداء الأجر", (60, 506, 260, 533)),
    TextRole("otherLine1", "بيانات أخرى — السطر 1", (75, 530, 470, 558)),
    TextRole("otherLine2", "بيانات أخرى — السطر 2", (75, 550, 470, 578)),
    TextRole("otherLine3", "بيانات أخرى — السطر 3", (75, 570, 470, 598)),
    TextRole("signatureName", "الاسم أسفل التوقيع", (55, 644, 285, 674)),
]

# صفوف التواريخ. النوع "20" يعني وجود «20» مطبوعة قبل شرطتي السنة (نمط تاريخ الإصابة).
# وفي غيابها تُرسم شرطتا السنة داخل مربع مرجعي صغير يسار أقصى شرطة.
DATE_ROWS = [
    {"source": "injuryDate", "label": "تاريخ الإصابة", "region": (410, 410, 585, 445), "mode": "with20"},
    {"source": "releaseDate", "label": "تحريرًا في", "region": (410, 595, 585, 630), "mode": "with20"},
    {"source": "sickLeaveStartDate", "label": "تاريخ بداية الإجازة المرضي", "region": (330, 440, 585, 470), "mode": "plain",
     "label_span": (436, 585), "year_box": 26.0},
    {"source": "sickLeaveEndDate", "label": "تاريخ نهاية الإجازة المرضي", "region": (60, 440, 340, 470), "mode": "plain",
     "label_span": (185, 335), "year_box": 26.0},
]


@dataclass
class Anchor:
    key: str
    label: str
    x0: float
    y0: float
    x1: float
    y1: float
    box_count: int | None = None

    def as_ts(self) -> str:
        boxes = f", boxCount: {self.box_count}" if self.box_count else ""
        return (f"  {self.key}: {{ x0: {self.x0:7.2f}, y0: {self.y0:6.2f}, "
                f"x1: {self.x1:7.2f}, y1: {self.y1:6.2f}{boxes} }}, // {self.label}")


def center_in(rect, region) -> bool:
    cx, cy = (rect.x0 + rect.x1) / 2, (rect.y0 + rect.y1) / 2
    return region[0] <= cx <= region[2] and region[1] <= cy <= region[3]


def page_glyphs(page):
    items = []
    for block in page.get_text("rawdict")["blocks"]:
        for line in block.get("lines", []):
            for span in line["spans"]:
                for char in span.get("chars", []):
                    origin = char.get("origin") or (char["bbox"][0], char["bbox"][3])
                    items.append({"c": char["c"], "bbox": char["bbox"], "origin": origin,
                                  "font": span["font"], "size": span["size"]})
    return items


def ink_columns(page, clip, dpi=600):
    pix = page.get_pixmap(dpi=dpi, clip=clip, colorspace=pymupdf.csGRAY)
    arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width)
    return pix, arr


def detect_grid_from_drawings(page, role: GridRole, scale) -> Anchor | None:
    """يستخرج حدود الشبكة من مستطيلات الرسم المتجهي المتلاصقة داخل منطقة الدور."""
    boxes = []
    for drawing in page.get_drawings():
        for item in drawing["items"]:
            if item[0] != "re" or item[1].height < 5 or item[1].width < 5:
                continue
            boxes.append(item[1])
    region = (role.region[0] / scale, role.region[1] / scale, role.region[2] / scale, role.region[3] / scale)
    inside = [box for box in boxes if center_in(box, region)]
    if len(inside) < role.box_count:
        return None
    # نجمع المربعات المتقاربة في نفس الصف حسب تطابق الإحداثيات الرأسية
    def cluster(values, tolerance=2.0):
        groups: list[list[float]] = []
        for value in sorted(values):
            if groups and value - groups[-1][-1] <= tolerance:
                groups[-1].append(value)
            else:
                groups.append([value])
        return groups
    rows: list[list] = []
    for box in sorted(inside, key=lambda b: b.x0):
        for row in rows:
            if abs(row[0].y0 - box.y0) <= 2 and abs(row[0].y1 - box.y1) <= 2:
                row.append(box)
                break
        else:
            rows.append([box])
    row = max(rows, key=len)
    # نستبعد المستطيلات المكرّرة (القالب يرسم حدود الخانات مرتين)
    uniq: list = []
    for box in sorted(row, key=lambda b: b.x0):
        if uniq and abs(box.x0 - uniq[-1].x0) < 1.2:
            continue
        uniq.append(box)
    if len(uniq) != role.box_count:
        print(f"      تنبيه: عدد خانات «{role.label}» {len(uniq)} والمتوقع {role.box_count}")
    x0 = uniq[0].x0 / scale
    x1 = uniq[-1].x1 / scale
    return Anchor(role.key, f"{role.label} [{len(uniq)} خانة]", round(x0, 2), round(uniq[0].y0 / scale, 2),
                  round(x1, 2), round(uniq[0].y1 / scale, 2), role.box_count)


def detect_grid_from_pixels(page, role: GridRole, scale) -> Anchor | None:
    region = pymupdf.Rect(role.region[0] / scale, role.region[1] / scale, role.region[2] / scale, role.region[3] / scale)
    pix, arr = ink_columns(page, region, dpi=900)
    profile = (arr < 140).sum(axis=0)
    if profile.max() == 0:
        return None
    columns = np.where(profile > profile.max() * 0.5)[0]
    groups: list[list[int]] = []
    for column in columns:
        if groups and column - groups[-1][-1] <= 4:
            groups[-1].append(column)
        else:
            groups.append([column])
    centers = [region.x0 + (g[0] + g[-1]) / 2 / (900 / 72) for g in groups]
    if len(centers) < 2:
        return None
    # امتداد الشبكة رأسيًا من الخطوط الرأسية نفسها (أدق من كل الحبر في المنطقة)
    vertical_rows = np.where((arr < 140)[:, groups[0][0]:groups[-1][-1] + 1].sum(axis=1) >= (len(groups) - 2))[0]
    if not len(vertical_rows):
        return None
    y0 = region.y0 + vertical_rows.min() / (900 / 72)
    y1 = region.y0 + vertical_rows.max() / (900 / 72)
    if role.box_count and len(centers) - 1 != role.box_count:
        print(f"      تنبيه: عدد خانات «{role.label}» {len(centers) - 1} والمتوقع {role.box_count}")
    return Anchor(role.key, f"{role.label} [{len(centers) - 1} خانة]", round(centers[0] * scale, 2),
                  round(y0 * scale, 2), round(centers[-1] * scale, 2), round(y1 * scale, 2), role.box_count)


def cluster_dot_runs(glyphs, scale):
    """يجمع نقاط الخطوط المنقّطة في أسطر منفصلة (نطاق أفقي متصل على نفس السطر)."""
    dots = sorted((g for g in glyphs if g["c"] == "."), key=lambda g: (round(g["bbox"][1], 0), g["bbox"][0]))
    runs: list[list[dict]] = []
    for dot in dots:
        placed = False
        for run in runs:
            last = run[-1]
            if abs(last["bbox"][1] - dot["bbox"][1]) <= 3 and dot["bbox"][0] - last["bbox"][2] <= 6:
                run.append(dot)
                placed = True
                break
        if not placed:
            runs.append([dot])
    return [run for run in runs if len(run) >= 4]


def detect_text_anchor(page, role: TextRole, scale, dot_runs) -> Anchor | None:
    """خانة النص تُقاس من سطر النقاط المنقّطة نفسه: المدى الأفقي + حبر النقاط رأسيًا."""
    region = (role.region[0], role.region[1], role.region[2], role.region[3])
    dot_runs = [run for run in dot_runs if center_in(pymupdf.Rect(run[0]["bbox"][0] * scale, run[0]["bbox"][1] * scale, run[-1]["bbox"][2] * scale, run[-1]["bbox"][3] * scale), region)]
    if not dot_runs:
        return None
    run = max(dot_runs, key=len)
    x0_raw = run[0]["bbox"][0]
    x1_raw = run[-1]["bbox"][2]
    # نقطة أساس النقاط (origin) هي المرجع الرأسي الدقيق: النقاط نفسها تجلس عليها،
    # فيُبنى سطر الكتابة أعلاها مباشرة دون الاعتماد على قياس الحبر أو صناديق الحروف.
    baseline = sum(d["origin"][1] for d in run) / len(run)
    return Anchor(role.key, role.label, round(x0_raw * scale + 1, 2),
                  round((baseline - BASELINE_TO_BAND) * scale, 2), round(x1_raw * scale - 1, 2),
                  round((baseline - BASELINE_TO_BAND + LINE_HEIGHT) * scale, 2))


def detect_date_row(page, spec, scale, glyphs) -> list[Anchor]:
    region = spec["region"]
    scaled = (region[0] / scale, region[1] / scale, region[2] / scale, region[3] / scale)
    in_row = [g for g in glyphs if scaled[0] <= g["bbox"][0] <= scaled[2] and scaled[1] <= g["bbox"][1] <= scaled[3]]
    slashes = sorted((g for g in in_row if g["c"] == "/"), key=lambda g: -g["bbox"][0])
    colons = [g for g in in_row if g["c"] == ":"]
    if len(slashes) < 2 or not colons:
        print(f"    تحذير: تعذّر استخراج شرحات «{spec['label']}» (شرطات={len(slashes)})")
        return []
    line = min(g["bbox"][1] for g in in_row)
    y0 = round(line * scale, 2)
    y1 = round((line + DATE_LINE_HEIGHT) * scale, 2)
    right_slash, left_slash = slashes[0], slashes[1]
    if spec["mode"] == "with20":
        prefix = [g for g in in_row if g["c"] in "20٢٠" and "imes" in g["font"] and g["bbox"][0] < left_slash["bbox"][0]]
        colon_start = min(g["bbox"][0] for g in colons if g["bbox"][0] > right_slash["bbox"][2])
        if not prefix:
            print(f"    تحذير: لا توجد «20» مطبوعة في سطر «{spec['label']}»")
            return []
        colon_start = min(g["bbox"][0] for g in colons if g["bbox"][0] > right_slash["bbox"][2])
        year_x1 = left_slash["bbox"][0] - 1.5
        year_x0 = max(g["bbox"][2] for g in prefix) + 1.5
    else:
        year_x1 = left_slash["bbox"][0] - 2.0
        year_x0 = year_x1 - spec["year_box"]
        colon_candidates = [g["bbox"][0] for g in colons
                            if right_slash["bbox"][2] < g["bbox"][0] <= spec["label_span"][1] / scale + 4]
        if not colon_candidates:
            print(f"    تحذير: لم أجد نقطتين التسمية في سطر «{spec['label']}»")
            return []
        colon_start = min(colon_candidates)
    segments = [
        ("year", year_x0, year_x1),
        ("month", left_slash["bbox"][2] + 1.5, right_slash["bbox"][0] - 1.5),
        ("day", right_slash["bbox"][2] + 1.5, colon_start - 1.5),
    ]
    out = []
    for name, x0, x1 in segments:
        out.append(Anchor(f"{spec['source']}_{name}", f"{spec['label']} — {name}", round(x0 * scale, 2), y0, round(x1 * scale, 2), y1, 2))
    return out


def main() -> int:
    if len(sys.argv) < 2:
        print(__doc__)
        return 1
    path = sys.argv[1]
    doc = pymupdf.open(path)
    page = doc[0]
    width, height = page.rect.width, page.rect.height
    print(f"الملف: {path}")
    print(f"مقاس الصفحة: {width:.0f} × {height:.0f} نقطة (المرجع {REFERENCE_WIDTH:.0f} × {REFERENCE_HEIGHT:.0f})")
    scale = REFERENCE_WIDTH / width
    if abs(height * scale - REFERENCE_HEIGHT) > 1:
        print("تنبيه: نسبة العرض إلى الارتفاع مختلفة عن المرجع — راجع الصورة الناتجة.")
    print(f"حقول AcroForm: {len(list(page.widgets() or []))} (لا يُعتمد عليها في هذه المعايرة)")
    glyphs = page_glyphs(page)
    dot_runs = cluster_dot_runs(glyphs, scale)

    anchors: list[Anchor] = []
    print("\n— الشبكات الرقمية —")
    for role in GRID_ROLES:
        found = detect_grid_from_drawings(page, role, scale) or detect_grid_from_pixels(page, role, scale)
        if not found:
            print(f"    ✗ {role.label}: لم تُكتشف")
            continue
        anchors.append(found)
        print(f"    ✓ {found.label}: x {found.x0} → {found.x1} ، y {found.y0} → {found.y1}")

    print("\n— خانات النصوص —")
    for role in TEXT_ROLES:
        found = detect_text_anchor(page, role, scale, dot_runs)
        if not found:
            print(f"    ✗ {role.label}: لم يُعثر على سطر نقاط")
            continue
        anchors.append(found)
        print(f"    ✓ {role.label}: x {found.x0} → {found.x1} ، y {found.y0} → {found.y1}")

    print("\n— شرحات التواريخ —")
    date_anchors: list[Anchor] = []
    for spec in DATE_ROWS:
        found = detect_date_row(page, spec, scale, glyphs)
        for anchor in found:
            print(f"    ✓ {anchor.label}: x {anchor.x0} → {anchor.x1} ، y {anchor.y0} → {anchor.y1}")
        date_anchors.extend(found)

    for anchor in anchors:
        target = pymupdf.Rect(anchor.x0 / scale, anchor.y0 / scale, anchor.x1 / scale, anchor.y1 / scale)
        color = (1, 0, 0) if anchor.box_count else (0.1, 0.1, 0.85)
        page.draw_rect(target, color=color, width=0.7)
        page.insert_text((target.x0 + 0.5, target.y0 + 4), anchor.key, fontsize=4.6, color=color)
    for anchor in date_anchors:
        target = pymupdf.Rect(anchor.x0 / scale, anchor.y0 / scale, anchor.x1 / scale, anchor.y1 / scale)
        page.draw_rect(target, color=(0, 0.5, 0), width=0.7)
        page.insert_text((target.x0, target.y0 - 1), anchor.key.replace("Date_", "-"), fontsize=4, color=(0, 0.5, 0))
    page.get_pixmap(dpi=150).save("form31-calibration.png")
    print("\n✔ تم إنشاء form31-calibration.png — تأكد أن كل مستطيل يغطي خانته المطبوعة.")

    print("\n========== انسخ هذا القسم داخل client/src/lib/form31-map.ts ==========")
    for anchor in anchors + date_anchors:
        print(anchor.as_ts())
    print("========== نهاية القسم ==========")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
