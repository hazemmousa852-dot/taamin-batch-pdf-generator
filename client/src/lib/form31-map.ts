import { normalizeDigits, normalizeText, type PersonRecord } from "./form";

/**
 * خريطة حقول نموذج رقم (31): «طلب صرف تعويض الأجر ومصاريف الانتقال في حالة الإصابة / المرض».
 *
 * النموذج الرسمي ملف Word مصدَّر إلى PDF لا يحتوي خانات AcroForm مفيدة، لذلك
 * تُكتب القيم فوق الخانات المطبوعة (overlay) باستخدام الخط العربي المضمّن، ثم
 * تُثبَّت النتيجة داخل الصفحة (flatten) لتظهر في كل قارئات PDF والطابعات.
 *
 * الإحداثيات هنا بنظام PyMuPDF (المحور Y من أعلى الصفحة) وبمقاس مرجعي
 * 612 × 792 نقطة، وتُحوَّل وقت التشغيل إلى نظام pdf-lib. أي قالب جديد بنفس
 * التخطيط يمكن معايرته عبر: python scripts/form31-calibrate.py <path-to-pdf>
 */

export type TopAnchor = {
  x0: number;
  y0: number;
  x1: number;
  y1: number;
  /** عند وجوده يعني أن الخانة شبكة أرقام مقسّمة إلى هذا العدد من المربعات. */
  boxCount?: number;
};

export const FORM31_REFERENCE_PAGE = { width: 612, height: 792 } as const;

/** خانات النص المكتوب فوق السطور المنقّطة. */
export const FORM31_TEXT_ANCHORS: Partial<Record<keyof PersonRecord, TopAnchor>> = {
  office: { x0: 467.3, y0: 39.94, x1: 549.28, y1: 55.04 },
  applicantName: { x0: 361.31, y0: 183.36, x1: 512.7, y1: 198.46 },
  applicantRole: { x0: 54.04, y0: 183.36, x1: 233, y1: 198.46 },
  applicantPhone: { x0: 58.24, y0: 213.12, x1: 256.77, y1: 228.22 },
  establishmentName: { x0: 353.99, y0: 333.03, x1: 463.17, y1: 348.13 },
  insuredName: { x0: 109.38, y0: 359.79, x1: 462.75, y1: 374.89 },
  companion: { x0: 121.5, y0: 415.13, x1: 239.9, y1: 430.23 },
  treatmentPlace: { x0: 79.26, y0: 474.29, x1: 163.65, y1: 489.39 },
  travelExpenses: { x0: 373.55, y0: 474.29, x1: 448.71, y1: 489.39 },
  wagePaymentMethod: { x0: 78.54, y0: 506.69, x1: 240.21, y1: 521.79 },
  monthlyWage: { x0: 364.67, y0: 506.69, x1: 427.47, y1: 521.79 },
  otherLine1: { x0: 84.54, y0: 529.73, x1: 462.71, y1: 544.83 },
  otherLine2: { x0: 84.54, y0: 550.01, x1: 462.71, y1: 565.11 },
  otherLine3: { x0: 84.54, y0: 570.32, x1: 462.71, y1: 585.42 },
  signatureName: { x0: 65.1, y0: 649.64, x1: 272.93, y1: 664.74 },
};

/** الشبكات الرقمية: كل رقم يُوزَّع على مربّعاته من اليمين إلى اليسار. */
export const FORM31_BOX_ANCHORS: Array<{ id: string; source: keyof PersonRecord; anchor: TopAnchor }> = [
  // id يطابق اسم الخانة في ناتج سكربت المعايرة، ويُستخدم في فحص الانحدار
  { id: "insuredInsuranceGrid", source: "insuranceNumber", anchor: { x0: 279.4, y0: 393.55, x1: 458.2, y1: 413.4, boxCount: 9 } },
  { id: "headerInsurance", source: "insuranceNumber", anchor: { x0: 99.64, y0: 49.44, x1: 204.44, y1: 67.92, boxCount: 9 } },
  { id: "headerNationalId", source: "nationalId", anchor: { x0: 52, y0: 62, x1: 207.32, y1: 81.52, boxCount: 14 } },
  { id: "applicantInsuranceGrid", source: "applicantInsuranceNumber", anchor: { x0: 326.35, y0: 218.52, x1: 505.15, y1: 238.37, boxCount: 9 } },
  { id: "applicantNationalIdGrid", source: "applicantNationalId", anchor: { x0: 227.05, y0: 250.8, x1: 504.8, y1: 270.65, boxCount: 14 } },
  { id: "establishmentNumberGrid", source: "establishmentNumber", anchor: { x0: 93.4, y0: 337.74, x1: 263.8, y1: 359.04, boxCount: 8 } },
];

/**
 * كل تواريخ النموذج تُكتب كنص متصل بجوار العنوان. القيمة المخزنة هي
 * YYYY-MM-DD، ويُحافظ العرض على الترتيب المطلوب YYYY/MM/DD من اليسار لليمين؛
 * وبعد تحويل الأرقام تصبح مثل: ٢٠٢٦/٠١/٣٠.
 */
export const FORM31_INLINE_DATE_ANCHORS: Array<{ source: keyof PersonRecord; anchor: TopAnchor }> = [
  { source: "injuryDate", anchor: { x0: 406, y0: 416.02, x1: 502, y1: 436.66 } },
  { source: "sickLeaveStartDate", anchor: { x0: 323, y0: 447, x1: 427, y1: 468 } },
  { source: "sickLeaveEndDate", anchor: { x0: 58, y0: 451, x1: 179, y1: 472 } },
  { source: "releaseDate", anchor: { x0: 418, y0: 601.21, x1: 519, y1: 621.85 } },
];

export type Form31Rect = { x: number; y: number; width: number; height: number };

/**
 * يحوّل إحداثيات «من أعلى الصفحة» إلى إحداثيات pdf-lib «من أسفل الصفحة»،
 * مع تكييفها مع أي مقاس صفحة أو دوران 180 درجة.
 */
export function toPdfLibRect(anchor: TopAnchor, pageWidth: number, pageHeight: number, rotation = 0): Form31Rect {
  const scaleX = pageWidth / FORM31_REFERENCE_PAGE.width;
  const scaleY = pageHeight / FORM31_REFERENCE_PAGE.height;
  const x0 = anchor.x0 * scaleX;
  const x1 = anchor.x1 * scaleX;
  const top = anchor.y0 * scaleY;
  const bottom = anchor.y1 * scaleY;
  const normalized = ((rotation % 360) + 360) % 360;
  if (normalized === 180) {
    return { x: pageWidth - x1, y: top, width: x1 - x0, height: bottom - top };
  }
  if (normalized !== 0) {
    throw new Error("قالب نموذج 31 مُدوَّر بزاوية غير مدعومة؛ أعد تصدير الملف بدون تدوير الصفحة.");
  }
  return { x: x0, y: pageHeight - bottom, width: x1 - x0, height: bottom - top };
}

function displayDate(value: string) {
  const match = normalizeDigits(value).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return match ? `${match[3]}/${match[2]}/${match[1]}` : normalizeText(value);
}

const YES_VALUES = new Set(["نعم", "يوجد", "true", "yes", "y", "1", "✓"]);
const NO_VALUES = new Set(["لا", "لا يوجد", "false", "no", "n", "0", "✗"]);

/** توحيد إجابات «نعم / لا» المطبوعة حتى لو جاءت من Excel بصيغ مختلفة. */
export function normalizeYesNo(value: string) {
  const text = normalizeText(value);
  if (!text) return "";
  const key = text.toLocaleLowerCase("ar-EG");
  if (YES_VALUES.has(key)) return "نعم";
  if (NO_VALUES.has(key)) return "لا";
  return text;
}

/** توحيد إجابة «العلاج داخل أم خارج البلاد» بحسب الكلمة الأساسية. */
export function normalizeTreatmentPlace(value: string) {
  const text = normalizeText(value);
  if (!text) return "";
  if (text.includes("خارج")) return "خارج البلاد";
  if (text.includes("داخل") || text.includes("الداخل")) return "داخل البلاد";
  return text;
}

/** القيمة النهائية كما تُطبع في كل خانة من خانات نموذج 31. */
export function form31DisplayValue(key: keyof PersonRecord, record: PersonRecord): string {
  const raw = record[key] ?? "";
  switch (key) {
    case "signatureName":
      // خانة «الاسم» أسفل توقيع مقدم الطلب: تُترك فارغة لاستخدام اسم مقدم الطلب.
      return normalizeText(record.signatureName) || normalizeText(record.applicantName);
    case "companion":
      return normalizeYesNo(raw);
    case "travelExpenses":
      return normalizeYesNo(raw);
    case "treatmentPlace":
      return normalizeTreatmentPlace(raw);
    case "injuryDate":
    case "sickLeaveStartDate":
    case "sickLeaveEndDate":
    case "releaseDate":
      // هذه الحقول لها شرحات مطبوعة في القالب، وتُكتب نصًّا كاملًا فقط إن تعذّر ذلك
      return displayDate(raw);
    case "monthlyWage":
      return normalizeDigits(raw).trim();
    default:
      return normalizeText(raw);
  }
}

/** كل الحقول التي يقرأها نموذج 31 من السجل — تُستخدم في فحص التغطية والاختبارات. */
export const FORM31_FIELDS: Array<keyof PersonRecord> = [
  ...(Object.keys(FORM31_TEXT_ANCHORS) as Array<keyof PersonRecord>),
  ...FORM31_BOX_ANCHORS.map((item) => item.source),
  ...FORM31_INLINE_DATE_ANCHORS.map((item) => item.source),
];

/** اسم ملف PDF آمن لنموذج 31. */
export function form31FileName(record: PersonRecord, index = 1) {
  const person = (record.insuredName || "سجل").replace(/[^\u0600-\u06ff\w\s-]/g, "").trim().replace(/\s+/g, "-");
  const identity = normalizeDigits(record.insuranceNumber || record.nationalId).replace(/\D/g, "").slice(-10);
  return `نموذج31-${person || "سجل"}-${identity || String(index).padStart(3, "0")}-${String(index).padStart(3, "0")}.pdf`;
}