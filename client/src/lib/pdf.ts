// Design reminder: exported documents preserve the original official form; the app adds only the user's data and keeps the workflow quiet and auditable.

import { PDFDocument, TextAlignment, type PDFTextField, type PDFFont } from "pdf-lib";
import fontkit from "@pdf-lib/fontkit";
import * as reshaperPackage from "arabic-persian-reshaper";
import JSZip from "jszip";
import type { PersonRecord, TemplateId } from "./form";

const configuredBase = import.meta.env?.BASE_URL ?? "/";
const ASSET_BASE = configuredBase.endsWith("/") ? configuredBase : `${configuredBase}/`;
const assetUrl = (name: string) => `${ASSET_BASE}assets/${name}`;

export const TEMPLATE_URLS: Record<TemplateId, string> = {
  s1: assetUrl("taamin-template.pdf"),
  s6: assetUrl("taamin-s6-template.pdf"),
};
const ARABIC_FONT_URL = assetUrl("NotoNaskhArabic-Regular.ttf");

const ArabicShaper = (reshaperPackage as unknown as { ArabicShaper?: { convertArabic: (value: string) => string } }).ArabicShaper;

const s1TextBindings: Partial<Record<keyof PersonRecord, string[]>> = {
  office: ["بتكم"], applicantName: [":بلطلا مدقم"], applicantRole: [": ةفصبلطلا مدقم_1"], applicantPhone: ["يلتلا مقر:نوف"],
  insuredName: ["fill_3"], qualification: ["ةنهلما"], profession: ["ةنهلما_1"], gender: ["ةيــــــــــــسنلجا"],
  country: ["ةيــــــــــــسنلجا_1"], governorate: [": ةظفامح"], district: ["ةيرق"], street: ["ةيرق_1"], center: [":زكرم / مسق"],
  establishmentName: ["fill_4"], releaseDate: ["fill_7"], increasePercent: ["%"], contributionCode: [": عاطقلا"],
};
const s1MoneyBindings: Partial<Record<keyof PersonRecord, string[]>> = {
  basicWage: ["Text Field4"], totalWage: ["Text Field5"],
};
const s1BoxBindings: Partial<Record<keyof PersonRecord, string[]>> = {
  establishmentNumber: ["أشنلما مقرة"], insuranceNumber: [":يــــنيمأتلا مـــقرلا"], nationalId: ["ةيــــــــــــسنلجا_1"], applicantNationalId: ["مقر: ىموق"],
};
const s1DateBindings: Partial<Record<keyof PersonRecord, string[]>> = {
  startDate: ["Text Field8", "Text Field9", "Text Field10"],
  increaseDate: ["Text Field0", "Text Field1", "Text Field2"],
};
const s6TextBindings: Partial<Record<keyof PersonRecord, string[]>> = {
  office: ["بتكم"], applicantRole: [":بلطلا مدقم ةفص"], applicantName: [":بلطلا مدقم ةفص_1"], applicantPhone: [":نوفيلتلا مقر"], insuredName: ["Text Field1"], establishmentName: [":هأشنلما مسا", ":هأشنلما مسا_1", "Text Field5"], address: ["Text Field6"], endReason: ["fill_1"],
};
const s6BoxBindings: Partial<Record<keyof PersonRecord, string[]>> = {
  applicantNationalId: ["موقلا مقري"], insuranceNumber: [": نييمأتلا مقرلا", ":نييمأتلا اهمقر", "Text Field7"], nationalId: [": يـموقلا مقرلا"],
};
const checkboxBindings: Partial<Record<keyof PersonRecord, Record<string, string>>> = { category: { "عاملين لدى الغير": "يرغلا ىدل ينلماع", "أصحاب أعمال": "لامعأ باحصأشنم ملهآت", "عمالة غير منتظمة": "زباخلماب ينلماعلا" } };

function hasArabic(value: string) { return /[\u0600-\u06ff]/.test(value); }
function toPdfText(value: string) {
  if (!value || !hasArabic(value) || !ArabicShaper) return value;
  return value.split(/([0-9A-Za-z%/().,\- ]+)/).map((part) => hasArabic(part) ? ArabicShaper.convertArabic(part).split("").reverse().join("") : part).join("");
}
function cleanDigits(value: string) { return value.replace(/[^0-9٠-٩]/g, "").replace(/[٠-٩]/g, (digit) => String("٠١٢٣٤٥٦٧٨٩".indexOf(digit))); }
function field<T>(form: ReturnType<PDFDocument["getForm"]>, name: string, getter: (form: ReturnType<PDFDocument["getForm"]>, name: string) => T) { try { return getter(form, name); } catch { return null; } }
function setText(form: ReturnType<PDFDocument["getForm"]>, name: string, value: string, font: PDFFont, alignment: TextAlignment = TextAlignment.Right) {
  const textField = field<PDFTextField>(form, name, (current, fieldName) => current.getTextField(fieldName));
  if (!textField) return;
  try {
    textField.setText(toPdfText(value));
    textField.setAlignment(alignment);
    textField.setFontSize(Math.max(7, Math.min(11, value.length > 28 ? 7 : value.length > 20 ? 8 : 10)));
    textField.updateAppearances(font);
  } catch { /* continue with the remaining fields */ }
}
function setBoxes(form: ReturnType<PDFDocument["getForm"]>, names: string[], rawValue: string, font: PDFFont) { const digits = cleanDigits(rawValue); if (names.length === 1) { setText(form, names[0], digits, font, TextAlignment.Center); return; } names.forEach((name, index) => setText(form, name, digits[index] ?? "", font, TextAlignment.Center)); }
function setDateFields(form: ReturnType<PDFDocument["getForm"]>, names: string[], rawValue: string, font: PDFFont) { const parts = rawValue.includes("-") ? rawValue.split("-").reverse() : cleanDigits(rawValue).match(/\d{1,4}/g) ?? []; names.forEach((name, index) => setText(form, name, parts[index] ?? "", font, TextAlignment.Center)); }
function setMoneyFields(form: ReturnType<PDFDocument["getForm"]>, names: string[], rawValue: string, font: PDFFont) { const [whole, fraction = ""] = String(rawValue).replace(/[,،]/g, ".").split("."); setText(form, names[0], whole, font, TextAlignment.Center); if (names[1]) setText(form, names[1], fraction.padEnd(2, "0").slice(0, 2), font, TextAlignment.Center); }
function setCheckboxes(form: ReturnType<PDFDocument["getForm"]>, category: string) { Object.entries(checkboxBindings.category ?? {}).forEach(([label, fieldName]) => { const checkbox = field(form, fieldName, (current, name) => current.getCheckBox(name)); if (!checkbox) return; try { label === category ? checkbox.check() : checkbox.uncheck(); } catch { /* non-standard appearance state */ } }); }

export async function fillPdf(record: PersonRecord, template: TemplateId = "s1") {
  const [templateResponse, fontResponse] = await Promise.all([fetch(TEMPLATE_URLS[template]), fetch(ARABIC_FONT_URL)]);
  if (!templateResponse.ok) throw new Error("تعذر تحميل قالب النموذج");
  if (!fontResponse.ok) throw new Error("تعذر تحميل الخط العربي");
  const pdfDoc = await PDFDocument.load(await templateResponse.arrayBuffer());
  pdfDoc.registerFontkit(fontkit);
  const arabicFont = await pdfDoc.embedFont(await fontResponse.arrayBuffer(), { subset: true });
  const form = pdfDoc.getForm();
  const textBindings = template === "s6" ? s6TextBindings : s1TextBindings;
  const boxBindings = template === "s6" ? s6BoxBindings : s1BoxBindings;
  (Object.keys(textBindings) as Array<keyof PersonRecord>).forEach((key) => { const value = record[key]; if (value) (textBindings[key] ?? []).forEach((name) => setText(form, name, value, arabicFont)); });
  (Object.keys(boxBindings) as Array<keyof PersonRecord>).forEach((key) => { const value = record[key]; if (value) setBoxes(form, boxBindings[key] ?? [], value, arabicFont); });
  if (template === "s1") {
    (Object.keys(s1DateBindings) as Array<keyof PersonRecord>).forEach((key) => { const value = record[key]; if (value) setDateFields(form, s1DateBindings[key] ?? [], value, arabicFont); });
    (Object.keys(s1MoneyBindings) as Array<keyof PersonRecord>).forEach((key) => { const value = record[key]; if (value) setMoneyFields(form, s1MoneyBindings[key] ?? [], value, arabicFont); });
  }
  if (template === "s6") {
    setText(form, "موقلا مقري", record.applicantNationalId, arabicFont);
    setText(form, ": يـموقلا مقرلا", record.nationalId, arabicFont);
    setText(form, ": نييمأتلا مقرلا", record.insuranceNumber, arabicFont);
    setText(form, "Text Field7", record.insuranceNumber, arabicFont);
    const [year, month, day] = record.endDate.split("-");
    if (day && month && year) {
      setText(form, "Text Field4", day, arabicFont);
      setText(form, "Text Field3", month, arabicFont);
      setText(form, "Text Field2", year.slice(-2), arabicFont);
    }
  }
  if (template === "s1") setCheckboxes(form, record.category);
  // Refresh all widgets with the embedded Arabic font so untouched fields never fall back to WinAnsi.
  // Values were already reshaped before setText, so this preserves RTL glyph order while avoiding encoding errors.
  form.updateFieldAppearances(arabicFont);
  return pdfDoc.save({ useObjectStreams: false });
}

export function safeFileName(record: PersonRecord, index = 1, template: TemplateId = "s1") { const base = (record.insuredName || `سجل-${String(index).padStart(2, "0")}`).replace(/[^\u0600-\u06ff\w\s-]/g, "").trim().replace(/\s+/g, "-"); return `${template === "s6" ? "س6" : "س1"}-${base || `سجل-${String(index).padStart(2, "0")}`}.pdf`; }
export function downloadBlob(data: BlobPart, fileName: string, type = "application/octet-stream") { const blob = new Blob([data], { type }); const href = URL.createObjectURL(blob); const anchor = document.createElement("a"); anchor.href = href; anchor.download = fileName; document.body.appendChild(anchor); anchor.click(); anchor.remove(); window.setTimeout(() => URL.revokeObjectURL(href), 1000); }
export async function createZip(records: PersonRecord[], template: TemplateId = "s1") { const zip = new JSZip(); for (let index = 0; index < records.length; index += 1) { const record = records[index]; zip.file(safeFileName(record, index + 1, template), await fillPdf(record, template)); } return zip.generateAsync({ type: "blob", compression: "DEFLATE", compressionOptions: { level: 6 } }); }
