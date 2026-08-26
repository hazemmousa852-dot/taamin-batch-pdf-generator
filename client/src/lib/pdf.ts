// Design reminder: exported documents preserve the original official form; the app adds only the user's data and keeps the workflow quiet and auditable.

import { PDFDocument, type PDFTextField, type PDFFont } from "pdf-lib";
import fontkit from "@pdf-lib/fontkit";
import * as reshaperPackage from "arabic-persian-reshaper";
import JSZip from "jszip";
import type { PersonRecord, TemplateId } from "./form";

export const TEMPLATE_URLS: Record<TemplateId, string> = {
  s1: new URL("assets/taamin-template.pdf", import.meta.env.BASE_URL).toString(),
  s6: new URL("assets/taamin-s6-template.pdf", import.meta.env.BASE_URL).toString(),
};
const ARABIC_FONT_URL = new URL("assets/NotoNaskhArabic-Regular.ttf", import.meta.env.BASE_URL).toString();

const ArabicShaper = (reshaperPackage as unknown as { ArabicShaper?: { convertArabic: (value: string) => string } }).ArabicShaper;

const s1TextBindings: Partial<Record<keyof PersonRecord, string[]>> = {
  office: ["Text55"], establishmentName: ["Text6"], insuredName: ["Text11"], qualification: ["Text14"], profession: ["Text17"], professionCode: ["Text16"], startDate: ["Text15"], country: ["Text31"], city: ["Text30"], governorate: ["Text1"], district: ["Text28"], street: ["Text27"], center: ["Text26"], phone: ["Text45"], email: ["Text44"], employer: ["Text33"], manager: ["Text32"], releaseDate: ["Text50"], basicWage: ["Text19", "Text18"], variableWage: ["Text25"], totalWage: ["Text20"], increaseDate: ["Text22", "Text21"], increasePercent: ["Text23"],
};
const s1BoxBindings: Partial<Record<keyof PersonRecord, string[]>> = {
  establishmentNumber: ["Text61", "Text60", "Text59", "Text58", "Text57", "Text56", "Text54", "Text53"], insuranceNumber: ["Text69", "Text68", "Text67", "Text66", "Text65", "Text64", "Text63", "Text62", "Text7"], nationalId: ["Text82", "Text81", "Text80", "Text79", "Text78", "Text77", "Text76", "Text75", "Text74", "Text73", "Text72", "Text71", "Text70", "Text8"],
};
const s6TextBindings: Partial<Record<keyof PersonRecord, string[]>> = {
  office: ["بتكم"], applicantRole: [":بلطلا مدقم ةفص"], applicantName: [":بلطلا مدقم ةفص_1"], applicantPhone: [":نوفيلتلا مقر"], insuredName: ["Text Field1"], establishmentName: [":هأشنلما مسا", ":هأشنلما مسا_1", "Text Field5"], address: ["Text Field6"], endReason: ["fill_1"],
};
const s6BoxBindings: Partial<Record<keyof PersonRecord, string[]>> = {
  applicantNationalId: ["موقلا مقري"], insuranceNumber: [": نييمأتلا مقرلا", ":نييمأتلا اهمقر", "Text Field7"], nationalId: [": يـموقلا مقرلا"],
};
const checkboxBindings: Partial<Record<keyof PersonRecord, Record<string, string>>> = { category: { "عاملين لدى الغير": "Check Box1", "المصريين بالخارج": "Check Box2", "أصحاب أعمال": "Check Box3", "عمالة غير منتظمة": "Check Box4" } };

function hasArabic(value: string) { return /[\u0600-\u06ff]/.test(value); }
function toPdfText(value: string) {
  if (!value || !hasArabic(value) || !ArabicShaper) return value;
  return value.split(/([0-9A-Za-z%/().,\- ]+)/).map((part) => hasArabic(part) ? ArabicShaper.convertArabic(part).split("").reverse().join("") : part).join("");
}
function cleanDigits(value: string) { return value.replace(/[^0-9٠-٩]/g, "").replace(/[٠-٩]/g, (digit) => String("٠١٢٣٤٥٦٧٨٩".indexOf(digit))); }
function field<T>(form: ReturnType<PDFDocument["getForm"]>, name: string, getter: (form: ReturnType<PDFDocument["getForm"]>, name: string) => T) { try { return getter(form, name); } catch { return null; } }
function setText(form: ReturnType<PDFDocument["getForm"]>, name: string, value: string, font: PDFFont) {
  const textField = field<PDFTextField>(form, name, (current, fieldName) => current.getTextField(fieldName));
  if (!textField) return;
  try { textField.setText(toPdfText(value)); textField.setFontSize(Math.max(7, Math.min(11, value.length > 22 ? 8 : 10))); textField.updateAppearances(font); } catch { /* continue with the remaining fields */ }
}
function setBoxes(form: ReturnType<PDFDocument["getForm"]>, names: string[], rawValue: string, font: PDFFont) { const digits = cleanDigits(rawValue); names.forEach((name, index) => setText(form, name, digits[index] ?? "", font)); }
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
  form.updateFieldAppearances(arabicFont);
  return pdfDoc.save({ useObjectStreams: false });
}

export function safeFileName(record: PersonRecord, index = 1, template: TemplateId = "s1") { const base = (record.insuredName || `سجل-${String(index).padStart(2, "0")}`).replace(/[^\u0600-\u06ff\w\s-]/g, "").trim().replace(/\s+/g, "-"); return `${template === "s6" ? "س6" : "س1"}-${base || `سجل-${String(index).padStart(2, "0")}`}.pdf`; }
export function downloadBlob(data: BlobPart, fileName: string, type = "application/octet-stream") { const blob = new Blob([data], { type }); const href = URL.createObjectURL(blob); const anchor = document.createElement("a"); anchor.href = href; anchor.download = fileName; document.body.appendChild(anchor); anchor.click(); anchor.remove(); window.setTimeout(() => URL.revokeObjectURL(href), 1000); }
export async function createZip(records: PersonRecord[], template: TemplateId = "s1") { const zip = new JSZip(); for (let index = 0; index < records.length; index += 1) { const record = records[index]; zip.file(safeFileName(record, index + 1, template), await fillPdf(record, template)); } return zip.generateAsync({ type: "blob", compression: "DEFLATE", compressionOptions: { level: 6 } }); }
