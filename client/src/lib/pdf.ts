// Design reminder: exported documents preserve the original official form; the app adds only the user's data and keeps the workflow quiet and auditable.

import { PDFDocument, type PDFTextField, type PDFFont } from "pdf-lib";
import fontkit from "@pdf-lib/fontkit";
import * as reshaperPackage from "arabic-persian-reshaper";
import JSZip from "jszip";
import type { PersonRecord } from "./form";

export const TEMPLATE_URL = "/manus-storage/taamin-template_d1277e6e.pdf";
const ARABIC_FONT_URL = "/manus-storage/NotoNaskhArabic-Regular_4e88e6ff.ttf";

const ArabicShaper = (reshaperPackage as unknown as { ArabicShaper?: { convertArabic: (value: string) => string } }).ArabicShaper;

const textBindings: Partial<Record<keyof PersonRecord, string[]>> = {
  office: ["Text55"],
  establishmentName: ["Text6"],
  insuredName: ["Text11"],
  qualification: ["Text14"],
  profession: ["Text17"],
  professionCode: ["Text16"],
  startDate: ["Text15"],
  country: ["Text31"],
  city: ["Text30"],
  governorate: ["Text1"],
  district: ["Text28"],
  street: ["Text27"],
  center: ["Text26"],
  phone: ["Text45"],
  email: ["Text44"],
  employer: ["Text33"],
  manager: ["Text32"],
  releaseDate: ["Text50"],
  basicWage: ["Text19", "Text18"],
  variableWage: ["Text25"],
  totalWage: ["Text20"],
  increaseDate: ["Text22", "Text21"],
  increasePercent: ["Text23"],
};

const boxBindings: Partial<Record<keyof PersonRecord, string[]>> = {
  establishmentNumber: ["Text61", "Text60", "Text59", "Text58", "Text57", "Text56", "Text54", "Text53"],
  insuranceNumber: ["Text69", "Text68", "Text67", "Text66", "Text65", "Text64", "Text63", "Text62", "Text7"],
  nationalId: ["Text82", "Text81", "Text80", "Text79", "Text78", "Text77", "Text76", "Text75", "Text74", "Text73", "Text72", "Text71", "Text70", "Text8"],
};

const checkboxBindings: Partial<Record<keyof PersonRecord, Record<string, string>>> = {
  category: {
    "عاملين لدى الغير": "Check Box1",
    "المصريين بالخارج": "Check Box2",
    "أصحاب أعمال": "Check Box3",
    "عمالة غير منتظمة": "Check Box4",
  },
};

function hasArabic(value: string) {
  return /[\u0600-\u06ff]/.test(value);
}

function toPdfText(value: string) {
  if (!value || !hasArabic(value) || !ArabicShaper) return value;
  return value
    .split(/([0-9A-Za-z%/().,\- ]+)/)
    .map((part) => {
      if (!hasArabic(part)) return part;
      return ArabicShaper.convertArabic(part).split("").reverse().join("");
    })
    .join("");
}

function cleanDigits(value: string) {
  return value.replace(/[^0-9٠-٩]/g, "").replace(/[٠-٩]/g, (digit) => String("٠١٢٣٤٥٦٧٨٩".indexOf(digit)));
}

function field<T>(form: ReturnType<PDFDocument["getForm"]>, name: string, getter: (form: ReturnType<PDFDocument["getForm"]>, name: string) => T) {
  try {
    return getter(form, name);
  } catch {
    return null;
  }
}

function setText(form: ReturnType<PDFDocument["getForm"]>, name: string, value: string, font: PDFFont) {
  const textField = field<PDFTextField>(form, name, (current, fieldName) => current.getTextField(fieldName));
  if (!textField) return;
  try {
    textField.setText(toPdfText(value));
    textField.setFontSize(Math.max(7, Math.min(11, value.length > 22 ? 8 : 10)));
    textField.updateAppearances(font);
  } catch {
    // A malformed field should not block the remaining fields in a batch.
  }
}

function setBoxes(form: ReturnType<PDFDocument["getForm"]>, names: string[], rawValue: string, font: PDFFont) {
  const digits = cleanDigits(rawValue);
  names.forEach((name, index) => setText(form, name, digits[index] ?? "", font));
}

function setCheckboxes(form: ReturnType<PDFDocument["getForm"]>, category: string) {
  const categoryMap = checkboxBindings.category ?? {};
  Object.entries(categoryMap).forEach(([label, fieldName]) => {
    const checkbox = field(form, fieldName, (current, name) => current.getCheckBox(name));
    if (!checkbox) return;
    try {
      if (label === category) checkbox.check();
      else checkbox.uncheck();
    } catch {
      // Keep going if a checkbox has a non-standard appearance state.
    }
  });
}

export async function fillPdf(record: PersonRecord) {
  const [templateResponse, fontResponse] = await Promise.all([fetch(TEMPLATE_URL), fetch(ARABIC_FONT_URL)]);
  if (!templateResponse.ok) throw new Error("تعذر تحميل قالب النموذج");
  if (!fontResponse.ok) throw new Error("تعذر تحميل الخط العربي");

  const pdfDoc = await PDFDocument.load(await templateResponse.arrayBuffer());
  pdfDoc.registerFontkit(fontkit);
  const arabicFont = await pdfDoc.embedFont(await fontResponse.arrayBuffer(), { subset: true });
  const form = pdfDoc.getForm();

  (Object.keys(textBindings) as Array<keyof PersonRecord>).forEach((key) => {
    const value = record[key];
    if (!value) return;
    (textBindings[key] ?? []).forEach((name) => setText(form, name, value, arabicFont));
  });

  (Object.keys(boxBindings) as Array<keyof PersonRecord>).forEach((key) => {
    const value = record[key];
    if (!value) return;
    setBoxes(form, boxBindings[key] ?? [], value, arabicFont);
  });

  setCheckboxes(form, record.category);
  form.updateFieldAppearances(arabicFont);
  return pdfDoc.save({ useObjectStreams: false });
}

export function safeFileName(record: PersonRecord, index = 1) {
  const base = (record.insuredName || `سجل-${String(index).padStart(2, "0")}`)
    .replace(/[^\u0600-\u06ff\w\s-]/g, "")
    .trim()
    .replace(/\s+/g, "-");
  return `${base || `سجل-${String(index).padStart(2, "0")}`}.pdf`;
}

export function downloadBlob(data: BlobPart, fileName: string, type = "application/octet-stream") {
  const blob = new Blob([data], { type });
  const href = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = href;
  anchor.download = fileName;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(href), 1000);
}

export async function createZip(records: PersonRecord[]) {
  const zip = new JSZip();
  for (let index = 0; index < records.length; index += 1) {
    const record = records[index];
    zip.file(safeFileName(record, index + 1), await fillPdf(record));
  }
  return zip.generateAsync({ type: "blob", compression: "DEFLATE", compressionOptions: { level: 6 } });
}
