import { PDFDocument, PDFHexString, PDFTextField, TextAlignment, type PDFFont } from "pdf-lib";
import fontkit from "@pdf-lib/fontkit";
import * as reshaperPackage from "arabic-persian-reshaper";
import JSZip from "jszip";
import { normalizeDigits, normalizeText, type PersonRecord, type TemplateId } from "./form";

const configuredBase = import.meta.env?.BASE_URL ?? "/";
const ASSET_BASE = configuredBase.endsWith("/") ? configuredBase : `${configuredBase}/`;
const assetUrl = (name: string) => `${ASSET_BASE}assets/${name}`;
export const TEMPLATE_URLS: Record<TemplateId, string> = { s1: assetUrl("taamin-template.pdf"), s6: assetUrl("taamin-s6-template.pdf") };
const ARABIC_FONT_URL = assetUrl("NotoNaskhArabic-Regular.ttf");
const ArabicShaper = (reshaperPackage as unknown as { ArabicShaper?: { convertArabic: (value: string) => string } }).ArabicShaper;

const assetCache = new Map<string, Promise<ArrayBuffer>>();
function fetchAsset(url: string, label: string) {
  let pending = assetCache.get(url);
  if (!pending) {
    pending = fetch(url).then(async (response) => {
      if (!response.ok) throw new Error(`تعذر تحميل ${label}`);
      return response.arrayBuffer();
    });
    assetCache.set(url, pending);
  }
  return pending;
}

// The official PDFs use visually reversed names. These maps were verified by
// widget coordinates; changing a template must be followed by the PDF tests.
const s1TextBindings: Partial<Record<keyof PersonRecord, string[]>> = {
  office: ["بتكم"], applicantName: [": ةفصبلطلا مدقم_1"], applicantRole: ["Text Field6"], applicantPhone: ["يلتلا مقر:نوف"],
  insuredName: ["fill_3", "انأ رقأ"], qualification: ["ةنهلما_1"], profession: ["ةنهلما"], country: ["ةيــــــــــــسنلجا"],
  governorate: [": ةظفامح"], district: ["ةيرق"], street: ["ةيرق_1"], center: [": ةظفامح_1"],
  establishmentName: ["fill_4", ":نييمأتلا اهمقر_1"], address: [":ناوــــــــــــنعلا"], buildingNumber: ["fill_5"], workType: ["fill_6"], contributionCode: ["fill_7"], increasePercent: ["%"], phone: ["fill_2"], sector: [": عاطقلا"],
};
const s1MoneyBindings: Partial<Record<keyof PersonRecord, string[]>> = { basicWage: ["Text Field4"], totalWage: ["Text Field5"] };
const s1BoxBindings: Partial<Record<keyof PersonRecord, string[]>> = {
  establishmentNumber: ["أشنلما مقرة", ":نييمأتلا اهمقر"], applicantInsuranceNumber: ["Text Field3"], insuranceNumber: [":يــــنيمأتلا مـــقرلا"],
  nationalId: ["ةيــــــــــــسنلجا_1"], applicantNationalId: ["مقر: ىموق"],
};
const s1DateBindings: Partial<Record<keyof PersonRecord, string[]>> = {
  startDate: ["Text Field10", "Text Field9", "Text Field8"], increaseDate: ["Text Field2", "Text Field1", "Text Field0"],
};

const S6_PAGE2_INSURANCE = "ا: هيلع نمؤلما مــــــــس---------------------------------------:نييمأتلا همقر";
const s6TextBindings: Partial<Record<keyof PersonRecord, string[]>> = {
  office: ["بتكم"], applicantRole: [":بلطلا مدقم ةفص"], applicantName: [":بلطلا مدقم ةفص_1"], applicantPhone: [":نوفيلتلا مقر"],
  insuredName: ["Text Field1", "Text Field7"], establishmentName: [":هأشنلما مسا", "Text Field5"],
  applicantInsuranceNumber: ["Text Field0"], insuranceNumber: [": نييمأتلا مقرلا", S6_PAGE2_INSURANCE], address: ["Text Field6"], endReason: ["fill_1"],
};
const s6BoxBindings: Partial<Record<keyof PersonRecord, string[]>> = {
  applicantNationalId: ["موقلا مقري"], establishmentNumber: [":هأشنلما مسا_1", ":نييمأتلا اهمقر"], nationalId: [": يـموقلا مقرلا"],
};
const checkboxBindings: Partial<Record<keyof PersonRecord, Record<string, string>>> = {
  category: { "عاملين لدى الغير": "يرغلا ىدل ينلماع", "أصحاب أعمال": "لامعأ باحصأشنم ملهآت", "عمالة غير منتظمة": "زباخلماب ينلماعلا" },
  medicalExam: { "نعم": "ئادتبلاا بيطلا فشكلا ءافيتساي", "لا": "لا" },
  establishmentType: { "نمطي": "أشنلما عونةطنم :ي", "سيارة": "رايسة", "مركب صيد": "ديص بكرم", "مخابز بلدية": "ةيدلب زبامخ" },
};

function hasArabic(value: string) { return /[\u0600-\u06ff]/.test(value); }
function toArabicNumerals(value: string) {
  return value.replace(/\d/g, (digit) => "٠١٢٣٤٥٦٧٨٩"[Number(digit)]);
}
export function toPdfText(rawValue: string) {
  const value = normalizeText(rawValue);
  if (!value || !hasArabic(value) || !ArabicShaper) return value;
  // PDF appearance streams paint glyphs from left to right. Shape the whole
  // Arabic sentence, then reverse the complete visual run (including word
  // order). Keep Latin text and numbers protected so their characters do not
  // get reversed, e.g. 2026 and phone numbers remain readable.
  const protectedTokens: string[] = [];
  const masked = value.replace(/[0-9A-Za-z@%/().,\-+]+/g, (token) => {
    const marker = String.fromCharCode(0xe000 + protectedTokens.length);
    protectedTokens.push(token);
    return marker;
  });
  const shaped = ArabicShaper.convertArabic(masked);
  const visualOrder = Array.from(shaped).reverse().join("");
  return visualOrder.replace(/[\ue000-\uf8ff]/g, (marker) => protectedTokens[marker.charCodeAt(0) - 0xe000] ?? marker);
}
function cleanDigits(value: string) { return normalizeDigits(value).replace(/\D/g, ""); }
function getTextField(form: ReturnType<PDFDocument["getForm"]>, name: string) {
  try { return form.getTextField(name); } catch { throw new Error(`حقل PDF غير موجود: ${name}`); }
}
function fittedFontSize(textField: ReturnType<typeof getTextField>, text: string, font: PDFFont) {
  const widget = textField.acroField.getWidgets()[0];
  const width = widget?.getRectangle().width ?? 120;
  const availableWidth = Math.max(8, width - 5);
  let size = 12;
  while (size > 7.5 && font.widthOfTextAtSize(text, size) > availableWidth) size -= 0.5;
  return size;
}
function setText(form: ReturnType<PDFDocument["getForm"]>, name: string, value: string, font: PDFFont, alignment: TextAlignment = TextAlignment.Right) {
  const textField = getTextField(form, name);
  const normalized = normalizeText(value);
  const maxLength = typeof textField.getMaxLength === "function" ? textField.getMaxLength() : undefined;
  // The official PDFs contain legacy character limits that are sometimes
  // shorter than valid insurance and phone numbers. Preserve the user's full
  // value and let the generated appearance reduce the font size to fit.
  if (maxLength && normalized.length > maxLength) textField.removeMaxLength();
  const shapedText = toPdfText(normalized);
  const pdfText = typeof document === "undefined" ? shapedText : toArabicNumerals(shapedText);
  textField.setText(pdfText);
  textField.setAlignment(alignment);
  // Match the printed form's text scale. Start at 12 pt and shrink only when
  // the actual shaped Arabic glyphs would exceed this specific field's width.
  const fontSize = fittedFontSize(textField, pdfText, font);
  // Several fields in the official form have no /DA entry. Without it,
  // pdf-lib silently falls back to auto-sizing and Arabic becomes extremely
  // small. Seed a valid appearance first, then apply a consistent readable size.
  if (!textField.acroField.getDefaultAppearance()) textField.acroField.setDefaultAppearance("/Helv 12 Tf 0 g");
  try {
    textField.setFontSize(fontSize);
  } catch {
    // A few official fields contain a /DA string but omit the required Tf
    // operator. Replace that malformed appearance and retry with our size.
    textField.acroField.setDefaultAppearance(`/Helv ${fontSize} Tf 0 g`);
    textField.setFontSize(fontSize);
  }
  textField.updateAppearances(font);
  // Keep the canonical field value as normal Unicode Arabic. The appearance
  // stream above retains the shaped visual glyphs required by PDF renderers,
  // while copying/searching/reading the field now returns "اسم المؤمن عليه"
  // instead of Arabic Presentation Forms such as "ﻢﺳﺍ".
  textField.acroField.setValue(PDFHexString.fromText(normalized));
}
function setBoxes(form: ReturnType<PDFDocument["getForm"]>, names: string[], rawValue: string, font: PDFFont) {
  const digits = cleanDigits(rawValue);
  names.forEach((name) => setText(form, name, digits, font, TextAlignment.Center));
}
function setDateFields(form: ReturnType<PDFDocument["getForm"]>, names: string[], rawValue: string, font: PDFFont) {
  const match = normalizeDigits(rawValue).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) throw new Error(`صيغة التاريخ غير صحيحة: ${rawValue}`);
  const [, year, month, day] = match;
  [day, month, year].forEach((part, index) => setText(form, names[index], part, font, TextAlignment.Center));
}
function setMoneyFields(form: ReturnType<PDFDocument["getForm"]>, names: string[], rawValue: string, font: PDFFont) {
  const normalized = normalizeDigits(rawValue).replace(/[,،]/g, ".");
  if (!/^\d+(\.\d{1,2})?$/.test(normalized)) throw new Error(`قيمة الأجر غير صحيحة: ${rawValue}`);
  const [whole, fraction = ""] = normalized.split(".");
  setText(form, names[0], whole, font, TextAlignment.Center);
  if (names[1]) setText(form, names[1], fraction.padEnd(2, "0").slice(0, 2), font, TextAlignment.Center);
}
function setCheckboxes(form: ReturnType<PDFDocument["getForm"]>, key: keyof PersonRecord, value: string) {
  for (const [label, fieldName] of Object.entries(checkboxBindings[key] ?? {})) {
    let checkbox;
    try { checkbox = form.getCheckBox(fieldName); } catch { throw new Error(`خانة اختيار PDF غير موجودة: ${fieldName}`); }
    label === value ? checkbox.check() : checkbox.uncheck();
  }
}

async function bakeArabicText(
  pdfDoc: PDFDocument,
  form: ReturnType<PDFDocument["getForm"]>,
  font: PDFFont,
  fontBytes: ArrayBuffer,
) {
  if (typeof document === "undefined" || typeof FontFace === "undefined") return false;
  const pages = pdfDoc.getPages();
  const overlays: Array<{ pageIndex: number; x: number; y: number; width: number; height: number; text: string; alignment: TextAlignment }> = [];

  for (const candidate of form.getFields()) {
    if (!(candidate instanceof PDFTextField)) continue;
    const field = candidate;
    const rawText = normalizeText(field.getText() ?? "");
    if (!hasArabic(rawText)) continue;
    const text = toArabicNumerals(rawText);
    const alignment = field.getAlignment();
    let placed = false;
    for (const widget of field.acroField.getWidgets()) {
      const pageRef = widget.P()?.toString();
      let pageIndex = pages.findIndex((page) => page.ref.toString() === pageRef);
      if (pageIndex < 0) {
        pageIndex = pages.findIndex((page) => {
          const annots = page.node.Annots();
          return Boolean(annots?.asArray().some((annot) => pdfDoc.context.lookup(annot) === widget.dict));
        });
      }
      if (pageIndex < 0) continue;
      const { x, y, width, height } = widget.getRectangle();
      overlays.push({ pageIndex, x, y, width, height, text, alignment });
      placed = true;
    }
    // Flatten a blank appearance; the correctly shaped Arabic is painted on
    // top afterwards by the browser's native Arabic text engine.
    if (placed) {
      field.setText("");
      field.updateAppearances(font);
    }
  }

  form.flatten({ updateFieldAppearances: false });
  const face = new FontFace("TaaminArabic", fontBytes.slice(0));
  await face.load();
  document.fonts.add(face);
  const scale = 4;

  for (const overlay of overlays) {
    const canvas = document.createElement("canvas");
    canvas.width = Math.max(1, Math.ceil(overlay.width * scale));
    canvas.height = Math.max(1, Math.ceil(overlay.height * scale));
    const context = canvas.getContext("2d");
    if (!context) throw new Error("تعذر تجهيز محرك رسم العربية");
    context.direction = "rtl";
    context.textBaseline = "middle";
    context.fillStyle = "#000";
    let fontSize = Math.min(14, Math.max(10, overlay.height * 0.84));
    context.font = `${fontSize * scale}px TaaminArabic`;
    const maxWidth = Math.max(8, (overlay.width - 5) * scale);
    while (fontSize > 7.5 && context.measureText(overlay.text).width > maxWidth) {
      fontSize -= 0.5;
      context.font = `${fontSize * scale}px TaaminArabic`;
    }
    const centered = overlay.alignment === TextAlignment.Center;
    context.textAlign = centered ? "center" : "right";
    context.fillText(overlay.text, centered ? canvas.width / 2 : canvas.width - (2 * scale), canvas.height / 2, maxWidth);
    const pngBlob = await new Promise<Blob>((resolve, reject) => canvas.toBlob((blob) => blob ? resolve(blob) : reject(new Error("تعذر رسم النص العربي")), "image/png"));
    const png = await pdfDoc.embedPng(await pngBlob.arrayBuffer());
    pages[overlay.pageIndex].drawImage(png, { x: overlay.x, y: overlay.y, width: overlay.width, height: overlay.height });
  }
  return true;
}

export async function fillPdf(record: PersonRecord, template: TemplateId = "s1") {
  const [templateBytes, fontBytes] = await Promise.all([fetchAsset(TEMPLATE_URLS[template], "قالب النموذج"), fetchAsset(ARABIC_FONT_URL, "الخط العربي")]);
  const pdfDoc = await PDFDocument.load(templateBytes.slice(0));
  pdfDoc.registerFontkit(fontkit);
  const arabicFont = await pdfDoc.embedFont(fontBytes.slice(0), { subset: true });
  const form = pdfDoc.getForm();
  const textBindings = template === "s6" ? s6TextBindings : s1TextBindings;
  const boxBindings = template === "s6" ? s6BoxBindings : s1BoxBindings;
  for (const key of Object.keys(textBindings) as Array<keyof PersonRecord>) if (record[key]) for (const name of textBindings[key] ?? []) setText(form, name, record[key], arabicFont);
  for (const key of Object.keys(boxBindings) as Array<keyof PersonRecord>) if (record[key]) setBoxes(form, boxBindings[key] ?? [], record[key], arabicFont);
  if (template === "s1") {
    for (const key of Object.keys(s1DateBindings) as Array<keyof PersonRecord>) if (record[key] && (key !== "increaseDate" || /^\d{4}-\d{2}-\d{2}$/.test(record[key]))) setDateFields(form, s1DateBindings[key] ?? [], record[key], arabicFont);
    for (const key of Object.keys(s1MoneyBindings) as Array<keyof PersonRecord>) if (record[key]) setMoneyFields(form, s1MoneyBindings[key] ?? [], record[key], arabicFont);
    if (record.category) setCheckboxes(form, "category", record.category);
    if (record.medicalExam) setCheckboxes(form, "medicalExam", record.medicalExam);
    if (record.establishmentType) setCheckboxes(form, "establishmentType", record.establishmentType);
  } else if (record.endDate) setDateFields(form, ["Text Field2", "Text Field3", "Text Field4"], record.endDate, arabicFont);
  const baked = await bakeArabicText(pdfDoc, form, arabicFont, fontBytes);
  if (!baked) for (const field of form.getFields()) field.enableReadOnly();
  // Every populated text field is updated in setText. Avoid updating every
  // field in the source PDF here: some unused official fields have no /DA
  // font operator and make pdf-lib throw "No Tf operator found".
  return pdfDoc.save({ useObjectStreams: false });
}

export function safeFileName(record: PersonRecord, index = 1, template: TemplateId = "s1") {
  const person = (record.insuredName || "سجل").replace(/[^\u0600-\u06ff\w\s-]/g, "").trim().replace(/\s+/g, "-");
  const identity = cleanDigits(record.insuranceNumber || record.nationalId).slice(-10) || String(index).padStart(3, "0");
  return `${template === "s6" ? "س6" : "س1"}-${person || "سجل"}-${identity}-${String(index).padStart(3, "0")}.pdf`;
}
export function downloadBlob(data: BlobPart, fileName: string, type = "application/octet-stream") {
  const blob = new Blob([data], { type }); const href = URL.createObjectURL(blob); const anchor = document.createElement("a");
  anchor.href = href; anchor.download = fileName; document.body.appendChild(anchor); anchor.click(); anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(href), 1000);
}
export async function createZip(records: PersonRecord[], template: TemplateId = "s1", onProgress?: (done: number, total: number) => void) {
  const zip = new JSZip();
  for (let index = 0; index < records.length; index += 1) {
    zip.file(safeFileName(records[index], index + 1, template), await fillPdf(records[index], template));
    onProgress?.(index + 1, records.length);
  }
  return zip.generateAsync({ type: "blob", compression: "DEFLATE", compressionOptions: { level: 6 } });
}

export async function createMergedPdf(records: PersonRecord[], template: TemplateId = "s1", onProgress?: (done: number, total: number) => void) {
  const merged = await PDFDocument.create();
  for (let index = 0; index < records.length; index += 1) {
    const source = await PDFDocument.load(await fillPdf(records[index], template));
    const pages = await merged.copyPages(source, source.getPageIndices());
    pages.forEach((page) => merged.addPage(page));
    onProgress?.(index + 1, records.length);
  }
  return merged.save({ useObjectStreams: false });
}

