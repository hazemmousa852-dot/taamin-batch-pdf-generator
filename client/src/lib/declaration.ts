import * as XLSX from "xlsx";
import { PDFDocument } from "pdf-lib";
import fontkit from "@pdf-lib/fontkit";
import { toPdfText } from "./pdf";
import { normalizeDigits } from "./form";

export const DECLARATION_SHEET = "إقرار صاحب العمل";
export const DECLARATION_HEADERS = ["رقم المنشأة", "اسم المنشأة", "المكتب", "مقدم الطلب", "صفة مقدم الطلب", "تليفون مقدم الطلب", "الرقم التأميني لمقدم الطلب", "الرقم القومي لمقدم الطلب", "رقم المركز الرئيسي", "رقم التسجيل الضريبي", "رقم السجل التجاري", "رقم السجل التجاري الموحد", "اسم صاحب العمل أو المدير المسؤول ثلاثي", "صفة صاحب العمل أو المدير المسؤول", "تليفون صاحب العمل أو المدير المسؤول", "تاريخ تحرير الإقرار (DD/MM/YYYY)"];
export type Declaration = Record<string, string>;
export function addDeclarationSheet(workbook: XLSX.WorkBook) {
  if (workbook.Sheets[DECLARATION_SHEET]) return;
  const sheet = XLSX.utils.aoa_to_sheet([DECLARATION_HEADERS]);
  sheet["!cols"] = DECLARATION_HEADERS.map(() => ({ wch: 32 }));
  // Identifiers remain text, including leading zeros.
  for (let r = 1; r <= 100; r++) for (let c = 0; c < DECLARATION_HEADERS.length; c++) sheet[XLSX.utils.encode_cell({r,c})] = { t: "s", v: "", z: "@" };
  sheet["!ref"] = "A1:P101";
  XLSX.utils.book_append_sheet(workbook, sheet, DECLARATION_SHEET);
  const instructions = XLSX.utils.aoa_to_sheet([["تعليمات إقرار صاحب العمل"], ["اكتب سطرًا واحدًا لكل منشأة في ورقة إقرار صاحب العمل. لا تكرر البيانات لكل عامل."], ["رقم المنشأة هو مفتاح الربط. احفظ الأرقام كنص للحفاظ على الأصفار في أول الرقم."], ["التاريخ بصيغة يوم/شهر/سنة. اترك رقم المركز الرئيسي فارغًا إذا لم تكن المنشأة فرعًا."], ["من تبويب س٢ استخدم زر تحميل إقرار صاحب العمل. ملفه منفصل عن س٢ ويشمل تعليمات النموذج."], ["التوقيع والختم وبيانات استلام مكتب التأمينات تُستكمل يدويًا."]]);
  instructions["!cols"] = [{wch: 110}];
  XLSX.utils.book_append_sheet(workbook, instructions, "تعليمات الإقرار");
}
export function readDeclarations(workbook: XLSX.WorkBook): Declaration[] {
  const sheet = workbook.Sheets[DECLARATION_SHEET];
  if (!sheet) return [];
  const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, {defval: "", raw: false});
  const unique = new Map<string, Declaration>();
  for (const row of rows) {
    const d = Object.fromEntries(DECLARATION_HEADERS.map(h => [h, String(row[h] ?? "").trim()]));
    if (!Object.values(d).some(Boolean)) continue;
    const key = normalizeDigits(d["رقم المنشأة"]).replace(/\s/g, "");
    if (!/^\d{1,8}$/.test(key)) throw new Error("أدخل رقم المنشأة (حتى ٨ أرقام) لكل إقرار في ورقة إقرار صاحب العمل.");
    d["رقم المنشأة"] = key;
    for (const required of ["اسم المنشأة", "المكتب", "مقدم الطلب", "صفة مقدم الطلب", "اسم صاحب العمل أو المدير المسؤول ثلاثي", "صفة صاحب العمل أو المدير المسؤول"]) {
      if (!d[required]) throw new Error(`أكمل ${required} في إقرار المنشأة ${key}`);
    }
    const date = d["تاريخ تحرير الإقرار (DD/MM/YYYY)"];
    if (date) {
      const match = normalizeDigits(date).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      if (!match) throw new Error("تاريخ تحرير الإقرار يجب أن يكون يوم/شهر/سنة");
      const check = new Date(+match[3], +match[2] - 1, +match[1]);
      if (check.getDate() !== +match[1] || check.getMonth() !== +match[2] - 1 || check.getFullYear() !== +match[3]) throw new Error("تاريخ تحرير الإقرار غير صحيح");
    }
    const previous = unique.get(key);
    if (previous && JSON.stringify(previous) !== JSON.stringify(d)) throw new Error(`بيانات إقرار متعارضة للمنشأة ${key}. اترك سطرًا واحدًا لكل منشأة.`);
    unique.set(key, d);
  }
  return Array.from(unique.values());
}
export async function createDeclarationPdf(data: Declaration) {
  const base = import.meta.env?.BASE_URL ?? "/";
  const load = async (name: string) => { const r = await fetch(`${base}assets/${name}`); if (!r.ok) throw new Error("تعذر تحميل قالب الإقرار أو الخط"); return r.arrayBuffer(); };
  const [source, fontBytes] = await Promise.all([load("employer-declaration.pdf"), load("NotoNaskhArabic-Regular.ttf")]);
  const pdf = await PDFDocument.load(source);
  pdf.registerFontkit(fontkit);
  const font = await pdf.embedFont(fontBytes);
  const page = pdf.getPages()[0];
  const scale = page.getWidth() / 849;
  const text = (header: string, x: number, y: number, width: number, max = 13) => {
    const value = toPdfText(data[header] || ""); if (!value) return;
    const size = Math.min(max * scale, width * scale / Math.max(font.widthOfTextAtSize(value, 1), 1));
    page.drawText(value, {x: (x + width) * scale - font.widthOfTextAtSize(value, size), y: page.getHeight() - y * scale, size, font});
  };
  const boxes = (header: string, x: number, top: number, width: number, count: number) => {
    const value = normalizeDigits(data[header] || "").replace(/\s/g, "");
    if (value && (!/^\d+$/.test(value) || value.length > count)) throw new Error(`راجع ${header}: أرقام فقط وبحد أقصى ${count} خانة`);
    for (const [i, digit] of Array.from(value).map((digit, i) => [i, digit] as const)) {
      const size = 13 * scale; const cell = width / count;
      page.drawText(digit, {x: (x + i * cell + cell / 2) * scale - font.widthOfTextAtSize(digit,size)/2, y: page.getHeight() - (top + 20) * scale, size, font});
    }
  };
  text("المكتب", 630, 68, 160);
  text("مقدم الطلب", 450, 189, 255); text("صفة مقدم الطلب", 40, 189, 268);
  text("تليفون مقدم الطلب", 40, 228, 295); boxes("الرقم التأميني لمقدم الطلب",445,207,255,9);
  boxes("الرقم القومي لمقدم الطلب",297,248,396,14);
  text("اسم المنشأة",423,344,292); boxes("رقم المنشأة",91,328,228,8);
  boxes("رقم المركز الرئيسي",244,362,228,8); boxes("رقم التسجيل الضريبي",375,397,255,9);
  boxes("رقم السجل التجاري",375,428,255,9); boxes("رقم السجل التجاري الموحد",173,459,418,15);
  text("اسم صاحب العمل أو المدير المسؤول ثلاثي",327,557,390); text("صفة صاحب العمل أو المدير المسؤول",42,557,230);
  text("اسم صاحب العمل أو المدير المسؤول ثلاثي",485,782,232); text("صفة صاحب العمل أو المدير المسؤول",489,815,258);
  text("تليفون صاحب العمل أو المدير المسؤول",510,849,190);
  const date = normalizeDigits(data["تاريخ تحرير الإقرار (DD/MM/YYYY)"] || "").split("/");
  if (date.length === 3) {
    for (const [value, x, width] of [[date[0], 695, 25], [date[1], 663, 20], [date[2], 609, 40]] as const) {
      const size = 13 * scale;
      page.drawText(value, {x: (x + width / 2) * scale - font.widthOfTextAtSize(value,size)/2, y: page.getHeight()-881*scale, size, font});
    }
  }
  return pdf.save();
}
