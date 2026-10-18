import * as XLSX from "xlsx";
import { PDFDocument, rgb, StandardFonts } from "pdf-lib";
import fontkit from "@pdf-lib/fontkit";
import { normalizeDigits, type PersonRecord } from "./form";

export const DECLARATION_SHEET = "إقرار صاحب العمل";
export const DECLARATION_HEADERS = ["رقم المنشأة", "اسم المنشأة", "المكتب", "مقدم الطلب", "صفة مقدم الطلب", "تليفون مقدم الطلب", "الرقم التأميني لمقدم الطلب", "الرقم القومي لمقدم الطلب", "رقم المركز الرئيسي", "رقم التسجيل الضريبي", "رقم السجل التجاري", "رقم السجل التجاري الموحد", "اسم صاحب العمل أو المدير المسؤول ثلاثي", "صفة صاحب العمل أو المدير المسؤول", "تليفون صاحب العمل أو المدير المسؤول", "تاريخ تحرير الإقرار (DD/MM/YYYY)"];
export type Declaration = Record<string, string>;
export function formatTaxRegistration(value: string) {
  const normalized = normalizeDigits(value).trim();
  if (!normalized) return "";
  if (!/^(?:\d{9}|\d{3}\s*-\s*\d{3}\s*-\s*\d{3})$/.test(normalized)) throw new Error("رقم التسجيل الضريبي يجب أن يكون ٩ أرقام بصيغة مثل 123-456-789 أو بدون شرطات.");
  const digits = normalized.replace(/\s|-/g, "");
  return `${digits.slice(0, 3)}-${digits.slice(3, 6)}-${digits.slice(6)}`;
}
export function addDeclarationSheet(workbook: XLSX.WorkBook) {
  if (workbook.Sheets[DECLARATION_SHEET]) return;
  const sheet = XLSX.utils.aoa_to_sheet([DECLARATION_HEADERS]);
  sheet["!cols"] = DECLARATION_HEADERS.map(() => ({ wch: 32 }));
  // Identifiers remain text, including leading zeros.
  for (let r = 1; r <= 100; r++) for (let c = 0; c < DECLARATION_HEADERS.length; c++) sheet[XLSX.utils.encode_cell({r,c})] = { t: "s", v: "", z: "@" };
  sheet["!ref"] = "A1:P101";
  XLSX.utils.book_append_sheet(workbook, sheet, DECLARATION_SHEET);
  const instructions = XLSX.utils.aoa_to_sheet([["تعليمات إقرار صاحب العمل"], ["اكتب سطرًا واحدًا لكل منشأة في ورقة إقرار صاحب العمل. يمكنك رفع بيانات الإقرار فقط دون تعبئة بيانات العاملين أو س٢."], ["رقم المنشأة هو مفتاح الربط. احفظ الأرقام كنص للحفاظ على الأصفار في أول الرقم."], ["التاريخ بصيغة يوم/شهر/سنة. اترك رقم المركز الرئيسي فارغًا إذا لم تكن المنشأة فرعًا."], ["من تبويب س٢ استخدم زر تحميل إقرار صاحب العمل. ملفه منفصل عن س٢ ويشمل تعليمات النموذج."], ["رقم التسجيل الضريبي يقبل أي ٩ أرقام، مثل 123-456-789. الأصفار مجرد مثال للتنسيق. يُعرض الرقم بالشرطات في الإقرار، وتُقبل الأرقام العربية أو الإدخال بدون شرطات أيضًا."], ["التوقيع والختم وبيانات استلام مكتب التأمينات تُستكمل يدويًا."]]);
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
    d["رقم التسجيل الضريبي"] = formatTaxRegistration(d["رقم التسجيل الضريبي"]);
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
export function validateDeclarationEstablishments(declarations: Declaration[], records: PersonRecord[]) {
  const numberKey = (value: string) => normalizeDigits(value).replace(/\s/g, "").replace(/^0+(?=\d)/, "");
  for (const declaration of declarations) {
    const key = numberKey(declaration["رقم المنشأة"]);
    const matches = records.filter(record => numberKey(record.establishmentNumber) === key);
    if (!matches.length) throw new Error(`رقم المنشأة ${declaration["رقم المنشأة"]} في الإقرار غير موجود في بيانات س٢. راجع ورقتي الإكسيل.`);
    const nameKey = (value: string) => value.trim().replace(/\s+/g, " ");
    if (matches.some(record => record.establishmentName && nameKey(record.establishmentName) !== nameKey(declaration["اسم المنشأة"]))) throw new Error(`اسم المنشأة ${declaration["رقم المنشأة"]} مختلف بين الإقرار وبيانات س٢. وحّد الاسم في ورقتي الإكسيل.`);
  }
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
  const overlays: Array<{value: string; x: number; y: number; width: number; max: number}> = [];
  const text = (header: string, x: number, y: number, width: number, max = 13) => {
    const value = (data[header] || ""); if (!value) return;
    overlays.push({value, x, y, width, max});
    if (typeof document !== "undefined") return;
    const size = Math.min(max * scale, width * scale / Math.max(font.widthOfTextAtSize(value, 1), 1));
    page.drawText(value, {x: (x + width) * scale - font.widthOfTextAtSize(value, size), y: page.getHeight() - y * scale, size, font});
  };
  const boxes = (header: string, x: number, top: number, width: number, count: number) => {
    const raw = header === "رقم التسجيل الضريبي" ? formatTaxRegistration(data[header] || "").replace(/-/g, "") : data[header] || "";
    const value = normalizeDigits(raw).replace(/\s/g, "");
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
  boxes("رقم المركز الرئيسي",244,362,228,8);
  const taxNumber = formatTaxRegistration(data["رقم التسجيل الضريبي"] || "");
  if (taxNumber) {
    // Keep the outer border, replacing the nine inner cells with one formatted value.
    page.drawRectangle({x: 376 * scale, y: page.getHeight() - 424 * scale, width: 253 * scale, height: 26 * scale, color: rgb(1, 1, 1)});
    const taxFont = await pdf.embedFont(StandardFonts.Helvetica);
    const size = 16 * scale;
    page.drawText(taxNumber, {x: 502.5 * scale - taxFont.widthOfTextAtSize(taxNumber, size) / 2, y: page.getHeight() - 417 * scale, size, font: taxFont});
  }
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
  if (typeof document !== "undefined") {
    const face = new FontFace("DeclarationArabic", fontBytes);
    await face.load(); document.fonts.add(face);
    try {
      for (const overlay of overlays) {
        const canvas = document.createElement("canvas");
        const resolution = 3;
        const width = overlay.width * scale;
        const fontSize = overlay.max * scale;
        const height = fontSize * 2.5;
        canvas.width = Math.ceil(width * resolution); canvas.height = Math.ceil(height * resolution);
        const ctx = canvas.getContext("2d");
        if (!ctx) throw new Error("تعذر تجهيز النص العربي للإقرار");
        ctx.scale(resolution, resolution);
        ctx.direction = "rtl"; ctx.textAlign = "right";
        ctx.font = `${fontSize}px DeclarationArabic`;
        const fitted = Math.min(fontSize, fontSize * (width - 4) / Math.max(ctx.measureText(overlay.value).width, 1));
        if (fitted < 6) throw new Error(`النص طويل جدًا لخانة الإقرار: ${overlay.value}. اختصره للحفاظ على وضوح الطباعة.`);
        ctx.font = `${fitted}px DeclarationArabic`;
        // Match the PDF baseline; leave room above and below for Arabic marks.
        ctx.fillText(overlay.value, width - 2, fontSize * 1.5);
        const png = await pdf.embedPng(canvas.toDataURL("image/png"));
        page.drawImage(png, {x: overlay.x * scale, y: page.getHeight() - (overlay.y - 3) * scale - fontSize, width, height});
      }
    } finally { document.fonts.delete(face); }
  }
  return pdf.save();
}
