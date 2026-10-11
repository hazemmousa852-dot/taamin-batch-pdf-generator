import JSZip from "jszip";
import * as XLSX from "xlsx";
import { FORM31_SHEET_COLUMNS, FORM31_SHEET_NAME } from "./form";

/**
 * يبني قالب Excel لنموذج 31 انطلاقًا من القالب المرن الحالي: يحتفظ بورقة
 * «بيانات المنشأة والمفوض» (تُملأ مرة واحدة وتُضاف تلقائيًا لكل سجل) وبورقة
 * «قوائم الأكواد»، ويضيف ورقة «نموذج 31» بصف واحد لكل إصابة أو إجازة مرضية.
 */

const SHARED_SHEET_NAME = "بيانات المنشأة والمفوض";
const PEOPLE_SHEET_NAME = "بيانات المؤمن عليهم";
const CODES_SHEET_NAME = "قوائم الأكواد";

/** قوائم منسدلة تُضاف إلى ورقة نموذج 31 لتقليل الأخطاء في حقول نعم/لا. */
const LIST_RULES: Array<{ key: string; values: string[] }> = [
  { key: "companion", values: ["نعم", "لا"] },
  { key: "travelExpenses", values: ["نعم", "لا"] },
  { key: "treatmentPlace", values: ["داخل البلاد", "خارج البلاد"] },
];

function escapeXml(value: string) {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

/** يعثر على مسار ملف XML الخاص بورقة معيّنة داخل ملف xlsx. */
async function findSheetPath(zip: JSZip, sheetName: string) {
  const workbookEntry = zip.file("xl/workbook.xml");
  const relsEntry = zip.file("xl/_rels/workbook.xml.rels");
  if (!workbookEntry || !relsEntry) return undefined;
  const workbookXml = await workbookEntry.async("string");
  const escaped = sheetName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const match = workbookXml.match(new RegExp(`<sheet[^>]*name="${escaped}"[^>]*r:id="([^"]+)"`));
  if (!match) return undefined;
  const relsXml = await relsEntry.async("string");
  const relMatch = relsXml.match(new RegExp(`<Relationship[^>]*Id="${match[1]}"[^>]*Target="([^"]+)"`));
  if (!relMatch) return undefined;
  const target = relMatch[1].startsWith("/") ? relMatch[1].slice(1) : `xl/${relMatch[1].replace(/^\.\//, "")}`;
  return target;
}

async function decorateForm31Sheet(data: ArrayBuffer) {
  const zip = await JSZip.loadAsync(data);
  const path = await findSheetPath(zip, FORM31_SHEET_NAME);
  if (!path) return data;
  const entry = zip.file(path);
  if (!entry) return data;
  let xml = await entry.async("string");

  // تثبيت صف العناوين حتى يبقى واضحًا أثناء الكتابة
  xml = xml.replace(/<sheetView([^>]*)>/, (opening) =>
    xml.includes("<pane ") ? opening : `${opening}<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>`);

  // قوائم منسدلة لحقول نعم/لا ومكان العلاج
  const validations = LIST_RULES.flatMap((rule) => {
    const index = FORM31_SHEET_COLUMNS.findIndex((column) => column.key === rule.key);
    if (index < 0) return [];
    const column = XLSX.utils.encode_col(index);
    const list = rule.values.map(escapeXml).join(",");
    return [
      `<dataValidation type="list" allowBlank="1" showErrorMessage="1" errorTitle="قيمة غير صحيحة" error="اختر قيمة من القائمة" sqref="${column}2:${column}1000"><formula1>"${list}"</formula1></dataValidation>`,
    ];
  });
  if (validations.length && !xml.includes("<dataValidations")) {
    xml = xml.replace("</worksheet>", `<dataValidations count="${validations.length}">${validations.join("")}</dataValidations></worksheet>`);
  }
  zip.file(path, xml);
  return zip.generateAsync({ type: "arraybuffer", compression: "DEFLATE" });
}

/** يبني ملف Excel جاهزًا للتنزيل: أوراق مشتركة + ورقة نموذج 31. */
export async function buildForm31Workbook(templateBytes: ArrayBuffer | Uint8Array) {
  // SheetJS يتوقع Uint8Array عند القراءة من الذاكرة
  const source = XLSX.read(
    templateBytes instanceof Uint8Array ? templateBytes : new Uint8Array(templateBytes),
    { type: "array" },
  );
  const headerRow = FORM31_SHEET_COLUMNS.map((column) => column.label);
  const form31Sheet = XLSX.utils.aoa_to_sheet([headerRow]);
  form31Sheet["!cols"] = FORM31_SHEET_COLUMNS.map((column) => ({
    wch: Math.min(42, Math.max(20, column.label.length + 6)),
  }));
  form31Sheet["!freeze"] = { xSplit: 0, ySplit: 1 };

  const output = XLSX.utils.book_new();
  const order = [SHARED_SHEET_NAME, FORM31_SHEET_NAME, PEOPLE_SHEET_NAME, CODES_SHEET_NAME];
  for (const name of order) {
    const worksheet = name === FORM31_SHEET_NAME ? form31Sheet : source.Sheets[name];
    if (worksheet) XLSX.utils.book_append_sheet(output, worksheet, name);
  }
  // أي ورقة إضافية في القالب الأصلي تبقى محفوظة في النهاية
  for (const name of source.SheetNames) {
    if (!order.includes(name)) XLSX.utils.book_append_sheet(output, source.Sheets[name], name);
  }

  const bytes = XLSX.write(output, { bookType: "xlsx", type: "array" }) as ArrayBuffer;
  return decorateForm31Sheet(bytes);
}
