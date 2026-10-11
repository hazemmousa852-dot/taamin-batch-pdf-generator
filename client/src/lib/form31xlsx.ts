import JSZip from "jszip";
import * as XLSX from "xlsx";
import { FORM31_SHEET_COLUMNS, FORM31_SHEET_NAME } from "./form";

/**
 * يبني قالب Excel لنموذج 31 في **ورقة واحدة** تضم كل البيانات المطلوبة في النموذج:
 * بيانات المكتب ومقدم الطلب والمنشأة (تُكتب مرة واحدة في أول صف وتُورَّث تلقائيًا
 * لكل صف تالٍ يُترك فارغًا)، ثم بيانات كل حالة على حدة.
 */

/** قوائم منسدلة تقلّل الأخطاء في الحقول المغلقة. */
const LIST_RULES: Array<{ key: string; values: string[] }> = [
  { key: "companion", values: ["نعم", "لا"] },
  { key: "travelExpenses", values: ["نعم", "لا"] },
  { key: "treatmentPlace", values: ["داخل البلاد", "خارج البلاد"] },
];

function escapeXml(value: string) {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/** يثبّت صف العناوين ويضيف القوائم المنسدلة إلى ورقة القالب. */
async function decorateSheet(bytes: ArrayBuffer) {
  const zip = await JSZip.loadAsync(bytes);
  const sheetPath = zip.file("xl/worksheets/sheet1.xml") ? "xl/worksheets/sheet1.xml" : undefined;
  if (!sheetPath) return bytes;
  const entry = zip.file(sheetPath);
  if (!entry) return bytes;
  let xml = await entry.async("string");

  xml = xml.replace(/<sheetView([^>]*)>/, (opening) =>
    xml.includes("<pane ") ? opening : `${opening}<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>`);

  if (!xml.includes("<dataValidations")) {
    const validations = LIST_RULES.flatMap((rule) => {
      const index = FORM31_SHEET_COLUMNS.findIndex((column) => column.key === rule.key);
      if (index < 0) return [];
      const column = XLSX.utils.encode_col(index);
      const list = rule.values.map(escapeXml).join(",");
      return [
        `<dataValidation type="list" allowBlank="1" showErrorMessage="1" errorTitle="قيمة غير صحيحة" error="اختر قيمة من القائمة" sqref="${column}2:${column}1000"><formula1>"${list}"</formula1></dataValidation>`,
      ];
    });
    if (validations.length) {
      xml = xml.replace("</worksheet>", `<dataValidations count="${validations.length}">${validations.join("")}</dataValidations></worksheet>`);
    }
  }
  zip.file(sheetPath, xml);
  return zip.generateAsync({ type: "arraybuffer", compression: "DEFLATE" });
}

/** يبني ملف Excel جاهزًا للتنزيل: ورقة واحدة بعنوان «بيان كل الداتا المطلوبة». */
export async function buildForm31Workbook() {
  const headerRow = FORM31_SHEET_COLUMNS.map((column) => column.label);
  const sheet = XLSX.utils.aoa_to_sheet([headerRow]);
  sheet["!cols"] = FORM31_SHEET_COLUMNS.map((column) => ({
    wch: Math.min(34, Math.max(16, Math.max(column.label.length, column.hint?.length ?? 0) + 4)),
  }));
  sheet["!rows"] = [{ hpt: 26 }];

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, sheet, FORM31_SHEET_NAME);
  const bytes = XLSX.write(workbook, { bookType: "xlsx", type: "array" }) as ArrayBuffer;
  return decorateSheet(bytes);
}
