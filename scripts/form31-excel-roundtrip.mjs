/**
 * اختبار دورة عمل نموذج 31 كاملة بدون متصفح:
 *   1) يبني قالب Excel بالورقة الواحدة (نفس كود التطبيق).
 *   2) يضيف صفين تجريبيين: البيانات المشتركة في الأول فقط.
 *   3) يعيد قراءة الملف بنفس منطق الاستيراد في الواجهة (مع توريث البيانات المشتركة).
 *   4) يتحقق من تطابق القيم مع الأعمدة المتوقعة.
 *
 * الاستخدام: npx tsx scripts/form31-excel-roundtrip.mjs
 */
import { readFile, writeFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import * as XLSX from "xlsx";

const root = new URL("../", import.meta.url);
const asset = (name) => fileURLToPath(new URL(`client/public/assets/${name}`, root));

const { buildForm31Workbook } = await import("../client/src/lib/form31xlsx.ts");
const { FORM31_SHEET_COLUMNS, FORM31_SHEET_NAME, carryForwardShared, mapExcelRow } = await import("../client/src/lib/form.ts");

const b64 = (await readFile(asset("taamin-flexible-template.b64"), "utf8")).trim();
const templateBytes = Uint8Array.from(Buffer.from(b64, "base64"));
console.log("حجم قالب Excel:", templateBytes.length, "بايت");
const workbookBytes = await buildForm31Workbook(templateBytes);

// --- املأ الملف كما سيفعل المستخدم ---
const workbook = XLSX.read(workbookBytes, { type: "array", cellDates: true });
// أعمدة الورقة الواحدة: المكتب، مقدم الطلب، صفته، رقمه التأميني، تليفونه، رقمه القومي،
// اسم المنشأة، رقم المنشأة، ثم اسم المؤمن عليه، الرقم القومي، الرقم التأميني،
// تاريخ الإصابة، مرافق، بداية الإجازة، نهاية الإجازة، مصاريف انتقال، مكان العلاج،
// أجر الاشتراك، طريقة الأداء، بيانات أخرى 1-3، تحريرًا في، اسم التوقيع.
const sheet = workbook.Sheets[FORM31_SHEET_NAME];
XLSX.utils.sheet_add_aoa(sheet, [
  ["مكتب المنصورة", "حازم موسى", "مسؤول شؤون العاملين", "001234567", "01012345678", "29001010123456",
   "شركة النيل للأغذية", "12345678",
   "محمد أحمد عبد الرحمن", "29501010123456", "001234567", "14/09/2026", "نعم", "15/09/2026", "29/09/2026",
   "لا", "داخل البلاد", "4500", "بالحساب البنكي", "ملاحظة 1", "", "", "01/10/2026", ""],
  // الصف الثاني يترك البيانات المشتركة فارغة لتُورَّث من الصف الأول
  ["", "", "", "", "", "", "", "",
   "سعاد إبراهيم حسن", "29001010123456", "009876543", "", "لا", "01/10/2026", "10/10/2026",
   "نعم", "خارج البلاد", "3200.50", "نقدًا", "", "ملاحظة 2", "", "05/10/2026", "سعاد إبراهيم"],
], { origin: "A2" });
const filled = XLSX.write(workbook, { bookType: "xlsx", type: "buffer" });
await writeFile("form31-template-filled.xlsx", filled);

// --- أعد القراءة بنفس منطق الواجهة ---
const reloaded = XLSX.read(filled, { type: "buffer", cellDates: true });
console.log("أوراق الملف:", reloaded.SheetNames.join(" | "));
const rows = XLSX.utils.sheet_to_json(reloaded.Sheets[FORM31_SHEET_NAME], { defval: "" });
let previous;
const records = rows.map((row) => {
  const record = mapExcelRow(row);
  carryForwardShared(record, previous);
  previous = record;
  return record;
});

const assertions = [
  ["عدد أوراق الملف", 1, reloaded.SheetNames.length],
  ["اسم الورقة", FORM31_SHEET_NAME, reloaded.SheetNames[0]],
  ["عدد الأعمدة في الورقة", FORM31_SHEET_COLUMNS.length, XLSX.utils.sheet_to_json(reloaded.Sheets[FORM31_SHEET_NAME], { header: 1 })[0].length],
  ["عدد السجلات", 2, records.length],
  ["اسم أول مؤمن عليه", "محمد أحمد عبد الرحمن", records[0].insuredName],
  ["تاريخ الإصابة", "2026-09-14", records[0].injuryDate],
  ["بداية الإجازة", "2026-09-15", records[0].sickLeaveStartDate],
  ["نهاية الإجازة", "2026-09-29", records[0].sickLeaveEndDate],
  ["أجر الاشتراك الشهري", "4500", records[0].monthlyWage],
  ["اسم المنشأة من نفس الورقة", "شركة النيل للأغذية", records[0].establishmentName],
  ["المكتب من نفس الورقة", "مكتب المنصورة", records[0].office],
  ["اسم المنشأة موروث للصف الثاني", "شركة النيل للأغذية", records[1].establishmentName],
  ["المكتب موروث للصف الثاني", "مكتب المنصورة", records[1].office],
  ["مقدم الطلب موروث للصف الثاني", "حازم موسى", records[1].applicantName],
  ["رقم المنشأة موروث للصف الثاني", "12345678", records[1].establishmentNumber],
  ["رقم المنشأة", "12345678", records[0].establishmentNumber],
  ["اسم التوقيع المخصص", "سعاد إبراهيم", records[1].signatureName],
  ["السجل الثاني بلا تاريخ إصابة", "", records[1].injuryDate],
];

let failed = false;
for (const [label, expected, actual] of assertions) {
  const ok = String(expected) === String(actual);
  if (!ok) failed = true;
  console.log(`${ok ? "✓" : "✗"} ${label}: ${actual}${ok ? "" : ` (المتوقع ${expected})`}`);
}
console.log(failed ? "\nهناك قيم غير مطابقة!" : "\nكل القيم مطابقة — دورة Excel سليمة.");
process.exit(failed ? 1 : 0);
