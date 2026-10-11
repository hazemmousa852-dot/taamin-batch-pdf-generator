/**
 * اختبار بصري لنموذج 31: يشغّل نفس مسار التصدير الموجود في التطبيق
 * (fillPdf → overlay canvas → pdf-lib) داخل Node مع محاكي canvas، ثم يحفظ
 * ملف PDF حقيقي لمراجعته ورؤيته. لا يُستخدم في الإنتاج، بل للتحقق اليدوي.
 *
 * الاستخدام:
 *   npm i canvas --no-save
 *   npx tsx scripts/form31-render-node.mjs [output.pdf]
 */
import { registerFont, createCanvas, Canvas } from "canvas";
import { readFile, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

const root = new URL("../", import.meta.url);
const assetPath = (name) => fileURLToPath(new URL(`client/public/assets/${name}`, root));

// --- محاكي بيئة المتصفح الذي يستخدمه كود التصدير ---
Canvas.prototype.toBlob = function toBlob(callback, type = "image/png") {
  callback(new Blob([this.toBuffer(type)], { type }));
};
class FontFaceShim {
  constructor(family, source) {
    this.family = family;
    this.buffer = Buffer.from(source);
  }
  async load() {
    // node-canvas يتطلب مسار ملف لتسجيل الخط، بخلاف FontFace في المتصفح
    const path = join(tmpdir(), `taamin-form31-${this.family}-${this.buffer.length}.ttf`);
    await writeFile(path, this.buffer);
    registerFont(path, { family: this.family });
    return this;
  }
}
globalThis.FontFace = FontFaceShim;
globalThis.document = {
  fonts: { add() {} },
  createElement(tag) {
    if (tag !== "canvas") throw new Error(`عنصر غير مدعوم في المحاكي: ${tag}`);
    return createCanvas(1, 1);
  },
};
const assets = new Map();
globalThis.fetch = async (url) => {
  const name = String(url).split("/").pop();
  if (!assets.has(name)) assets.set(name, await readFile(assetPath(name)));
  // كود pdf-lib يقرأ البايتات عبر arrayBuffer()
  return { ok: true, arrayBuffer: async () => assets.get(name).buffer.slice(assets.get(name).byteOffset, assets.get(name).byteOffset + assets.get(name).byteLength) };
};

const { fillPdf } = await import("../client/src/lib/pdf.ts");

const record = {
  id: "sample-1",
  office: "مكتب تأمينات المنصورة",
  applicantName: "حازم موسى إبراهيم",
  applicantRole: "مسؤول شؤون العاملين",
  applicantPhone: "01012345678",
  applicantInsuranceNumber: "001234567",
  applicantNationalId: "29001010123456",
  insuredName: "محمد أحمد عبد الرحمن",
  insuranceNumber: "001234567",
  nationalId: "29501010123456",
  establishmentName: "شركة النيل للصناعات الغذائية",
  establishmentNumber: "12345678",
  injuryDate: "2026-09-14",
  companion: "نعم",
  sickLeaveStartDate: "2026-09-15",
  sickLeaveEndDate: "2026-09-29",
  travelExpenses: "لا",
  treatmentPlace: "داخل البلاد",
  monthlyWage: "4500",
  wagePaymentMethod: "بالحساب البنكي",
  otherLine1: "تم إرفاق التقرير الطبي المعتمد",
  otherLine2: "برجاء الصرف لمقدم الطلب",
  signatureName: "",
  releaseDate: "2026-10-01",
};

const output = process.argv[2] ?? "form31-sample.pdf";
const bytes = await fillPdf(record, "s31");
await writeFile(output, bytes);
console.log(`تم إنشاء ${output} بحجم ${(bytes.length / 1024).toFixed(1)} كيلوبايت`);
