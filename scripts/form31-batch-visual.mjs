/** يبني نموذجين PDF من ملف Excel المعبأ وينتج صورًا للمراجعة البصرية. */
import { readFile, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import * as XLSX from "xlsx";
import { registerFont, createCanvas, Canvas } from "canvas";

Canvas.prototype.toBlob = function toBlob(cb, type = "image/png") { cb(new Blob([this.toBuffer(type)], { type })); };
class FontFaceShim {
  constructor(family, source) { this.family = family; this.buffer = Buffer.from(source); }
  async load() {
    const path = join(tmpdir(), `taamin-${this.family}-${this.buffer.length}.ttf`);
    await writeFile(path, this.buffer);
    registerFont(path, { family: this.family });
    return this;
  }
}
globalThis.FontFace = FontFaceShim;
globalThis.document = { fonts: { add() {} }, createElement(tag) { return createCanvas(1, 1); } };
const cache = new Map();
globalThis.fetch = async (url) => {
  const name = String(url).split("/").pop();
  if (!cache.has(name)) cache.set(name, await readFile(new URL(`../client/public/assets/${name}`, import.meta.url)));
  const buf = cache.get(name);
  return { ok: true, arrayBuffer: async () => buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) };
};

const { fillPdf } = await import("../client/src/lib/pdf.ts");
const { carryForwardShared, mapExcelRow } = await import("../client/src/lib/form.ts");
const { createZip } = await import("../client/src/lib/pdf.ts");
const JSZip = (await import("jszip")).default;

const workbook = XLSX.read(await readFile("form31-template-filled.xlsx"), { type: "buffer", cellDates: true });
const sheetName = workbook.SheetNames[0];
const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "" });
// نفس منطق الواجهة: البيانات المشتركة تُورَّث من الصف السابق عند تركها فارغة
let previous;
const records = rows.map((row) => {
  const record = mapExcelRow(row);
  carryForwardShared(record, previous);
  previous = record;
  return record;
});

for (let index = 0; index < records.length; index += 1) {
  const bytes = await fillPdf(records[index], "s31");
  await writeFile(`/tmp/form31-batch-${index + 1}.pdf`, bytes);
  console.log(`نموذج ${index + 1}: ${(bytes.length / 1024).toFixed(1)} كيلوبايت — ${records[index].insuredName}`);
}
const zip = await createZip(records, "s31");
await writeFile("/tmp/form31-batch.zip", Buffer.from(await zip.arrayBuffer()));
const opened = await JSZip.loadAsync(await readFile("/tmp/form31-batch.zip"));
console.log("الملفات داخل ZIP:", Object.keys(opened.files).join(" | "));
