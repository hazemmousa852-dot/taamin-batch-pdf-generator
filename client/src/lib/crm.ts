import JSZip from "jszip";
import { normalizeDigits, type PersonRecord } from "@/lib/form";

const XLSX_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

function escapeXml(value: string) {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function replaceCell(xml: string, reference: string, value: string, numeric = false) {
  const pattern = new RegExp(
    `<c\\b([^>]*\\br="${reference}"[^>]*)\\/>|<c\\b([^>]*\\br="${reference}"[^>]*)>[\\s\\S]*?<\\/c>`,
  );
  const match = xml.match(pattern);
  if (!match) throw new Error(`تعذر العثور على الخلية ${reference} داخل قالب CRM`);

  const attributes = (match[1] ?? match[2]).replace(/\s+t="[^"]*"/g, "");
  const content = numeric
    ? `<v>${escapeXml(normalizeDigits(value).replace(/[^0-9.-]/g, ""))}</v>`
    : `<is><t xml:space="preserve">${escapeXml(value)}</t></is>`;
  return xml.replace(pattern, `<c${attributes}${numeric ? "" : ' t="inlineStr"'}>${content}</c>`);
}

function insuranceSortValue(value: string) {
  return normalizeDigits(value).replace(/\D/g, "").replace(/^0+(?=\d)/, "");
}

export function isS2CrmComplete(record: PersonRecord) {
  return [record.insuranceNumber, record.insuredName, record.basicWage, record.totalWage]
    .every((value) => value.trim().length > 0);
}

export function sortS2CrmRecords(records: PersonRecord[]) {
  return [...records].sort((left, right) => {
    const a = insuranceSortValue(left.insuranceNumber);
    const b = insuranceSortValue(right.insuranceNumber);
    return a.length - b.length || a.localeCompare(b, "en");
  });
}

async function loadOriginalTemplate(partsBaseUrl: string) {
  const responses = await Promise.all(
    Array.from({ length: 14 }, (_, index) =>
      fetch(`${partsBaseUrl}/part-${String(index).padStart(2, "0")}.bin`),
    ),
  );
  if (responses.some((response) => !response.ok)) {
    throw new Error("تعذر تحميل قالب س2 الأصلي كاملًا");
  }

  const parts = await Promise.all(responses.map((response) => response.arrayBuffer()));
  const totalLength = parts.reduce((total, part) => total + part.byteLength, 0);
  const template = new Uint8Array(totalLength);
  let offset = 0;
  for (const part of parts) {
    template.set(new Uint8Array(part), offset);
    offset += part.byteLength;
  }
  return template;
}

export async function createS2CrmWorkbook(records: PersonRecord[], partsBaseUrl: string) {
  if (records.length > 5_000) throw new Error("الحد الأقصى لملف CRM الواحد هو 5000 موظف");
  const zip = await JSZip.loadAsync(await loadOriginalTemplate(partsBaseUrl));
  const sheetPath = "xl/worksheets/sheet1.xml";
  const sheetFile = zip.file(sheetPath);
  if (!sheetFile) throw new Error("ورقة CRM الأساسية غير موجودة في القالب");

  let xml = await sheetFile.async("string");
  const company = records[0];
  xml = replaceCell(xml, "B6", company.establishmentName.trim());
  xml = replaceCell(xml, "H6", normalizeDigits(company.establishmentNumber).trim());
  xml = replaceCell(xml, "B7", "");
  xml = replaceCell(xml, "H7", company.office.trim());

  records.forEach((record, index) => {
    const row = 15 + index;
    xml = replaceCell(xml, `A${row}`, normalizeDigits(record.insuranceNumber).trim());
    xml = replaceCell(xml, `B${row}`, record.insuredName.trim());
    xml = replaceCell(xml, `G${row}`, record.basicWage.trim(), true);
    xml = replaceCell(xml, `H${row}`, record.totalWage.trim(), true);
  });

  zip.file(sheetPath, xml);
  return zip.generateAsync({ type: "blob", mimeType: XLSX_TYPE, compression: "DEFLATE" });
}

