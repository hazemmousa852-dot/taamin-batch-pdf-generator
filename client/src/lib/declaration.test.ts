import { makeEmptyRecord } from "./form";
import { describe, it, expect, vi } from "vitest";
import * as XLSX from "xlsx";
import { readFileSync, writeFileSync } from "node:fs";
import { PDFDocument } from "pdf-lib";
import { addDeclarationSheet, readDeclarations, createDeclarationPdf, DECLARATION_SHEET, DECLARATION_HEADERS, formatTaxRegistration, validateDeclarationEstablishments } from "./declaration";

describe("employer declarations", () => {
  const data = Object.fromEntries(DECLARATION_HEADERS.map(h => [h, ""]));
  Object.assign(data, {"رقم المنشأة":"00000001", "اسم المنشأة":"شركة الاختبار", "المكتب":"مكتب تجريبي", "مقدم الطلب":"مقدم طلب تجريبي", "صفة مقدم الطلب":"مفوض", "اسم صاحب العمل أو المدير المسؤول ثلاثي":"صاحب عمل تجريبي", "صفة صاحب العمل أو المدير المسؤول":"مدير مسؤول", "الرقم القومي لمقدم الطلب":"00000000000000", "الرقم التأميني لمقدم الطلب":"000000000", "تاريخ تحرير الإقرار (DD/MM/YYYY)":"18/10/2026"});
  it("round trips identifiers and deduplicates each establishment", () => {
    const wb = XLSX.utils.book_new(); addDeclarationSheet(wb);
    wb.Sheets[DECLARATION_SHEET] = XLSX.utils.json_to_sheet([data,data]);
    const read = XLSX.read(XLSX.write(wb,{type:"array",bookType:"xlsx"}),{type:"array"});
    expect(readDeclarations(read)).toEqual([data]);
    wb.Sheets[DECLARATION_SHEET] = XLSX.utils.json_to_sheet([data,{...data,"رقم المنشأة":"00000002"}]);
    expect(readDeclarations(wb)).toHaveLength(2);
    wb.Sheets[DECLARATION_SHEET] = XLSX.utils.json_to_sheet([data,{...data,"اسم المنشأة":"مختلف"}]);
    expect(() => readDeclarations(wb)).toThrow("متعارضة");
    expect(readDeclarations(XLSX.utils.book_new())).toEqual([]);
  });
  it("accepts formatted tax numbers, preserves zeros and normalizes Excel input", () => {
    expect(formatTaxRegistration("123456789")).toBe("123-456-789");
    expect(formatTaxRegistration("987-654-321")).toBe("987-654-321");
    expect(formatTaxRegistration("000-000-000")).toBe("000-000-000");
    expect(formatTaxRegistration("٠٠٠-٠٠٠-٠٠١")).toBe("000-000-001");
    expect(formatTaxRegistration("000000001")).toBe("000-000-001");
    expect(() => formatTaxRegistration("000-00-000")).toThrow("٩ أرقام");
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet([{...data, "رقم التسجيل الضريبي":"000-000-001"}, {...data, "رقم التسجيل الضريبي":"000000001"}]), DECLARATION_SHEET);
    const roundtrip = XLSX.read(XLSX.write(wb, {type:"array",bookType:"xlsx"}), {type:"array"});
    expect(readDeclarations(roundtrip)).toHaveLength(1);
    expect(readDeclarations(roundtrip)[0]["رقم التسجيل الضريبي"]).toBe("000-000-001");
  });
  it("checks the establishment number and name against S2", () => {
    const record = {...makeEmptyRecord(), establishmentNumber: "00000001", establishmentName: "شركة الاختبار"};
    expect(() => validateDeclarationEstablishments([data], [record])).not.toThrow();
    expect(() => validateDeclarationEstablishments([data], [{...record, establishmentNumber: "00000002"}])).toThrow("غير موجود");
    expect(() => validateDeclarationEstablishments([data], [{...record, establishmentName: "اسم مختلف"}])).toThrow("مختلف");
  });
  it("preserves the two-page source and fills Arabic and identifiers", async () => {
    vi.stubGlobal("fetch", async (url: string) => { const buffer = readFileSync(`client/public/assets/${url.split('/').pop()}`); return {ok:true, arrayBuffer:async()=>buffer.buffer.slice(buffer.byteOffset,buffer.byteOffset+buffer.byteLength)}; });
    try {
      const bytes = await createDeclarationPdf({...data, "رقم التسجيل الضريبي":"000-000-001"});
      expect((await PDFDocument.load(bytes)).getPageCount()).toBe(2);
      writeFileSync('/tmp/declaration-filled.pdf',bytes);
      await expect(createDeclarationPdf({...data,"رقم المنشأة":"000000000"})).rejects.toThrow("8");
    } finally { vi.unstubAllGlobals(); }
  });
});
