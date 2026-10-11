import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { buildForm31Workbook } from "./form31xlsx";

/**
 * اختبارات سلامة بنية قالب Excel — الهدف منع تكرار العطل الذي كان يجعل
 * Excel يرفض ورقة القالب ويفتحها فارغة (تثبيت الصفوف يجب أن يكون داخل
 * <sheetView>، والقوائم المنسدلة في ترتيبها الصحيح من مخطط OOXML).
 */

async function buildSheetXml() {
  const bytes = await buildForm31Workbook();
  const zip = await JSZip.loadAsync(bytes);
  const file = zip.file("xl/worksheets/sheet1.xml");
  if (!file) throw new Error("sheet1.xml مفقود من القالب");
  return file.async("string");
}

describe("قالب Excel لنموذج 31 — سلامة البنية", () => {
  it(" يحتوي ورقة واحدة فقط", async () => {
    const bytes = await buildForm31Workbook();
    const zip = await JSZip.loadAsync(bytes);
    const names = Object.keys(zip.files);
    expect(names).toContain("xl/worksheets/sheet1.xml");
    expect(names).not.toContain("xl/worksheets/sheet2.xml");
  });

  it("تثبيت صف العناوين داخل sheetView (وليس خارجه)", async () => {
    const xml = await buildSheetXml();
    expect(xml).toMatch(/<sheetView[^>]*><pane [^>]*state="frozen"[^>]*\/><\/sheetView>/);
  });

  it("القوائم المنسدلة الثلاثة بعد sheetData وفي موضعها الصحيح", async () => {
    const xml = await buildSheetXml();
    const dvIndex = xml.indexOf("<dataValidations");
    expect(dvIndex).toBeGreaterThan(xml.indexOf("</sheetData>"));
    const ignoredIndex = xml.indexOf("<ignoredErrors");
    if (ignoredIndex >= 0) expect(dvIndex).toBeLessThan(ignoredIndex);
    expect(xml.match(/<dataValidation /g)?.length).toBe(3);
    expect(xml).toContain('sqref="M2:M1000"');
    expect(xml).toContain('sqref="P2:P1000"');
    expect(xml).toContain('sqref="Q2:Q1000"');
  });
});
