import { readFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { PDFDocument } from "pdf-lib";
import { beforeAll, describe, expect, it, vi } from "vitest";
import { makeEmptyRecord } from "./form";
import { createMergedPdf, fillPdf, safeFileName, toPdfText } from "./pdf";

beforeAll(() => {
  vi.stubGlobal("fetch", async (url: string | URL) => {
    const name = String(url).split("/").pop()!;
    const bytes = await readFile(fileURLToPath(new URL(`../../public/assets/${name}`, import.meta.url)));
    return new Response(bytes, { status: 200 });
  });
});

const record = {
  ...makeEmptyRecord(), office: "مكتب القاهرة", applicantName: "أحمد محمد", applicantRole: "صاحب العمل",
  applicantInsuranceNumber: "009876543210", applicantPhone: "01012345678", applicantNationalId: "29001010123456", insuredName: "محمد أحمد علي",
  insuranceNumber: "001234567", nationalId: "29501010123456", qualification: "بكالوريوس تجارة",
  profession: "محاسب", sector: "خاص", country: "مصري", category: "عاملين لدى الغير", medicalExam: "نعم", contributionCode: "12",
  workType: "دائمة", startDate: "2024-05-19", basicWage: "2500", totalWage: "4000",
  increaseDate: "2025-06-20", increasePercent: "10", establishmentName: "شركة الاختبار",
  establishmentNumber: "001234567", establishmentType: "نمطي", governorate: "القاهرة", buildingNumber: "12", district: "الزمالك", street: "النيل",
  center: "قصر النيل", phone: "01011111111", address: "١ شارع النيل - القاهرة", endDate: "2025-07-21", endReason: "انتهاء الخدمة",
};

describe("official PDF field maps", () => {
  it("exports partial and empty forms without required-data errors", async () => {
    for (const template of ["s1", "s2", "s6"] as const) {
      const pdf = await PDFDocument.load(await fillPdf(makeEmptyRecord(), template));
      expect(pdf.getPageCount()).toBe(template === "s2" ? 1 : 2);
    }
  });
  it("fills the approved additional S1 and S6 destinations", async () => {
    const extra = { ...makeEmptyRecord(), releaseDate: "2026-09-21", signatureMatchDate: "2026-09-22",
      receiptNumber: "12345", receiptDate: "2026-09-23", insuredSignDate: "2026-09-24",
      managerSignDate: "2026-09-25", signatureVerifier: "اسم المطابق", disputeManagerSignDate: "2026-09-26",
      disputeInsuredSignDate: "2026-09-27", applicantSignDate: "2026-09-28", category: "العاملين بالمخابز" };
    const s1Bytes = await fillPdf(extra, "s1");
    const s1 = (await PDFDocument.load(s1Bytes)).getForm();
    expect(s1.getTextField("optional_releaseDate").getText()).toBe("21/09/2026");
    expect(s1.getTextField("optional_signatureMatchDate").getText()).toBe("22/09/2026");
    expect(s1.getCheckBox("زباخلماب ينلماعلا").isChecked()).toBe(true);
    const s6Bytes = await fillPdf(extra, "s6");
    const s6 = (await PDFDocument.load(s6Bytes)).getForm();
    for (const key of ["receiptNumber", "receiptDate", "insuredSignDate", "managerSignDate", "signatureVerifier", "disputeManagerSignDate", "disputeInsuredSignDate", "applicantSignDate"] as const) {
      const expected = key.endsWith("Date") ? extra[key].split("-").reverse().join("/") : extra[key];
      expect(s6.getTextField(`optional_${key}`).getText()).toBe(expected);
    }
  });
  it("orders complete Arabic phrases correctly for PDF appearance streams", () => {
    expect(toPdfText("اسم المؤمن عليه")).toBe("ﻪﻴﻠﻋ ﻦﻣﺆﻤﻟﺍ ﻢﺳﺍ");
    expect(toPdfText("محمد 2026")).toContain("2026");
  });

  it("places S1 Arabic text and dates in their visual fields", async () => {
    const doc = await PDFDocument.load(await fillPdf(record, "s1"));
    const form = doc.getForm();
    expect(form.getTextField(": ةفصبلطلا مدقم_1").getText()).toBe(record.applicantName);
    expect(form.getTextField("Text Field6").getText()).toBe(record.applicantRole);
    expect(form.getTextField("Text Field3").getText()).toBe(record.applicantInsuranceNumber);
    expect(form.getTextField("ةنهلما_1").getText()).toBe(record.qualification);
    expect(form.getTextField("ةنهلما").getText()).toBe(record.profession);
    expect(form.getTextField("Text Field2").getText()).toBe("19");
    expect(form.getTextField("Text Field1").getText()).toBe("05");
    expect(form.getTextField("Text Field0").getText()).toBe("2024");
    expect(form.getTextField("Text Field4").getText()).toBe("2500");
    expect(form.getTextField("Text Field5").getText()).toBe("4000");
    expect(form.getTextField("fill_6").getText()).toBe(record.workType);
    expect(form.getTextField("fill_7").getText()).toBe("12");
    expect(form.getTextField(": عاطقلا").getText()).toBe(record.sector);
    expect(form.getTextField("fill_5").getText()).toBe(record.buildingNumber);
    expect(form.getCheckBox("ئادتبلاا بيطلا فشكلا ءافيتساي").isChecked()).toBe(true);
    expect(form.getCheckBox("أشنلما عونةطنم :ي").isChecked()).toBe(true);
    expect(form.getTextField("fill_4").isReadOnly()).toBe(true);
  });

  it("places S6 repeated values and date in the correct fields", async () => {
    const doc = await PDFDocument.load(await fillPdf(record, "s6"));
    const form = doc.getForm();
    expect(form.getTextField("Text Field0").getText()).toBe(record.applicantInsuranceNumber);
    expect(form.getTextField("Text Field2").getText()).toBe("21");
    expect(form.getTextField("Text Field3").getText()).toBe("07");
    expect(form.getTextField("Text Field4").getText()).toBe("2025");
    expect(form.getTextField("Text Field6").getText()).toBe("1 شارع النيل - القاهرة");
    expect(form.getTextField(":نييمأتلا اهمقر").getText()).toBe(record.establishmentNumber);
    expect(form.getTextField("ا: هيلع نمؤلما مــــــــس---------------------------------------:نييمأتلا همقر").getText()).toBe(record.insuranceNumber);
  });

  it("always creates unique, filesystem-safe names", () => {
    expect(safeFileName(record, 1, "s1")).not.toBe(safeFileName(record, 2, "s1"));
    expect(safeFileName(record, 1, "s1")).toMatch(/^س1-/);
  });

  it("merges every generated form in record order", async () => {
    const merged = await PDFDocument.load(await createMergedPdf([record, { ...record, insuredName: "شخص ثان" }], "s6"));
    expect(merged.getPageCount()).toBe(4);
  });

  it("paginates S2 at eleven employees per page", async () => {
    const records = Array.from({ length: 12 }, (_, index) => ({ ...record, id: `s2-${index}`, insuranceNumber: String(12 - index) }));
    const merged = await PDFDocument.load(await createMergedPdf(records, "s2"));
    expect(merged.getPageCount()).toBe(2);
  });
});


