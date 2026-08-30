import { describe, expect, it } from "vitest";
import { makeEmptyRecord, mapExcelRow, normalizeDigits, recordStatus, validateRecord } from "./form";

function validRecord() {
  return {
    ...makeEmptyRecord(),
    office: "مكتب القاهرة", applicantName: "أحمد محمد", applicantRole: "صاحب العمل",
    applicantPhone: "01012345678", applicantNationalId: "٢٩٠٠١٠١٠١٢٣٤٥٦",
    insuredName: "محمد أحمد علي", insuranceNumber: "001234567", nationalId: "٢٩٥٠١٠١٠١٢٣٤٥٦",
    qualification: "بكالوريوس", profession: "محاسب", category: "عاملين لدى الغير",
    contributionCode: "12", startDate: "2024-05-19", basicWage: "2500.50", totalWage: "4000",
    establishmentName: "شركة الاختبار", establishmentNumber: "001234567", country: "مصري",
    governorate: "القاهرة", district: "الزمالك", street: "النيل", center: "قصر النيل",
    phone: "01011111111", workType: "دائمة", endDate: "2025-06-20", endReason: "انتهاء الخدمة", address: "القاهرة",
  };
}

describe("Arabic form validation", () => {
  it("normalizes Arabic and Persian digits", () => expect(normalizeDigits("١٢٣۴۵")).toBe("12345"));
  it("accepts a complete S1 record", () => {
    const record = validRecord();
    expect(validateRecord(record, "s1")).toEqual([]);
    expect(recordStatus(record, "s1")).toBe("جاهز");
  });
  it("rejects invalid national IDs and reversed dates", () => {
    const record = { ...validRecord(), nationalId: "123", endDate: "2020-01-01" };
    const messages = validateRecord(record, "s6").map((issue) => issue.message).join(" ");
    expect(messages).toContain("14 رقمًا");
    expect(messages).toContain("يسبق تاريخ البدء");
  });
  it("normalizes imported dates and Arabic digits", () => {
    const record = mapExcelRow({ "الرقم القومي": "٢٩٥٠١٠١٠١٢٣٤٥٦", "تاريخ بدء الاشتراك": new Date(2024, 4, 19) });
    expect(record.nationalId).toBe("29501010123456");
    expect(record.startDate).toBe("2024-05-19");
  });
});

