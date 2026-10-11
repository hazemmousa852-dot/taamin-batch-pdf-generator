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
  it("allows blank and partial forms for all three templates", () => {
    for (const template of ["s1", "s2", "s6"] as const) {
      expect(validateRecord(makeEmptyRecord(), template)).toEqual([]);
      const partial = { ...makeEmptyRecord(), insuredName: "اسم فقط" };
      expect(validateRecord(partial, template)).toEqual([]);
      expect(recordStatus(partial, template)).toBe("جاهز");
    }
  });
  it("imports the approved optional headings without imposing list membership", () => {
    const value = mapExcelRow({
      "تاريخ الالتحاق (س2) (DD/MM/YYYY)": "21/09/2026",
      "رقم وصل الخطاب المسجل بعلم الوصول (س6)": "وصل 123",
      "اسم القائم بمطابقة التوقيع (س6)": "اسم المطابق",
      "الفئة": "فئة مخصصة",
    });
    expect(value.hireDate).toBe("2026-09-21");
    expect(value.receiptNumber).toBe("وصل 123");
    expect(value.signatureVerifier).toBe("اسم المطابق");
    expect(value.category).toBe("فئة مخصصة");
    expect(validateRecord(value, "s6")).toEqual([]);
  });
  it("normalizes Arabic and Persian digits", () => expect(normalizeDigits("١٢٣۴۵")).toBe("12345"));
  it("imports the نموذج 31 sheet columns with their exact Excel headings", () => {
    const record = mapExcelRow({
      "اسم المؤمن عليه": "محمد أحمد",
      "الرقم القومي": "29501010123456",
      "الرقم التأميني": "001234567",
      "تاريخ الإصابة (DD/MM/YYYY)": "14/09/2026",
      "هل يوجد مرافق": "نعم",
      "تاريخ بداية الإجازة المرضي (DD/MM/YYYY)": "15/09/2026",
      "تاريخ نهاية الإجازة المرضي (DD/MM/YYYY)": "29/09/2026",
      "هل توجد مصاريف انتقال": "لا",
      "العلاج يتم داخل أم خارج البلاد": "داخل البلاد",
      "أجر الاشتراك الشهري للمصاب": "4500",
      "طريقة أداء الأجر": "بالحساب البنكي",
      "بيانات أخرى 1": "ملاحظة أولى",
      "تحريرًا في (DD/MM/YYYY)": "01/10/2026",
    });
    expect(record.injuryDate).toBe("2026-09-14");
    expect(record.sickLeaveStartDate).toBe("2026-09-15");
    expect(record.sickLeaveEndDate).toBe("2026-09-29");
    expect(record.releaseDate).toBe("2026-10-01");
    expect(record.companion).toBe("نعم");
    expect(record.travelExpenses).toBe("لا");
    expect(record.treatmentPlace).toBe("داخل البلاد");
    expect(record.monthlyWage).toBe("4500");
    expect(record.otherLine1).toBe("ملاحظة أولى");
    expect(validateRecord(record, "s31")).toEqual([]);
  });

  it("accepts blank and partial نموذج 31 records", () => {
    expect(validateRecord(makeEmptyRecord(), "s31")).toEqual([]);
    const partial = { ...makeEmptyRecord(), insuredName: "اسم فقط", sickLeaveStartDate: "not-a-date" };
    expect(validateRecord(partial, "s31")).toHaveLength(1);
  });

  it("accepts a complete S1 record", () => {
    const record = validRecord();
    expect(validateRecord(record, "s1")).toEqual([]);
    expect(recordStatus(record, "s1")).toBe("جاهز");
  });
  it("allows partial identifiers and date chronology", () => {
    const record = { ...validRecord(), nationalId: "123", endDate: "2020-01-01" };
    expect(validateRecord(record, "s6")).toEqual([]);
  });
  it("normalizes imported dates and Arabic digits", () => {
    const record = mapExcelRow({ "الرقم القومي": "٢٩٥٠١٠١٠١٢٣٤٥٦", "تاريخ بدء الاشتراك": new Date(2024, 4, 19) });
    expect(record.nationalId).toBe("29501010123456");
    expect(record.startDate).toBe("2024-05-19");
  });
  it("reads day/month/year text dates from Excel", () => {
    const record = mapExcelRow({ "تاريخ بداية الاشتراك": "31/08/2026", "تاريخ بداية العجز": "٠٥/٠٩/٢٠٢٦" });
    expect(record.startDate).toBe("2026-08-31");
    expect(record.increaseDate).toBe("2026-09-05");
  });
  it("keeps insurance numbers exactly as entered in Excel", () => {
    const record = mapExcelRow({ "الرقم التأميني": 123456789, "الرقم التأميني لمقدم الطلب": 987654321 });
    expect(record.insuranceNumber).toBe("123456789");
    expect(record.applicantInsuranceNumber).toBe("987654321");
  });
  it("accepts decimal wages for S2", () => {
    const record = { ...validRecord(), basicWage: "2500.50" };
    expect(validateRecord(record, "s2")).toEqual([]);
    expect(validateRecord({ ...record, basicWage: "2500" }, "s2")).toEqual([]);
  });
});
