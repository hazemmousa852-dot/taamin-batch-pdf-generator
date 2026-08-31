// Design reminder: form vocabulary follows the “Registry Desk” philosophy—quiet, precise, and close to the printed document.

export type TemplateId = "s1" | "s6";
export const TEMPLATE_LABELS: Record<TemplateId, string> = { s1: "نموذج س1", s6: "نموذج س6" };
export type RecordStatus = "مسودة" | "جاهز" | "ناقص";
export type ValidationIssue = { key: keyof PersonRecord; message: string };

export type PersonRecord = {
  id: string;
  office: string;
  establishmentName: string;
  establishmentNumber: string;
  establishmentType: string;
  insuredName: string;
  insuranceNumber: string;
  nationalId: string;
  qualification: string;
  profession: string;
  professionCode: string;
  sector: string;
  startDate: string;
  birthDate: string;
  gender: string;
  category: string;
  medicalExam: string;
  contributionCode: string;
  workType: string;
  basicWage: string;
  variableWage: string;
  totalWage: string;
  increaseDate: string;
  increasePercent: string;
  country: string;
  city: string;
  governorate: string;
  buildingNumber: string;
  district: string;
  street: string;
  center: string;
  phone: string;
  email: string;
  employer: string;
  manager: string;
  releaseDate: string;
  applicantName: string;
  applicantRole: string;
  applicantInsuranceNumber: string;
  applicantPhone: string;
  applicantNationalId: string;
  endDate: string;
  endReason: string;
  address: string;
};

export const EXCEL_HEADERS: Array<{ key: keyof PersonRecord; label: string }> = [
  { key: "insuredName", label: "اسم المؤمن عليه" },
  { key: "nationalId", label: "الرقم القومي" },
  { key: "insuranceNumber", label: "الرقم التأميني" },
  { key: "startDate", label: "تاريخ بداية الاشتراك (YYYY-MM-DD)" },
  { key: "establishmentName", label: "اسم المنشأة" },
  { key: "establishmentNumber", label: "رقم المنشأة" },
  { key: "establishmentType", label: "نوع المنشأة" },
  { key: "office", label: "المكتب" },
  { key: "qualification", label: "المؤهل" },
  { key: "profession", label: "المهنة" },
  { key: "professionCode", label: "كود المهنة" },
  { key: "sector", label: "القطاع" },
  { key: "birthDate", label: "تاريخ الميلاد" },
  { key: "gender", label: "النوع" },
  { key: "category", label: "الفئة" },
  { key: "medicalExam", label: "استيفاء الكشف الطبي الابتدائي" },
  { key: "contributionCode", label: "كود الاشتراك" },
  { key: "workType", label: "نوع المدة" },
  { key: "basicWage", label: "أجر / دخل الاشتراك" },
  { key: "variableWage", label: "الأجر المتغير" },
  { key: "totalWage", label: "الأجر الشامل" },
  { key: "increaseDate", label: "تاريخ بداية العجز" },
  { key: "increasePercent", label: "نسبة العجز" },
  { key: "country", label: "الجنسية" },
  { key: "city", label: "المدينة" },
  { key: "governorate", label: "المحافظة" },
  { key: "buildingNumber", label: "رقم العقار" },
  { key: "district", label: "الشياخة / القرية" },
  { key: "street", label: "الشارع" },
  { key: "center", label: "القسم / المركز" },
  { key: "phone", label: "التليفون" },
  { key: "email", label: "البريد الإلكتروني" },
  { key: "employer", label: "جهة العمل" },
  { key: "manager", label: "المدير المسؤول" },
  { key: "releaseDate", label: "تحريرًا في" },
  { key: "applicantName", label: "مقدم الطلب" },
  { key: "applicantRole", label: "صفة مقدم الطلب" },
  { key: "applicantInsuranceNumber", label: "الرقم التأميني لمقدم الطلب" },
  { key: "applicantPhone", label: "رقم تليفون مقدم الطلب" },
  { key: "applicantNationalId", label: "الرقم القومي لمقدم الطلب" },
  { key: "endDate", label: "تاريخ انتهاء الاشتراك" },
  { key: "endReason", label: "سبب انتهاء الاشتراك" },
  { key: "address", label: "العنوان" },
];

const EMPTY_VALUES: Omit<PersonRecord, "id"> = {
  office: "",
  establishmentName: "",
  establishmentNumber: "",
  establishmentType: "",
  insuredName: "",
  insuranceNumber: "",
  nationalId: "",
  qualification: "",
  profession: "",
  professionCode: "",
  sector: "",
  startDate: "",
  birthDate: "",
  gender: "",
  category: "",
  medicalExam: "",
  contributionCode: "",
  workType: "",
  basicWage: "",
  variableWage: "",
  totalWage: "",
  increaseDate: "",
  increasePercent: "",
  country: "",
  city: "",
  governorate: "",
  buildingNumber: "",
  district: "",
  street: "",
  center: "",
  phone: "",
  email: "",
  employer: "",
  manager: "",
  releaseDate: "",
  applicantName: "",
  applicantRole: "",
  applicantInsuranceNumber: "",
  applicantPhone: "",
  applicantNationalId: "",
  endDate: "",
  endReason: "",
  address: "",
};

export function makeEmptyRecord(): PersonRecord {
  return { id: `${Date.now()}-${Math.random().toString(16).slice(2)}`, ...EMPTY_VALUES };
}

const REQUIRED_FIELDS: Record<TemplateId, Array<keyof PersonRecord>> = {
  s1: [
    "office", "applicantName", "applicantRole", "applicantInsuranceNumber", "applicantPhone", "applicantNationalId",
    "insuredName", "insuranceNumber", "nationalId", "qualification", "profession", "sector", "category", "medicalExam",
    "contributionCode", "startDate", "basicWage", "totalWage", "establishmentName",
    "establishmentNumber", "establishmentType", "country", "governorate", "buildingNumber", "street", "center", "phone", "workType", "address",
  ],
  s6: [
    "office", "applicantName", "applicantRole", "applicantInsuranceNumber", "applicantPhone", "applicantNationalId",
    "insuredName", "insuranceNumber", "nationalId", "establishmentName", "establishmentNumber",
    "endDate", "endReason", "address",
  ],
};

const FIELD_LABELS = new Map<keyof PersonRecord, string>(EXCEL_HEADERS.map(({ key, label }) => [key, label]));
const digitsOnlyFields: Array<keyof PersonRecord> = [
  "nationalId", "applicantNationalId", "insuranceNumber", "applicantInsuranceNumber", "establishmentNumber",
  "professionCode", "contributionCode", "phone", "applicantPhone", "buildingNumber",
];
const dateFields: Array<keyof PersonRecord> = ["startDate", "birthDate", "releaseDate", "endDate"];
const moneyFields: Array<keyof PersonRecord> = ["basicWage", "variableWage", "totalWage"];

export function normalizeDigits(value: unknown) {
  return String(value ?? "")
    .replace(/[٠-٩]/g, (digit) => String("٠١٢٣٤٥٦٧٨٩".indexOf(digit)))
    .replace(/[۰-۹]/g, (digit) => String("۰۱۲۳۴۵۶۷۸۹".indexOf(digit)));
}

export function normalizeText(value: unknown) {
  return normalizeDigits(value)
    .normalize("NFC")
    .replace(/[\u200e\u200f\u202a-\u202e\u2066-\u2069]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function isValidDate(value: string) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) return false;
  const [year, month, day] = value.split("-").map(Number);
  const date = new Date(Date.UTC(year, month - 1, day));
  return date.getUTCFullYear() === year && date.getUTCMonth() === month - 1 && date.getUTCDate() === day;
}

export function validateRecord(record: PersonRecord, template: TemplateId = "s1"): ValidationIssue[] {
  const issues: ValidationIssue[] = [];
  for (const key of REQUIRED_FIELDS[template]) {
    if (!record[key].trim()) issues.push({ key, message: `${FIELD_LABELS.get(key) ?? key}: حقل مطلوب` });
  }
  for (const key of digitsOnlyFields) {
    const value = record[key].trim();
    if (value && !/^\d+$/.test(normalizeDigits(value))) issues.push({ key, message: `${FIELD_LABELS.get(key) ?? key}: استخدم أرقامًا فقط` });
  }
  if (record.nationalId && normalizeDigits(record.nationalId).length !== 14) issues.push({ key: "nationalId", message: "الرقم القومي: يجب أن يتكون من 14 رقمًا" });
  if (record.applicantNationalId && normalizeDigits(record.applicantNationalId).length !== 14) issues.push({ key: "applicantNationalId", message: "الرقم القومي لمقدم الطلب: يجب أن يتكون من 14 رقمًا" });
  if (record.phone && normalizeDigits(record.phone).length !== 11) issues.push({ key: "phone", message: "التليفون: يجب أن يتكون من 11 رقمًا" });
  if (record.applicantPhone && normalizeDigits(record.applicantPhone).length !== 11) issues.push({ key: "applicantPhone", message: "تليفون مقدم الطلب: يجب أن يتكون من 11 رقمًا" });
  for (const key of dateFields) {
    const value = record[key].trim();
    if (value && !isValidDate(value)) issues.push({ key, message: `${FIELD_LABELS.get(key) ?? key}: تاريخ غير صحيح` });
  }
  for (const key of moneyFields) {
    const value = normalizeDigits(record[key]).replace(",", ".");
    if (value && (!/^\d+(\.\d{1,2})?$/.test(value) || Number(value) < 0)) issues.push({ key, message: `${FIELD_LABELS.get(key) ?? key}: قيمة رقمية غير صحيحة` });
  }
  if (record.email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(record.email)) issues.push({ key: "email", message: "البريد الإلكتروني: صيغة غير صحيحة" });
  if (record.startDate && record.birthDate && record.startDate < record.birthDate) issues.push({ key: "startDate", message: "تاريخ بدء الاشتراك يسبق تاريخ الميلاد" });
  if (record.endDate && record.startDate && record.endDate < record.startDate) issues.push({ key: "endDate", message: "تاريخ انتهاء الاشتراك يسبق تاريخ البدء" });
  return issues;
}

export function relevantFieldCount(template: TemplateId = "s1") {
  return REQUIRED_FIELDS[template].length;
}

export function recordStatus(record: PersonRecord, template: TemplateId = "s1"): RecordStatus {
  const filled = REQUIRED_FIELDS[template].filter((key) => record[key].trim()).length;
  if (filled === 0) return "مسودة";
  if (validateRecord(record, template).length === 0) return "جاهز";
  return "ناقص";
}

export function filledCount(record: PersonRecord) {
  return Object.values(record).filter(Boolean).length - 1;
}

function normalizeKey(value: unknown) {
  return normalizeText(value)
    .trim()
    .toLocaleLowerCase("ar-EG")
    .replace(/[\s_\-–—:/\\().]+/g, "");
}

const aliases: Record<keyof PersonRecord, string[]> = {
  id: ["id", "رقم"],
  insuredName: ["اسم المؤمن عليه", "الاسم", "اسم العامل", "insuredname", "name"],
  nationalId: ["الرقم القومي", "رقم قومي", "الرقم القومى", "nationalid", "national id"],
  insuranceNumber: ["الرقم التأميني", "الرقم التامينى", "رقم تأميني", "insurance number", "insurancenumber"],
  establishmentName: ["اسم المنشأة", "اسم المنشأه", "المنشأة", "اسم الشركة", "establishmentname"],
  establishmentNumber: ["رقم المنشأة", "رقم المنشأه", "establishmentnumber"],
  establishmentType: ["نوع المنشأة", "نوع المنشأه", "establishmenttype"],
  office: ["المكتب", "مكتب التأمينات", "office"],
  qualification: ["المؤهل", "المؤهل الدراسي", "qualification"],
  profession: ["المهنة", "المهنه", "الوظيفة", "profession", "job"],
  professionCode: ["كود المهنة", "كود المهنه", "professioncode", "jobcode"],
  sector: ["القطاع", "قطاع", "sector"],
  startDate: ["تاريخ بدء الاشتراك", "تاريخ بدء الاشتراك التأميني", "startdate", "subscriptionstartdate"],
  birthDate: ["تاريخ الميلاد", "birthdate", "dateofbirth"],
  gender: ["النوع", "الجنس", "gender", "sex"],
  category: ["الفئة", "الفئه", "category"],
  medicalExam: ["استيفاء الكشف الطبي الابتدائي", "الكشف الطبي", "medicalexam"],
  contributionCode: ["كود الاشتراك", "كود الاشتراك التأميني", "contributioncode"],
  workType: ["نوع المدة", "نوع المده", "نوع العمل", "worktype"],
  basicWage: ["أجر / دخل الاشتراك", "أجر الاشتراك", "دخل الاشتراك", "الأجر الأساسي", "الاجر الاساسي", "basicwage"],
  variableWage: ["الأجر المتغير", "الاجر المتغير", "variablewage"],
  totalWage: ["الأجر الشامل", "الاجر الشامل", "totalwage", "wage"],
  increaseDate: ["تاريخ بداية العجز", "تاريخ الزيادة", "increasedate"],
  increasePercent: ["نسبة العجز", "نسبة الزيادة", "increasepercent"],
  country: ["الجنسية", "الجنسيه", "الدولة", "الدوله", "nationality", "country"],
  city: ["المدينة", "المدينه", "city"],
  governorate: ["المحافظة", "المحافظه", "governorate"],
  buildingNumber: ["رقم العقار", "عقار رقم", "buildingnumber"],
  district: ["الشياخة / القرية", "الشياخة", "القرية", "district", "village"],
  street: ["الشارع", "street"],
  center: ["القسم / المركز", "القسم", "المركز", "center", "district center"],
  phone: ["التليفون", "الهاتف", "رقم الهاتف", "phone", "telephone"],
  email: ["البريد الإلكتروني", "البريد الالكتروني", "email"],
  employer: ["جهة العمل", "جهة العمل / المنشأة", "employer"],
  manager: ["المدير المسؤول", "المدير المسئول", "manager"],
  releaseDate: ["تحريرًا في", "تحريرا في", "تاريخ التحرير", "releasedate"],
  applicantName: ["مقدم الطلب", "اسم مقدم الطلب", "applicantname"],
  applicantRole: ["صفة مقدم الطلب", "صفة مقدم الطلب", "applicantrole"],
  applicantInsuranceNumber: ["الرقم التأميني لمقدم الطلب", "رقم تأميني مقدم الطلب", "applicantinsurancenumber"],
  applicantPhone: ["رقم تليفون مقدم الطلب", "تليفون مقدم الطلب", "applicantphone"],
  applicantNationalId: ["الرقم القومي لمقدم الطلب", "رقم قومي مقدم الطلب", "applicantnationalid"],
  endDate: ["تاريخ انتهاء الاشتراك", "تاريخ الانتهاء", "enddate", "subscriptionenddate"],
  endReason: ["سبب انتهاء الاشتراك", "سبب الانتهاء", "endreason", "subscriptionendreason"],
  address: ["العنوان", "عنوان مقدم الطلب", "address"],
};

export function mapExcelRow(row: Record<string, unknown>): PersonRecord {
  const normalized = new Map(Object.entries(row).map(([key, value]) => [normalizeKey(key), value]));
  const result = makeEmptyRecord();
  (Object.keys(aliases) as Array<keyof PersonRecord>).forEach((key) => {
    const found = aliases[key].find((candidate) => normalized.has(normalizeKey(candidate)));
    if (!found) return;
    const rawValue = normalized.get(normalizeKey(found));
    if (rawValue instanceof Date && !Number.isNaN(rawValue.getTime())) {
      const year = rawValue.getFullYear();
      const month = String(rawValue.getMonth() + 1).padStart(2, "0");
      const day = String(rawValue.getDate()).padStart(2, "0");
      result[key] = `${year}-${month}-${day}`;
    } else {
      result[key] = normalizeText(rawValue);
      // Excel often stores Egyptian phone numbers as numeric cells and drops
      // the leading zero. Restore it while importing the fixed 11-digit fields.
      if ((key === "phone" || key === "applicantPhone") && /^\d{1,10}$/.test(normalizeDigits(result[key]))) {
        result[key] = normalizeDigits(result[key]).padStart(11, "0");
      }
    }
  });
  return result;
}

export function displayValue(value: string | undefined, empty = "—") {
  return value?.trim() ? value : empty;
}

