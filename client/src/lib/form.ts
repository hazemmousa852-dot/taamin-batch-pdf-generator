// Design reminder: form vocabulary follows the “Registry Desk” philosophy—quiet, precise, and close to the printed document.

export type TemplateId = "s1" | "s6";
export const TEMPLATE_LABELS: Record<TemplateId, string> = { s1: "نموذج س1", s6: "نموذج س6" };
export type RecordStatus = "مسودة" | "جاهز" | "ناقص";

export type PersonRecord = {
  id: string;
  office: string;
  establishmentName: string;
  establishmentNumber: string;
  insuredName: string;
  insuranceNumber: string;
  nationalId: string;
  qualification: string;
  profession: string;
  professionCode: string;
  startDate: string;
  birthDate: string;
  gender: string;
  category: string;
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
  { key: "establishmentName", label: "اسم المنشأة" },
  { key: "establishmentNumber", label: "رقم المنشأة" },
  { key: "office", label: "المكتب" },
  { key: "qualification", label: "المؤهل" },
  { key: "profession", label: "المهنة" },
  { key: "professionCode", label: "كود المهنة" },
  { key: "startDate", label: "تاريخ بدء الاشتراك" },
  { key: "birthDate", label: "تاريخ الميلاد" },
  { key: "gender", label: "النوع" },
  { key: "category", label: "الفئة" },
  { key: "contributionCode", label: "كود الاشتراك" },
  { key: "workType", label: "نوع المدة" },
  { key: "basicWage", label: "الأجر الأساسي" },
  { key: "variableWage", label: "الأجر المتغير" },
  { key: "totalWage", label: "الأجر الشامل" },
  { key: "increaseDate", label: "تاريخ بداية العجز" },
  { key: "increasePercent", label: "نسبة العجز" },
  { key: "country", label: "الدولة" },
  { key: "city", label: "المدينة" },
  { key: "governorate", label: "المحافظة" },
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
  insuredName: "",
  insuranceNumber: "",
  nationalId: "",
  qualification: "",
  profession: "",
  professionCode: "",
  startDate: "",
  birthDate: "",
  gender: "",
  category: "",
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
  applicantPhone: "",
  applicantNationalId: "",
  endDate: "",
  endReason: "",
  address: "",
};

export function makeEmptyRecord(): PersonRecord {
  return { id: `${Date.now()}-${Math.random().toString(16).slice(2)}`, ...EMPTY_VALUES };
}

export function recordStatus(record: PersonRecord): RecordStatus {
  const core = [record.insuredName, record.nationalId, record.establishmentName];
  if (core.every(Boolean)) return "جاهز";
  if (core.some(Boolean)) return "ناقص";
  return "مسودة";
}

export function filledCount(record: PersonRecord) {
  return Object.values(record).filter(Boolean).length - 1;
}

function normalizeKey(value: unknown) {
  return String(value ?? "")
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
  office: ["المكتب", "مكتب التأمينات", "office"],
  qualification: ["المؤهل", "المؤهل الدراسي", "qualification"],
  profession: ["المهنة", "المهنه", "الوظيفة", "profession", "job"],
  professionCode: ["كود المهنة", "كود المهنه", "professioncode", "jobcode"],
  startDate: ["تاريخ بدء الاشتراك", "تاريخ بدء الاشتراك التأميني", "startdate", "subscriptionstartdate"],
  birthDate: ["تاريخ الميلاد", "birthdate", "dateofbirth"],
  gender: ["النوع", "الجنس", "gender", "sex"],
  category: ["الفئة", "الفئه", "category"],
  contributionCode: ["كود الاشتراك", "كود الاشتراك التأميني", "contributioncode"],
  workType: ["نوع المدة", "نوع المده", "نوع العمل", "worktype"],
  basicWage: ["الأجر الأساسي", "الاجر الاساسي", "basicwage"],
  variableWage: ["الأجر المتغير", "الاجر المتغير", "variablewage"],
  totalWage: ["الأجر الشامل", "الاجر الشامل", "totalwage", "wage"],
  increaseDate: ["تاريخ بداية العجز", "تاريخ الزيادة", "increasedate"],
  increasePercent: ["نسبة العجز", "نسبة الزيادة", "increasepercent"],
  country: ["الدولة", "الدوله", "country"],
  city: ["المدينة", "المدينه", "city"],
  governorate: ["المحافظة", "المحافظه", "governorate"],
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
    if (found) result[key] = String(normalized.get(normalizeKey(found)) ?? "").trim();
  });
  return result;
}

export function displayValue(value: string | undefined, empty = "—") {
  return value?.trim() ? value : empty;
}
