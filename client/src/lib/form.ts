// Design reminder: form vocabulary follows the “Registry Desk” philosophy—quiet, precise, and close to the printed document.

export type TemplateId = "s1" | "s2" | "s6" | "s31";
export const TEMPLATE_LABELS: Record<TemplateId, string> = { s1: "نموذج س1", s2: "نموذج س2", s6: "نموذج س6", s31: "نموذج 31" };
export type RecordStatus = "مسودة" | "جاهز" | "ناقص";
export type ValidationIssue = { key: keyof PersonRecord; message: string };

export type CodedOption = { value: string; label: string };
export const CODED_OPTIONS = {
  sector: [
    { value: "1", label: "العاملين بالقطاع الحكومي والهيئات والمؤسسات العامة التابعة للحكومة" },
    { value: "2", label: "العاملين بالقطاع العام وقطاع الأعمال" },
    { value: "3", label: "العاملين بالقطاع الخاص" },
    { value: "4", label: "العاملين بقطاع المقاولات" },
    { value: "5", label: "العاملين بقطاع النقل البري" },
    { value: "6", label: "العاملين بقطاع المخابز" },
    { value: "7", label: "العاملين بالخارج" },
    { value: "8", label: "أصحاب الأعمال ومن في حكمهم" },
    { value: "9", label: "العاملون بقطاع العمالة غير المنتظمة" },
  ],
  contributionCode: [
    { value: "11", label: "عادية" }, { value: "12", label: "أعمال صعبة" },
    { value: "13", label: "أعمال خطرة" }, { value: "14", label: "متدرج بدون أجر" },
    { value: "15", label: "متدرج بأجر" },
  ],
  workType: [
    { value: "01", label: "فعلية" },
    { value: "02", label: "بعد سنتين من تأمين البطالة" },
    { value: "03", label: "بعد سن إيقاف انتفاع تأمين الشيخوخة والعجز والوفاة بالداخل" },
    { value: "04", label: "بعد سن إيقاف انتفاع تأمين الشيخوخة والعجز والوفاة بالخارج" },
    { value: "05", label: "تأمين مع إيقاف تأمين المرضى" },
    { value: "06", label: "بعد سن الـ 60" },
    { value: "21", label: "خدمة تجنيد إجبارية" },
  ],
} satisfies Record<"sector" | "contributionCode" | "workType", CodedOption[]>;

/** اسم ورقة «نموذج 31» داخل قالب Excel — ورقة واحدة تضم كل البيانات المطلوبة. */
export const FORM31_SHEET_NAME = "بيان كل الداتا المطلوبة";
/** أسماء أوراق قديمة في ملفات سابقة — تُقرأ للتوافق ولا تُكتب في القالب الجديد. */
export const FORM31_LEGACY_SHEET_NAMES = ["نموذج 31"];

export type PersonRecord = {
  hireDate: string;
  receiptNumber: string;
  receiptDate: string;
  insuredSignDate: string;
  managerSignDate: string;
  signatureVerifier: string;
  disputeManagerSignDate: string;
  disputeInsuredSignDate: string;
  applicantSignDate: string;
  signatureMatchDate: string;

  // نموذج رقم (31): طلب صرف تعويض الأجر ومصاريف الانتقال في حالة الإصابة / المرض
  injuryDate: string;
  companion: string;
  treatmentPlace: string;
  travelExpenses: string;
  wagePaymentMethod: string;
  monthlyWage: string;
  sickLeaveStartDate: string;
  sickLeaveEndDate: string;
  otherLine1: string;
  otherLine2: string;
  otherLine3: string;
  signatureName: string;

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
  noticeDate: string;
  taxRegistrationNumber: string;
  commercialRegistrationNumber: string;
  unifiedCommercialRegistrationNumber: string;
  declarationRole: string;
};

/** يحدد الورقة التي يُقرأ منها أو يُكتب فيها كل عمود داخل قالب Excel. */
export type ExcelSheet = "shared" | "people" | "form31";

export const EXCEL_HEADERS: Array<{ key: keyof PersonRecord; label: string; sheet?: ExcelSheet }> = [
  { key: "hireDate", label: "تاريخ الالتحاق (س2) (DD/MM/YYYY)" },
  { key: "receiptNumber", label: "رقم وصل الخطاب المسجل بعلم الوصول (س6)" },
  { key: "receiptDate", label: "تاريخ وصل الخطاب المسجل (س6) (DD/MM/YYYY)" },
  { key: "insuredSignDate", label: "تاريخ توقيع المؤمن عليه (س6) (DD/MM/YYYY)" },
  { key: "managerSignDate", label: "تاريخ توقيع المدير المسؤول (س6) (DD/MM/YYYY)" },
  { key: "signatureVerifier", label: "اسم القائم بمطابقة التوقيع (س6)" },
  { key: "disputeManagerSignDate", label: "تاريخ توقيع المدير المسؤول في حالة النزاع (س6) (DD/MM/YYYY)" },
  { key: "disputeInsuredSignDate", label: "تاريخ توقيع المؤمن عليه في حالة النزاع (س6) (DD/MM/YYYY)" },
  { key: "applicantSignDate", label: "تاريخ توقيع مقدم الطلب (س6) (DD/MM/YYYY)" },
  { key: "signatureMatchDate", label: "تاريخ مطابقة التوقيع (س1) (DD/MM/YYYY)" },

  // أعمدة ورقة «نموذج 31». القيم المشتركة (المكتب، المنشأة، مقدم الطلب) تُملأ
  // من ورقة «بيانات المنشأة والمفوض» تلقائيًا عند الاستيراد.
  { key: "injuryDate", label: "تاريخ الإصابة (نموذج 31) (DD/MM/YYYY)", sheet: "form31" },
  { key: "companion", label: "هل يوجد مرافق (نموذج 31)", sheet: "form31" },
  { key: "sickLeaveStartDate", label: "تاريخ بداية الإجازة المرضي (نموذج 31) (DD/MM/YYYY)", sheet: "form31" },
  { key: "sickLeaveEndDate", label: "تاريخ نهاية الإجازة المرضي (نموذج 31) (DD/MM/YYYY)", sheet: "form31" },
  { key: "travelExpenses", label: "هل توجد مصاريف انتقال (نموذج 31)", sheet: "form31" },
  { key: "treatmentPlace", label: "العلاج يتم داخل أم خارج البلاد (نموذج 31)", sheet: "form31" },
  { key: "monthlyWage", label: "أجر الاشتراك الشهري للمصاب (نموذج 31)", sheet: "form31" },
  { key: "wagePaymentMethod", label: "طريقة أداء الأجر (نموذج 31)", sheet: "form31" },
  { key: "otherLine1", label: "بيانات أخرى 1 (نموذج 31)", sheet: "form31" },
  { key: "otherLine2", label: "بيانات أخرى 2 (نموذج 31)", sheet: "form31" },
  { key: "otherLine3", label: "بيانات أخرى 3 (نموذج 31)", sheet: "form31" },
  { key: "signatureName", label: "الاسم أسفل توقيع مقدم الطلب (نموذج 31)", sheet: "form31" },

  { key: "insuredName", label: "اسم المؤمن عليه" },
  { key: "nationalId", label: "الرقم القومي" },
  { key: "insuranceNumber", label: "الرقم التأميني" },
  { key: "startDate", label: "تاريخ بداية الاشتراك (DD/MM/YYYY)" },
  { key: "establishmentName", label: "اسم المنشأة" },
  { key: "establishmentNumber", label: "رقم المنشأة" },
  { key: "establishmentType", label: "نوع المنشأة" },
  { key: "office", label: "المكتب" },
  { key: "qualification", label: "المؤهل" },
  { key: "profession", label: "المهنة" },
  { key: "professionCode", label: "كود المهنة" },
  { key: "sector", label: "القطاع" },
  { key: "birthDate", label: "تاريخ الميلاد (DD/MM/YYYY)" },
  { key: "gender", label: "النوع" },
  { key: "category", label: "الفئة" },
  { key: "medicalExam", label: "استيفاء الكشف الطبي الابتدائي" },
  { key: "contributionCode", label: "كود الاشتراك" },
  { key: "workType", label: "نوع المدة" },
  { key: "basicWage", label: "أجر / دخل الاشتراك" },
  { key: "variableWage", label: "الأجر المتغير" },
  { key: "totalWage", label: "الأجر الشامل" },
  { key: "increaseDate", label: "تاريخ بداية العجز - خانة مستقلة (DD/MM/YYYY)" },
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
  { key: "releaseDate", label: "تحريرًا في (DD/MM/YYYY)" },
  { key: "applicantName", label: "مقدم الطلب" },
  { key: "applicantRole", label: "صفة مقدم الطلب" },
  { key: "applicantInsuranceNumber", label: "الرقم التأميني لمقدم الطلب" },
  { key: "applicantPhone", label: "رقم تليفون مقدم الطلب" },
  { key: "applicantNationalId", label: "الرقم القومي لمقدم الطلب" },
  { key: "endDate", label: "تاريخ انتهاء الاشتراك (DD/MM/YYYY)" },
  { key: "endReason", label: "سبب انتهاء الاشتراك" },
  { key: "address", label: "العنوان" },
  { key: "noticeDate", label: "تاريخ الإخطار (DD/MM/YYYY)" },
  { key: "taxRegistrationNumber", label: "رقم التسجيل الضريبي" },
  { key: "commercialRegistrationNumber", label: "رقم السجل التجاري" },
  { key: "unifiedCommercialRegistrationNumber", label: "رقم السجل التجاري الموحد" },
  { key: "declarationRole", label: "الصفة في الإقرار" },
];

const EMPTY_VALUES: Omit<PersonRecord, "id"> = {
  hireDate: "",
  receiptNumber: "",
  receiptDate: "",
  insuredSignDate: "",
  managerSignDate: "",
  signatureVerifier: "",
  disputeManagerSignDate: "",
  disputeInsuredSignDate: "",
  applicantSignDate: "",
  signatureMatchDate: "",
  injuryDate: "",
  companion: "",
  treatmentPlace: "",
  travelExpenses: "",
  wagePaymentMethod: "",
  monthlyWage: "",
  sickLeaveStartDate: "",
  sickLeaveEndDate: "",
  otherLine1: "",
  otherLine2: "",
  otherLine3: "",
  signatureName: "",

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
  noticeDate: "",
  taxRegistrationNumber: "",
  commercialRegistrationNumber: "",
  unifiedCommercialRegistrationNumber: "",
  declarationRole: "",
};

export function makeEmptyRecord(): PersonRecord {
  return { id: `${Date.now()}-${Math.random().toString(16).slice(2)}`, ...EMPTY_VALUES };
}

const REQUIRED_FIELDS: Record<TemplateId, Array<keyof PersonRecord>> = {
  s1: [
    "office", "applicantName", "applicantRole", "applicantInsuranceNumber", "applicantPhone", "applicantNationalId",
    "insuredName", "insuranceNumber", "nationalId", "qualification", "profession", "category", "medicalExam",
    "startDate", "basicWage", "totalWage", "establishmentName",
    "establishmentNumber", "establishmentType", "phone",
  ],
  s6: [
    "office", "applicantName", "applicantRole", "applicantInsuranceNumber", "applicantPhone", "applicantNationalId",
    "insuredName", "insuranceNumber", "nationalId", "establishmentName", "establishmentNumber",
    "endDate", "endReason",
  ],
  s2: ["insuredName", "insuranceNumber", "nationalId", "startDate", "basicWage", "totalWage"],
  // نموذج 31: أعمدة الورقة الخاصة + البيانات المشتركة التي تظهر في النموذج المطبوع
  s31: [
    "insuredName", "nationalId", "insuranceNumber", "injuryDate", "companion",
    "sickLeaveStartDate", "sickLeaveEndDate", "travelExpenses", "treatmentPlace",
    "monthlyWage", "wagePaymentMethod", "otherLine1", "otherLine2", "otherLine3",
    "releaseDate", "signatureName",
    "office", "establishmentName", "establishmentNumber",
    "applicantName", "applicantRole", "applicantInsuranceNumber", "applicantPhone", "applicantNationalId",
  ],
};

const FIELD_LABELS = new Map<keyof PersonRecord, string>(EXCEL_HEADERS.map(({ key, label }) => [key, label]));

/**
 * أعمدة الورقة الواحدة لنموذج 31 في قالب Excel — بترتيب مطابق لسير النموذج المطبوع:
 * بيانات المكتب ومقدم الطلب والمنشأة أولًا (تُكتب مرة واحدة وتُورَّث تلقائيًا لكل صف
 * تالٍ يُترك فارغًا)، ثم بيانات كل حالة إصابة أو إجازة مرضية على حدة.
 */
export const FORM31_SHEET_COLUMNS: Array<{ key: keyof PersonRecord; label: string; hint?: string }> = [
  { key: "office", label: "المكتب", hint: "مكتب التأمينات" },
  { key: "applicantName", label: "مقدم الطلب" },
  { key: "applicantRole", label: "صفة مقدم الطلب", hint: "مثال: المدير المسؤول" },
  { key: "applicantInsuranceNumber", label: "الرقم التأميني لمقدم الطلب" },
  { key: "applicantPhone", label: "رقم تليفون مقدم الطلب" },
  { key: "applicantNationalId", label: "الرقم القومي لمقدم الطلب", hint: "14 رقمًا" },
  { key: "establishmentName", label: "اسم المنشأة" },
  { key: "establishmentNumber", label: "رقم المنشأة" },
  { key: "insuredName", label: "اسم المؤمن عليه" },
  { key: "nationalId", label: "الرقم القومي", hint: "14 رقمًا" },
  { key: "insuranceNumber", label: "الرقم التأميني", hint: "كما هو بدون فواصل" },
  { key: "injuryDate", label: "تاريخ الإصابة (DD/MM/YYYY)", hint: "يوم/شهر/سنة — اتركه فارغًا في حالات المرض" },
  { key: "companion", label: "هل يوجد مرافق", hint: "نعم / لا" },
  { key: "sickLeaveStartDate", label: "تاريخ بداية الإجازة المرضي (DD/MM/YYYY)" },
  { key: "sickLeaveEndDate", label: "تاريخ نهاية الإجازة المرضي (DD/MM/YYYY)" },
  { key: "travelExpenses", label: "هل توجد مصاريف انتقال", hint: "نعم / لا" },
  { key: "treatmentPlace", label: "العلاج يتم داخل أم خارج البلاد", hint: "داخل البلاد / خارج البلاد" },
  { key: "monthlyWage", label: "أجر الاشتراك الشهري للمصاب" },
  { key: "wagePaymentMethod", label: "طريقة أداء الأجر" },
  { key: "otherLine1", label: "بيانات أخرى 1" },
  { key: "otherLine2", label: "بيانات أخرى 2" },
  { key: "otherLine3", label: "بيانات أخرى 3" },
  { key: "releaseDate", label: "تحريرًا في (DD/MM/YYYY)" },
  { key: "signatureName", label: "الاسم أسفل توقيع مقدم الطلب", hint: "اتركه فارغًا لاستخدام اسم مقدم الطلب" },
];

/**
 * الحقول المشتركة في الورقة الواحدة: تُكتب مرة واحدة في أول صف، ويورّثها النظام
 * تلقائيًا لكل صف تالٍ تركه المستخدم فارغًا.
 */
export const FORM31_CARRY_FIELDS: Array<keyof PersonRecord> = [
  "office", "applicantName", "applicantRole", "applicantInsuranceNumber", "applicantPhone",
  "applicantNationalId", "establishmentName", "establishmentNumber",
];

/** يورّث الحقول المشتركة من سجل سابق إلى سجل جديد عند تركها فارغة. */
export function carryForwardShared(target: PersonRecord, source?: PersonRecord) {
  if (!source) return target;
  for (const key of FORM31_CARRY_FIELDS) {
    if (!target[key] && source[key]) target[key] = source[key];
  }
  return target;
}

const dateFields: Array<keyof PersonRecord> = ["hireDate", "receiptDate", "insuredSignDate", "managerSignDate", "disputeManagerSignDate", "disputeInsuredSignDate", "applicantSignDate", "signatureMatchDate", "startDate", "birthDate", "increaseDate", "releaseDate", "endDate", "noticeDate", "injuryDate", "sickLeaveStartDate", "sickLeaveEndDate"];
const moneyFields: Array<keyof PersonRecord> = ["basicWage", "variableWage", "totalWage", "monthlyWage"];

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
  // All fields are optional. Validate only supplied values.
  for (const key of dateFields) {
    const value = record[key].trim();
    if (value && !isValidDate(value)) issues.push({ key, message: `${FIELD_LABELS.get(key) ?? key}: تاريخ غير صحيح` });
  }
  for (const key of moneyFields) {
    const value = normalizeDigits(record[key]).replace(",", ".");
    if (value && (!/^\d+(\.\d{1,2})?$/.test(value) || Number(value) < 0)) issues.push({ key, message: `${FIELD_LABELS.get(key) ?? key}: قيمة رقمية غير صحيحة` });
  }
  if (record.email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(record.email)) issues.push({ key: "email", message: "البريد الإلكتروني: صيغة غير صحيحة" });
  return issues;
}

export function relevantFieldCount(template: TemplateId = "s1") {
  return REQUIRED_FIELDS[template].length;
}

export function recordStatus(record: PersonRecord, template: TemplateId = "s1"): RecordStatus {
  const filled = filledCount(record);
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
  hireDate: ["تاريخ الالتحاق (س2) (DD/MM/YYYY)", "hireDate"],
  receiptNumber: ["رقم وصل الخطاب المسجل بعلم الوصول (س6)", "receiptNumber"],
  receiptDate: ["تاريخ وصل الخطاب المسجل (س6) (DD/MM/YYYY)", "receiptDate"],
  insuredSignDate: ["تاريخ توقيع المؤمن عليه (س6) (DD/MM/YYYY)", "insuredSignDate"],
  managerSignDate: ["تاريخ توقيع المدير المسؤول (س6) (DD/MM/YYYY)", "managerSignDate"],
  signatureVerifier: ["اسم القائم بمطابقة التوقيع (س6)", "signatureVerifier"],
  disputeManagerSignDate: ["تاريخ توقيع المدير المسؤول في حالة النزاع (س6) (DD/MM/YYYY)", "disputeManagerSignDate"],
  disputeInsuredSignDate: ["تاريخ توقيع المؤمن عليه في حالة النزاع (س6) (DD/MM/YYYY)", "disputeInsuredSignDate"],
  applicantSignDate: ["تاريخ توقيع مقدم الطلب (س6) (DD/MM/YYYY)", "applicantSignDate"],
  signatureMatchDate: ["تاريخ مطابقة التوقيع (س1) (DD/MM/YYYY)", "signatureMatchDate"],

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
  startDate: ["تاريخ بداية الاشتراك", "تاريخ بداية الاشتراك (DD/MM/YYYY)", "تاريخ بداية الاشتراك (YYYY-MM-DD)", "تاريخ بدء الاشتراك", "تاريخ بدء الاشتراك التأميني", "startdate", "subscriptionstartdate"],
  birthDate: ["تاريخ الميلاد", "تاريخ الميلاد (DD/MM/YYYY)", "birthdate", "dateofbirth"],
  gender: ["النوع", "الجنس", "gender", "sex"],
  category: ["الفئة", "الفئه", "category"],
  medicalExam: ["استيفاء الكشف الطبي الابتدائي", "الكشف الطبي", "medicalexam"],
  contributionCode: ["كود الاشتراك", "كود الاشتراك التأميني", "contributioncode"],
  workType: ["نوع المدة", "نوع المده", "نوع العمل", "worktype"],
  basicWage: ["أجر / دخل الاشتراك", "أجر الاشتراك", "دخل الاشتراك", "الأجر الأساسي", "الاجر الاساسي", "basicwage"],
  variableWage: ["الأجر المتغير", "الاجر المتغير", "variablewage"],
  totalWage: ["الأجر الشامل", "الاجر الشامل", "totalwage", "wage"],
  increaseDate: ["تاريخ بداية العجز", "تاريخ بداية العجز (DD/MM/YYYY)", "تاريخ بداية العجز - خانة مستقلة (DD/MM/YYYY)", "تاريخ الزيادة", "increasedate"],
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
  releaseDate: ["تحريرًا في", "تحريرًا في (DD/MM/YYYY)", "تحريرا في", "تاريخ التحرير", "releasedate"],
  applicantName: ["مقدم الطلب", "اسم مقدم الطلب", "applicantname"],
  applicantRole: ["صفة مقدم الطلب", "صفة مقدم الطلب", "applicantrole"],
  applicantInsuranceNumber: ["الرقم التأميني لمقدم الطلب", "رقم تأميني مقدم الطلب", "applicantinsurancenumber"],
  applicantPhone: ["رقم تليفون مقدم الطلب", "تليفون مقدم الطلب", "applicantphone"],
  applicantNationalId: ["الرقم القومي لمقدم الطلب", "رقم قومي مقدم الطلب", "applicantnationalid"],
  endDate: ["تاريخ انتهاء الاشتراك", "تاريخ انتهاء الاشتراك (DD/MM/YYYY)", "تاريخ الانتهاء", "enddate", "subscriptionenddate"],
  endReason: ["سبب انتهاء الاشتراك", "سبب الانتهاء", "endreason", "subscriptionendreason"],
  address: ["العنوان", "عنوان مقدم الطلب", "address"],
  noticeDate: ["تاريخ الإخطار", "تاريخ الإخطار (DD/MM/YYYY)", "noticedate"],
  taxRegistrationNumber: ["رقم التسجيل الضريبي", "رقم التسجيل الضريبى", "taxregistrationnumber"],
  commercialRegistrationNumber: ["رقم السجل التجاري", "رقم السجل التجارى", "commercialregistrationnumber"],
  unifiedCommercialRegistrationNumber: ["رقم السجل التجاري الموحد", "رقم السجل التجارى الموحد", "unifiedcommercialregistrationnumber"],
  declarationRole: ["الصفة في الإقرار", "الصفة فى الإقرار", "صفة المقر", "declarationrole"],

  injuryDate: ["تاريخ الإصابة", "تاريخ الإصابة (DD/MM/YYYY)", "تاريخ الإصابة (نموذج 31) (DD/MM/YYYY)", "تاريخ الاصابة", "injurydate"],
  companion: ["هل يوجد مرافق", "هل يوجد مرافق (نموذج 31)", "يوجد مرافق", "المرافق", "companion"],
  treatmentPlace: ["العلاج يتم داخل أم خارج البلاد", "العلاج يتم داخل أم خارج البلاد (نموذج 31)", "العلاج داخل أم خارج البلاد", "مكان العلاج", "treatmentplace"],
  travelExpenses: ["هل توجد مصاريف انتقال", "هل توجد مصاريف انتقال (نموذج 31)", "مصاريف انتقال", "مصاريف الانتقال", "travelexpenses"],
  wagePaymentMethod: ["طريقة أداء الأجر", "طريقة أداء الأجر (نموذج 31)", "طريقة أداء الاجر", "طريقة صرف الأجر", "wagepaymentmethod"],
  monthlyWage: ["أجر الاشتراك الشهري للمصاب", "أجر الاشتراك الشهري للمصاب (نموذج 31)", "أجر الاشتراك الشهري", "الأجر التأميني الشهري", "monthlywage"],
  sickLeaveStartDate: ["تاريخ بداية الإجازة المرضي", "تاريخ بداية الإجازة المرضي (DD/MM/YYYY)", "تاريخ بداية الإجازة المرضي (نموذج 31) (DD/MM/YYYY)", "بداية الإجازة المرضية", "sickleavestartdate"],
  sickLeaveEndDate: ["تاريخ نهاية الإجازة المرضي", "تاريخ نهاية الإجازة المرضي (DD/MM/YYYY)", "تاريخ نهاية الإجازة المرضي (نموذج 31) (DD/MM/YYYY)", "نهاية الإجازة المرضية", "sickleaveenddate"],
  otherLine1: ["بيانات أخرى 1", "بيانات أخرى 1 (نموذج 31)", "otherline1"],
  otherLine2: ["بيانات أخرى 2", "بيانات أخرى 2 (نموذج 31)", "otherline2"],
  otherLine3: ["بيانات أخرى 3", "بيانات أخرى 3 (نموذج 31)", "otherline3"],
  signatureName: ["الاسم أسفل توقيع مقدم الطلب", "الاسم أسفل توقيع مقدم الطلب (نموذج 31)", "اسم الموقّع", "signaturename"],
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
      if (dateFields.includes(key)) {
        const dateMatch = normalizeDigits(result[key]).match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
        if (dateMatch) {
          const [, day, month, year] = dateMatch;
          result[key] = `${year}-${month.padStart(2, "0")}-${day.padStart(2, "0")}`;
        }
      }
      // Excel often stores Egyptian phone numbers as numeric cells and drops
      // the leading zero. Restore it while importing the fixed 11-digit fields.
      if ((key === "phone" || key === "applicantPhone") && /^\d{1,10}$/.test(normalizeDigits(result[key]))) {
        result[key] = normalizeDigits(result[key]).padStart(11, "0");
      }
    }
  });
  for (const key of Object.keys(CODED_OPTIONS) as Array<keyof typeof CODED_OPTIONS>) {
    const current = normalizeText(result[key]);
    if (!current) continue;
    const option = CODED_OPTIONS[key].find(({ value, label }) => current === value || current === label || current.startsWith(`${value} -`) || current.startsWith(`${value}-`));
    if (option) result[key] = option.value;
  }
  return result;
}

export function displayValue(value: string | undefined, empty = "—") {
  return value?.trim() ? value : empty;
}
