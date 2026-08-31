// Design reminder: this page is a neo-editorial registry desk—warm paper surfaces, amber actions, charcoal ink, and document-first hierarchy.

import { useEffect, useMemo, useRef, useState, type ChangeEvent } from "react";
import * as XLSX from "xlsx";
import JSZip from "jszip";
import { getDocument, GlobalWorkerOptions } from "pdfjs-dist";
import workerSrc from "pdfjs-dist/build/pdf.worker.min.mjs?url";
import {
  ArrowDownToLine,
  ArrowLeft,
  BadgeCheck,
  BriefcaseBusiness,
  Building2,
  Check,
  ChevronDown,
  CircleHelp,
  ClipboardList,
  CloudUpload,
  Download,
  FileSpreadsheet,
  FileText,
  FolderOpen,
  Hash,
  LayoutDashboard,
  Loader2,
  Menu,
  MoreHorizontal,
  Plus,
  Printer,
  Search,
  Settings2,
  ShieldCheck,
  Trash2,
  Upload,
  UserRound,
  UsersRound,
  X,
} from "lucide-react";
import { toast } from "sonner";
import { Button } from "@/components/ui/button";
import { Badge } from "@/components/ui/badge";
import { Card } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Tabs, TabsList, TabsTrigger } from "@/components/ui/tabs";
import {
  createMergedPdf,
  createZip,
  downloadBlob,
  fillPdf,
  safeFileName,
} from "@/lib/pdf";
import {
  EXCEL_HEADERS,
  filledCount,
  makeEmptyRecord,
  mapExcelRow,
  normalizeDigits,
  recordStatus,
  relevantFieldCount,
  validateRecord,
  type PersonRecord,
  TEMPLATE_LABELS,
  type TemplateId,
} from "@/lib/form";

const ASSET_BASE = import.meta.env.BASE_URL.endsWith("/") ? import.meta.env.BASE_URL : `${import.meta.env.BASE_URL}/`;
const assetUrl = (name: string) => `${ASSET_BASE}assets/${name}`;
const MARK_URL = assetUrl("taamin-mark.webp");
const PAPER_URL = assetUrl("taamin-paper-surface.webp");
const RIBBON_URL = assetUrl("taamin-spreadsheet-ribbon.webp");
const FORM_URL = assetUrl("taamin-form-preview.webp");
const S6_PREVIEW_URL = assetUrl("taamin-s6-preview.png");
GlobalWorkerOptions.workerSrc = workerSrc;

type EditableKey = Exclude<keyof PersonRecord, "id">;
type Mode = "manual" | "bulk";

type FieldSpec = {
  key: EditableKey;
  label: string;
  placeholder?: string;
  type?: "text" | "date" | "number" | "email";
  wide?: boolean;
};

type FieldGroup = {
  id: string;
  number: string;
  title: string;
  note: string;
  icon: typeof UserRound;
  fields: FieldSpec[];
};

const fieldGroups: FieldGroup[] = [
  {
    id: "identity",
    number: "01",
    title: "بيانات المؤمن عليه",
    note: "المعلومات الأساسية كما ستظهر في بداية النموذج",
    icon: UserRound,
    fields: [
      { key: "insuredName", label: "اسم المؤمن عليه", placeholder: "اكتب الاسم رباعيًا" },
      { key: "nationalId", label: "الرقم القومي", placeholder: "14 رقمًا", type: "number" },
      { key: "insuranceNumber", label: "الرقم التأميني", placeholder: "رقم الملف التأميني", type: "number" },
      { key: "birthDate", label: "تاريخ الميلاد", placeholder: "يوم / شهر / سنة", type: "date" },
      { key: "gender", label: "النوع", placeholder: "ذكر / أنثى" },
      { key: "category", label: "الفئة", placeholder: "اختر الفئة" },
    ],
  },
  {
    id: "work",
    number: "02",
    title: "بيانات العمل والاشتراك",
    note: "الحقل الذي يربط الشخص بالمنشأة والتغطية التأمينية",
    icon: BriefcaseBusiness,
    fields: [
      { key: "establishmentName", label: "اسم المنشأة", placeholder: "اسم الشركة أو المنشأة" },
      { key: "establishmentNumber", label: "رقم المنشأة", placeholder: "رقم المنشأة", type: "number" },
      { key: "establishmentType", label: "نوع المنشأة", placeholder: "اختر نوع المنشأة" },
      { key: "office", label: "المكتب", placeholder: "مكتب التأمينات" },
      { key: "profession", label: "المهنة", placeholder: "المسمى الوظيفي" },
      { key: "professionCode", label: "كود المهنة", placeholder: "الكود" },
      { key: "sector", label: "القطاع", placeholder: "القطاع التابع له" },
      { key: "qualification", label: "المؤهل", placeholder: "المؤهل الدراسي" },
      { key: "startDate", label: "تاريخ بدء الاشتراك", placeholder: "تاريخ البدء", type: "date" },
      { key: "contributionCode", label: "كود الاشتراك", placeholder: "كود الاشتراك" },
      { key: "workType", label: "نوع المدة", placeholder: "نوع المدة" },
      { key: "medicalExam", label: "الكشف الطبي الابتدائي", placeholder: "نعم / لا" },
    ],
  },
  {
    id: "wage",
    number: "03",
    title: "الأجر والعجز",
    note: "الأرقام تُحفظ كما أدخلتها وتُنسخ إلى خانات النموذج",
    icon: Hash,
    fields: [
      { key: "basicWage", label: "أجر / دخل الاشتراك", placeholder: "0.00", type: "number" },
      { key: "variableWage", label: "الأجر المتغير", placeholder: "0.00", type: "number" },
      { key: "totalWage", label: "الأجر الشامل", placeholder: "0.00", type: "number" },
      { key: "increaseDate", label: "تاريخ بداية العجز (اختياري)", placeholder: "اتركه فارغًا إن لم يوجد", type: "date" },
      { key: "increasePercent", label: "نسبة العجز (اختياري)", placeholder: "اتركها فارغة إن لم توجد", type: "number" },
    ],
  },
  {
    id: "address",
    number: "04",
    title: "العنوان والتواصل",
    note: "بيانات محل الإقامة وجهة العمل والتواصل",
    icon: Building2,
    fields: [
      { key: "country", label: "الجنسية", placeholder: "مثال: مصري" },
      { key: "city", label: "المدينة", placeholder: "المدينة" },
      { key: "governorate", label: "المحافظة", placeholder: "المحافظة" },
      { key: "buildingNumber", label: "رقم العقار", placeholder: "رقم العقار", type: "number" },
      { key: "district", label: "الشياخة / القرية (اختياري)", placeholder: "اتركها فارغة إن لم توجد" },
      { key: "street", label: "الشارع", placeholder: "اسم الشارع" },
      { key: "center", label: "القسم / المركز", placeholder: "القسم أو المركز" },
      { key: "phone", label: "التليفون", placeholder: "رقم الهاتف", type: "number" },
      { key: "email", label: "البريد الإلكتروني", placeholder: "name@company.com", type: "email" },
      { key: "employer", label: "جهة العمل", placeholder: "جهة العمل" },
      { key: "manager", label: "المدير المسؤول", placeholder: "اسم المسؤول" },
      { key: "releaseDate", label: "تحريرًا في", placeholder: "تاريخ التحرير", type: "date" },
    ],
    },
  {
    id: "s6",
    number: "05",
    title: "بيانات مقدم الطلب",
    note: "بيانات مقدم الطلب وحقول انتهاء الاشتراك في نموذج س6",
    icon: FileText,
    fields: [
      { key: "applicantName", label: "مقدم الطلب", placeholder: "اسم مقدم الطلب" },
      { key: "applicantRole", label: "صفة مقدم الطلب", placeholder: "صاحب العمل / المسؤول" },
      { key: "applicantInsuranceNumber", label: "الرقم التأميني لمقدم الطلب", placeholder: "الرقم التأميني", type: "number" },
      { key: "applicantPhone", label: "رقم تليفون مقدم الطلب", placeholder: "رقم الهاتف", type: "number" },
      { key: "applicantNationalId", label: "الرقم القومي لمقدم الطلب", placeholder: "الرقم القومي", type: "number" },
      { key: "endDate", label: "تاريخ انتهاء الاشتراك", placeholder: "تاريخ الانتهاء", type: "date" },
      { key: "endReason", label: "سبب انتهاء الاشتراك", placeholder: "سبب الانتهاء" },
      { key: "address", label: "العنوان", placeholder: "العنوان بالتفصيل", wide: true },
    ],
  },
];

// Only show fields that have a verified destination in the selected official PDF.
// This prevents users from entering data that would never appear in the export.
const TEMPLATE_FIELDS: Record<TemplateId, Set<EditableKey>> = {
  s1: new Set([
    "insuredName", "nationalId", "insuranceNumber", "category", "country", "establishmentName",
    "establishmentNumber", "establishmentType", "office", "profession", "sector", "qualification", "startDate", "contributionCode",
    "workType", "medicalExam", "basicWage", "totalWage", "increaseDate", "increasePercent", "governorate",
    "buildingNumber", "district", "street", "center", "phone", "address", "applicantName", "applicantRole", "applicantInsuranceNumber", "applicantPhone", "applicantNationalId",
  ]),
  s6: new Set([
    "insuredName", "nationalId", "insuranceNumber", "establishmentName", "establishmentNumber", "office",
    "applicantName", "applicantRole", "applicantInsuranceNumber", "applicantPhone", "applicantNationalId", "endDate", "endReason", "address",
  ]),
};
const IDENTIFIER_FIELDS = new Set<EditableKey>([
  "nationalId", "applicantNationalId", "insuranceNumber", "applicantInsuranceNumber", "establishmentNumber",
  "professionCode", "contributionCode", "phone", "applicantPhone", "buildingNumber",
]);
const SHARED_EXCEL_FIELDS = new Set<EditableKey>([
  "office", "establishmentName", "establishmentNumber", "establishmentType", "address",
  "applicantName", "applicantRole", "applicantInsuranceNumber", "applicantPhone", "applicantNationalId",
]);
const SELECT_OPTIONS: Partial<Record<EditableKey, string[]>> = {
  applicantRole: ["صاحب العمل", "المسؤول", "مفوض"],
  category: ["عاملين لدى الغير", "أصحاب أعمال", "عمالة غير منتظمة"],
  medicalExam: ["نعم", "لا"],
  establishmentType: ["نمطي", "سيارة", "مركب صيد", "مخابز بلدية"],
  workType: ["دائمة", "مؤقتة", "موسمية"],
  sector: ["حكومي", "عام / أعمال عام", "خاص"],
  gender: ["ذكر", "أنثى"],
};
const EXCEL_LIST_OPTIONS: Partial<Record<EditableKey, string[]>> = {
  applicantRole: ["صاحب العمل", "المسؤول", "مفوض"],
  establishmentType: ["نمطي", "سيارة", "مركب صيد", "مخابز بلدية"],
  category: ["عاملين لدى الغير", "أصحاب أعمال", "عمالة غير منتظمة"],
  medicalExam: ["نعم", "لا"],
  workType: ["دائمة", "مؤقتة", "موسمية"],
  sector: ["حكومي", "عام / أعمال عام", "خاص"],
  gender: ["ذكر", "أنثى"],
};
function excelColumn(index: number) {
  let value = index + 1;
  let result = "";
  while (value) { value -= 1; result = String.fromCharCode(65 + (value % 26)) + result; value = Math.floor(value / 26); }
  return result;
}
function xmlEscape(value: string) {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
async function addExcelValidations(data: ArrayBuffer, sheets: Array<{ path: string; headers: typeof EXCEL_HEADERS }>) {
  const zip = await JSZip.loadAsync(data);
  for (const sheet of sheets) {
    const entry = zip.file(sheet.path);
    if (!entry) continue;
    let xml = await entry.async("string");
    const rules: string[] = [];
    sheet.headers.forEach((header, index) => {
      const column = excelColumn(index);
      const options = EXCEL_LIST_OPTIONS[header.key as EditableKey];
      if (options) rules.push(`<dataValidation type="list" allowBlank="1" showErrorMessage="1" errorTitle="اختيار غير صحيح" error="اختر قيمة من القائمة" sqref="${column}2:${column}1000"><formula1>&quot;${xmlEscape(options.join(","))}&quot;</formula1></dataValidation>`);
      const length = header.key === "nationalId" || header.key === "applicantNationalId" ? 14 : header.key === "phone" || header.key === "applicantPhone" ? 11 : 0;
      if (length) rules.push(`<dataValidation type="textLength" operator="equal" allowBlank="1" showErrorMessage="1" errorTitle="عدد أرقام غير صحيح" error="يجب إدخال ${length} رقمًا" sqref="${column}2:${column}1000"><formula1>${length}</formula1></dataValidation>`);
    });
    if (rules.length) {
      const validations = `<dataValidations count="${rules.length}">${rules.join("")}</dataValidations>`;
      // OOXML requires dataValidations before print/page settings,
      // ignoredErrors and other trailing worksheet elements. Appending it at
      // the very end makes Excel repair the sheet and discard its contents.
      const trailingElement = /<(hyperlinks|printOptions|pageMargins|pageSetup|headerFooter|rowBreaks|colBreaks|customProperties|cellWatches|ignoredErrors|smartTags|drawing|legacyDrawing|picture|oleObjects|controls|webPublishItems|tableParts|extLst)\b/;
      const match = trailingElement.exec(xml);
      xml = match ? `${xml.slice(0, match.index)}${validations}${xml.slice(match.index)}` : xml.replace("</worksheet>", `${validations}</worksheet>`);
    }
    zip.file(sheet.path, xml);
  }
  return zip.generateAsync({ type: "arraybuffer", compression: "DEFLATE" });
}
function initials(value: string) {
  return value
    .trim()
    .split(/\s+/)
    .slice(0, 2)
    .map((part) => part[0])
    .join("") || "—";
}

function statusClass(status: ReturnType<typeof recordStatus>) {
  if (status === "جاهز") return "status-ready";
  if (status === "ناقص") return "status-warning";
  return "status-draft";
}

export default function Home() {
  const [records, setRecords] = useState<PersonRecord[]>(() => [makeEmptyRecord()]);
  const [activeId, setActiveId] = useState(() => records[0].id);
  const [mode, setMode] = useState<Mode>("manual");
  const [template, setTemplate] = useState<TemplateId>("s1");
  const [activeGroup, setActiveGroup] = useState("identity");
  const [search, setSearch] = useState("");
  const [fileName, setFileName] = useState("");
  const [isProcessing, setIsProcessing] = useState(false);
    const [sidebarOpen, setSidebarOpen] = useState(false);
  const [previewImage, setPreviewImage] = useState("");
  const [previewState, setPreviewState] = useState<"loading" | "ready" | "error">("loading");
  const fileInputRef = useRef<HTMLInputElement>(null);
  const activeRecord = records.find((record) => record.id === activeId) ?? records[0];
  useEffect(() => {
    let cancelled = false;
    const timer = window.setTimeout(async () => {
      setPreviewState("loading");
      try {
        const data = await fillPdf(activeRecord, template);
        if (cancelled) return;
        const pdf = await getDocument({ data: new Uint8Array(data) }).promise;
        const page = await pdf.getPage(1);
        const baseViewport = page.getViewport({ scale: 1 });
        const targetWidth = 460;
        const viewport = page.getViewport({ scale: targetWidth / baseViewport.width });
        const canvas = document.createElement("canvas");
        canvas.width = Math.ceil(viewport.width);
        canvas.height = Math.ceil(viewport.height);
        const context = canvas.getContext("2d");
        if (!context) throw new Error("تعذر تجهيز سطح المعاينة");
        await page.render({ canvas, canvasContext: context, viewport }).promise;
        if (cancelled) return;
        setPreviewImage(canvas.toDataURL("image/png"));
        setPreviewState("ready");
      } catch (error) {
        console.error("PDF preview generation failed", error);
        if (!cancelled) setPreviewState("error");
      }
    }, 250);
    return () => { cancelled = true; window.clearTimeout(timer); };
  }, [activeRecord, template]);
  const readyCount = records.filter((record) => recordStatus(record, template) === "جاهز").length;
  const draftCount = records.filter((record) => recordStatus(record, template) !== "جاهز").length;
  const activeIssues = useMemo(() => validateRecord(activeRecord, template), [activeRecord, template]);
  const issueByField = useMemo(() => new Map(activeIssues.map((issue) => [issue.key, issue.message])), [activeIssues]);
  const availableGroups = useMemo(() => fieldGroups
    .map((group) => ({ ...group, fields: group.fields.filter((field) => TEMPLATE_FIELDS[template].has(field.key)) }))
    .filter((group) => group.fields.length > 0), [template]);
  const visibleRecords = useMemo(() => {
    const query = search.trim().toLocaleLowerCase("ar-EG");
    if (!query) return records;
    return records.filter((record) =>
      [record.insuredName, record.nationalId, record.establishmentName]
        .join(" ")
        .toLocaleLowerCase("ar-EG")
        .includes(query),
    );
  }, [records, search]);
  function changeTemplate(next: TemplateId) {
    setTemplate(next);
    setActiveGroup("identity");
    toast.success(`تم اختيار ${TEMPLATE_LABELS[next]}`, { description: "سيُستخدم هذا القالب عند المعاينة والتصدير." });
  }

  function updateField(key: EditableKey, value: string) {
    const nextValue = IDENTIFIER_FIELDS.has(key) ? normalizeDigits(value).replace(/\D/g, "") : value;
    setRecords((current) =>
      current.map((record) => (record.id === activeId ? { ...record, [key]: nextValue } : record)),
    );
  }

  function addRecord() {
    const next = makeEmptyRecord();
    setRecords((current) => [...current, next]);
    setActiveId(next.id);
    setMode("manual");
    toast.success("أضيف سجل جديد", { description: "يمكنك بدء تعبئة بيانات المؤمن عليه الآن." });
  }

  function removeRecord(id: string) {
    if (records.length === 1) {
      toast.error("لا يمكن حذف السجل الوحيد", { description: "أضف سجلًا آخر قبل الحذف." });
      return;
    }
    const next = records.filter((record) => record.id !== id);
    setRecords(next);
    if (id === activeId) setActiveId(next[0].id);
    toast.success("تم حذف السجل");
  }

  async function handleExcel(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;
    setFileName(file.name);
    setIsProcessing(true);
    try {
      const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true });
      const peopleSheet = workbook.Sheets["بيانات المؤمن عليهم"] ?? workbook.Sheets[workbook.SheetNames.at(-1)!];
      const sharedSheet = workbook.Sheets["بيانات المنشأة والمفوض"];
      const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(peopleSheet, { defval: "" });
      const sharedRow = sharedSheet ? XLSX.utils.sheet_to_json<Record<string, unknown>>(sharedSheet, { defval: "" })[0] : undefined;
      const sharedRecord = sharedRow ? mapExcelRow(sharedRow) : undefined;
      const imported = rows.map((row) => {
        const record = mapExcelRow(row);
        if (sharedRecord) for (const key of SHARED_EXCEL_FIELDS) if (!record[key] && sharedRecord[key]) record[key] = sharedRecord[key];
        return record;
      }).filter((record) => filledCount(record) > 0);
      if (!imported.length) {
        toast.error("لم نجد صفوفًا قابلة للاستيراد", { description: "تأكد من أن الصف الأول يحتوي على أسماء الأعمدة." });
        return;
      }
      setRecords(imported);
      setActiveId(imported[0].id);
      setMode("bulk");
      const invalidCount = imported.filter((record) => validateRecord(record, template).length > 0).length;
      toast.success(`تم استيراد ${imported.length} سجل`, { description: invalidCount ? `${invalidCount} سجل يحتاج مراجعة قبل التصدير.` : "كل السجلات جاهزة للتصدير." });
    } catch {
      toast.error("تعذر قراءة ملف Excel", { description: "استخدم ملف .xlsx أو .xls سليمًا." });
    } finally {
      setIsProcessing(false);
      event.target.value = "";
    }
  }

  async function downloadTemplate() {
    const sharedHeaders = EXCEL_HEADERS.filter((item) => item.key !== "id" && SHARED_EXCEL_FIELDS.has(item.key as EditableKey));
    const personHeaders = EXCEL_HEADERS.filter((item) => item.key !== "id" && TEMPLATE_FIELDS[template].has(item.key as EditableKey) && !SHARED_EXCEL_FIELDS.has(item.key as EditableKey));
    const sharedSheet = XLSX.utils.aoa_to_sheet([
      sharedHeaders.map((item) => item.label),
      sharedHeaders.map(() => ""),
    ]);
    const peopleSheet = XLSX.utils.aoa_to_sheet([
      personHeaders.map((item) => item.label),
      personHeaders.map(() => ""),
    ]);
    sharedSheet["!cols"] = sharedHeaders.map(() => ({ wch: 24 }));
    peopleSheet["!cols"] = personHeaders.map(() => ({ wch: 22 }));
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, sharedSheet, "بيانات المنشأة والمفوض");
    XLSX.utils.book_append_sheet(workbook, peopleSheet, "بيانات المؤمن عليهم");
    const rawData = XLSX.write(workbook, { bookType: "xlsx", type: "array" }) as ArrayBuffer;
    const data = await addExcelValidations(rawData, [
      { path: "xl/worksheets/sheet1.xml", headers: sharedHeaders },
      { path: "xl/worksheets/sheet2.xml", headers: personHeaders },
    ]);
    downloadBlob(data, "قالب-بيانات-التأمينات.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    toast.success("تم تنزيل قالب Excel");
  }

  async function downloadSelected() {
    const issues = validateRecord(activeRecord, template);
    if (issues.length) {
      toast.error("السجل يحتاج مراجعة", { description: issues.slice(0, 3).map((issue) => issue.message).join("، ") });
      return;
    }
    setIsProcessing(true);
    try {
      const pdf = await fillPdf(activeRecord, template);
      downloadBlob(pdf, safeFileName(activeRecord, 1, template), "application/pdf");
      toast.success("تم تجهيز النموذج", { description: "الملف جاهز للطباعة أو المراجعة." });
    } catch (error) {
      toast.error("تعذر تجهيز PDF", { description: error instanceof Error ? error.message : "حاول مرة أخرى." });
    } finally {
      setIsProcessing(false);
    }
  }

  function validBatch() {
    return records.filter((record) => recordStatus(record, template) === "جاهز");
  }

  async function downloadAll() {
    const validRecords = validBatch();
    if (!validRecords.length) {
      toast.error("لا توجد سجلات مكتملة", { description: "أكمل الاسم والرقم القومي واسم المنشأة أولًا." });
      return;
    }
    setIsProcessing(true);
    try {
      const pdf = await createMergedPdf(validRecords, template);
      downloadBlob(pdf, `نماذج-التأمينات-مجمعة-${new Date().toISOString().slice(0, 10)}.pdf`, "application/pdf");
      const skipped = records.length - validRecords.length;
      toast.success(`تم دمج ${validRecords.length} نموذجًا`, { description: skipped ? `تم استبعاد ${skipped} سجل ناقص.` : "كل النماذج مرتبة داخل PDF واحد." });
    } catch (error) {
      toast.error("تعذر تجهيز الدفعة", { description: error instanceof Error ? error.message : "حاول مرة أخرى." });
    } finally {
      setIsProcessing(false);
    }
  }

  async function downloadSeparateZip() {
    const validRecords = validBatch();
    if (!validRecords.length) { toast.error("لا توجد سجلات مكتملة"); return; }
    setIsProcessing(true);
    try {
      const zip = await createZip(validRecords, template);
      downloadBlob(zip, `نماذج-التأمينات-منفصلة-${new Date().toISOString().slice(0, 10)}.zip`, "application/zip");
      toast.success(`تم تجهيز ${validRecords.length} ملف PDF منفصل`);
    } catch (error) {
      toast.error("تعذر تجهيز الملفات المنفصلة", { description: error instanceof Error ? error.message : "حاول مرة أخرى." });
    } finally { setIsProcessing(false); }
  }

  return (
    <div className="app-shell" dir="rtl">
      <header className="topbar">
        <div className="brand-lockup">
          <button className="mobile-menu" onClick={() => setSidebarOpen((open) => !open)} aria-label="فتح القائمة">
            <Menu size={18} />
          </button>
          <img src={MARK_URL} alt="" className="brand-mark" />
          <div>
            <p className="brand-name">معبّي</p>
            <p className="brand-caption">مكتب النماذج الذكي</p>
          </div>
        </div>
        <div className="topbar-actions">
          <div className="connection-state"><span className="state-dot" /> يعمل محليًا على جهازك</div>
          <button className="icon-button" aria-label="المساعدة"><CircleHelp size={18} /></button>
          <button className="avatar-button" aria-label="الحساب">م</button>
        </div>
      </header>

      <div className="workspace">
        <aside className={`sidebar ${sidebarOpen ? "sidebar-open" : ""}`}>
          <div className="sidebar-inner">
            <div className="sidebar-kicker">مساحة العمل</div>
            <nav className="main-nav" aria-label="التنقل الرئيسي">
              <button className="nav-item nav-item-active"><LayoutDashboard size={17} /> لوحة النماذج <span className="nav-count">{records.length}</span></button>
              <button className="nav-item"><FolderOpen size={17} /> أرشيف الملفات <span className="nav-soon">قريبًا</span></button>
              <button className="nav-item"><Settings2 size={17} /> إعدادات القالب <span className="nav-soon">قريبًا</span></button>
            </nav>
            <div className="sidebar-rule" />
            <div className="sidebar-note">
              <div className="note-seal"><ShieldCheck size={16} /></div>
              <div>
                <p>بياناتك تبقى هنا</p>
                <span>المعالجة تتم داخل المتصفح ولا نرفع الملفات إلى خادم.</span>
              </div>
            </div>
            <div className="sidebar-bottom">
              <p className="template-label">القالب الحالي</p>
              <div className="template-switcher" aria-label="اختيار القالب">
                {(Object.keys(TEMPLATE_LABELS) as TemplateId[]).map((id) => <button key={id} className={template === id ? "template-option template-option-active" : "template-option"} onClick={() => changeTemplate(id)}><FileText size={14} /><span>{TEMPLATE_LABELS[id]}</span>{template === id && <BadgeCheck size={14} />}</button>)}
              </div>
            </div>
          </div>
        </aside>

        <main className="main-content">
          <section className="page-intro">
            <div className="intro-copy">
              <div className="breadcrumb"><span>لوحة النماذج</span><ArrowLeft size={13} /><strong>{TEMPLATE_LABELS[template]}</strong></div>
              <p className="eyebrow"><span className="eyebrow-line" /> {template === "s6" ? "إخطار انتهاء اشتراك مؤمن عليه" : "نموذج الاشتراك المؤمن عليه"}</p>
              <h1>من جدول واحد<br /><em>إلى نماذج جاهزة.</em></h1>
              <p className="intro-description">اكتب بيانات شخص واحد أو ارفع ملف Excel كاملًا. راجع كل سجل، ثم احصل على ملفات PDF مرتبة وجاهزة للطباعة.</p>
              <div className="intro-actions">
                <Button className="button-primary" onClick={downloadAll} disabled={isProcessing}>
                  {isProcessing ? <Loader2 className="spin" size={16} /> : <Printer size={16} />}
                  تنزيل PDF مجمّع
                </Button>
                <Button variant="outline" className="button-outline" onClick={downloadSeparateZip} disabled={isProcessing}><Download size={16} /> ZIP ملفات منفصلة</Button>
                <Button variant="outline" className="button-outline" onClick={downloadTemplate}><Download size={16} /> قالب Excel</Button>
              </div>
            </div>
            <div className="intro-visual">
              <img className="visual-paper" src={PAPER_URL} alt="" />
              <div className="hero-form-artifact"><img src={template === "s6" ? S6_PREVIEW_URL : FORM_URL} alt="" /><span>{TEMPLATE_LABELS[template]}</span></div>
              <div className="visual-overlay"><span>{template === "s6" ? "س6" : "س1"}</span><b>{template === "s6" ? <>إخطار<br />اشتراك</> : <>وثيقة<br />تأمين</>}</b><small>جاهز للطباعة</small></div>
              <div className="visual-ribbon"><FileSpreadsheet size={17} /><span>Excel → PDF</span></div>
              <div className="visual-stamp"><Check size={16} /> جاهز</div>
            </div>
          </section>

          <section className="summary-strip" aria-label="ملخص السجلات">
            <span className="summary-spine" aria-hidden="true" />
            <div className="summary-item"><span className="summary-icon amber"><ClipboardList size={17} /></span><div><strong>{records.length}</strong><span>إجمالي السجلات</span></div></div>
            <div className="summary-item"><span className="summary-icon green"><BadgeCheck size={17} /></span><div><strong>{readyCount}</strong><span>جاهز للطباعة</span></div></div>
            <div className="summary-item"><span className="summary-icon gray"><MoreHorizontal size={17} /></span><div><strong>{draftCount}</strong><span>يحتاج مراجعة</span></div></div>
            <div className="summary-divider" />
            <div className="summary-hint"><span className="hint-pin" /> آخر تعديل محفوظ في هذه الجلسة</div>
          </section>

          <section className="workbench-section registry-section">
            <div className="section-heading-row">
              <div><p className="section-overline">مسار العمل <span>01 / 03</span></p><h2>أدخل البيانات</h2></div>
              <div className="heading-tools"><span className="template-pill"><FileText size={14} /> {TEMPLATE_LABELS[template]} <ChevronDown size={14} /></span><span className="autosave"><span /> حفظ تلقائي</span></div>
            </div>

            <Card className="mode-card registry-card">
              <div className="mode-copy"><div className="mode-icon"><ClipboardList size={18} /></div><div><h3>كيف تريد البدء؟</h3><p>اختر إدخال سجل يدويًا أو حمّل دفعة من Excel.</p></div></div>
              <Tabs value={mode} onValueChange={(value) => setMode(value as Mode)} className="mode-tabs">
                <TabsList><TabsTrigger value="manual"><UserRound size={15} /> سجل يدوي</TabsTrigger><TabsTrigger value="bulk"><UsersRound size={15} /> دفعة Excel</TabsTrigger></TabsList>
              </Tabs>
            </Card>

            {mode === "bulk" ? (
              <Card className="upload-card registry-card">
                <div className="upload-art"><img src={RIBBON_URL} alt="" /><div className="upload-art-label"><FileSpreadsheet size={17} /><span>صفوف منظمة<br /><b>إلى ملفات</b></span></div></div>
                <div className="upload-copy"><div className="upload-title-row"><div><h3>ارفع ملف البيانات</h3><p>نقرأ بيانات المنشأة والمفوّض مرة واحدة، ثم ننشئ نموذجًا لكل صف في شيت المؤمن عليهم.</p></div><span className="supported-formats">.XLSX / .XLS</span></div>
                  <div className="upload-actions"><Button className="button-primary" onClick={() => fileInputRef.current?.click()} disabled={isProcessing}>{isProcessing ? <Loader2 className="spin" size={16} /> : <CloudUpload size={16} />} اختر ملف Excel</Button><button className="text-link" onClick={downloadTemplate}>نزّل قالب الأعمدة <ArrowLeft size={14} /></button></div>
                  {fileName && <div className="uploaded-file"><FileSpreadsheet size={16} /><span>{fileName}</span><Check size={15} /></div>}
                  <input ref={fileInputRef} type="file" accept=".xlsx,.xls" onChange={handleExcel} aria-label="رفع ملف Excel" className="file-input-accessible" />
                </div>
              </Card>
            ) : (
              <div className="manual-callout"><span className="callout-number">A</span><div><strong>ابدأ بسجل واحد</strong><p>اكتب الحقول الأساسية، وستظهر المعاينة الورقية على اليسار لحظة بلحظة.</p></div><ArrowLeft size={16} /></div>
            )}

            <div className="records-toolbar">
              <div className="records-title"><UsersRound size={17} /><strong>سجل البيانات</strong><span>{records.length} صف</span></div>
              <div className="records-actions"><div className="search-wrap"><Search size={15} /><input value={search} onChange={(event) => setSearch(event.target.value)} placeholder="ابحث بالاسم أو الرقم..." /></div><Button variant="outline" className="add-button" onClick={addRecord}><Plus size={15} /> إضافة سجل</Button></div>
            </div>

            <div className="records-table-wrap registry-table">
              <table className="records-table">
                <thead><tr><th className="check-cell"><input type="checkbox" aria-label="تحديد الكل" /></th><th>المؤمن عليه</th><th>الرقم القومي</th><th>اسم المنشأة</th><th>اكتمال البيانات</th><th>الحالة</th><th /></tr></thead>
                <tbody>
                  {visibleRecords.map((record, index) => {
                    const status = recordStatus(record, template);
                    const isActive = record.id === activeId;
                    return <tr key={record.id} className={isActive ? "record-row-active" : ""} onClick={() => setActiveId(record.id)}>
                      <td className="check-cell"><input type="checkbox" checked={isActive} onChange={() => setActiveId(record.id)} onClick={(event) => event.stopPropagation()} aria-label={`تحديد ${record.insuredName || `السجل ${index + 1}`}`} /></td>
                      <td><div className="person-cell"><span className="person-avatar">{initials(record.insuredName)}</span><div><strong>{record.insuredName || "سجل جديد"}</strong><small>{record.insuranceNumber || "لم يُدخل الرقم التأميني"}</small></div></div></td>
                      <td className="mono-cell">{record.nationalId || "—"}</td><td>{record.establishmentName || "—"}</td>
                      <td><div className="progress-cell"><div className="progress-track"><span style={{ width: `${Math.min(100, Math.round((filledCount(record) / relevantFieldCount(template)) * 100))}%` }} /></div><small>{Math.min(filledCount(record), relevantFieldCount(template))} / {relevantFieldCount(template)}</small></div></td>
                      <td><span className={`status-badge ${statusClass(status)}`}><span />{status}</span></td>
                      <td><button className="row-menu" onClick={(event) => { event.stopPropagation(); removeRecord(record.id); }} aria-label="حذف السجل"><Trash2 size={15} /></button></td>
                    </tr>;
                  })}
                  {!visibleRecords.length && <tr><td colSpan={7} className="empty-table">لا توجد نتائج بهذا البحث.</td></tr>}
                </tbody>
              </table>
            </div>
          </section>

          <section className="editor-layout">
            <div className="editor-panel">
              <div className="editor-header"><div><p className="section-overline">السجل المحدد</p><h2>{activeRecord.insuredName || "سجل جديد"}</h2></div><Badge className={`status-badge ${statusClass(recordStatus(activeRecord, template))}`}><span />{recordStatus(activeRecord, template)}</Badge></div>
              <div className="group-tabs" role="tablist" aria-label="أقسام النموذج">
                {availableGroups.map((group) => { const Icon = group.icon; return <button key={group.id} className={activeGroup === group.id ? "group-tab-active" : ""} onClick={() => setActiveGroup(group.id)}><span className="group-tab-number">{group.number}</span><Icon size={16} /><span>{group.title}</span></button>; })}
              </div>
              <div className="field-groups">
                {availableGroups.map((group) => {
                  const Icon = group.icon;
                  const visible = activeGroup === group.id;
                  return <div key={group.id} className={`field-group ${visible ? "field-group-visible" : ""}`}>
                    <div className="field-group-heading"><span className="field-group-icon"><Icon size={17} /></span><div><h3>{group.title}</h3><p>{group.note}</p></div></div>
                    <div className="fields-grid">
                      {group.fields.map((field) => { const issue = issueByField.get(field.key); const isIdentifier = IDENTIFIER_FIELDS.has(field.key); const options = SELECT_OPTIONS[field.key]; const maxLength = field.key === "nationalId" || field.key === "applicantNationalId" ? 14 : field.key === "phone" || field.key === "applicantPhone" ? 11 : undefined; return <div className={`field-shell ${field.wide ? "field-wide" : ""}`} key={field.key}><Label htmlFor={field.key}>{field.label}</Label>{options ? <select id={field.key} value={activeRecord[field.key]} onChange={(event) => updateField(field.key, event.target.value)} aria-invalid={Boolean(issue)}><option value="">{field.placeholder}</option>{options.map((option) => <option key={option}>{option}</option>)}</select> : <Input id={field.key} type={isIdentifier ? "text" : field.type || "text"} inputMode={isIdentifier ? "numeric" : undefined} maxLength={maxLength} value={activeRecord[field.key]} onChange={(event) => updateField(field.key, event.target.value)} placeholder={field.placeholder} dir={field.key === "email" || isIdentifier || field.type === "number" || field.type === "date" ? "ltr" : "rtl"} aria-invalid={Boolean(issue)} />}{issue && activeRecord[field.key] && <span className="field-error">{issue}</span>}</div>; })}
                    </div>
                  </div>;
                })}
              </div>
              <div className="editor-footer"><span><ShieldCheck size={15} /> المعالجة محلية وآمنة</span><div><Button variant="outline" className="button-outline" onClick={() => setRecords((current) => current.map((record) => record.id === activeId ? { ...makeEmptyRecord(), id: record.id } : record))}>مسح الحقول</Button><Button className="button-primary" onClick={downloadSelected} disabled={isProcessing}>{isProcessing ? <Loader2 className="spin" size={16} /> : <ArrowDownToLine size={16} />} تنزيل PDF</Button></div></div>
            </div>

            <aside className="preview-panel">
              <div className="preview-header"><div><p className="section-overline">المعاينة الحية</p><h2>صفحة النموذج</h2></div><span className="preview-page">1 / 2</span></div>
              <div className="preview-frame"><div className="preview-paper preview-pdf-paper">{previewImage ? <img className="preview-rendered-image" src={previewImage} alt={`معاينة ${TEMPLATE_LABELS[template]} بالبيانات الحالية`} /> : <div className="preview-loading">{previewState === "error" ? "تعذر إنشاء المعاينة" : "جاري تجهيز النموذج..."}</div>}</div><div className="preview-control"><button aria-label="تصغير المعاينة"><ChevronDown size={14} /></button><span>100%</span><button aria-label="تكبير المعاينة"><Plus size={14} /></button></div></div>
              <div className="preview-note"><div className="note-seal small"><Check size={14} /></div><div><strong>معاينة القالب الأصلي</strong><p>تُعرض الصفحة الأولى من PDF الفعلي بالبيانات الحالية، ويمكن تنزيل الملف كاملًا من الزر أدناه.</p></div></div>
              <Button className="full-preview-button" variant="outline" onClick={downloadSelected}><Printer size={15} /> معاينة / تنزيل الصفحة</Button>
            </aside>
          </section>

          <footer className="app-footer"><span>معبّي © 2026</span><span>{TEMPLATE_LABELS[template]} · النسخة الأولى</span><span>مصمم لتسهيل العمل المكتبي</span></footer>
        </main>
      </div>
    </div>
  );
}

