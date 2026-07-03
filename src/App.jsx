import React, { useEffect, useMemo, useRef, useState } from "react";
import ExcelJS from "exceljs";
import {
  LayoutDashboard,
  FolderKanban,
  ListChecks,
  Trophy,
  Users,
  Building2,
  FileText,
  Search,
  Filter,
  Plus,
  ChevronDown,
  ChevronRight,
  CalendarDays,
  Dumbbell,
  Goal,
  Waves,
  PersonStanding,
  Volleyball,
  CheckCircle2,
  AlertTriangle,
  Library,
  Clock3,
  BarChart3,
  Edit3,
  Trash2,
  Target,
  Bell,
  Mail,
  Megaphone,
  Send,
  Upload,
  Download,
  FileSpreadsheet,
  PartyPopper,
  MoreHorizontal,
  Copy,
  X,
  EyeOff,
  Layers,
} from "lucide-react";

const STORAGE_KEY = "jam_saas_v19_clean";

const BASE_SPORTS = [
  { id: "football", name: "Football", icon: Goal },
  { id: "athletisme", name: "Athlétisme", icon: PersonStanding },
  { id: "developpe_couche", name: "Développé couché", icon: Dumbbell },
  { id: "natation", name: "Natation", icon: Waves },
  { id: "volley", name: "Volley", icon: Volleyball },
];

const STATUS_OPTIONS = ["À faire", "En cours", "En attente", "Terminé"];
const PRIORITY_OPTIONS = ["Basse", "Moyenne", "Haute", "Critique"];
const DOC_MODES = ["all", "internal", "external"];
const DOC_STATUS = ["Brouillon", "À valider", "Validé", "Publié", "Archivé"];
const ATHLETE_STATUS = ["Actif", "À relancer", "Blessé", "Absent"];
const ATHLETE_PROFILE_OPTIONS = ["Homme", "Femme", "Junior"];
const CLUB_STATUS = ["À contacter", "Invité", "Inscrit", "Validé"];
const COMM_STATUS = ["Brouillon", "Prêt à envoyer", "Envoyée"];
const CUSTOM_FIELD_TYPES = ["texte", "nombre", "pourcentage", "date", "financier", "statut", "case_a_cocher", "personne"];
const DASHBOARD_WIDGET_TYPES = [
  { id: "number", label: "Nombre" },
  { id: "bar", label: "Barres" },
  { id: "donut", label: "Donut" },
];
const DASHBOARD_SOURCES = [
  { id: "tasks", label: "Tâches" },
  { id: "projects", label: "Projets" },
  { id: "athletes", label: "Athlètes" },
  { id: "clubs", label: "Sections" },
];
const DASHBOARD_FILTERS = [
  { id: "all", label: "Tous" },
  { id: "completed", label: "Terminés" },
  { id: "open", label: "Non terminés" },
  { id: "overdue", label: "En retard" },
  { id: "upcoming", label: "À venir 14 jours" },
  { id: "highPriority", label: "Haute priorité" },
  { id: "unassigned", label: "Non assignés" },
  { id: "incomplete", label: "À compléter" },
  { id: "budget", label: "Budget" },
];
const DASHBOARD_GROUPS = [
  { id: "status", label: "Statut" },
  { id: "priority", label: "Priorité" },
  { id: "section", label: "Section" },
  { id: "profile", label: "Profil" },
  { id: "assignee", label: "Responsable" },
];
const DEFAULT_DASHBOARD_WIDGETS = [
  { id: "dw_completed_tasks", type: "number", title: "Tâches terminées", source: "tasks", filter: "completed", tone: "green" },
  { id: "dw_open_tasks", type: "number", title: "Tâches non terminées", source: "tasks", filter: "open", tone: "slate" },
  { id: "dw_overdue_tasks", type: "number", title: "Tâches en retard", source: "tasks", filter: "overdue", tone: "red" },
  { id: "dw_total_budget", type: "number", title: "Budget total", source: "projects", filter: "budget", tone: "green" },
  { id: "dw_tasks_status_bar", type: "bar", title: "Tâches par statut", source: "tasks", filter: "all", groupBy: "status", tone: "violet", size: "wide" },
  { id: "dw_athletes_section_bar", type: "bar", title: "Athlètes par section", source: "athletes", filter: "all", groupBy: "section", tone: "blue", size: "wide" },
  { id: "dw_tasks_status_donut", type: "donut", title: "Achèvement des tâches", source: "tasks", filter: "all", groupBy: "status", tone: "violet", size: "wide" },
];

const uid = () => crypto.randomUUID();

const seedProjects = [
  {
    id: uid(),
    name: "Challenge National JAM 2026",
    description: "Organisation générale de la compétition.",
    owner: "Samy",
    sport: "developpe_couche",
    status: "En cours",
    priority: "Haute",
    customFieldValues: {},
    startDate: "2026-03-01",
    endDate: "2026-05-30",
    durationDays: 90,
    parentId: null,
    workspaceId: "ws_jam",
  },
  {
    id: uid(),
    name: "Logistique générale",
    description: "Gymnase, matériel, coordination.",
    owner: "Samy",
    sport: "developpe_couche",
    status: "En cours",
    priority: "Haute",
    customFieldValues: {},
    startDate: "2026-03-10",
    endDate: "2026-05-20",
    durationDays: 72,
    parentId: null,
    workspaceId: "ws_jam",
  },
  {
    id: uid(),
    name: "Portail inscriptions clubs",
    description: "Suivi clubs et exports.",
    owner: "Samy",
    sport: "developpe_couche",
    status: "En cours",
    priority: "Haute",
    customFieldValues: {},
    startDate: "2026-03-05",
    endDate: "2026-04-30",
    durationDays: 56,
    parentId: null,
    workspaceId: "ws_jam",
  },
];

const ROOT_ID = seedProjects[0].id;
const LOG_ID = seedProjects[1].id;
const PORTAL_ID = seedProjects[2].id;

const defaultData = {
  workspaces: [
    { id: "ws_jam", name: "JAM", color: "#0f172a", initial: "J" },
    { id: "ws_propcept", name: "Propcept Event", color: "#7c3aed", initial: "P" },
  ],
  currentWorkspaceId: "ws_jam",
  customFields: [
    { id: uid(), name: "Budget prévisionnel", type: "financier", appliesTo: ["projects"], position: "right" },
    { id: uid(), name: "Date comité", type: "date", appliesTo: ["projects", "tasks"], position: "right" },
    { id: uid(), name: "Charge estimée", type: "nombre", appliesTo: ["tasks"], position: "right" },
  ],
  convocations: [
    {
      id: uid(),
      title: "Convocation match U13",
      sport: "football",
      type: "Match",
      date: "2026-04-12",
      time: "14:00",
      location: "Stade municipal",
      recipients: [],
      message: "Merci de vous présenter 45 minutes avant le coup d’envoi.",
      attachmentTitle: "Feuille de match U13",
      status: "Brouillon",
      workspaceId: "ws_jam",
    },
  ],
  notifications: [
    {
      id: uid(),
      title: "Partage feuille d'entraînement",
      sport: "developpe_couche",
      channel: "Email",
      audience: "Groupe sport",
      targetSport: "developpe_couche",
      message: "Nouvelle feuille d'entraînement disponible pour la semaine.",
      attachmentTitle: "Programme semaine force",
      status: "Prêt à envoyer",
      sentAt: "",
      workspaceId: "ws_jam",
    },
  ],
  customSports: [],
  projects: [
    ...seedProjects,
    {
      id: uid(),
      name: "Communication clubs",
      description: "Mails et relances.",
      owner: "Samy",
      sport: "developpe_couche",
      status: "En cours",
      priority: "Moyenne",
      customFieldValues: {},
      startDate: "2026-03-15",
      endDate: "2026-04-15",
      durationDays: 31,
      parentId: ROOT_ID,
      workspaceId: "ws_jam",
    },
    {
      id: uid(),
      name: "Pesée et planning sportif",
      description: "Horaires et flux athlètes.",
      owner: "Samy",
      sport: "developpe_couche",
      status: "À faire",
      priority: "Haute",
      customFieldValues: {},
      startDate: "2026-04-20",
      endDate: "2026-05-15",
      durationDays: 25,
      parentId: ROOT_ID,
      workspaceId: "ws_jam",
    },
    {
      id: uid(),
      name: "Restauration",
      description: "Plateaux repas et régimes.",
      owner: "Samy",
      sport: "developpe_couche",
      status: "À faire",
      priority: "Moyenne",
      customFieldValues: {},
      startDate: "2026-04-01",
      endDate: "2026-05-10",
      durationDays: 39,
      parentId: LOG_ID,
      workspaceId: "ws_jam",
    },
    {
      id: uid(),
      name: "Export inscriptions",
      description: "Contrôle et export CSV.",
      owner: "Samy",
      sport: "developpe_couche",
      status: "En cours",
      priority: "Moyenne",
      customFieldValues: {},
      startDate: "2026-03-20",
      endDate: "2026-04-25",
      durationDays: 36,
      parentId: PORTAL_ID,
      workspaceId: "ws_jam",
    },
  ],
  tasks: [
    {
      id: uid(),
      title: "Envoyer le mail d'ouverture aux clubs",
      description: "Annonce officielle.",
      assignee: "Samy",
      projectId: PORTAL_ID,
      status: "Terminé",
      priority: "Haute",
      customFieldValues: {},
      progress: 100,
      startDate: "2026-03-25",
      endDate: "2026-03-31",
      durationDays: 6,
      workspaceId: "ws_jam",
    },
    {
      id: uid(),
      title: "Préparer la trame des repas",
      description: "Quantités et options.",
      assignee: "Samy",
      projectId: LOG_ID,
      status: "En cours",
      priority: "Moyenne",
      customFieldValues: {},
      progress: 45,
      startDate: "2026-04-10",
      endDate: "2026-04-18",
      durationDays: 8,
      workspaceId: "ws_jam",
    },
    {
      id: uid(),
      title: "Valider la feuille de pesée",
      description: "Version club et arbitres.",
      assignee: "Samy",
      projectId: ROOT_ID,
      status: "À faire",
      priority: "Haute",
      customFieldValues: {},
      progress: 10,
      startDate: "2026-04-15",
      endDate: "2026-04-20",
      durationDays: 5,
      workspaceId: "ws_jam",
    },
  ],
  athletes: [
    { id: uid(), name: "Athlète exemple 1", sport: "developpe_couche", subSection: "-88 kg", category: "Homme", coach: "Coach JAM", email: "", phone: "", status: "Actif", notes: "Suivi force.", workspaceId: "ws_jam" },
    { id: uid(), name: "Athlète exemple 2", sport: "football", subSection: "U13", category: "Junior", coach: "Coach Terrain", email: "", phone: "", status: "À relancer", notes: "Feuille de match.", workspaceId: "ws_jam" },
  ],
  clubs: [
    { id: uid(), name: "Développé couché", sport: "developpe_couche", city: "Montrouge", contactName: "Samy", email: "", phone: "", delegationSize: "12", status: "Validé", notes: "Section interne JAM.", subSections: ["-59 kg", "-66 kg", "-74 kg", "-83 kg", "-93 kg", "-105 kg", "+105 kg"], workspaceId: "ws_jam" },
    { id: uid(), name: "Football", sport: "football", city: "Montrouge", contactName: "Responsable section", email: "", phone: "", delegationSize: "18", status: "Inscrit", notes: "Section football JAM.", subSections: ["U7", "U9", "U11", "U13", "U15"], workspaceId: "ws_jam" },
  ],
  documents: [
    { id: uid(), sport: "developpe_couche", title: "Feuille d'entraînement type", type: "Suivi sportif", status: "Validé", notes: "Modèle à transmettre.", url: "", content: "Exemple de contenu interne pour suivi sportif.", workspaceId: "ws_jam" },
    { id: uid(), sport: "football", title: "Feuille de match", type: "Match", status: "À valider", notes: "Document joueur.", url: "", content: "Composition, remplaçants, consignes.", workspaceId: "ws_jam" },
    { id: uid(), sport: "developpe_couche", title: "Portail inscriptions", type: "Lien", status: "Publié", notes: "Accès externe au portail.", url: "https://challenge-national-jam.netlify.app/", content: "", workspaceId: "ws_jam" },
  ],
  tableConfig: {
    projects: { customCols: [], hiddenCols: [], renamed: {}, typeOverrides: {}, order: [], groupBy: null, filters: {} },
    tasks: { customCols: [], hiddenCols: [], renamed: {}, typeOverrides: {}, order: [], groupBy: null, filters: {} },
    athletes: { customCols: [], hiddenCols: [], renamed: {}, typeOverrides: {}, order: [], groupBy: null, filters: {} },
    clubs: { customCols: [], hiddenCols: [], renamed: {}, typeOverrides: {}, order: [], groupBy: null, filters: {} },
    documents: { customCols: [], hiddenCols: [], renamed: {}, typeOverrides: {}, order: [], groupBy: null, filters: {} },
  },
  dashboardWidgets: DEFAULT_DASHBOARD_WIDGETS,
};

function loadData() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return defaultData;
    const parsed = JSON.parse(raw);
    const addWs = (arr) => (arr || []).map((item) => item.workspaceId ? item : { ...item, workspaceId: "ws_jam" });
    const addDocDefaults = (arr) => addWs(arr).map((doc) => ({ status: doc.url ? "Publié" : "Brouillon", ...doc }));
    const clubs = addWs(parsed.clubs || defaultData.clubs).map((section) => ({
      ...section,
      subSections: Array.isArray(section.subSections) ? section.subSections : String(section.subSections || "").split(",").map((item) => item.trim()).filter(Boolean),
    }));
    const athletes = addWs(parsed.athletes || defaultData.athletes).map((athlete) => {
      const legacyCategory = String(athlete.category || "").trim();
      const categoryIsProfile = ATHLETE_PROFILE_OPTIONS.includes(legacyCategory);
      const fallbackSection = clubs.find((section) => section.workspaceId === athlete.workspaceId && section.sport === athlete.sport);
      return {
        ...athlete,
        sectionId: athlete.sectionId || fallbackSection?.id || "",
        subSection: athlete.subSection || (categoryIsProfile ? "" : legacyCategory),
        category: categoryIsProfile ? legacyCategory : "",
      };
    });
    return {
      ...defaultData,
      ...parsed,
      workspaces: parsed.workspaces || defaultData.workspaces,
      currentWorkspaceId: parsed.currentWorkspaceId || "ws_jam",
      projects: addWs(parsed.projects || defaultData.projects),
      tasks: addWs(parsed.tasks || defaultData.tasks),
      athletes,
      clubs,
      documents: addDocDefaults(parsed.documents || defaultData.documents),
      convocations: addWs(parsed.convocations || defaultData.convocations),
      notifications: addWs(parsed.notifications || defaultData.notifications),
      tableConfig: { ...defaultData.tableConfig, ...(parsed.tableConfig || {}) },
      dashboardWidgets: parsed.dashboardWidgets?.length ? parsed.dashboardWidgets : DEFAULT_DASHBOARD_WIDGETS,
    };
  } catch {
    return defaultData;
  }
}

function saveData(data) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
}

function badgeClass(kind, value) {
  if (kind === "status") {
    if (value === "Terminé") return "bg-green-100 text-green-700 border-green-200";
    if (value === "En cours") return "bg-blue-100 text-blue-700 border-blue-200";
    if (value === "En attente") return "bg-amber-100 text-amber-700 border-amber-200";
    return "bg-slate-100 text-slate-700 border-slate-200";
  }
  if (value === "Critique") return "bg-red-100 text-red-700 border-red-200";
  if (value === "Haute") return "bg-orange-100 text-orange-700 border-orange-200";
  if (value === "Moyenne") return "bg-violet-100 text-violet-700 border-violet-200";
  return "bg-slate-100 text-slate-700 border-slate-200";
}

function statusCellClass(value) {
  if (value === "Terminé") return "bg-emerald-500 text-white";
  if (value === "En cours") return "bg-blue-500 text-white";
  if (value === "En attente") return "bg-amber-400 text-slate-950";
  return "bg-slate-200 text-slate-700";
}

function priorityDotClass(value) {
  if (value === "Critique") return "bg-red-500";
  if (value === "Haute") return "bg-orange-500";
  if (value === "Moyenne") return "bg-violet-500";
  return "bg-slate-400";
}

function progressTone(value) {
  if (value >= 100) return "bg-green-500";
  if (value >= 70) return "bg-blue-500";
  if (value >= 40) return "bg-amber-500";
  return "bg-red-500";
}

function descendants(projectId, projects) {
  const ids = new Set([projectId]);
  let changed = true;
  while (changed) {
    changed = false;
    for (const p of projects) {
      if (p.parentId && ids.has(p.parentId) && !ids.has(p.id)) {
        ids.add(p.id);
        changed = true;
      }
    }
  }
  return ids;
}

function parentOptionsForProject(projectId, projects) {
  if (!projectId) return projects;
  const blocked = descendants(projectId, projects);
  return projects.filter((project) => !blocked.has(project.id));
}

function clampProgress(value) {
  return Math.max(0, Math.min(100, Number(value) || 0));
}

function normalizeTaskProgress(progress, status) {
  const raw = clampProgress(progress);
  if (status === "Terminé") return 100;
  if (status === "À faire") return raw === 100 ? 0 : raw;
  if (status === "En attente") return raw === 0 ? 5 : Math.min(raw, 95);
  return raw === 0 ? 5 : Math.min(raw, 99);
}

function statusFromProgress(progress, currentStatus) {
  const value = clampProgress(progress);
  if (value >= 100) return "Terminé";
  if (value <= 0) return "À faire";
  if (currentStatus === "En attente") return "En attente";
  return "En cours";
}

function ownProgress(projectId, tasks) {
  const rows = tasks.filter((t) => t.projectId === projectId);
  if (!rows.length) return 0;
  return Math.round(rows.reduce((sum, row) => sum + normalizeTaskProgress(row.progress, row.status), 0) / rows.length);
}

function consolidatedProgress(projectId, tasks, projects) {
  const ids = descendants(projectId, projects);
  const rows = tasks.filter((t) => ids.has(t.projectId));
  if (!rows.length) return 0;
  return Math.round(rows.reduce((sum, row) => sum + normalizeTaskProgress(row.progress, row.status), 0) / rows.length);
}

function deriveProjectStatus(progress, currentStatus) {
  if (progress >= 100) return "Terminé";
  if (progress <= 0) return "À faire";
  if (currentStatus === "En attente") return "En attente";
  return "En cours";
}

function buildTree(projects) {
  const map = new Map(projects.map((p) => [p.id, { ...p, children: [] }]));
  const roots = [];
  map.forEach((p) => {
    if (p.parentId && map.has(p.parentId)) map.get(p.parentId).children.push(p);
    else roots.push(p);
  });
  return roots;
}

function levelOf(projectId, projects) {
  const map = new Map(projects.map((p) => [p.id, p]));
  let level = 0;
  let cur = map.get(projectId);
  while (cur?.parentId && map.has(cur.parentId)) {
    level += 1;
    cur = map.get(cur.parentId);
  }
  return level + 1;
}

function daysBetween(startDate, endDate) {
  if (!startDate || !endDate) return "—";
  const start = new Date(startDate);
  const end = new Date(endDate);
  const diff = Math.round((end - start) / (1000 * 60 * 60 * 24));
  return Number.isFinite(diff) ? diff : "—";
}

function dateOnlyTime(value) {
  if (!value) return null;
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return null;
  date.setHours(0, 0, 0, 0);
  return date.getTime();
}

function daysFromToday(value, todayTime) {
  const time = dateOnlyTime(value);
  if (time === null) return null;
  return Math.round((time - todayTime) / (1000 * 60 * 60 * 24));
}

function formatCustomFieldValue(type, value) {
  if (value === undefined || value === null || value === "") return "—";
  if (type === "financier") {
    const num = Number(value);
    return Number.isFinite(num) ? new Intl.NumberFormat("fr-FR", { style: "currency", currency: "EUR" }).format(num) : value;
  }
  if (type === "nombre") {
    const num = Number(value);
    return Number.isFinite(num) ? new Intl.NumberFormat("fr-FR").format(num) : value;
  }
  if (type === "pourcentage") {
    const num = Number(value);
    return Number.isFinite(num) ? `${new Intl.NumberFormat("fr-FR").format(num)} %` : value;
  }
  if (type === "case_a_cocher") return value ? "Oui" : "Non";
  return value;
}

function formatMoney(value) {
  const num = Number(value);
  return Number.isFinite(num) ? new Intl.NumberFormat("fr-FR", { style: "currency", currency: "EUR" }).format(num) : "0 €";
}

function normalizeKey(value) {
  return String(value || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9%€]+/g, "");
}

function firstValue(row, aliases) {
  const wanted = aliases.map(normalizeKey);
  const key = Object.keys(row || {}).find((candidate) => wanted.includes(normalizeKey(candidate)));
  return key ? row[key] : "";
}

function excelDateToIso(value) {
  if (!value) return "";
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value.toISOString().slice(0, 10);
  if (typeof value === "number") {
    const date = new Date(Math.round((value - 25569) * 86400 * 1000));
    if (!Number.isNaN(date.getTime())) return date.toISOString().slice(0, 10);
  }
  const raw = String(value).trim();
  const slash = raw.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})$/);
  if (slash) {
    const year = slash[3].length === 2 ? `20${slash[3]}` : slash[3];
    return `${year}-${slash[2].padStart(2, "0")}-${slash[1].padStart(2, "0")}`;
  }
  const date = new Date(raw);
  return Number.isNaN(date.getTime()) ? raw : date.toISOString().slice(0, 10);
}

function parseNumberValue(value) {
  if (value === undefined || value === null || value === "") return "";
  const normalized = String(value).replace(/\s/g, "").replace("€", "").replace("%", "").replace(",", ".");
  const num = Number(normalized);
  return Number.isFinite(num) ? num : "";
}

function safeSheetName(name, existing = []) {
  const base = String(name).slice(0, 31) || "Feuille";
  if (!existing.includes(base)) return base;
  let i = 2;
  while (existing.includes(`${base.slice(0, 27)} ${i}`)) i += 1;
  return `${base.slice(0, 27)} ${i}`;
}

function appendSheet(workbook, name, rows) {
  const sheetName = safeSheetName(name, workbook.worksheets.map((sheet) => sheet.name));
  const worksheet = workbook.addWorksheet(sheetName);
  const headers = Array.from(rows.reduce((set, row) => {
    Object.keys(row || {}).forEach((key) => set.add(key));
    return set;
  }, new Set()));
  worksheet.addRow(headers);
  rows.forEach((row) => worksheet.addRow(headers.map((header) => row?.[header] ?? "")));
  worksheet.getRow(1).font = { bold: true };
  worksheet.columns.forEach((column) => {
    column.width = Math.min(42, Math.max(14, ...column.values.slice(1).map((value) => String(value ?? "").length + 2)));
  });
  return worksheet;
}

function applyImportSheetFormat(worksheet, { required = [], widths = {}, listValidations = {}, dateHeaders = [], numericHeaders = [], rows = 1200, color = "0F172A" } = {}) {
  const headerRow = worksheet.getRow(1);
  const headers = headerRow.values.slice(1).map((value) => String(value || ""));
  worksheet.views = [{ state: "frozen", ySplit: 1 }];
  worksheet.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: Math.max(2, worksheet.rowCount), column: Math.max(1, headers.length) },
  };
  worksheet.properties.defaultRowHeight = 22;

  headerRow.height = 28;
  headerRow.eachCell((cell) => {
    const isRequired = required.includes(String(cell.value || ""));
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: isRequired ? "FFB91C1C" : `FF${color}` } };
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    cell.border = {
      top: { style: "thin", color: { argb: "FFE2E8F0" } },
      bottom: { style: "thin", color: { argb: "FFCBD5E1" } },
      left: { style: "thin", color: { argb: "FFE2E8F0" } },
      right: { style: "thin", color: { argb: "FFE2E8F0" } },
    };
  });

  headers.forEach((header, index) => {
    const column = worksheet.getColumn(index + 1);
    column.width = widths[header] || Math.min(34, Math.max(15, header.length + 4));
    if (dateHeaders.includes(header)) column.numFmt = "yyyy-mm-dd";
    if (numericHeaders.includes(header)) column.numFmt = "0";
    for (let row = 2; row <= rows; row += 1) {
      const cell = worksheet.getCell(row, index + 1);
      cell.alignment = { vertical: "middle" };
      if (row % 2 === 0) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };
      if (listValidations[header]) {
        cell.dataValidation = {
          type: "list",
          allowBlank: !required.includes(header),
          formulae: [`"${listValidations[header].join(",")}"`],
          showErrorMessage: true,
          errorTitle: "Valeur non valide",
          error: "Choisis une valeur dans la liste.",
        };
      }
    }
  });

  worksheet.getRow(2).font = { italic: true, color: { argb: "FF475569" } };
}

function applyGuideSheetFormat(worksheet) {
  worksheet.views = [{ state: "frozen", ySplit: 1 }];
  worksheet.getRow(1).height = 28;
  worksheet.getRow(1).eachCell((cell) => {
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0F172A" } };
    cell.alignment = { vertical: "middle", wrapText: true };
  });
  worksheet.getColumn(1).width = 18;
  worksheet.getColumn(2).width = 34;
  worksheet.getColumn(3).width = 90;
  worksheet.getColumn(4).width = 18;
  worksheet.eachRow((row) => {
    row.alignment = { vertical: "top", wrapText: true };
  });
}

function readSheetRows(workbook, names) {
  const normalizedNames = names.map(normalizeKey);
  const worksheet = workbook.worksheets.find((sheet) => normalizedNames.includes(normalizeKey(sheet.name)));
  if (!worksheet) return [];
  const headers = [];
  worksheet.getRow(1).eachCell({ includeEmpty: true }, (cell, colNumber) => {
    headers[colNumber] = cell.value ? String(cell.value) : "";
  });
  const rows = [];
  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) return;
    const item = {};
    headers.forEach((header, colNumber) => {
      if (!header) return;
      const cell = row.getCell(colNumber);
      const value = cell.value;
      item[header] = value && typeof value === "object" && "text" in value ? value.text : value ?? "";
    });
    if (Object.values(item).some((value) => value !== "")) rows.push(item);
  });
  return rows;
}

async function downloadWorkbook(workbook, filename) {
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function customFieldValuesFromRow(row, fields) {
  const values = {};
  fields.forEach((field) => {
    const key = Object.keys(row || {}).find((candidate) => normalizeKey(candidate) === normalizeKey(field.name));
    if (!key) return;
    const raw = row[key];
    if (raw === "") return;
    if (field.type === "date") values[field.id] = excelDateToIso(raw);
    else if (field.type === "nombre" || field.type === "financier" || field.type === "pourcentage") values[field.id] = parseNumberValue(raw);
    else if (field.type === "case_a_cocher") values[field.id] = ["oui", "yes", "true", "1", "x"].includes(String(raw).trim().toLowerCase());
    else values[field.id] = raw;
  });
  return values;
}

function tableColumnsToFields(cols = [], appliesTo = []) {
  return (cols || []).map((col) => ({
    id: col.key,
    name: col.label,
    type: col.type,
    appliesTo,
    source: "table",
  }));
}

function mergeCustomFields(...groups) {
  const seen = new Set();
  return groups.flat().filter((field) => {
    if (!field?.id || seen.has(field.id)) return false;
    seen.add(field.id);
    return true;
  });
}

function makeSportId(name) {
  return name
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "");
}

function Card({ children, className = "", ...props }) {
  return <div className={`bg-white rounded-2xl border border-slate-200 shadow-sm ${className}`} {...props}>{children}</div>;
}

function Badge({ children, className = "" }) {
  return <span className={`text-xs px-2 py-1 rounded-full border ${className}`}>{children}</span>;
}

function ProgressBar({ value, label }) {
  return (
    <div>
      <div className="flex justify-between text-xs text-slate-500 mb-1">
        <span>{label}</span>
        <span>{value}%</span>
      </div>
      <div className="w-full h-2 rounded-full bg-slate-200 overflow-hidden">
        <div className={`h-full rounded-full ${progressTone(value)}`} style={{ width: `${value}%` }} />
      </div>
    </div>
  );
}

function ViewToggle({ options, value, onChange }) {
  return (
    <div className="flex rounded-2xl border bg-white p-1 gap-1">
      {options.map(({ id, label, Icon }) => (
        <button
          key={id}
          type="button"
          onClick={() => onChange(id)}
          className={`flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-sm font-medium transition-colors ${value === id ? "bg-slate-900 text-white" : "text-slate-600 hover:bg-slate-100"}`}
        >
          <Icon className="w-4 h-4" /> {label}
        </button>
      ))}
    </div>
  );
}

function KanbanBoard({ rows, statuses, getStatus = (row) => row.status, onStatusChange, renderCard, emptyLabel = "Aucun élément.", celebrateStatus }) {
  const [dragId, setDragId] = useState(null);
  const [celebration, setCelebration] = useState(null);
  const normalizedStatus = (row) => {
    const value = getStatus(row);
    return statuses.includes(value) ? value : statuses[0];
  };

  if (!rows.length) return <p className="text-center text-slate-500 py-12">{emptyLabel}</p>;

  return (
    <div className="overflow-x-auto pb-3">
      <div className="grid auto-cols-[minmax(260px,1fr)] grid-flow-col gap-4 min-w-full">
        {statuses.map((status) => {
          const items = rows.filter((row) => normalizedStatus(row) === status);
          return (
            <section
              key={status}
              className="min-h-[360px] rounded-2xl border bg-slate-100/70 p-3"
              onDragOver={(e) => onStatusChange && e.preventDefault()}
              onDrop={(e) => {
                e.preventDefault();
                if (dragId && onStatusChange) {
                  onStatusChange(dragId, status);
                  if (celebrateStatus && status === celebrateStatus) {
                    setCelebration(Date.now());
                    window.setTimeout(() => setCelebration(null), 1200);
                  }
                }
                setDragId(null);
              }}
            >
              <div className="flex items-center justify-between gap-3 mb-3">
                <h2 className="text-sm font-bold text-slate-800">{status}</h2>
                <Badge className="bg-white text-slate-600 border-slate-200">{items.length}</Badge>
              </div>
              <div className="space-y-3">
                {items.map((row) => (
                  <div
                    key={row.id}
                    draggable={Boolean(onStatusChange)}
                    onDragStart={() => setDragId(row.id)}
                    onDragEnd={() => setDragId(null)}
                    className="rounded-2xl border bg-white p-4 shadow-sm cursor-grab active:cursor-grabbing"
                  >
                    {renderCard(row)}
                  </div>
                ))}
                {!items.length && (
                  <div className="rounded-2xl border border-dashed bg-white/60 p-4 text-center text-xs text-slate-400">
                    Dépose ici
                  </div>
                )}
              </div>
            </section>
          );
        })}
      </div>
      {celebration && (
        <div key={celebration} className="pointer-events-none fixed inset-0 z-[70] overflow-hidden">
          {Array.from({ length: 32 }).map((_, index) => (
            <span
              key={index}
              className="confetti-piece"
              style={{
                left: `${12 + ((index * 73) % 76)}%`,
                backgroundColor: ["#22c55e", "#3b82f6", "#f59e0b", "#ef4444", "#8b5cf6"][index % 5],
                animationDelay: `${(index % 9) * 0.04}s`,
                transform: `rotate(${index * 29}deg)`,
              }}
            />
          ))}
          <div className="absolute left-1/2 top-24 -translate-x-1/2 rounded-2xl bg-white px-5 py-3 text-sm font-bold text-slate-900 shadow-xl border">
            <PartyPopper className="mr-2 inline h-4 w-4 text-amber-500" /> Projet terminé
          </div>
        </div>
      )}
    </div>
  );
}

function StatCard({ title, value, subtitle, Icon, tone = "slate", onClick }) {
  const tones = {
    slate: "bg-slate-100 text-slate-700",
    blue: "bg-blue-100 text-blue-700",
    green: "bg-green-100 text-green-700",
    amber: "bg-amber-100 text-amber-700",
    red: "bg-red-100 text-red-700",
    violet: "bg-violet-100 text-violet-700",
  };
  const content = (
    <>
      <div className="flex items-start justify-between">
        <div className="min-w-0">
          <p className="text-sm text-slate-500">{title}</p>
          <p className="text-2xl font-semibold mt-2 truncate">{value}</p>
          <p className="text-xs text-slate-500 mt-2">{subtitle}</p>
        </div>
        <div className={`p-2.5 rounded-xl ${tones[tone] || tones.slate}`}><Icon className="w-5 h-5" /></div>
      </div>
      {onClick && <p className="mt-3 text-[11px] font-semibold uppercase tracking-wide text-slate-400">Ouvrir</p>}
    </>
  );
  if (onClick) {
    return (
      <button
        type="button"
        onClick={onClick}
        className="group w-full text-left bg-white rounded-2xl border border-slate-200 shadow-sm p-4 transition hover:-translate-y-0.5 hover:border-slate-300 hover:shadow-md focus:outline-none focus:ring-2 focus:ring-slate-900/20"
      >
        {content}
      </button>
    );
  }
  return (
    <Card className="p-4">
      {content}
    </Card>
  );
}

function Modal({ open, title, children, onClose, onSave, saveLabel = "Enregistrer" }) {
  if (!open) return null;
  return (
    <div className="fixed inset-0 z-50 bg-black/40 flex items-center justify-center p-4">
      <div className="bg-white rounded-3xl shadow-2xl w-full max-w-3xl p-6 max-h-[92vh] overflow-auto">
        <div className="flex items-center justify-between mb-4">
          <h3 className="text-xl font-semibold">{title}</h3>
          <button onClick={onClose} className="text-slate-500 text-xl">✕</button>
        </div>
        {children}
        <div className="flex justify-end gap-2 mt-4">
          <button onClick={onClose} className="px-4 py-2 rounded-2xl border">Annuler</button>
          <button onClick={onSave} className="px-4 py-2 rounded-2xl bg-slate-900 text-white">{saveLabel}</button>
        </div>
      </div>
    </div>
  );
}

function CustomFieldsEditor({ fields, values, onChange }) {
  if (!fields.length) {
    return <p className="text-sm text-slate-500">Aucune colonne personnalisée pour ce type d’élément.</p>;
  }

  return (
    <div className="grid md:grid-cols-2 gap-4">
      {fields.map((field) => {
        const value = values?.[field.id] ?? "";
        const isNumeric = field.type === "nombre" || field.type === "financier" || field.type === "pourcentage";
        if (field.type === "case_a_cocher") {
          return (
            <label key={field.id} className="flex items-center gap-3 rounded-2xl border px-4 py-3 bg-white">
              <input type="checkbox" checked={Boolean(value)} onChange={(e) => onChange(field.id, e.target.checked)} />
              <span className="text-sm font-medium text-slate-700">{field.name}</span>
            </label>
          );
        }

        if (field.type === "statut") {
          return (
            <label key={field.id} className="block">
              <span className="mb-1.5 block text-xs font-bold uppercase tracking-wide text-slate-500">{field.name}</span>
              <select className="w-full rounded-2xl border px-4 py-3" value={value} onChange={(e) => onChange(field.id, e.target.value)}>
                <option value="">Non renseigné</option>
                <option value="À faire">À faire</option>
                <option value="En cours">En cours</option>
                <option value="Bloqué">Bloqué</option>
                <option value="Terminé">Terminé</option>
              </select>
            </label>
          );
        }

        return (
          <label key={field.id} className="block">
            <span className="mb-1.5 block text-xs font-bold uppercase tracking-wide text-slate-500">
              {field.name}{field.type === "financier" ? " (€)" : field.type === "pourcentage" ? " (%)" : ""}
            </span>
            <input
              type={field.type === "date" ? "date" : isNumeric ? "number" : "text"}
              step={field.type === "financier" ? "0.01" : field.type === "nombre" || field.type === "pourcentage" ? "1" : undefined}
              min={field.type === "pourcentage" ? "0" : undefined}
              max={field.type === "pourcentage" ? "100" : undefined}
              className="w-full rounded-2xl border px-4 py-3"
              placeholder={field.type === "date" ? "" : field.name}
              value={value}
              onChange={(e) => onChange(field.id, e.target.value)}
            />
          </label>
        );
      })}
    </div>
  );
}

function ProjectTreeNode({ node, tasks, projects, expanded, setExpanded, onQuickTask, onNewSubProject, onEditProject, requestDelete, onDuplicate, projectCustomFields }) {
  const [menuOpen, setMenuOpen] = useState(false);
  const menuRef = useRef(null);

  useEffect(() => {
    if (!menuOpen) return;
    const handler = (e) => { if (menuRef.current && !menuRef.current.contains(e.target)) setMenuOpen(false); };
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, [menuOpen]);

  const own = node.computedOwnProgress ?? ownProgress(node.id, tasks);
  const total = node.computedTotalProgress ?? consolidatedProgress(node.id, tasks, projects);
  const isOpen = expanded[node.id] ?? true;
  const childCount = node.children.length;
  const directTasks = tasks.filter((t) => t.projectId === node.id);

  return (
    <div className="border rounded-2xl p-4 bg-white shadow-sm">
      <div className="flex items-start justify-between gap-4">
        <div className="flex gap-3 flex-1 min-w-0">
          <button onClick={() => setExpanded((prev) => ({ ...prev, [node.id]: !isOpen }))} className="mt-1 text-slate-500 shrink-0">
            {childCount ? (isOpen ? <ChevronDown className="w-4 h-4" /> : <ChevronRight className="w-4 h-4" />) : <div className="w-4" />}
          </button>
          <div className="flex-1 min-w-0">
            <div className="flex flex-wrap items-center gap-2">
              <h3 className="font-semibold truncate">{node.name}</h3>
              <Badge>Niv. {levelOf(node.id, projects)}</Badge>
              <Badge className={badgeClass("status", node.status)}>{node.status}</Badge>
              <Badge className={badgeClass("priority", node.priority)}>{node.priority}</Badge>
              {node.sportName && <Badge>{node.sportName}</Badge>}
            </div>
            {node.description && <p className="text-sm text-slate-500 mt-1">{node.description}</p>}
            <div className="grid md:grid-cols-2 gap-4 mt-3 max-w-3xl">
              <ProgressBar value={own} label="Progression directe" />
              <ProgressBar value={total} label="Progression consolidée" />
            </div>
            <div className="flex flex-wrap gap-3 text-xs text-slate-500 mt-2">
              {node.owner && <span>Responsable : {node.owner}</span>}
              {node.startDate && <span>Début : {node.startDate}</span>}
              {node.endDate && <span>Fin : {node.endDate}</span>}
              <span>Sous-projets : {childCount}</span>
              <span>Tâches : {directTasks.length}</span>
            </div>

            {projectCustomFields.length > 0 && (
              <div className="flex flex-wrap gap-2 mt-2">
                {projectCustomFields.map((field) => (
                  <Badge key={field.id}>{field.name} : {formatCustomFieldValue(field.type, node.customFieldValues?.[field.id])}</Badge>
                ))}
              </div>
            )}

            {directTasks.length > 0 && (
              <div className="mt-4 rounded-2xl bg-slate-50 border border-slate-200 p-3">
                <p className="text-xs font-semibold text-slate-500 uppercase tracking-wide mb-2">Tâches ({directTasks.length})</p>
                <div className="space-y-2">
                  {directTasks.map((task) => {
                    const prog = normalizeTaskProgress(task.progress, task.status);
                    return (
                      <div key={task.id} className="rounded-xl border bg-white p-3">
                        <div className="flex items-start justify-between gap-3">
                          <div className="flex items-center gap-2 min-w-0">
                            <CheckCircle2 className={`w-4 h-4 shrink-0 ${task.status === "Terminé" ? "text-green-500" : "text-slate-300"}`} />
                            <p className="text-sm font-medium text-slate-900 truncate">{task.title}</p>
                          </div>
                          <div className="flex items-center gap-1.5 shrink-0">
                            <Badge className={badgeClass("status", task.status)}>{task.status}</Badge>
                            <Badge className={badgeClass("priority", task.priority)}>{task.priority}</Badge>
                          </div>
                        </div>
                        <div className="flex items-center gap-3 mt-2 pl-6">
                          <span className="text-xs text-slate-500">{task.assignee || "Non assigné"}</span>
                          {task.endDate && <span className="text-xs text-slate-400">· {task.endDate}</span>}
                        </div>
                        <div className="mt-2 pl-6">
                          <div className="flex items-center gap-2">
                            <div className="flex-1 h-1.5 rounded-full bg-slate-200 overflow-hidden">
                              <div className={`h-full rounded-full ${progressTone(prog)}`} style={{ width: `${prog}%` }} />
                            </div>
                            <span className="text-xs font-bold text-slate-600 w-7 text-right">{prog}%</span>
                          </div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}
          </div>
        </div>

        {/* ··· context menu */}
        <div className="relative shrink-0" ref={menuRef}>
          <button
            onClick={(e) => { e.stopPropagation(); setMenuOpen((v) => !v); }}
            className="p-2 rounded-xl hover:bg-slate-100 text-slate-500"
          >
            <MoreHorizontal className="w-4 h-4" />
          </button>
          {menuOpen && (
            <div className="absolute right-0 top-full z-40 mt-1 bg-white rounded-xl shadow-xl border py-1 min-w-[180px]">
              <button onClick={() => { onEditProject(node); setMenuOpen(false); }} className="w-full text-left px-3 py-2 text-sm hover:bg-slate-50 flex items-center gap-2">
                <Edit3 className="w-4 h-4 text-slate-400" /> Modifier
              </button>
              <button onClick={() => { onQuickTask(node.id); setMenuOpen(false); }} className="w-full text-left px-3 py-2 text-sm hover:bg-slate-50 flex items-center gap-2">
                <Plus className="w-4 h-4 text-slate-400" /> Ajouter une tâche
              </button>
              <button onClick={() => { onNewSubProject(node); setMenuOpen(false); }} className="w-full text-left px-3 py-2 text-sm hover:bg-slate-50 flex items-center gap-2">
                <FolderKanban className="w-4 h-4 text-slate-400" /> Créer un sous-projet
              </button>
              <button onClick={() => { onDuplicate(node); setMenuOpen(false); }} className="w-full text-left px-3 py-2 text-sm hover:bg-slate-50 flex items-center gap-2">
                <Copy className="w-4 h-4 text-slate-400" /> Dupliquer
              </button>
              <div className="border-t my-1" />
              <button onClick={() => { requestDelete("project", node); setMenuOpen(false); }} className="w-full text-left px-3 py-2 text-sm hover:bg-red-50 text-red-600 flex items-center gap-2">
                <Trash2 className="w-4 h-4" /> Supprimer
              </button>
            </div>
          )}
        </div>
      </div>

      {isOpen && childCount > 0 && (
        <div className="pl-8 mt-4 border-l-2 border-slate-100 space-y-3">
                <ProjectTree nodes={node.children} tasks={tasks} projects={projects} expanded={expanded} setExpanded={setExpanded} onQuickTask={onQuickTask} onNewSubProject={onNewSubProject} onEditProject={onEditProject} requestDelete={requestDelete} onDuplicate={onDuplicate} projectCustomFields={projectCustomFields} />
        </div>
      )}
    </div>
  );
}

function ProjectTree({ nodes, tasks, projects, expanded, setExpanded, onQuickTask, onNewSubProject, onEditProject, requestDelete, onDuplicate, projectCustomFields }) {
  return (
    <div className="space-y-3">
      {nodes.map((node) => (
        <ProjectTreeNode key={node.id} node={node} tasks={tasks} projects={projects} expanded={expanded} setExpanded={setExpanded} onQuickTask={onQuickTask} onNewSubProject={onNewSubProject} onEditProject={onEditProject} requestDelete={requestDelete} onDuplicate={onDuplicate} projectCustomFields={projectCustomFields} />
      ))}
    </div>
  );
}

// ===== Generic table column helpers =====
const COL_TYPES = [
  { id: "texte", label: "Texte" },
  { id: "nombre", label: "Nombre" },
  { id: "pourcentage", label: "Pourcentage (%)" },
  { id: "date", label: "Date" },
  { id: "financier", label: "Financier (€)" },
  { id: "statut", label: "Statut" },
  { id: "case_a_cocher", label: "Case à cocher" },
  { id: "personne", label: "Personne" },
];

const COL_TYPE_STYLES = {
  texte: { badge: "T", className: "bg-amber-400 text-white" },
  nombre: { badge: "123", className: "bg-yellow-400 text-white text-[10px]" },
  pourcentage: { badge: "%", className: "bg-emerald-500 text-white" },
  date: { badge: "D", className: "bg-violet-500 text-white" },
  financier: { badge: "€", className: "bg-green-500 text-white" },
  statut: { badge: "=", className: "bg-emerald-500 text-white" },
  case_a_cocher: { badge: "✓", className: "bg-blue-500 text-white" },
  personne: { badge: "P", className: "bg-sky-400 text-white" },
};

function ColumnTypeMark({ type }) {
  const style = COL_TYPE_STYLES[type] || COL_TYPE_STYLES.texte;
  return (
    <span className={`inline-flex h-7 w-7 shrink-0 items-center justify-center rounded-lg text-xs font-bold ${style.className}`}>
      {style.badge}
    </span>
  );
}

function passesFilter(value, filter, type) {
  if (!filter || filter.value === undefined || filter.value === null || filter.value === "") return true;
  const v = value;
  if (type === "texte" || type === "personne") {
    return String(v ?? "").toLowerCase().includes(String(filter.value).toLowerCase());
  }
  if (type === "nombre" || type === "financier" || type === "pourcentage") {
    const a = Number(v); const b = Number(filter.value);
    if (!Number.isFinite(a) || !Number.isFinite(b)) return true;
    if (filter.op === ">") return a > b;
    if (filter.op === "<") return a < b;
    return a === b;
  }
  if (type === "date") {
    if (!v) return false;
    const a = new Date(v).getTime();
    if (filter.op === "before") return a < new Date(filter.value).getTime();
    if (filter.op === "after") return a > new Date(filter.value).getTime();
    if (filter.op === "between" && filter.value2) {
      const t1 = new Date(filter.value).getTime();
      const t2 = new Date(filter.value2).getTime();
      return a >= t1 && a <= t2;
    }
    return true;
  }
  if (type === "statut") {
    if (Array.isArray(filter.value)) return filter.value.length === 0 || filter.value.includes(v);
    return v === filter.value;
  }
  if (type === "case_a_cocher") {
    return Boolean(v) === Boolean(filter.value);
  }
  return true;
}

function ColumnFilterUI({ col, filter, onChange }) {
  const t = col.type;
  if (t === "statut" && col.options) {
    const sel = Array.isArray(filter?.value) ? filter.value : (filter?.value ? [filter.value] : []);
    const toggle = (opt) => {
      const next = sel.includes(opt) ? sel.filter((x) => x !== opt) : [...sel, opt];
      onChange(next.length ? { value: next } : null);
    };
    return (
      <div className="space-y-1">
        {col.options.map((opt) => (
          <button key={opt} onClick={() => toggle(opt)}
            className={`w-full text-left px-2 py-1.5 rounded-lg text-sm flex items-center gap-2 ${sel.includes(opt) ? "bg-slate-900 text-white" : "hover:bg-slate-50"}`}>
            {sel.includes(opt) && <CheckCircle2 className="w-3 h-3 shrink-0" />} {opt}
          </button>
        ))}
      </div>
    );
  }
  if (t === "case_a_cocher") {
    return (
      <div className="grid grid-cols-2 gap-1">
        <button onClick={() => onChange({ value: true })} className={`px-2 py-1.5 rounded-lg text-sm border ${filter?.value === true ? "bg-slate-900 text-white" : "hover:bg-slate-50"}`}>Oui</button>
        <button onClick={() => onChange({ value: false })} className={`px-2 py-1.5 rounded-lg text-sm border ${filter?.value === false ? "bg-slate-900 text-white" : "hover:bg-slate-50"}`}>Non</button>
      </div>
    );
  }
  if (t === "nombre" || t === "financier" || t === "pourcentage") {
    return (
      <div className="flex gap-1">
        <select className="rounded-lg border px-2 py-1.5 text-sm" value={filter?.op || "="} onChange={(e) => onChange({ ...filter, op: e.target.value, value: filter?.value ?? "" })}>
          <option value="=">=</option><option value=">">{">"}</option><option value="<">{"<"}</option>
        </select>
        <input type="number" className="flex-1 rounded-lg border px-2 py-1.5 text-sm" placeholder="valeur"
          value={filter?.value ?? ""} onChange={(e) => onChange({ op: filter?.op || "=", value: e.target.value })} />
      </div>
    );
  }
  if (t === "date") {
    return (
      <div className="space-y-1">
        <select className="w-full rounded-lg border px-2 py-1.5 text-sm" value={filter?.op || "before"} onChange={(e) => onChange({ ...filter, op: e.target.value, value: filter?.value ?? "" })}>
          <option value="before">Avant</option><option value="after">Après</option><option value="between">Entre</option>
        </select>
        <input type="date" className="w-full rounded-lg border px-2 py-1.5 text-sm" value={filter?.value ?? ""} onChange={(e) => onChange({ op: filter?.op || "before", value: e.target.value, value2: filter?.value2 })} />
        {filter?.op === "between" && (
          <input type="date" className="w-full rounded-lg border px-2 py-1.5 text-sm" value={filter?.value2 ?? ""} onChange={(e) => onChange({ ...filter, value2: e.target.value })} />
        )}
      </div>
    );
  }
  // texte / personne / fallback
  return (
    <input type="text" autoFocus className="w-full rounded-lg border px-2 py-1.5 text-sm" placeholder="contient…"
      value={filter?.value ?? ""} onChange={(e) => onChange({ value: e.target.value })} />
  );
}

function AddColumnModal({ open, onClose, onCreate, position }) {
  const [name, setName] = useState("");
  const [type, setType] = useState("texte");
  if (!open) return null;
  return (
    <div className="fixed inset-0 z-[60] bg-black/40 flex items-center justify-center" onClick={onClose}>
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-md" onClick={(e) => e.stopPropagation()}>
        <div className="px-5 py-4 border-b flex items-center justify-between">
          <h3 className="font-bold">Ajout de colonne {position === "left" ? "à gauche" : position === "right" ? "à droite" : ""}</h3>
          <button onClick={onClose}><X className="w-5 h-5" /></button>
        </div>
        <div className="p-5 space-y-4">
          <div>
            <label className="block text-xs font-bold text-slate-700 mb-1.5 uppercase tracking-wide">Nom</label>
            <input autoFocus type="text" value={name} onChange={(e) => setName(e.target.value)} placeholder="Ex : Budget"
              className="w-full rounded-xl border px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-slate-900" />
          </div>
          <div>
            <label className="block text-xs font-bold text-slate-700 mb-1.5 uppercase tracking-wide">Type</label>
            <div className="grid grid-cols-2 gap-2">
              {COL_TYPES.map((t) => (
                <button key={t.id} onClick={() => setType(t.id)}
                  className={`flex items-center gap-2 px-3 py-2 rounded-xl border text-sm text-left ${type === t.id ? "bg-slate-900 text-white border-slate-900" : "hover:bg-slate-50"}`}>
                  <ColumnTypeMark type={t.id} />
                  <span>{t.label}</span>
                </button>
              ))}
            </div>
          </div>
        </div>
        <div className="px-5 py-4 border-t flex justify-end gap-2">
          <button onClick={onClose} className="px-4 py-2 rounded-xl text-sm font-semibold hover:bg-slate-100">Annuler</button>
          <button disabled={!name.trim()} onClick={() => { onCreate({ name: name.trim(), type }); setName(""); setType("texte"); }}
            className="px-4 py-2 rounded-xl text-sm font-semibold bg-slate-900 text-white disabled:opacity-40">Créer</button>
        </div>
      </div>
    </div>
  );
}

function ImportPreviewModal({ preview, onClose, onConfirm, onTabChange }) {
  if (!preview) return null;
  const tabs = [
    { id: "athletes", label: "Athlètes", rows: preview.athleteRows || [], required: "Nom complet" },
    { id: "clubs", label: "Sections", rows: preview.sectionRows || [], required: "Section" },
    { id: "projects", label: "Projets", rows: preview.projectRows || [], required: "Nom du projet" },
    { id: "tasks", label: "Tâches", rows: preview.taskRows || [], required: "Titre" },
  ];
  const activeTab = tabs.find((tab) => tab.id === preview.activeTab) || tabs.find((tab) => tab.rows.length) || tabs[0];
  const sampleRows = activeTab.rows.slice(0, 8);
  const headers = Array.from(new Set(sampleRows.flatMap((row) => Object.keys(row || {})))).slice(0, 8);
  const totalRows = tabs.reduce((sum, tab) => sum + tab.rows.length, 0);

  return (
    <div className="fixed inset-0 z-[70] bg-slate-950/50 flex items-center justify-center p-4" onClick={onClose}>
      <div className="bg-white rounded-3xl shadow-2xl w-full max-w-5xl overflow-hidden" onClick={(e) => e.stopPropagation()}>
        <div className="px-6 py-5 border-b flex items-start justify-between gap-4">
          <div>
            <div className="flex items-center gap-2 text-slate-500 text-sm font-semibold mb-1">
              <FileSpreadsheet className="w-4 h-4" />
              {preview.fileName}
            </div>
            <h3 className="text-xl font-bold text-slate-950">Prévisualisation de l'import Excel</h3>
            <p className="text-sm text-slate-600 mt-1">Vérifie les lignes détectées avant d'ajouter ou mettre à jour les données.</p>
          </div>
          <button onClick={onClose} className="p-2 rounded-2xl hover:bg-slate-100 text-slate-500"><X className="w-5 h-5" /></button>
        </div>

        <div className="p-6 space-y-5">
          <div className="grid md:grid-cols-4 gap-3">
            {tabs.map((tab) => (
              <button
                key={tab.id}
                type="button"
                onClick={() => onTabChange(tab.id)}
                className={`rounded-2xl border p-4 text-left transition ${activeTab.id === tab.id ? "border-slate-900 bg-slate-900 text-white" : "bg-white hover:bg-slate-50"}`}
              >
                <span className={`text-xs font-bold uppercase tracking-wide ${activeTab.id === tab.id ? "text-slate-300" : "text-slate-500"}`}>{tab.label}</span>
                <span className="block text-2xl font-bold mt-1">{tab.rows.length}</span>
                <span className={`text-xs ${activeTab.id === tab.id ? "text-slate-300" : "text-slate-500"}`}>champ obligatoire : {tab.required}</span>
              </button>
            ))}
          </div>

          {preview.warnings?.length > 0 && (
            <div className="rounded-2xl border border-amber-200 bg-amber-50 p-4">
              <div className="flex items-center gap-2 text-amber-800 font-semibold text-sm mb-2">
                <AlertTriangle className="w-4 h-4" />
                Points à vérifier
              </div>
              <ul className="space-y-1 text-sm text-amber-800">
                {preview.warnings.map((warning, index) => <li key={`${warning}-${index}`}>{warning}</li>)}
              </ul>
            </div>
          )}

          <div className="rounded-2xl border overflow-hidden">
            <div className="px-4 py-3 bg-slate-50 border-b flex items-center justify-between gap-3">
              <div>
                <p className="font-semibold text-slate-900">{activeTab.label}</p>
                <p className="text-xs text-slate-500">{activeTab.rows.length} ligne(s) détectée(s), aperçu des 8 premières.</p>
              </div>
              <Badge>{totalRows} ligne(s) au total</Badge>
            </div>
            {sampleRows.length ? (
              <div className="overflow-x-auto max-h-80">
                <table className="w-full text-sm">
                  <thead className="bg-white sticky top-0">
                    <tr>
                      {headers.map((header) => <th key={header} className="text-left px-4 py-3 border-b font-bold text-slate-700 whitespace-nowrap">{header}</th>)}
                    </tr>
                  </thead>
                  <tbody>
                    {sampleRows.map((row, rowIndex) => (
                      <tr key={rowIndex} className="odd:bg-white even:bg-slate-50/70">
                        {headers.map((header) => (
                          <td key={header} className="px-4 py-3 border-b text-slate-700 whitespace-nowrap">
                            {String(row[header] ?? "")}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            ) : (
              <div className="p-8 text-center text-sm text-slate-500">Aucune ligne détectée dans cet onglet.</div>
            )}
          </div>
        </div>

        <div className="px-6 py-4 border-t bg-slate-50 flex flex-wrap justify-end gap-2">
          <button onClick={onClose} className="px-4 py-2 rounded-2xl border bg-white text-sm font-semibold hover:bg-slate-100">Annuler</button>
          <button
            disabled={!totalRows}
            onClick={onConfirm}
            className="px-4 py-2 rounded-2xl bg-slate-900 text-white text-sm font-semibold disabled:opacity-40"
          >
            Importer maintenant
          </button>
        </div>
      </div>
    </div>
  );
}

// ===== Generic flat EntityTable (Tasks, Athletes, Clubs, Documents) =====
function EntityTable({ rows, builtinCols, config, onConfigChange, onEdit, onDelete, nameKey = "name", emptyLabel = "Aucun élément.", customFields = [] }) {
  const [openColMenu, setOpenColMenu] = useState(null);
  const [submenu, setSubmenu] = useState(null);
  const [addColModal, setAddColModal] = useState(null);

  const cfg = config || { customCols: [], hiddenCols: [], renamed: {}, typeOverrides: {}, order: [], groupBy: null, filters: {}, sortCol: nameKey, sortDir: "asc" };
  const setCfg = (patch) => onConfigChange({ ...cfg, ...(typeof patch === "function" ? patch(cfg) : patch) });
  const sortCol = cfg.sortCol || nameKey;
  const sortDir = cfg.sortDir || "asc";
  const closeMenu = () => { setOpenColMenu(null); setSubmenu(null); };

  // Build cols
  const customCols = [
    ...(customFields || []).map((field) => ({ key: field.id, label: field.name, type: field.type, builtin: false, sortable: true })),
    ...(cfg.customCols || []).filter((col) => !(customFields || []).some((field) => field.id === col.key)),
  ].map((c) => ({ ...c, builtin: false, sortable: true }));
  let allCols = [...builtinCols.map((c) => ({ ...c, builtin: true })), ...customCols].map((c) => ({
    ...c,
    type: cfg.typeOverrides?.[c.key] || c.type,
    label: cfg.renamed?.[c.key] || c.label,
  }));
  if (cfg.order?.length) {
    const ordered = cfg.order.map((k) => allCols.find((c) => c.key === k)).filter(Boolean);
    const remaining = allCols.filter((c) => !cfg.order.includes(c.key));
    allCols = [...ordered, ...remaining];
  }
  const visibleCols = allCols.filter((c) => !(cfg.hiddenCols || []).includes(c.key));

  const valueFor = (row, col) => {
    if (col.builtin) {
      if (col.accessor) return col.accessor(row);
      return row[col.key];
    }
    return row.customFieldValues?.[col.key];
  };

  const filtered = rows.filter((r) => visibleCols.every((c) => passesFilter(valueFor(r, c), cfg.filters?.[c.key], c.type)));
  const sortFn = (a, b) => {
    const col = allCols.find((c) => c.key === sortCol) || allCols[0];
    const va = valueFor(a, col); const vb = valueFor(b, col);
    if (col.type === "nombre" || col.type === "financier" || col.type === "pourcentage") {
      return sortDir === "asc" ? (Number(va) || 0) - (Number(vb) || 0) : (Number(vb) || 0) - (Number(va) || 0);
    }
    const sa = String(va ?? ""); const sb = String(vb ?? "");
    return sortDir === "asc" ? sa.localeCompare(sb) : sb.localeCompare(sa);
  };
  const sorted = [...filtered].sort(sortFn);
  const activeFilters = Object.entries(cfg.filters || {}).filter(([_, v]) => v && v.value !== "" && v.value !== undefined && v.value !== null);

  // Column actions
  const setFilter = (key, val) => setCfg((c) => ({ ...c, filters: { ...(c.filters || {}), [key]: val } }));
  const clearFilter = (key) => setCfg((c) => { const f = { ...(c.filters || {}) }; delete f[key]; return { ...c, filters: f }; });
  const hideColumn = (key) => setCfg((c) => ({ ...c, hiddenCols: [...new Set([...(c.hiddenCols || []), key])] }));
  const renameColumn = (key) => {
    const cur = cfg.renamed?.[key] || allCols.find((c) => c.key === key)?.label || "";
    const next = window.prompt("Nouveau nom :", cur);
    if (next && next.trim()) setCfg((c) => ({ ...c, renamed: { ...(c.renamed || {}), [key]: next.trim() } }));
  };
  const changeType = (key, type) => setCfg((c) => {
    if (allCols.find((x) => x.key === key)?.builtin) return { ...c, typeOverrides: { ...(c.typeOverrides || {}), [key]: type } };
    return { ...c, customCols: (c.customCols || []).map((x) => x.key === key ? { ...x, type } : x) };
  });
  const duplicateColumn = (key) => setCfg((c) => {
    const src = allCols.find((x) => x.key === key);
    if (!src) return c;
    const newKey = "cf_" + uid().slice(0, 8);
    return { ...c, customCols: [...(c.customCols || []), { key: newKey, label: `Copie de ${src.label}`, type: src.type }] };
  });
  const deleteColumn = (key) => {
    if (!window.confirm("Supprimer cette colonne et toutes ses valeurs ?")) return;
    setCfg((c) => ({ ...c, customCols: (c.customCols || []).filter((x) => x.key !== key) }));
  };
  const setGroupBy = (key) => setCfg({ groupBy: cfg.groupBy === key ? null : key });
  const handleSort = (key, dir) => { setCfg({ sortCol: key, sortDir: dir }); closeMenu(); };
  const addColumn = ({ name, type }) => {
    const newKey = "cf_" + uid().slice(0, 8);
    const anchor = addColModal?.anchorKey;
    setCfg((c) => {
      const existing = [...(c.customCols || []), { key: newKey, label: name, type }];
      let order = c.order?.length ? [...c.order] : allCols.map((x) => x.key);
      if (anchor) {
        const idx = order.indexOf(anchor);
        if (addColModal.position === "left") order.splice(idx, 0, newKey);
        else order.splice(idx + 1, 0, newKey);
      } else order.push(newKey);
      return { ...c, customCols: existing, order };
    });
    setAddColModal(null);
  };

  const renderCell = (row, col) => {
    if (col.render) return col.render(row);
    const v = valueFor(row, col);
    if (col.type === "case_a_cocher") return <span>{v ? "✓" : "—"}</span>;
    if (col.type === "financier" || col.type === "nombre" || col.type === "pourcentage") return <span className="text-slate-700">{formatCustomFieldValue(col.type, v)}</span>;
    if (col.type === "statut" && v && col.options) return <Badge className={badgeClass("status", v)}>{v}</Badge>;
    return <span className="text-slate-600">{v || "—"}</span>;
  };

  const ColHeader = ({ col }) => {
    const hasFilter = cfg.filters?.[col.key] && cfg.filters[col.key].value !== undefined && cfg.filters[col.key].value !== "";
    const isOpen = openColMenu === col.key;
    return (
      <th className="py-3 px-3 text-left font-semibold text-xs text-slate-600 uppercase tracking-wide whitespace-nowrap border-b border-r border-slate-200">
        <div className="relative inline-block">
          <button onClick={(e) => { e.stopPropagation(); setOpenColMenu(isOpen ? null : col.key); setSubmenu(null); }}
            className={`flex items-center gap-1 hover:text-slate-900 ${hasFilter ? "text-blue-600" : ""}`}>
            {col.label}
            {hasFilter && <span className="w-1.5 h-1.5 rounded-full bg-blue-500 inline-block ml-0.5" />}
            <ChevronDown className={`w-3 h-3 transition-transform ${isOpen ? "rotate-180" : ""}`} />
          </button>
          {isOpen && (
            <div className="absolute top-full left-0 z-30 mt-1 bg-white rounded-xl shadow-xl border py-1 min-w-64" onClick={(e) => e.stopPropagation()}>
              <div className="px-3 py-2">
                <p className="text-xs font-semibold text-slate-500 uppercase tracking-wide mb-1.5 flex items-center gap-1.5"><Filter className="w-3 h-3" /> Filtre</p>
                <ColumnFilterUI col={col} filter={cfg.filters?.[col.key]} onChange={(v) => v ? setFilter(col.key, v) : clearFilter(col.key)} />
                {hasFilter && <button onClick={() => clearFilter(col.key)} className="mt-1.5 text-xs text-red-600 hover:underline">✕ Effacer</button>}
              </div>
              <div className="border-t my-1" />
              {col.sortable !== false && (
                <>
                  <button onClick={() => handleSort(col.key, "asc")} className="w-full text-left px-3 py-2 text-sm hover:bg-slate-50 flex items-center gap-2">
                    <span className="text-slate-400 font-mono w-4 text-center">↑</span> Trier croissant {sortCol === col.key && sortDir === "asc" && <CheckCircle2 className="w-3 h-3 text-green-500 ml-auto" />}
                  </button>
                  <button onClick={() => handleSort(col.key, "desc")} className="w-full text-left px-3 py-2 text-sm hover:bg-slate-50 flex items-center gap-2">
                    <span className="text-slate-400 font-mono w-4 text-center">↓</span> Trier décroissant {sortCol === col.key && sortDir === "desc" && <CheckCircle2 className="w-3 h-3 text-green-500 ml-auto" />}
                  </button>
                  <div className="border-t my-1" />
                </>
              )}
              <button onClick={() => { setGroupBy(col.key); closeMenu(); }} className="w-full text-left px-3 py-2 text-sm hover:bg-slate-50 flex items-center gap-2">
                <Layers className="w-4 h-4 text-slate-400" /> {cfg.groupBy === col.key ? "Dégrouper" : "Grouper par"}
                {cfg.groupBy === col.key && <CheckCircle2 className="w-3 h-3 text-green-500 ml-auto" />}
              </button>
              <div className="border-t my-1" />
              <button onClick={() => { duplicateColumn(col.key); closeMenu(); }} className="w-full text-left px-3 py-2 text-sm hover:bg-slate-50 flex items-center gap-2">
                <Copy className="w-4 h-4 text-slate-400" /> Dupliquer la colonne
              </button>
              <button onClick={() => { setAddColModal({ position: "left", anchorKey: col.key }); closeMenu(); }} className="w-full text-left px-3 py-2 text-sm hover:bg-slate-50 flex items-center gap-2">
                <Plus className="w-4 h-4 text-slate-400" /> Ajouter une colonne à gauche
              </button>
              <button onClick={() => { setAddColModal({ position: "right", anchorKey: col.key }); closeMenu(); }} className="w-full text-left px-3 py-2 text-sm hover:bg-slate-50 flex items-center gap-2">
                <Plus className="w-4 h-4 text-slate-400" /> Ajouter une colonne à droite
              </button>
              <div className="relative">
                <button onClick={() => setSubmenu(submenu === "type" ? null : "type")} className="w-full text-left px-3 py-2 text-sm hover:bg-slate-50 flex items-center gap-2">
                  <Layers className="w-4 h-4 text-slate-400 rotate-90" /> Changer le type de colonne <ChevronRight className="w-3 h-3 ml-auto" />
                </button>
                {submenu === "type" && (
                  <div className="absolute left-full top-0 ml-1 bg-white rounded-xl shadow-xl border py-1 min-w-44 z-40">
                    {COL_TYPES.map((t) => (
                      <button key={t.id} onClick={() => { changeType(col.key, t.id); closeMenu(); }}
                        className={`w-full text-left px-3 py-2 text-sm hover:bg-slate-50 flex items-center gap-2 ${col.type === t.id ? "text-slate-900 font-semibold" : ""}`}>
                        {col.type === t.id && <CheckCircle2 className="w-3 h-3 text-green-500" />} {t.label}
                      </button>
                    ))}
                  </div>
                )}
              </div>
              <div className="border-t my-1" />
              <button onClick={() => { renameColumn(col.key); closeMenu(); }} className="w-full text-left px-3 py-2 text-sm hover:bg-slate-50 flex items-center gap-2">
                <Edit3 className="w-4 h-4 text-slate-400" /> Renommer
              </button>
              <button onClick={() => { hideColumn(col.key); closeMenu(); }} className="w-full text-left px-3 py-2 text-sm hover:bg-slate-50 flex items-center gap-2">
                <EyeOff className="w-4 h-4 text-slate-400" /> Masquer
              </button>
              <button onClick={() => { if (!col.builtin) { deleteColumn(col.key); closeMenu(); } }} disabled={col.builtin}
                className={`w-full text-left px-3 py-2 text-sm flex items-center gap-2 ${col.builtin ? "text-slate-300 cursor-not-allowed" : "hover:bg-red-50 text-red-600"}`}>
                <Trash2 className="w-4 h-4" /> Supprimer {col.builtin && <span className="ml-auto text-xs">(intégré)</span>}
              </button>
            </div>
          )}
        </div>
      </th>
    );
  };

  const groupCol = cfg.groupBy ? allCols.find((c) => c.key === cfg.groupBy) : null;
  const grouped = groupCol
    ? sorted.reduce((acc, r) => { const k = String(valueFor(r, groupCol) ?? "—"); (acc[k] = acc[k] || []).push(r); return acc; }, {})
    : { __all__: sorted };
  const groupKeys = Object.keys(grouped).sort();

  return (
    <>
      <div className="mb-3 flex justify-end">
        <button
          type="button"
          onClick={() => setAddColModal({ position: "right", anchorKey: visibleCols[visibleCols.length - 1]?.key })}
          className="inline-flex items-center gap-2 rounded-2xl border bg-white px-4 py-2 text-sm font-semibold text-slate-700 shadow-sm hover:bg-slate-50"
        >
          <Plus className="w-4 h-4" /> Ajouter une colonne
        </button>
      </div>

      {(activeFilters.length > 0 || cfg.groupBy || (cfg.hiddenCols || []).length > 0) && (
        <div className="flex flex-wrap gap-2 mb-3">
          {activeFilters.map(([k, v]) => {
            const c = allCols.find((x) => x.key === k);
            const display = Array.isArray(v.value) ? v.value.join(", ") : (v.op ? `${v.op} ${v.value}` : v.value);
            return (
              <span key={k} className="inline-flex items-center gap-1.5 bg-blue-50 text-blue-700 text-xs font-medium px-2.5 py-1 rounded-full border border-blue-200">
                <Filter className="w-3 h-3" /> {c?.label || k}: <span className="font-bold">{String(display)}</span>
                <button onClick={() => clearFilter(k)} className="hover:text-blue-900"><X className="w-3 h-3" /></button>
              </span>
            );
          })}
          {cfg.groupBy && (
            <span className="inline-flex items-center gap-1.5 bg-violet-50 text-violet-700 text-xs font-medium px-2.5 py-1 rounded-full border border-violet-200">
              <Layers className="w-3 h-3" /> Groupé par <span className="font-bold">{groupCol?.label}</span>
              <button onClick={() => setCfg({ groupBy: null })} className="hover:text-violet-900"><X className="w-3 h-3" /></button>
            </span>
          )}
          {(cfg.hiddenCols || []).length > 0 && (
            <span className="inline-flex items-center gap-1.5 bg-slate-100 text-slate-700 text-xs font-medium px-2.5 py-1 rounded-full border">
              <EyeOff className="w-3 h-3" /> {cfg.hiddenCols.length} colonne{cfg.hiddenCols.length > 1 ? "s" : ""} masquée{cfg.hiddenCols.length > 1 ? "s" : ""}
              <button onClick={() => setCfg({ hiddenCols: [] })} className="hover:text-slate-900"><X className="w-3 h-3" /></button>
            </span>
          )}
        </div>
      )}

      <div className="bg-white rounded-2xl border overflow-x-auto shadow-sm" onClick={() => openColMenu && closeMenu()}>
        <table className="w-full text-sm min-w-[800px]">
          <thead>
            <tr className="border-b bg-slate-50">
              {visibleCols.map((col) => <ColHeader key={col.key} col={col} />)}
              <th className="py-3 px-3 w-10">
                <button onClick={(e) => { e.stopPropagation(); setAddColModal({ position: "right", anchorKey: visibleCols[visibleCols.length - 1]?.key }); }}
                  className="p-1.5 rounded-lg hover:bg-slate-200 text-slate-500" title="Ajouter une colonne">
                  <Plus className="w-4 h-4" />
                </button>
              </th>
              <th className="w-20" />
            </tr>
          </thead>
          <tbody>
            {groupKeys.map((gk) => (
              <React.Fragment key={gk}>
                {groupCol && (
                  <tr className="bg-violet-50/50 border-b">
                    <td colSpan={visibleCols.length + 2} className="py-2 px-3 text-xs font-bold text-violet-700 uppercase tracking-wide">
                      {groupCol.label} : {gk} <span className="text-slate-400">({grouped[gk].length})</span>
                    </td>
                  </tr>
                )}
                {grouped[gk].map((row) => (
                  <tr key={row.id} className="border-b hover:bg-slate-50">
                    {visibleCols.map((col) => (
                      <td key={col.key} className="py-3 px-3">
                        {col.key === nameKey ? (
                          <p className="font-semibold text-slate-900">{row[nameKey] || row.title || "—"}</p>
                        ) : renderCell(row, col)}
                      </td>
                    ))}
                    <td />
                    <td className="py-3 px-3" onClick={(e) => e.stopPropagation()}>
                      <div className="flex gap-1.5 justify-end">
                        {onEdit && <button onClick={() => onEdit(row)} className="p-1.5 rounded-lg hover:bg-slate-100 text-slate-600" title="Modifier"><Edit3 className="w-3.5 h-3.5" /></button>}
                        {onDelete && <button onClick={() => onDelete(row)} className="p-1.5 rounded-lg hover:bg-red-100 text-red-600" title="Supprimer"><Trash2 className="w-3.5 h-3.5" /></button>}
                      </div>
                    </td>
                  </tr>
                ))}
              </React.Fragment>
            ))}
          </tbody>
        </table>
        {!sorted.length && <p className="text-center text-slate-500 py-12">{emptyLabel}</p>}
      </div>

      <AddColumnModal open={!!addColModal} onClose={() => setAddColModal(null)} onCreate={addColumn} position={addColModal?.position} />
    </>
  );
}

function GuidedField({ label, children }) {
  return (
    <div>
      <label className="block text-sm font-bold text-slate-800 mb-2">{label}</label>
      {children}
    </div>
  );
}

function StatusToggle({ options, value, onChange }) {
  return (
    <div className="grid grid-cols-2 gap-2">
      {options.map((s) => (
        <button
          key={s}
          type="button"
          onClick={() => onChange(s)}
          className={`py-2.5 rounded-xl border text-sm font-medium transition-colors ${value === s ? "bg-slate-900 text-white border-slate-900" : "hover:bg-slate-50"}`}
        >
          {s}
        </button>
      ))}
    </div>
  );
}

function ProjectTable({ projects, tasks, onEdit, requestDelete, onQuickTask, onNewSubProject, config, onConfigChange, customFields = [] }) {
  const [expandedRows, setExpandedRows] = useState({});
  const [openColMenu, setOpenColMenu] = useState(null);
  const [submenu, setSubmenu] = useState(null); // 'sort' | 'type'
  const [addColModal, setAddColModal] = useState(null); // { position, anchorKey }
  const [selectedIds, setSelectedIds] = useState(new Set());

  const cfg = config || { customCols: [], hiddenCols: [], renamed: {}, typeOverrides: {}, order: [], groupBy: null, filters: {}, sortCol: "name", sortDir: "asc" };
  const setCfg = (patch) => onConfigChange({ ...cfg, ...(typeof patch === "function" ? patch(cfg) : patch) });
  const sortCol = cfg.sortCol || "name";
  const sortDir = cfg.sortDir || "asc";

  const toggleRow = (id) => setExpandedRows((prev) => ({ ...prev, [id]: !prev[id] }));
  const closeMenu = () => { setOpenColMenu(null); setSubmenu(null); };
  const toggleSelected = (id) => setSelectedIds((prev) => {
    const next = new Set(prev);
    if (next.has(id)) next.delete(id);
    else next.add(id);
    return next;
  });
  const clearSelection = () => setSelectedIds(new Set());

  const handleSort = (key, dir) => { setCfg({ sortCol: key, sortDir: dir }); closeMenu(); };

  // ===== Column model =====
  const builtinCols = [
    { key: "name", label: "Projet", type: "texte", builtin: true, sortable: true },
    { key: "status", label: "Statut", type: "statut", builtin: true, sortable: true, options: STATUS_OPTIONS },
    { key: "priority", label: "Priorité", type: "statut", builtin: true, sortable: true, options: PRIORITY_OPTIONS },
    { key: "computedTotalProgress", label: "Avancement", type: "nombre", builtin: true, sortable: true },
    { key: "tasksCount", label: "Tâches", type: "nombre", builtin: true, sortable: false },
    { key: "owner", label: "Responsable", type: "personne", builtin: true, sortable: true },
    { key: "endDate", label: "Échéance", type: "date", builtin: true, sortable: true },
  ];
  const customCols = [
    ...(customFields || []).map((field) => ({ key: field.id, label: field.name, type: field.type, builtin: false, sortable: true })),
    ...(cfg.customCols || []).filter((col) => !(customFields || []).some((field) => field.id === col.key)),
  ].map((c) => ({ ...c, builtin: false, sortable: true }));
  // apply type overrides + renames
  let allCols = [...builtinCols, ...customCols].map((c) => ({
    ...c,
    type: cfg.typeOverrides?.[c.key] || c.type,
    label: cfg.renamed?.[c.key] || c.label,
  }));
  // apply order
  if (cfg.order?.length) {
    const ordered = cfg.order.map((k) => allCols.find((c) => c.key === k)).filter(Boolean);
    const remaining = allCols.filter((c) => !cfg.order.includes(c.key));
    allCols = [...ordered, ...remaining];
  }
  // hide
  const visibleCols = allCols.filter((c) => !(cfg.hiddenCols || []).includes(c.key));

  // ===== Filtering =====
  const valueFor = (project, col) => {
    if (col.builtin) {
      if (col.key === "tasksCount") return tasks.filter((t) => t.projectId === project.id).length;
      return project[col.key];
    }
    return project.customFieldValues?.[col.key];
  };
  const filtered = projects.filter((p) => visibleCols.every((c) => passesFilter(valueFor(p, c), cfg.filters?.[c.key], c.type)));

  // Build children index from the FULL projects list (so hierarchy is preserved)
  const childrenByParent = {};
  projects.forEach((p) => {
    const k = p.parentId || "__root__";
    (childrenByParent[k] = childrenByParent[k] || []).push(p);
  });

  const sortFn = (a, b) => {
    const col = allCols.find((c) => c.key === sortCol) || builtinCols[0];
    const va = valueFor(a, col); const vb = valueFor(b, col);
    if (col.type === "nombre" || col.type === "financier" || col.type === "pourcentage") {
      return sortDir === "asc" ? (Number(va) || 0) - (Number(vb) || 0) : (Number(vb) || 0) - (Number(va) || 0);
    }
    const sa = String(va ?? ""); const sb = String(vb ?? "");
    return sortDir === "asc" ? sa.localeCompare(sb) : sb.localeCompare(sa);
  };

  const filteredIds = new Set(filtered.map((p) => p.id));
  const roots = filtered.filter((p) => !p.parentId || !filteredIds.has(p.parentId)).sort(sortFn);
  const activeFilters = Object.entries(cfg.filters || {}).filter(([_, v]) => v && v.value !== "" && v.value !== undefined && v.value !== null);
  const completedCount = filtered.filter((project) => project.status === "Terminé").length;
  const visibleTaskCount = filtered.reduce((sum, project) => sum + tasks.filter((task) => task.projectId === project.id).length, 0);
  const selectedCount = selectedIds.size;

  // ===== Cell renderer =====
  const renderCell = (project, col) => {
    const v = valueFor(project, col);
    if (col.builtin && col.key === "name") return null; // handled in special name cell
    if (col.builtin && col.key === "status") return <span className={`block w-full rounded-xl px-3 py-2 text-center text-xs font-bold ${statusCellClass(project.status)}`}>{project.status}</span>;
    if (col.builtin && col.key === "priority") return (
      <span className="inline-flex items-center gap-2 text-slate-700">
        <span className={`h-2.5 w-2.5 rounded-full ${priorityDotClass(project.priority)}`} />
        <span className="text-sm font-medium">{project.priority || "—"}</span>
      </span>
    );
    if (col.builtin && col.key === "computedTotalProgress") return (
      <div className="flex items-center gap-2 min-w-32">
        <div className="flex-1 h-2 rounded-full bg-slate-200 overflow-hidden">
          <div className={`h-full rounded-full ${progressTone(project.computedTotalProgress)}`} style={{ width: `${project.computedTotalProgress}%` }} />
        </div>
        <span className="text-xs font-bold text-slate-700 w-9 text-right">{project.computedTotalProgress}%</span>
      </div>
    );
    if (col.builtin && col.key === "tasksCount") {
      const n = tasks.filter((t) => t.projectId === project.id).length;
      return <span className="text-slate-600"><span className="font-medium text-slate-800">{n}</span> tâche{n > 1 ? "s" : ""}</span>;
    }
    if (col.type === "case_a_cocher") return <span>{v ? "✓" : "—"}</span>;
    if (col.type === "financier" || col.type === "nombre" || col.type === "pourcentage") return <span className="text-slate-700">{formatCustomFieldValue(col.type, v)}</span>;
    if (col.type === "statut" && v) return <span className={`block w-full rounded-xl px-3 py-2 text-center text-xs font-bold ${statusCellClass(v)}`}>{v}</span>;
    return <span className="text-slate-600">{v || "—"}</span>;
  };

  // ===== Column actions =====
  const setFilter = (key, val) => setCfg((c) => ({ ...c, filters: { ...(c.filters || {}), [key]: val } }));
  const clearFilter = (key) => setCfg((c) => { const f = { ...(c.filters || {}) }; delete f[key]; return { ...c, filters: f }; });
  const hideColumn = (key) => setCfg((c) => ({ ...c, hiddenCols: [...new Set([...(c.hiddenCols || []), key])] }));
  const renameColumn = (key) => {
    const cur = cfg.renamed?.[key] || allCols.find((c) => c.key === key)?.label || "";
    const next = window.prompt("Nouveau nom :", cur);
    if (next && next.trim()) setCfg((c) => ({ ...c, renamed: { ...(c.renamed || {}), [key]: next.trim() } }));
  };
  const changeType = (key, type) => setCfg((c) => {
    if (allCols.find((x) => x.key === key)?.builtin) return { ...c, typeOverrides: { ...(c.typeOverrides || {}), [key]: type } };
    return { ...c, customCols: (c.customCols || []).map((x) => x.key === key ? { ...x, type } : x) };
  });
  const duplicateColumn = (key) => setCfg((c) => {
    const src = allCols.find((x) => x.key === key);
    if (!src) return c;
    const newKey = "cf_" + uid().slice(0, 8);
    const newCol = { key: newKey, label: `Copie de ${src.label}`, type: src.type };
    return { ...c, customCols: [...(c.customCols || []), newCol] };
  });
  const deleteColumn = (key) => {
    if (!window.confirm("Supprimer cette colonne et toutes ses valeurs ?")) return;
    setCfg((c) => ({ ...c, customCols: (c.customCols || []).filter((x) => x.key !== key) }));
  };
  const setGroupBy = (key) => setCfg({ groupBy: cfg.groupBy === key ? null : key });
  const addColumn = ({ name, type }) => {
    const newKey = "cf_" + uid().slice(0, 8);
    const newCol = { key: newKey, label: name, type };
    const anchor = addColModal?.anchorKey;
    setCfg((c) => {
      const existing = [...(c.customCols || []), newCol];
      let order = c.order?.length ? [...c.order] : allCols.map((x) => x.key);
      if (anchor) {
        const idx = order.indexOf(anchor);
        if (addColModal.position === "left") order.splice(idx, 0, newKey);
        else order.splice(idx + 1, 0, newKey);
      } else order.push(newKey);
      return { ...c, customCols: existing, order };
    });
    setAddColModal(null);
  };

  const ColHeader = ({ col }) => {
    const hasFilter = cfg.filters?.[col.key] && cfg.filters[col.key].value !== undefined && cfg.filters[col.key].value !== "";
    const isOpen = openColMenu === col.key;
    return (
      <th className="py-3 px-3 text-left font-semibold text-xs text-slate-600 uppercase tracking-wide whitespace-nowrap">
        <div className="relative inline-block">
          <button onClick={(e) => { e.stopPropagation(); setOpenColMenu(isOpen ? null : col.key); setSubmenu(null); }}
            className={`flex items-center gap-1 hover:text-slate-900 ${hasFilter ? "text-blue-600" : ""}`}>
            {col.label}
            {hasFilter && <span className="w-1.5 h-1.5 rounded-full bg-blue-500 inline-block ml-0.5" />}
            <ChevronDown className={`w-3 h-3 transition-transform ${isOpen ? "rotate-180" : ""}`} />
          </button>
          {isOpen && (
            <div className="absolute top-full left-0 z-30 mt-1 bg-white rounded-xl shadow-xl border py-1 min-w-64" onClick={(e) => e.stopPropagation()}>
              {/* Filtre */}
              <div className="px-3 py-2">
                <p className="text-xs font-semibold text-slate-500 uppercase tracking-wide mb-1.5 flex items-center gap-1.5"><Filter className="w-3 h-3" /> Filtre</p>
                <ColumnFilterUI col={col} filter={cfg.filters?.[col.key]} onChange={(v) => v ? setFilter(col.key, v) : clearFilter(col.key)} />
                {hasFilter && (
                  <button onClick={() => { clearFilter(col.key); }} className="mt-1.5 text-xs text-red-600 hover:underline">✕ Effacer</button>
                )}
              </div>
              <div className="border-t my-1" />
              {/* Trier */}
              {col.sortable && (
                <>
                  <button onClick={() => handleSort(col.key, "asc")} className="w-full text-left px-3 py-2 text-sm hover:bg-slate-50 flex items-center gap-2">
                    <span className="text-slate-400 font-mono w-4 text-center">↑</span> Trier croissant {sortCol === col.key && sortDir === "asc" && <CheckCircle2 className="w-3 h-3 text-green-500 ml-auto" />}
                  </button>
                  <button onClick={() => handleSort(col.key, "desc")} className="w-full text-left px-3 py-2 text-sm hover:bg-slate-50 flex items-center gap-2">
                    <span className="text-slate-400 font-mono w-4 text-center">↓</span> Trier décroissant {sortCol === col.key && sortDir === "desc" && <CheckCircle2 className="w-3 h-3 text-green-500 ml-auto" />}
                  </button>
                  <div className="border-t my-1" />
                </>
              )}
              {/* Grouper par */}
              <button onClick={() => { setGroupBy(col.key); closeMenu(); }} className="w-full text-left px-3 py-2 text-sm hover:bg-slate-50 flex items-center gap-2">
                <Layers className="w-4 h-4 text-slate-400" /> {cfg.groupBy === col.key ? "Dégrouper" : "Grouper par"}
                {cfg.groupBy === col.key && <CheckCircle2 className="w-3 h-3 text-green-500 ml-auto" />}
              </button>
              <div className="border-t my-1" />
              {/* Dupliquer */}
              <button onClick={() => { duplicateColumn(col.key); closeMenu(); }} className="w-full text-left px-3 py-2 text-sm hover:bg-slate-50 flex items-center gap-2">
                <Copy className="w-4 h-4 text-slate-400" /> Dupliquer la colonne
              </button>
              <button onClick={() => { setAddColModal({ position: "left", anchorKey: col.key }); closeMenu(); }} className="w-full text-left px-3 py-2 text-sm hover:bg-slate-50 flex items-center gap-2">
                <Plus className="w-4 h-4 text-slate-400" /> Ajouter une colonne à gauche
              </button>
              <button onClick={() => { setAddColModal({ position: "right", anchorKey: col.key }); closeMenu(); }} className="w-full text-left px-3 py-2 text-sm hover:bg-slate-50 flex items-center gap-2">
                <Plus className="w-4 h-4 text-slate-400" /> Ajouter une colonne à droite
              </button>
              {/* Changer le type */}
              <div className="relative">
                <button onClick={() => setSubmenu(submenu === "type" ? null : "type")} className="w-full text-left px-3 py-2 text-sm hover:bg-slate-50 flex items-center gap-2">
                  <Layers className="w-4 h-4 text-slate-400 rotate-90" /> Changer le type de colonne <ChevronRight className="w-3 h-3 ml-auto" />
                </button>
                {submenu === "type" && (
                  <div className="absolute left-full top-0 ml-1 bg-white rounded-xl shadow-xl border py-1 min-w-44 z-40">
                    {COL_TYPES.map((t) => (
                      <button key={t.id} onClick={() => { changeType(col.key, t.id); closeMenu(); }}
                        className={`w-full text-left px-3 py-2 text-sm hover:bg-slate-50 flex items-center gap-2 ${col.type === t.id ? "text-slate-900 font-semibold" : ""}`}>
                        {col.type === t.id && <CheckCircle2 className="w-3 h-3 text-green-500" />} {t.label}
                      </button>
                    ))}
                  </div>
                )}
              </div>
              <div className="border-t my-1" />
              <button onClick={() => { renameColumn(col.key); closeMenu(); }} className="w-full text-left px-3 py-2 text-sm hover:bg-slate-50 flex items-center gap-2">
                <Edit3 className="w-4 h-4 text-slate-400" /> Renommer
              </button>
              <button onClick={() => { hideColumn(col.key); closeMenu(); }} className="w-full text-left px-3 py-2 text-sm hover:bg-slate-50 flex items-center gap-2">
                <EyeOff className="w-4 h-4 text-slate-400" /> Masquer
              </button>
              <button onClick={() => { if (!col.builtin) { deleteColumn(col.key); closeMenu(); } }} disabled={col.builtin}
                className={`w-full text-left px-3 py-2 text-sm flex items-center gap-2 ${col.builtin ? "text-slate-300 cursor-not-allowed" : "hover:bg-red-50 text-red-600"}`}>
                <Trash2 className="w-4 h-4" /> Supprimer {col.builtin && <span className="ml-auto text-xs">(intégré)</span>}
              </button>
            </div>
          )}
        </div>
      </th>
    );
  };

  // ===== Render rows =====
  const nameCol = visibleCols.find((c) => c.key === "name");
  const otherCols = visibleCols.filter((c) => c.key !== "name");

  const renderProjectRow = (project, depth) => {
    const projectTasks = tasks.filter((t) => t.projectId === project.id);
    const childProjects = (childrenByParent[project.id] || []).filter((c) => filteredIds.has(c.id)).sort(sortFn);
    const hasChildren = childProjects.length > 0 || projectTasks.length > 0;
    const isOpen = expandedRows[project.id];
    const indentPx = depth * 24;
    const rows = [];

    rows.push(
      <tr key={project.id} className={`group border-b border-slate-200 hover:bg-slate-50/80 ${depth > 0 ? "bg-slate-50/30" : "bg-white"}`}>
        <td className="py-2.5 px-3 border-r border-slate-200">
          <input
            type="checkbox"
            checked={selectedIds.has(project.id)}
            onChange={() => toggleSelected(project.id)}
            className="h-4 w-4 rounded border-slate-300 text-slate-900"
            aria-label={`Sélectionner ${project.name}`}
          />
        </td>
        <td className="py-2.5 px-2 text-slate-400 cursor-pointer border-r border-slate-200" onClick={() => hasChildren && toggleRow(project.id)}>
          {hasChildren ? (isOpen ? <ChevronDown className="w-4 h-4" /> : <ChevronRight className="w-4 h-4" />) : <div className="w-4" />}
        </td>
        {nameCol && (
          <td className="py-2.5 px-3 cursor-pointer min-w-[280px] border-r border-slate-200" onClick={() => hasChildren && toggleRow(project.id)}>
            <div className="relative" style={{ paddingLeft: indentPx + 10 }}>
              <span className={`absolute left-0 top-1 h-8 w-1 rounded-full ${project.status === "Terminé" ? "bg-emerald-500" : project.status === "En cours" ? "bg-blue-500" : project.status === "En attente" ? "bg-amber-400" : "bg-slate-300"}`} />
              <div className="flex items-center gap-2">
                {depth > 0 && <span className="text-slate-300 text-xs">↳</span>}
                <p className="font-semibold text-slate-900">{project.name}</p>
                {childProjects.length > 0 && <span className="text-xs text-slate-400">({childProjects.length} sous-projet{childProjects.length > 1 ? "s" : ""})</span>}
              </div>
              <p className="text-xs text-slate-400 mt-0.5">{project.sportName || "Sport non renseigné"}</p>
            </div>
          </td>
        )}
        {otherCols.map((col) => <td key={col.key} className="py-2.5 px-3 border-r border-slate-200 align-middle">{renderCell(project, col)}</td>)}
        {/* spacer for + col */}
        <td className="border-r border-slate-200" />
        {/* Actions */}
        <td className="py-2.5 px-3 sticky right-0 bg-white/95 backdrop-blur border-l border-slate-200" onClick={(e) => e.stopPropagation()}>
          <div className="flex gap-1.5 justify-end opacity-70 group-hover:opacity-100">
            <button onClick={() => onEdit(project)} className="p-1.5 rounded-lg hover:bg-slate-100 text-slate-600" title="Modifier"><Edit3 className="w-3.5 h-3.5" /></button>
            <button onClick={() => onQuickTask(project.id)} className="p-1.5 rounded-lg hover:bg-blue-100 text-blue-700" title="Ajouter tâche"><Plus className="w-3.5 h-3.5" /></button>
            <button onClick={() => onNewSubProject(project)} className="p-1.5 rounded-lg hover:bg-violet-100 text-violet-700" title="Créer un sous-projet"><FolderKanban className="w-3.5 h-3.5" /></button>
            <button onClick={() => requestDelete("project", project)} className="p-1.5 rounded-lg hover:bg-red-100 text-red-600" title="Supprimer"><Trash2 className="w-3.5 h-3.5" /></button>
          </div>
        </td>
      </tr>
    );

    if (isOpen) {
      childProjects.forEach((child) => rows.push(...renderProjectRow(child, depth + 1)));
      projectTasks.forEach((task) => {
        const prog = normalizeTaskProgress(task.progress, task.status);
        rows.push(
          <tr key={task.id} className="border-b border-slate-200 bg-slate-50/70 text-sm">
            <td className="border-r border-slate-200" />
            <td className="border-r border-slate-200" />
            {nameCol && (
              <td className="py-2.5 px-3 border-r border-slate-200">
                <div style={{ paddingLeft: indentPx + 24 }}>
                  <div className="flex items-center gap-2">
                    <CheckCircle2 className={`w-3.5 h-3.5 shrink-0 ${task.status === "Terminé" ? "text-green-500" : "text-slate-300"}`} />
                    <span className="text-slate-800 font-medium">{task.title}</span>
                  </div>
                  <p className="text-xs text-slate-400 mt-0.5 pl-5">{task.assignee || "Non assigné"} · {task.endDate || "Sans échéance"}</p>
                </div>
              </td>
            )}
            {otherCols.map((col) => {
              if (col.key === "status") return <td key={col.key} className="py-2.5 px-3 border-r border-slate-200"><span className={`block w-full rounded-xl px-3 py-2 text-center text-xs font-bold ${statusCellClass(task.status)}`}>{task.status}</span></td>;
              if (col.key === "priority") return (
                <td key={col.key} className="py-2.5 px-3 border-r border-slate-200">
                  <span className="inline-flex items-center gap-2 text-slate-700">
                    <span className={`h-2.5 w-2.5 rounded-full ${priorityDotClass(task.priority)}`} />
                    <span className="text-sm font-medium">{task.priority || "—"}</span>
                  </span>
                </td>
              );
              if (col.key === "computedTotalProgress") return (
                <td key={col.key} className="py-2.5 px-3 min-w-32 border-r border-slate-200">
                  <div className="flex items-center gap-2">
                    <div className="flex-1 h-1.5 rounded-full bg-slate-200 overflow-hidden">
                      <div className={`h-full rounded-full ${progressTone(prog)}`} style={{ width: `${prog}%` }} />
                    </div>
                    <span className="text-xs font-bold text-slate-600 w-8 text-right">{prog}%</span>
                  </div>
                </td>
              );
              return <td key={col.key} className="border-r border-slate-200" />;
            })}
            <td className="border-r border-slate-200" /><td />
          </tr>
        );
      });
    }
    return rows;
  };

  // ===== Grouping =====
  const groupCol = cfg.groupBy ? allCols.find((c) => c.key === cfg.groupBy) : null;
  const groupedRoots = groupCol
    ? roots.reduce((acc, p) => { const k = String(valueFor(p, groupCol) ?? "—"); (acc[k] = acc[k] || []).push(p); return acc; }, {})
    : { __all__: roots };
  const groupKeys = Object.keys(groupedRoots).sort();

  return (
    <>
      <div className="mb-3 rounded-3xl border border-slate-200 bg-white p-3 shadow-sm">
        <div className="flex flex-wrap items-center justify-between gap-3">
          <div className="flex flex-wrap items-center gap-2">
            <span className="inline-flex h-9 w-9 items-center justify-center rounded-2xl bg-slate-900 text-white">
              <FolderKanban className="w-4 h-4" />
            </span>
            <div>
              <p className="text-sm font-bold text-slate-900">Plan de pilotage JAM</p>
              <p className="text-xs text-slate-500">{filtered.length} projet{filtered.length > 1 ? "s" : ""} · {visibleTaskCount} tâche{visibleTaskCount > 1 ? "s" : ""} · {completedCount} terminé{completedCount > 1 ? "s" : ""}</p>
            </div>
          </div>
          <div className="flex flex-wrap items-center gap-2">
            {selectedCount > 0 && (
              <button onClick={clearSelection} className="inline-flex items-center gap-2 rounded-2xl border border-slate-200 bg-slate-50 px-3 py-2 text-sm font-semibold text-slate-700 hover:bg-slate-100">
                {selectedCount} sélectionné{selectedCount > 1 ? "s" : ""}
                <X className="w-4 h-4" />
              </button>
            )}
            <button
              type="button"
              onClick={() => setCfg({ hiddenCols: [] })}
              className="inline-flex items-center gap-2 rounded-2xl border bg-white px-3 py-2 text-sm font-semibold text-slate-700 hover:bg-slate-50"
            >
              <EyeOff className="w-4 h-4" /> Afficher colonnes
            </button>
            <button
              type="button"
              onClick={() => setAddColModal({ position: "right", anchorKey: visibleCols[visibleCols.length - 1]?.key })}
              className="inline-flex items-center gap-2 rounded-2xl bg-slate-900 px-4 py-2 text-sm font-semibold text-white shadow-sm hover:bg-slate-800"
            >
              <Plus className="w-4 h-4" /> Ajouter colonne
            </button>
          </div>
        </div>
      </div>

      {/* Active filter chips */}
      {(activeFilters.length > 0 || cfg.groupBy) && (
        <div className="flex flex-wrap gap-2 mb-3">
          {activeFilters.map(([k, v]) => {
            const c = allCols.find((x) => x.key === k);
            const display = Array.isArray(v.value) ? v.value.join(", ") : (v.op ? `${v.op} ${v.value}` : v.value);
            return (
              <span key={k} className="inline-flex items-center gap-1.5 bg-blue-50 text-blue-700 text-xs font-medium px-2.5 py-1 rounded-full border border-blue-200">
                <Filter className="w-3 h-3" /> {c?.label || k}: <span className="font-bold">{String(display)}</span>
                <button onClick={() => clearFilter(k)} className="hover:text-blue-900"><X className="w-3 h-3" /></button>
              </span>
            );
          })}
          {cfg.groupBy && (
            <span className="inline-flex items-center gap-1.5 bg-violet-50 text-violet-700 text-xs font-medium px-2.5 py-1 rounded-full border border-violet-200">
              <Layers className="w-3 h-3" /> Groupé par <span className="font-bold">{groupCol?.label}</span>
              <button onClick={() => setCfg({ groupBy: null })} className="hover:text-violet-900"><X className="w-3 h-3" /></button>
            </span>
          )}
          {(cfg.hiddenCols || []).length > 0 && (
            <span className="inline-flex items-center gap-1.5 bg-slate-100 text-slate-700 text-xs font-medium px-2.5 py-1 rounded-full border">
              <EyeOff className="w-3 h-3" /> {cfg.hiddenCols.length} colonne{cfg.hiddenCols.length > 1 ? "s" : ""} masquée{cfg.hiddenCols.length > 1 ? "s" : ""}
              <button onClick={() => setCfg({ hiddenCols: [] })} className="hover:text-slate-900"><X className="w-3 h-3" /></button>
            </span>
          )}
        </div>
      )}

      <div className="bg-white rounded-3xl border border-slate-200 overflow-hidden shadow-sm" onClick={() => openColMenu && closeMenu()}>
        <div className="overflow-x-auto">
        <table className="w-full text-sm min-w-[980px] border-separate border-spacing-0">
          <thead className="sticky top-0 z-10">
            <tr className="bg-slate-100/95 backdrop-blur">
              <th className="w-10 py-3 px-3 border-b border-r border-slate-200">
                <span className="block h-4 w-4 rounded border border-slate-300 bg-white" />
              </th>
              <th className="w-8 py-3 px-2 border-b border-r border-slate-200" />
              {visibleCols.map((col) => <ColHeader key={col.key} col={col} />)}
              <th className="py-3 px-3 w-10 border-b border-r border-slate-200">
                <button onClick={(e) => { e.stopPropagation(); setAddColModal({ position: "right", anchorKey: visibleCols[visibleCols.length - 1]?.key }); }}
                  className="p-1.5 rounded-lg hover:bg-slate-200 text-slate-500" title="Ajouter une colonne">
                  <Plus className="w-4 h-4" />
                </button>
              </th>
              <th className="w-24 border-b border-slate-200" />
            </tr>
          </thead>
          <tbody>
            {groupKeys.map((gk) => (
              <React.Fragment key={gk}>
                {groupCol && (
                  <tr className="bg-violet-50/50 border-b">
                    <td colSpan={visibleCols.length + 4} className="py-2 px-3 text-xs font-bold text-violet-700 uppercase tracking-wide">
                      {groupCol.label} : {gk}  <span className="text-slate-400">({groupedRoots[gk].length})</span>
                    </td>
                  </tr>
                )}
                {groupedRoots[gk].flatMap((p) => renderProjectRow(p, 0))}
              </React.Fragment>
            ))}
          </tbody>
        </table>
        </div>
        {!roots.length && <p className="text-center text-slate-500 py-12">Aucun projet trouvé.</p>}
      </div>

      <AddColumnModal open={!!addColModal} onClose={() => setAddColModal(null)} onCreate={addColumn} position={addColModal?.position} />
    </>
  );
}

function GuidedInputPanel({ open, onClose, allSports, projects, sections = [], onSave }) {
  const [type, setType] = useState(null);
  const [form, setForm] = useState({});

  const ff = (key, val) => setForm((prev) => ({ ...prev, [key]: val }));
  const close = () => { setType(null); setForm({}); onClose(); };

  const types = [
    { id: "project", label: "Projet", Icon: FolderKanban, color: "bg-violet-100 text-violet-700" },
    { id: "task", label: "Tâche", Icon: ListChecks, color: "bg-blue-100 text-blue-700" },
    { id: "athlete", label: "Athlète", Icon: Users, color: "bg-green-100 text-green-700" },
    { id: "club", label: "Section", Icon: Building2, color: "bg-amber-100 text-amber-700" },
  ];

  const handleSave = () => {
    if (!type || !(form.name || form.title)) return;
    onSave(type, { status: "À faire", priority: "Moyenne", sport: allSports[0]?.id || "developpe_couche", customFieldValues: {}, ...form });
    close();
  };

  if (!open) return null;

  const inp = "w-full rounded-2xl border px-4 py-3 text-sm focus:outline-none focus:ring-2 focus:ring-slate-900";
  const hasProjectParent = projects.length > 0;

  return (
    <div className="fixed inset-0 z-50 flex">
      <div className="flex-1 bg-black/30" onClick={close} />
      <div className="w-[480px] bg-white h-full shadow-2xl flex flex-col">
        <div className="px-6 py-5 border-b flex items-center justify-between shrink-0">
          <h2 className="font-bold text-lg text-slate-900">Saisie guidée</h2>
          <button onClick={close} className="text-slate-400 hover:text-slate-700"><X className="w-5 h-5" /></button>
        </div>

        <div className="flex-1 overflow-auto p-6 space-y-7">
          {/* Type selector */}
          <GuidedField label="Type d'élément">
            <div className="grid grid-cols-2 gap-2">
              {types.map((t) => (
                <button key={t.id} type="button"
                  onClick={() => { setType(t.id); setForm({ status: "À faire", priority: "Moyenne", sport: allSports[0]?.id || "developpe_couche", parentId: null, customFieldValues: {} }); }}
                  className={`flex items-center gap-3 p-3 rounded-2xl border-2 text-left transition-all ${type === t.id ? "border-slate-900 bg-slate-50" : "border-slate-200 hover:border-slate-400"}`}>
                  <div className={`w-8 h-8 rounded-xl ${t.color} flex items-center justify-center shrink-0`}>
                    <t.Icon className="w-4 h-4" />
                  </div>
                  <span className="font-semibold text-sm text-slate-900">{t.label}</span>
                </button>
              ))}
            </div>
          </GuidedField>

          {type === "project" && <>
            <GuidedField label="Nom du projet *">
              <input autoFocus className={inp} placeholder="ex : Challenge régional 2026" value={form.name || ""} onChange={(e) => ff("name", e.target.value)} />
            </GuidedField>
            <GuidedField label="Description">
              <textarea className={`${inp} resize-none h-24`} placeholder="Objectif, contexte…" value={form.description || ""} onChange={(e) => ff("description", e.target.value)} />
            </GuidedField>
            <GuidedField label="Sport">
              <select className={inp} value={form.sport || ""} onChange={(e) => ff("sport", e.target.value)}>
                {allSports.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
              </select>
            </GuidedField>
            <GuidedField label="Priorité">
              <StatusToggle options={PRIORITY_OPTIONS} value={form.priority} onChange={(v) => ff("priority", v)} />
            </GuidedField>
            <div className="grid grid-cols-2 gap-4">
              <GuidedField label="Date de début"><input type="date" className={inp} value={form.startDate || ""} onChange={(e) => ff("startDate", e.target.value)} /></GuidedField>
              <GuidedField label="Date de fin"><input type="date" className={inp} value={form.endDate || ""} onChange={(e) => ff("endDate", e.target.value)} /></GuidedField>
            </div>
            <GuidedField label="Responsable">
              <input className={inp} placeholder="Nom du responsable" value={form.owner || ""} onChange={(e) => ff("owner", e.target.value)} />
            </GuidedField>
            <GuidedField label="Organisation du projet">
              <div className="grid grid-cols-2 gap-2">
                <button
                  type="button"
                  onClick={() => ff("parentId", null)}
                  className={`rounded-2xl border px-3 py-3 text-sm font-semibold ${!form.parentId ? "border-slate-900 bg-slate-900 text-white" : "bg-white hover:bg-slate-50"}`}
                >
                  Projet principal
                </button>
                <button
                  type="button"
                  disabled={!hasProjectParent}
                  onClick={() => ff("parentId", form.parentId || projects[0]?.id || null)}
                  className={`rounded-2xl border px-3 py-3 text-sm font-semibold disabled:cursor-not-allowed disabled:opacity-40 ${form.parentId ? "border-slate-900 bg-slate-900 text-white" : "bg-white hover:bg-slate-50"}`}
                >
                  Sous-projet
                </button>
              </div>
              {form.parentId && (
                <select className={`${inp} mt-2`} value={form.parentId || ""} onChange={(e) => ff("parentId", e.target.value || null)}>
                  {projects.map((p) => <option key={p.id} value={p.id}>Rattaché à : {p.name}</option>)}
                </select>
              )}
            </GuidedField>
          </>}

          {type === "task" && <>
            <GuidedField label="Titre de la tâche *">
              <input autoFocus className={inp} placeholder="ex : Envoyer les convocations" value={form.title || ""} onChange={(e) => ff("title", e.target.value)} />
            </GuidedField>
            <GuidedField label="Projet associé">
              <select className={inp} value={form.projectId || ""} onChange={(e) => ff("projectId", e.target.value)}>
                <option value="">Aucun projet</option>
                {projects.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
              </select>
            </GuidedField>
            <GuidedField label="Assigné à">
              <input className={inp} placeholder="Nom" value={form.assignee || ""} onChange={(e) => ff("assignee", e.target.value)} />
            </GuidedField>
            <GuidedField label="Priorité">
              <StatusToggle options={PRIORITY_OPTIONS} value={form.priority} onChange={(v) => ff("priority", v)} />
            </GuidedField>
            <GuidedField label="Statut">
              <StatusToggle options={STATUS_OPTIONS} value={form.status} onChange={(v) => ff("status", v)} />
            </GuidedField>
            <GuidedField label="Échéance">
              <input type="date" className={inp} value={form.endDate || ""} onChange={(e) => ff("endDate", e.target.value)} />
            </GuidedField>
          </>}

          {type === "athlete" && <>
            <GuidedField label="Nom complet *">
              <input autoFocus className={inp} placeholder="Prénom Nom" value={form.name || ""} onChange={(e) => ff("name", e.target.value)} />
            </GuidedField>
            <GuidedField label="Section">
              <select
                className={inp}
                value={form.sectionId || ""}
                onChange={(e) => {
                  const section = sections.find((item) => item.id === e.target.value);
                  setForm((prev) => ({ ...prev, sectionId: e.target.value, subSection: "", sport: section?.sport || prev.sport }));
                }}
              >
                <option value="">Section JAM</option>
                {sections.map((section) => <option key={section.id} value={section.id}>{section.name}</option>)}
              </select>
            </GuidedField>
            <GuidedField label="Sous-section">
              <select className={inp} value={form.subSection || ""} onChange={(e) => ff("subSection", e.target.value)} disabled={!form.sectionId}>
                <option value="">Sous-section</option>
                {(sections.find((section) => section.id === form.sectionId)?.subSections || []).map((sub) => <option key={sub} value={sub}>{sub}</option>)}
              </select>
            </GuidedField>
            <GuidedField label="Profil">
              <StatusToggle options={ATHLETE_PROFILE_OPTIONS} value={form.category} onChange={(v) => ff("category", v)} />
            </GuidedField>
            <GuidedField label="Statut">
              <StatusToggle options={ATHLETE_STATUS} value={form.status} onChange={(v) => ff("status", v)} />
            </GuidedField>
            <GuidedField label="Coach">
              <input className={inp} placeholder="Nom du coach" value={form.coach || ""} onChange={(e) => ff("coach", e.target.value)} />
            </GuidedField>
            <GuidedField label="Email">
              <input type="email" className={inp} value={form.email || ""} onChange={(e) => ff("email", e.target.value)} />
            </GuidedField>
          </>}

          {type === "club" && <>
            <GuidedField label="Nom de la section *">
              <input autoFocus className={inp} placeholder="ex : Football" value={form.name || ""} onChange={(e) => ff("name", e.target.value)} />
            </GuidedField>
            <GuidedField label="Lieu / ville">
              <input className={inp} placeholder="Montrouge" value={form.city || ""} onChange={(e) => ff("city", e.target.value)} />
            </GuidedField>
            <GuidedField label="Statut">
              <StatusToggle options={CLUB_STATUS} value={form.status} onChange={(v) => ff("status", v)} />
            </GuidedField>
            <GuidedField label="Responsable">
              <input className={inp} placeholder="Nom du responsable" value={form.contactName || ""} onChange={(e) => ff("contactName", e.target.value)} />
            </GuidedField>
            <GuidedField label="Sous-sections">
              <input className={inp} placeholder="U7, U9, U11…" value={form.subSectionsText || ""} onChange={(e) => ff("subSectionsText", e.target.value)} />
            </GuidedField>
            <GuidedField label="Email">
              <input type="email" className={inp} value={form.email || ""} onChange={(e) => ff("email", e.target.value)} />
            </GuidedField>
          </>}
        </div>

        <div className="px-6 py-4 border-t shrink-0">
          <button
            type="button"
            onClick={handleSave}
            disabled={!type || !(form.name || form.title)}
            className="w-full py-4 rounded-2xl bg-slate-900 text-white text-base font-bold disabled:opacity-40 hover:bg-slate-800 transition-colors"
          >
            Créer
          </button>
        </div>
      </div>
    </div>
  );
}

function WorkspaceSelector({ workspaces, currentWorkspaceId, onChange, onNew }) {
  const [open, setOpen] = useState(false);
  const ref = useRef(null);
  const current = workspaces.find((w) => w.id === currentWorkspaceId) || workspaces[0];

  useEffect(() => {
    if (!open) return;
    const h = (e) => { if (ref.current && !ref.current.contains(e.target)) setOpen(false); };
    document.addEventListener("mousedown", h);
    return () => document.removeEventListener("mousedown", h);
  }, [open]);

  return (
    <div className="relative px-3 py-3 border-b border-slate-200" ref={ref}>
      <button onClick={() => setOpen((v) => !v)}
        className="w-full flex items-center gap-2.5 px-2 py-2 rounded-2xl hover:bg-slate-50 transition-colors">
        <div className="w-7 h-7 rounded-lg flex items-center justify-center text-white text-xs font-bold shrink-0"
          style={{ backgroundColor: current?.color || "#0f172a" }}>
          {current?.initial || "J"}
        </div>
        <span className="text-sm font-semibold text-slate-900 flex-1 text-left truncate">{current?.name || "JAM"}</span>
        <ChevronDown className={`w-3.5 h-3.5 text-slate-400 transition-transform ${open ? "rotate-180" : ""}`} />
      </button>
      {open && (
        <div className="absolute top-full left-3 right-3 z-50 bg-white rounded-2xl shadow-xl border mt-1 py-1">
          {workspaces.map((ws) => (
            <button key={ws.id} onClick={() => { onChange(ws.id); setOpen(false); }}
              className={`w-full flex items-center gap-2.5 px-3 py-2.5 text-left transition-colors ${ws.id === currentWorkspaceId ? "bg-slate-50" : "hover:bg-slate-50"}`}>
              <div className="w-6 h-6 rounded-lg flex items-center justify-center text-white text-xs font-bold shrink-0"
                style={{ backgroundColor: ws.color || "#0f172a" }}>
                {ws.initial}
              </div>
              <span className="text-sm font-medium flex-1">{ws.name}</span>
              {ws.id === currentWorkspaceId && <CheckCircle2 className="w-4 h-4 text-green-500" />}
            </button>
          ))}
          <div className="border-t my-1" />
          <button onClick={() => { onNew(); setOpen(false); }}
            className="w-full flex items-center gap-2.5 px-3 py-2.5 text-left hover:bg-slate-50 text-slate-600">
            <div className="w-6 h-6 rounded-lg border-2 border-dashed border-slate-300 flex items-center justify-center shrink-0">
              <Plus className="w-3 h-3 text-slate-400" />
            </div>
            <span className="text-sm">Nouvel espace de travail</span>
          </button>
        </div>
      )}
    </div>
  );
}

export default function JamSaaSApp() {
  const [data, setData] = useState(loadData);
  const [view, setView] = useState("dashboard");
  const [expanded, setExpanded] = useState({});

  // Filters
  const [searchQ, setSearchQ] = useState("");
  const [filterSport, setFilterSport] = useState("all");
  const [filterStatus, setFilterStatus] = useState("all");
  const [filterPriority, setFilterPriority] = useState("all");
  const [docMode, setDocMode] = useState("all");

  // Modals
  const [projectModal, setProjectModal] = useState({ open: false, item: null });
  const [taskModal, setTaskModal] = useState({ open: false, item: null });
  const [athleteModal, setAthleteModal] = useState({ open: false, item: null });
  const [clubModal, setClubModal] = useState({ open: false, item: null });
  const [docModal, setDocModal] = useState({ open: false, item: null });
  const [convModal, setConvModal] = useState({ open: false, item: null });
  const [notifModal, setNotifModal] = useState({ open: false, item: null });
  const [cfModal, setCfModal] = useState({ open: false, item: null });
  const [deleteConfirm, setDeleteConfirm] = useState({ open: false, type: null, item: null });
  const [projectViewMode, setProjectViewMode] = useState("tree");
  const [taskViewMode, setTaskViewMode] = useState("table");
  const [athleteViewMode, setAthleteViewMode] = useState("table");
  const [clubViewMode, setClubViewMode] = useState("table");
  const [documentViewMode, setDocumentViewMode] = useState("table");
  const [convocationViewMode, setConvocationViewMode] = useState("kanban");
  const [notificationViewMode, setNotificationViewMode] = useState("kanban");
  const [guidedOpen, setGuidedOpen] = useState(false);
  const [wsModal, setWsModal] = useState(false);
  const [wsForm, setWsForm] = useState({});
  const [dashboardEditMode, setDashboardEditMode] = useState(false);
  const [dashboardWidgetModal, setDashboardWidgetModal] = useState({ open: false, item: null });
  const [dashboardWidgetForm, setDashboardWidgetForm] = useState({});
  const [dashboardDrilldown, setDashboardDrilldown] = useState(null);
  const [dashboardDataSearch, setDashboardDataSearch] = useState("");
  const [dashboardDataSort, setDashboardDataSort] = useState("name");

  // Form state
  const [form, setForm] = useState({});
  const [newSportName, setNewSportName] = useState("");
  const [globalColumnModal, setGlobalColumnModal] = useState(null);
  const [importPreview, setImportPreview] = useState(null);
  const [importScope, setImportScope] = useState("all");
  const importScopeRef = useRef("all");
  const importInputRef = useRef(null);

  const allSports = useMemo(() => [...BASE_SPORTS, ...data.customSports], [data.customSports]);
  const sportName = (id) => allSports.find((s) => s.id === id)?.name || id;
  const splitSubSections = (value) => String(value || "")
    .split(/\n|,/)
    .map((item) => item.trim())
    .filter(Boolean);
  const sectionName = (id) => wsClubs.find((section) => section.id === id)?.name || "";
  const subSectionsFor = (sectionId) => wsClubs.find((section) => section.id === sectionId)?.subSections || [];
  const sectionAthleteCount = (sectionId) => wsAthletes.filter((athlete) => athlete.sectionId === sectionId).length;

  useEffect(() => { saveData(data); }, [data]);

  const upd = (key, fn) => setData((prev) => ({ ...prev, [key]: fn(prev[key]) }));

  const legacyProjectCustomFields = useMemo(() => data.customFields.filter((field) => field.appliesTo.includes("projects")), [data.customFields]);
  const legacyTaskCustomFields = useMemo(() => data.customFields.filter((field) => field.appliesTo.includes("tasks")), [data.customFields]);
  const tableProjectCustomFields = useMemo(
    () => tableColumnsToFields(data.tableConfig?.projects?.customCols, ["projects"]),
    [data.tableConfig?.projects?.customCols]
  );
  const tableTaskCustomFields = useMemo(
    () => tableColumnsToFields(data.tableConfig?.tasks?.customCols, ["tasks"]),
    [data.tableConfig?.tasks?.customCols]
  );
  const athleteCustomFields = useMemo(
    () => tableColumnsToFields(data.tableConfig?.athletes?.customCols, ["athletes"]),
    [data.tableConfig?.athletes?.customCols]
  );
  const clubCustomFields = useMemo(
    () => tableColumnsToFields(data.tableConfig?.clubs?.customCols, ["clubs"]),
    [data.tableConfig?.clubs?.customCols]
  );
  const documentCustomFields = useMemo(
    () => tableColumnsToFields(data.tableConfig?.documents?.customCols, ["documents"]),
    [data.tableConfig?.documents?.customCols]
  );
  const projectCustomFields = useMemo(
    () => mergeCustomFields(legacyProjectCustomFields, tableProjectCustomFields),
    [legacyProjectCustomFields, tableProjectCustomFields]
  );
  const taskCustomFields = useMemo(
    () => mergeCustomFields(legacyTaskCustomFields, tableTaskCustomFields),
    [legacyTaskCustomFields, tableTaskCustomFields]
  );

  // Form helpers
  const f = (key, val) => setForm((prev) => ({ ...prev, [key]: val }));
  const fcf = (id, val) => setForm((prev) => ({ ...prev, customFieldValues: { ...prev.customFieldValues, [id]: val } }));

  const resetFilters = () => {
    setSearchQ(""); setFilterSport("all"); setFilterStatus("all");
    setFilterPriority("all"); setDocMode("all");
  };
  const goToView = (nextView) => {
    setView(nextView);
    if (nextView !== "dashboardData") setDashboardDrilldown(null);
    resetFilters();
  };
  const fileActions = (scope, label) => (
    <>
      <button onClick={() => downloadImportTemplate(scope)} className="flex items-center gap-2 px-4 py-2 rounded-2xl border bg-white text-slate-700 text-sm font-medium hover:bg-slate-50">
        <FileSpreadsheet className="w-4 h-4" /> Gabarit
      </button>
      <button onClick={() => triggerImport(scope)} className="flex items-center gap-2 px-4 py-2 rounded-2xl border bg-white text-slate-700 text-sm font-medium hover:bg-slate-50">
        <Upload className="w-4 h-4" /> Importer {label}
      </button>
      <button onClick={() => exportExcel(scope)} className="flex items-center gap-2 px-4 py-2 rounded-2xl border bg-white text-slate-700 text-sm font-medium hover:bg-slate-50">
        <Download className="w-4 h-4" /> Exporter
      </button>
    </>
  );

  // Handlers
  const handleQuickTask = (projectId) => {
    setForm({ projectId, title: "", assignee: "", status: "À faire", priority: "Moyenne", progress: 0, startDate: "", endDate: "", description: "", customFieldValues: {} });
    setTaskModal({ open: true, item: null });
  };

  const handleEditProject = (item) => {
    setForm({ ...item });
    setProjectModal({ open: true, item });
  };

  const handleEditSection = (section) => {
    setForm({ ...section, subSectionsText: (section.subSections || []).join("\n") });
    setClubModal({ open: true, item: section });
  };

  const requestDelete = (type, item) => setDeleteConfirm({ open: true, type, item });

  const confirmDelete = () => {
    const { type, item } = deleteConfirm;
    const map = { project: "projects", task: "tasks", athlete: "athletes", club: "clubs", document: "documents", convocation: "convocations", notification: "notifications", customField: "customFields" };
    if (map[type]) upd(map[type], (arr) => arr.filter((x) => x.id !== item.id));
    setDeleteConfirm({ open: false, type: null, item: null });
  };

  const saveProject = () => {
    const parentId = form.parentId && parentOptionsForProject(form.id, data.projects).some((project) => project.id === form.parentId) ? form.parentId : null;
    const item = { workspaceId: currentWsId, ...form, parentId, id: form.id || uid(), durationDays: daysBetween(form.startDate, form.endDate), customFieldValues: form.customFieldValues || {} };
    upd("projects", (arr) => form.id ? arr.map((p) => p.id === form.id ? item : p) : [...arr, item]);
    setProjectModal({ open: false, item: null }); setForm({});
  };

  const saveTask = () => {
    const progress = normalizeTaskProgress(form.progress, form.status);
    const status = form.status === "Terminé" ? "Terminé" : statusFromProgress(progress, form.status);
    const item = { workspaceId: currentWsId, ...form, id: form.id || uid(), progress, status, durationDays: daysBetween(form.startDate, form.endDate), customFieldValues: form.customFieldValues || {} };
    upd("tasks", (arr) => form.id ? arr.map((t) => t.id === form.id ? item : t) : [...arr, item]);
    setTaskModal({ open: false, item: null }); setForm({});
  };

  const saveAthlete = () => {
    const section = wsClubs.find((item) => item.id === form.sectionId);
    const item = { workspaceId: currentWsId, ...form, sport: section?.sport || form.sport || "", id: form.id || uid(), customFieldValues: form.customFieldValues || {} };
    upd("athletes", (arr) => form.id ? arr.map((a) => a.id === form.id ? item : a) : [...arr, item]);
    setAthleteModal({ open: false, item: null }); setForm({});
  };

  const saveClub = () => {
    const item = {
      workspaceId: currentWsId,
      ...form,
      id: form.id || uid(),
      sport: allSports.find((sport) => normalizeKey(sport.name) === normalizeKey(form.name))?.id || form.sport || "",
      subSections: splitSubSections(form.subSectionsText ?? (form.subSections || []).join("\n")),
      customFieldValues: form.customFieldValues || {},
    };
    delete item.subSectionsText;
    upd("clubs", (arr) => form.id ? arr.map((c) => c.id === form.id ? item : c) : [...arr, item]);
    setClubModal({ open: false, item: null }); setForm({});
  };

  const saveDoc = () => {
    const item = { workspaceId: currentWsId, ...form, id: form.id || uid(), status: form.status || (form.url ? "Publié" : "Brouillon"), customFieldValues: form.customFieldValues || {} };
    upd("documents", (arr) => form.id ? arr.map((d) => d.id === form.id ? item : d) : [...arr, item]);
    setDocModal({ open: false, item: null }); setForm({});
  };

  const saveConv = () => {
    const item = { ...form, id: form.id || uid() };
    upd("convocations", (arr) => form.id ? arr.map((c) => c.id === form.id ? item : c) : [...arr, item]);
    setConvModal({ open: false, item: null }); setForm({});
  };

  const saveNotif = () => {
    const item = { ...form, id: form.id || uid() };
    upd("notifications", (arr) => form.id ? arr.map((n) => n.id === form.id ? item : n) : [...arr, item]);
    setNotifModal({ open: false, item: null }); setForm({});
  };

  const saveCF = () => {
    const item = { ...form, id: form.id || uid(), appliesTo: form.appliesTo || ["projects"] };
    upd("customFields", (arr) => form.id ? arr.map((x) => x.id === form.id ? item : x) : [...arr, item]);
    setCfModal({ open: false, item: null }); setForm({});
  };

  const addSport = () => {
    if (!newSportName.trim()) return;
    const id = makeSportId(newSportName);
    if (allSports.find((s) => s.id === id)) return;
    upd("customSports", (arr) => [...arr, { id, name: newSportName.trim(), icon: Trophy }]);
    setNewSportName("");
  };

  const currentWsId = data.currentWorkspaceId || "ws_jam";
  const switchWorkspace = (wsId) => setData((prev) => ({ ...prev, currentWorkspaceId: wsId }));
  const saveWorkspace = () => {
    if (!wsForm.name?.trim()) return;
    const ws = { id: uid(), name: wsForm.name.trim(), color: wsForm.color || "#0f172a", initial: (wsForm.name.trim()[0] || "W").toUpperCase() };
    setData((prev) => ({ ...prev, workspaces: [...(prev.workspaces || []), ws], currentWorkspaceId: ws.id }));
    setWsModal(false); setWsForm({});
  };

  const handleDuplicate = (project) => {
    const newId = uid();
    const clone = { ...project, id: newId, name: `${project.name} (copie)`, parentId: project.parentId, workspaceId: currentWsId };
    upd("projects", (arr) => [...arr, clone]);
  };

  const handleGuidedSave = (type, guidedForm) => {
    if (type === "project") {
      const item = { workspaceId: currentWsId, ...guidedForm, id: uid(), durationDays: daysBetween(guidedForm.startDate, guidedForm.endDate), customFieldValues: {} };
      upd("projects", (arr) => [...arr, item]);
    }
    if (type === "task") {
      const progress = normalizeTaskProgress(0, guidedForm.status || "À faire");
      const item = { workspaceId: currentWsId, ...guidedForm, id: uid(), progress, durationDays: daysBetween(guidedForm.startDate, guidedForm.endDate), customFieldValues: {} };
      upd("tasks", (arr) => [...arr, item]);
    }
    if (type === "athlete") upd("athletes", (arr) => [...arr, { workspaceId: currentWsId, ...guidedForm, id: uid() }]);
    if (type === "club") upd("clubs", (arr) => {
      const item = {
        workspaceId: currentWsId,
        ...guidedForm,
        id: uid(),
        sport: allSports.find((sport) => normalizeKey(sport.name) === normalizeKey(guidedForm.name))?.id || guidedForm.sport || "",
        subSections: splitSubSections(guidedForm.subSectionsText),
      };
      delete item.subSectionsText;
      return [...arr, item];
    });
  };

  const updateProjectStatus = (id, status) => {
    upd("projects", (arr) => arr.map((project) => project.id === id ? { ...project, status } : project));
  };

  const updateTaskStatus = (id, status) => {
    upd("tasks", (arr) => arr.map((task) => {
      if (task.id !== id) return task;
      const progress = normalizeTaskProgress(task.progress, status);
      return { ...task, status: statusFromProgress(progress, status), progress };
    }));
  };

  const updateEntityStatus = (key, id, status) => {
    upd(key, (arr) => arr.map((item) => item.id === id ? { ...item, status } : item));
  };

  // Enriched + filtered data (workspace-scoped)
  const wsProjects = useMemo(() => data.projects.filter((p) => p.workspaceId === currentWsId), [data.projects, currentWsId]);
  const wsTasks = useMemo(() => data.tasks.filter((t) => t.workspaceId === currentWsId), [data.tasks, currentWsId]);
  const wsAthletes = useMemo(() => data.athletes.filter((a) => a.workspaceId === currentWsId), [data.athletes, currentWsId]);
  const wsClubs = useMemo(() => data.clubs.filter((c) => c.workspaceId === currentWsId), [data.clubs, currentWsId]);

  const enrichedProjects = useMemo(() =>
    wsProjects.map((p) => {
      const projectScope = descendants(p.id, wsProjects);
      const hasLinkedTasks = wsTasks.some((task) => projectScope.has(task.projectId));
      const computedOwnProgress = ownProgress(p.id, wsTasks);
      const computedTotalProgress = consolidatedProgress(p.id, wsTasks, wsProjects);
      return {
        ...p,
        status: hasLinkedTasks ? deriveProjectStatus(computedTotalProgress, p.status) : p.status,
        sportName: sportName(p.sport),
        computedOwnProgress,
        computedTotalProgress,
      };
    }), [wsProjects, wsTasks, allSports]);

  const filteredProjects = useMemo(() => enrichedProjects.filter((p) => {
    if (filterSport !== "all" && p.sport !== filterSport) return false;
    if (filterStatus !== "all" && p.status !== filterStatus) return false;
    if (searchQ && !p.name.toLowerCase().includes(searchQ.toLowerCase())) return false;
    return true;
  }), [enrichedProjects, filterSport, filterStatus, searchQ]);

  const filteredTasks = useMemo(() => wsTasks.filter((t) => {
    if (filterStatus !== "all" && t.status !== filterStatus) return false;
    if (filterPriority !== "all" && t.priority !== filterPriority) return false;
    if (searchQ && !t.title.toLowerCase().includes(searchQ.toLowerCase())) return false;
    return true;
  }), [wsTasks, filterStatus, filterPriority, searchQ]);

  const filteredAthletes = useMemo(() => wsAthletes.filter((a) => {
    if (filterSport !== "all" && a.sport !== filterSport) return false;
    if (filterStatus !== "all" && a.status !== filterStatus) return false;
    if (searchQ && !a.name.toLowerCase().includes(searchQ.toLowerCase())) return false;
    return true;
  }), [wsAthletes, filterSport, filterStatus, searchQ]);

  const filteredClubs = useMemo(() => wsClubs.filter((c) => {
    if (filterSport !== "all" && c.sport !== filterSport) return false;
    if (filterStatus !== "all" && c.status !== filterStatus) return false;
    if (searchQ && !c.name.toLowerCase().includes(searchQ.toLowerCase())) return false;
    return true;
  }), [wsClubs, filterSport, filterStatus, searchQ]);

  const wsDocs = useMemo(() => data.documents.filter((d) => d.workspaceId === currentWsId), [data.documents, currentWsId]);
  const wsConvocations = useMemo(() => data.convocations.filter((c) => c.workspaceId === currentWsId), [data.convocations, currentWsId]);
  const wsNotifications = useMemo(() => data.notifications.filter((n) => n.workspaceId === currentWsId), [data.notifications, currentWsId]);

  const filteredDocs = useMemo(() => wsDocs.filter((d) => {
    if (filterSport !== "all" && d.sport !== filterSport) return false;
    if (filterStatus !== "all" && d.status !== filterStatus) return false;
    if (docMode === "internal" && d.url) return false;
    if (docMode === "external" && !d.url) return false;
    if (searchQ && !d.title.toLowerCase().includes(searchQ.toLowerCase())) return false;
    return true;
  }), [wsDocs, filterSport, filterStatus, docMode, searchQ]);

  const filteredConvocations = useMemo(() => wsConvocations.filter((conv) => {
    if (filterSport !== "all" && conv.sport !== filterSport) return false;
    if (filterStatus !== "all" && conv.status !== filterStatus) return false;
    if (searchQ && !conv.title.toLowerCase().includes(searchQ.toLowerCase())) return false;
    return true;
  }), [wsConvocations, filterSport, filterStatus, searchQ]);

  const filteredNotifications = useMemo(() => wsNotifications.filter((notif) => {
    if (filterSport !== "all" && notif.sport !== filterSport) return false;
    if (filterStatus !== "all" && notif.status !== filterStatus) return false;
    if (searchQ && !notif.title.toLowerCase().includes(searchQ.toLowerCase())) return false;
    return true;
  }), [wsNotifications, filterSport, filterStatus, searchQ]);

  const projectTree = useMemo(() => buildTree(filteredProjects), [filteredProjects]);

  const projectParentChoices = parentOptionsForProject(form.id, wsProjects);
  const hasAvailableProjectParent = projectParentChoices.length > 0;
  const openNewProjectModal = () => {
    setForm({ status: "À faire", priority: "Moyenne", sport: allSports[0]?.id || "developpe_couche", parentId: null, customFieldValues: {} });
    setProjectModal({ open: true, item: null });
  };
  const handleNewSubProject = (parentProject) => {
    setForm({
      status: "À faire",
      priority: parentProject.priority || "Moyenne",
      sport: parentProject.sport || allSports[0]?.id || "developpe_couche",
      parentId: parentProject.id,
      customFieldValues: {},
    });
    setProjectModal({ open: true, item: null });
  };

  const resolveSportId = (value, sports, discoveredSports) => {
    if (!value) return allSports[0]?.id || "developpe_couche";
    const normalized = normalizeKey(value);
    const existing = sports.find((sport) => normalizeKey(sport.id) === normalized || normalizeKey(sport.name) === normalized);
    if (existing) return existing.id;
    const id = makeSportId(String(value));
    if (id && !sports.some((sport) => sport.id === id) && !discoveredSports.has(id)) {
      discoveredSports.set(id, { id, name: String(value).trim(), icon: Trophy });
    }
    return id || allSports[0]?.id || "developpe_couche";
  };

  const rowProjectName = (row) => firstValue(row, ["Nom", "Projet", "Nom du projet", "Project"]);
  const rowTaskTitle = (row) => firstValue(row, ["Titre", "Tâche", "Tache", "Nom de la tâche", "Task"]);
  const rowAthleteName = (row) => firstValue(row, ["Nom", "Athlète", "Athlete", "Nom complet", "Sportif"]);
  const rowSectionName = (row) => firstValue(row, ["Section", "Nom de la section", "Section *"]);

  const downloadImportTemplate = async (scope = "all") => {
    const wb = new ExcelJS.Workbook();
    wb.creator = "JAM SaaS";
    wb.created = new Date();
    wb.modified = new Date();

    const include = (entity) => scope === "all" || scope === entity;
    const sportsList = allSports.map((sport) => sport.name);
    const sectionsList = wsClubs.map((section) => section.name);
    const projectCustomExample = Object.fromEntries(projectCustomFields.map((field) => [field.name, field.type === "financier" ? 2500 : ""]));
    const taskCustomExample = Object.fromEntries(taskCustomFields.map((field) => [field.name, ""]));
    const athleteCustomExample = Object.fromEntries(athleteCustomFields.map((field) => [field.name, ""]));
    const sectionCustomExample = Object.fromEntries(clubCustomFields.map((field) => [field.name, ""]));

    if (include("projects")) {
      const projectSheet = appendSheet(wb, "Projets", [
        {
          "Nom du projet *": "Stage régional 2026",
          Description: "Préparation et organisation du stage",
          Responsable: "Samy",
          Sport: "Développé couché",
          Statut: "À faire",
          Priorité: "Moyenne",
          "Date début": "2026-06-01",
          "Date fin": "2026-06-03",
          "Projet parent": "",
          ...projectCustomExample,
        },
      ]);
      applyImportSheetFormat(projectSheet, {
        required: ["Nom du projet *"],
        rows: 500,
        color: "4338CA",
        widths: {
          "Nom du projet *": 30,
          Description: 42,
          Responsable: 22,
          Sport: 24,
          Statut: 18,
          Priorité: 16,
          "Date début": 16,
          "Date fin": 16,
          "Projet parent": 28,
        },
        listValidations: {
          Sport: sportsList,
          Statut: STATUS_OPTIONS,
          Priorité: PRIORITY_OPTIONS,
        },
        dateHeaders: ["Date début", "Date fin"],
      });
    }

    if (include("tasks")) {
      const taskSheet = appendSheet(wb, "Tâches", [
        {
          "Titre *": "Réserver la salle",
          Description: "Confirmer le créneau auprès de la section",
          "Projet associé": "Stage régional 2026",
          "Assigné à": "Samy",
          Statut: "À faire",
          Priorité: "Haute",
          Progression: 0,
          "Date début": "2026-05-20",
          Échéance: "2026-05-25",
          ...taskCustomExample,
        },
      ]);
      applyImportSheetFormat(taskSheet, {
        required: ["Titre *"],
        rows: 1800,
        color: "0369A1",
        widths: {
          "Titre *": 32,
          Description: 42,
          "Projet associé": 30,
          "Assigné à": 24,
          Statut: 18,
          Priorité: 16,
          Progression: 16,
          "Date début": 16,
          Échéance: 16,
        },
        listValidations: {
          Statut: STATUS_OPTIONS,
          Priorité: PRIORITY_OPTIONS,
        },
        dateHeaders: ["Date début", "Échéance"],
        numericHeaders: ["Progression"],
      });
    }

    if (include("clubs")) {
      const sectionSheet = appendSheet(wb, "Sections", [
        {
          "Section *": "Football",
          Ville: "Montrouge",
          Responsable: "Responsable section",
          "Sous-sections": "U7, U9, U11, U13, U15",
          Statut: "Inscrit",
          Email: "section@example.com",
          Téléphone: "0100000000",
          Notes: "Une sous-section par virgule ou par ligne.",
          ...sectionCustomExample,
        },
      ]);
      applyImportSheetFormat(sectionSheet, {
        required: ["Section *"],
        rows: 200,
        color: "7C3AED",
        widths: {
          "Section *": 26,
          Ville: 22,
          Responsable: 26,
          "Sous-sections": 42,
          Statut: 18,
          Email: 30,
          Téléphone: 18,
          Notes: 42,
        },
        listValidations: {
          Statut: CLUB_STATUS,
        },
      });
    }

    if (include("athletes")) {
      const athleteSheet = appendSheet(wb, "Athlètes", [
        {
          "Nom complet *": "Prénom Nom",
          Section: "Développé couché",
          "Sous-section": "-88 kg",
          Profil: "Homme",
          Statut: "Actif",
          Coach: "Coach référent",
          Email: "prenom.nom@example.com",
          Téléphone: "0600000000",
          Notes: "Ligne exemple à remplacer",
          ...athleteCustomExample,
        },
      ]);
      applyImportSheetFormat(athleteSheet, {
        required: ["Nom complet *"],
        rows: 1200,
        color: "15803D",
        widths: {
          "Nom complet *": 28,
          Section: 24,
          "Sous-section": 20,
          Profil: 16,
          Statut: 18,
          Coach: 24,
          Email: 30,
          Téléphone: 18,
          Notes: 42,
        },
        listValidations: {
          Section: sectionsList,
          Profil: ATHLETE_PROFILE_OPTIONS,
          Statut: ATHLETE_STATUS,
        },
      });
    }

    const guideSheet = appendSheet(wb, "Mode emploi", [
      { Étape: "1", Onglet: "Projets", À_faire: "Ajoute ou mets à jour les projets. Le nom du projet est obligatoire.", Important: "Projet parent doit correspondre exactement à un autre nom de projet." },
      { Étape: "2", Onglet: "Tâches", À_faire: "Ajoute les tâches et renseigne Projet associé avec le nom exact du projet.", Important: "Progression doit être entre 0 et 100." },
      { Étape: "3", Onglet: "Sections", À_faire: "Ajoute les sections sportives et leurs sous-sections.", Important: "Sépare les sous-sections par virgule ou retour ligne." },
      { Étape: "4", Onglet: "Athlètes", À_faire: "Colle ou saisis tes sportifs. La colonne Nom complet est obligatoire.", Important: "Section doit correspondre exactement à une section existante ou importée." },
    ]);
    applyGuideSheetFormat(guideSheet);

    wb.views = [{ activeTab: 0 }];
    const names = { projects: "projets", tasks: "taches", clubs: "sections", athletes: "athletes", all: "global" };
    await downloadWorkbook(wb, `gabarit-import-jam-saas-${names[scope] || "global"}.xlsx`);
  };

  const exportExcel = async (scope = "all") => {
    const wb = new ExcelJS.Workbook();
    wb.creator = "JAM SaaS";
    const include = (entity) => scope === "all" || scope === entity;
    if (scope === "all") {
      appendSheet(wb, "Tableau de bord", [
        { Indicateur: "Projets", Valeur: wsProjects.length },
        { Indicateur: "Tâches", Valeur: wsTasks.length },
        { Indicateur: "Athlètes", Valeur: wsAthletes.length },
        { Indicateur: "Sections", Valeur: wsClubs.length },
        { Indicateur: "Budget total", Valeur: totalBudget },
      ]);
    }
    if (include("projects")) {
      appendSheet(wb, "Projets", wsProjects.map((project) => ({
        "Nom du projet": project.name,
        Description: project.description || "",
        Responsable: project.owner || "",
        Sport: sportName(project.sport),
        Statut: project.status || "",
        Priorité: project.priority || "",
        "Date début": project.startDate || "",
        "Date fin": project.endDate || "",
        "Projet parent": wsProjects.find((parent) => parent.id === project.parentId)?.name || "",
        ...Object.fromEntries(projectCustomFields.map((field) => [field.name, project.customFieldValues?.[field.id] ?? ""])),
      })));
    }
    if (include("tasks")) {
      appendSheet(wb, "Tâches", wsTasks.map((task) => ({
        Titre: task.title,
        Description: task.description || "",
        "Projet associé": wsProjects.find((project) => project.id === task.projectId)?.name || "",
        "Assigné à": task.assignee || "",
        Statut: task.status || "",
        Priorité: task.priority || "",
        Progression: normalizeTaskProgress(task.progress, task.status),
        "Date début": task.startDate || "",
        Échéance: task.endDate || "",
        ...Object.fromEntries(taskCustomFields.map((field) => [field.name, task.customFieldValues?.[field.id] ?? ""])),
      })));
    }
    if (include("clubs")) {
      appendSheet(wb, "Sections", wsClubs.map((club) => ({
        Section: club.name,
        Ville: club.city || "",
        Responsable: club.contactName || "",
        "Sous-sections": (club.subSections || []).join(", "),
        Statut: club.status || "",
        Email: club.email || "",
        Téléphone: club.phone || "",
        Athlètes: sectionAthleteCount(club.id),
        Notes: club.notes || "",
        ...Object.fromEntries(clubCustomFields.map((field) => [field.name, club.customFieldValues?.[field.id] ?? ""])),
      })));
    }
    if (include("athletes")) {
      appendSheet(wb, "Athlètes", wsAthletes.map((athlete) => ({
        "Nom complet": athlete.name,
        Section: sectionName(athlete.sectionId),
        "Sous-section": athlete.subSection || "",
        Profil: athlete.category || "",
        Statut: athlete.status || "",
        Coach: athlete.coach || "",
        Email: athlete.email || "",
        Téléphone: athlete.phone || "",
        Notes: athlete.notes || "",
        ...Object.fromEntries(athleteCustomFields.map((field) => [field.name, athlete.customFieldValues?.[field.id] ?? ""])),
      })));
    }
    if (scope === "all") {
      appendSheet(wb, "Documents", wsDocs.map((doc) => ({
        Titre: doc.title,
        Sport: sportName(doc.sport),
        Type: doc.type || "",
        Statut: doc.status || "",
        URL: doc.url || "",
        Notes: doc.notes || "",
        ...Object.fromEntries(documentCustomFields.map((field) => [field.name, doc.customFieldValues?.[field.id] ?? ""])),
      })));
      appendSheet(wb, "Convocations", wsConvocations.map((conv) => ({ Titre: conv.title, Sport: sportName(conv.sport), Type: conv.type || "", Date: conv.date || "", Heure: conv.time || "", Lieu: conv.location || "", Statut: conv.status || "" })));
      appendSheet(wb, "Notifications", wsNotifications.map((notif) => ({ Titre: notif.title, Sport: sportName(notif.sport), Type: notif.type || "", Statut: notif.status || "", "Date envoi": notif.sentAt || "", Message: notif.message || "" })));
    }
    const names = { projects: "projets", tasks: "taches", clubs: "sections", athletes: "athletes", all: "global" };
    await downloadWorkbook(wb, `export-jam-saas-${names[scope] || "global"}.xlsx`);
  };

  const applyImportRows = ({ projectRows = [], taskRows = [], sectionRows = [], athleteRows = [] }) => {
    const counts = { projects: 0, tasks: 0, sections: 0, athletes: 0 };

    setData((prev) => {
      const discoveredSports = new Map();
      const sports = [...BASE_SPORTS, ...(prev.customSports || [])];
      const projectByName = new Map(prev.projects.filter((project) => project.workspaceId === currentWsId).map((project) => [normalizeKey(project.name), project]));
      const nextProjects = [...prev.projects];

      projectRows.forEach((row) => {
        const name = rowProjectName(row);
        if (!name) return;
        const existing = projectByName.get(normalizeKey(name));
        const id = existing?.id || uid();
        const startDate = excelDateToIso(firstValue(row, ["Date début", "Date debut", "Début", "Debut", "Start"])) || existing?.startDate || "";
        const endDate = excelDateToIso(firstValue(row, ["Date fin", "Fin", "Échéance", "Echeance", "End"])) || existing?.endDate || "";
        const project = {
          ...(existing || {}),
          id,
          workspaceId: currentWsId,
          name,
          description: firstValue(row, ["Description"]) || existing?.description || "",
          owner: firstValue(row, ["Responsable", "Owner"]) || existing?.owner || "",
          sport: resolveSportId(firstValue(row, ["Sport"]), sports, discoveredSports) || existing?.sport || "developpe_couche",
          status: firstValue(row, ["Statut", "Status"]) || existing?.status || "À faire",
          priority: firstValue(row, ["Priorité", "Priorite", "Priority"]) || existing?.priority || "Moyenne",
          startDate,
          endDate,
          customFieldValues: { ...(existing?.customFieldValues || {}), ...customFieldValuesFromRow(row, projectCustomFields) },
        };
        project.durationDays = daysBetween(project.startDate, project.endDate);
        projectByName.set(normalizeKey(name), project);
        const index = nextProjects.findIndex((item) => item.id === id);
        if (index >= 0) nextProjects[index] = project;
        else nextProjects.push(project);
        counts.projects += 1;
      });

      projectRows.forEach((row) => {
        const name = rowProjectName(row);
        const parentName = firstValue(row, ["Projet parent", "Parent", "Projet principal"]);
        const project = projectByName.get(normalizeKey(name));
        const parent = projectByName.get(normalizeKey(parentName));
        if (!project || !parent || project.id === parent.id) return;
        const index = nextProjects.findIndex((item) => item.id === project.id);
        if (index >= 0) nextProjects[index] = { ...nextProjects[index], parentId: parent.id };
      });

      const nextClubs = [...prev.clubs];
      const sectionByName = new Map(nextClubs.filter((section) => section.workspaceId === currentWsId).map((section) => [normalizeKey(section.name), section]));
      sectionRows.forEach((row) => {
        const name = rowSectionName(row);
        if (!name) return;
        const existing = sectionByName.get(normalizeKey(name));
        const subSectionsRaw = firstValue(row, ["Sous-sections", "Sous sections", "Sous-section", "Categories", "Catégories"]);
        const section = {
          ...(existing || {}),
          id: existing?.id || uid(),
          workspaceId: currentWsId,
          name,
          sport: existing?.sport || resolveSportId(name, sports, discoveredSports) || "developpe_couche",
          city: firstValue(row, ["Ville", "Lieu", "City"]) || existing?.city || "",
          contactName: firstValue(row, ["Responsable", "Contact", "Référent", "Referent"]) || existing?.contactName || "",
          subSections: subSectionsRaw ? splitSubSections(subSectionsRaw) : existing?.subSections || [],
          status: firstValue(row, ["Statut", "Status"]) || existing?.status || "À contacter",
          email: firstValue(row, ["Email", "E-mail", "Mail"]) || existing?.email || "",
          phone: firstValue(row, ["Téléphone", "Telephone", "Tel", "Phone"]) || existing?.phone || "",
          notes: firstValue(row, ["Notes", "Note"]) || existing?.notes || "",
          customFieldValues: { ...(existing?.customFieldValues || {}), ...customFieldValuesFromRow(row, clubCustomFields) },
        };
        sectionByName.set(normalizeKey(name), section);
        const index = nextClubs.findIndex((item) => item.id === section.id);
        if (index >= 0) nextClubs[index] = section;
        else nextClubs.push(section);
        counts.sections += 1;
      });

      const nextTasks = [...prev.tasks];
      taskRows.forEach((row) => {
        const title = rowTaskTitle(row);
        if (!title) return;
        const projectName = firstValue(row, ["Projet associé", "Projet associe", "Projet", "Project"]);
        const project = projectByName.get(normalizeKey(projectName));
        const existing = nextTasks.find((task) => task.workspaceId === currentWsId && normalizeKey(task.title) === normalizeKey(title) && (project ? task.projectId === project.id : true));
        const progress = parseNumberValue(firstValue(row, ["Progression", "Avancement", "%"])) || 0;
        const status = firstValue(row, ["Statut", "Status"]) || existing?.status || "À faire";
        const startDate = excelDateToIso(firstValue(row, ["Date début", "Date debut", "Début", "Debut", "Start"])) || existing?.startDate || "";
        const endDate = excelDateToIso(firstValue(row, ["Échéance", "Echeance", "Date fin", "Fin", "End"])) || existing?.endDate || "";
        const task = {
          ...(existing || {}),
          id: existing?.id || uid(),
          workspaceId: currentWsId,
          title,
          description: firstValue(row, ["Description"]) || existing?.description || "",
          projectId: project?.id || existing?.projectId || "",
          assignee: firstValue(row, ["Assigné à", "Assigne a", "Assigné", "Assignee"]) || existing?.assignee || "",
          status,
          priority: firstValue(row, ["Priorité", "Priorite", "Priority"]) || existing?.priority || "Moyenne",
          progress,
          startDate,
          endDate,
          customFieldValues: { ...(existing?.customFieldValues || {}), ...customFieldValuesFromRow(row, taskCustomFields) },
        };
        task.durationDays = daysBetween(task.startDate, task.endDate);
        const index = nextTasks.findIndex((item) => item.id === task.id);
        if (index >= 0) nextTasks[index] = task;
        else nextTasks.push(task);
        counts.tasks += 1;
      });

      const nextAthletes = [...prev.athletes];
      athleteRows.forEach((row) => {
        const name = rowAthleteName(row);
        if (!name) return;
        const email = firstValue(row, ["Email", "E-mail", "Mail"]);
        const sectionLabel = firstValue(row, ["Section"]);
        const section = sectionByName.get(normalizeKey(sectionLabel));
        const sport = section?.sport || resolveSportId(firstValue(row, ["Sport"]), sports, discoveredSports) || "developpe_couche";
        const existing = nextAthletes.find((athlete) => athlete.workspaceId === currentWsId && ((email && normalizeKey(athlete.email) === normalizeKey(email)) || (!email && normalizeKey(athlete.name) === normalizeKey(name) && athlete.sport === sport)));
        const athlete = {
          ...(existing || {}),
          id: existing?.id || uid(),
          workspaceId: currentWsId,
          name,
          sport,
          sectionId: section?.id || existing?.sectionId || "",
          subSection: firstValue(row, ["Sous-section", "Sous section", "Subsection"]) || existing?.subSection || "",
          category: firstValue(row, ["Profil", "Genre", "Sexe", "Catégorie", "Categorie", "Cat"]) || existing?.category || "",
          status: firstValue(row, ["Statut", "Status"]) || existing?.status || "Actif",
          coach: firstValue(row, ["Coach", "Entraîneur", "Entraineur"]) || existing?.coach || "",
          email: email || existing?.email || "",
          phone: firstValue(row, ["Téléphone", "Telephone", "Tel", "Phone"]) || existing?.phone || "",
          notes: firstValue(row, ["Notes", "Note"]) || existing?.notes || "",
          customFieldValues: { ...(existing?.customFieldValues || {}), ...customFieldValuesFromRow(row, athleteCustomFields) },
        };
        const index = nextAthletes.findIndex((item) => item.id === athlete.id);
        if (index >= 0) nextAthletes[index] = athlete;
        else nextAthletes.push(athlete);
        counts.athletes += 1;
      });

      const existingSportIds = new Set((prev.customSports || []).map((sport) => sport.id));
      const customSports = [
        ...(prev.customSports || []),
        ...Array.from(discoveredSports.values()).filter((sport) => !existingSportIds.has(sport.id) && !BASE_SPORTS.some((base) => base.id === sport.id)),
      ];

      return { ...prev, projects: nextProjects, tasks: nextTasks, clubs: nextClubs, athletes: nextAthletes, customSports };
    });

    return counts;
  };

  const buildImportWarnings = ({ projectRows, taskRows, sectionRows, athleteRows }) => {
    const warnings = [];
    const missingProjects = projectRows.filter((row) => !rowProjectName(row)).length;
    const missingTasks = taskRows.filter((row) => !rowTaskTitle(row)).length;
    const missingSections = sectionRows.filter((row) => !rowSectionName(row)).length;
    const missingAthletes = athleteRows.filter((row) => !rowAthleteName(row)).length;
    const projectNames = new Set([
      ...wsProjects.map((project) => project.name),
      ...projectRows.map((row) => rowProjectName(row)),
    ].filter(Boolean).map(normalizeKey));
    const sectionNames = new Set([
      ...wsClubs.map((section) => section.name),
      ...sectionRows.map((row) => rowSectionName(row)),
    ].filter(Boolean).map(normalizeKey));
    const unknownTaskProjects = taskRows.filter((row) => {
      const projectName = firstValue(row, ["Projet associé", "Projet associe", "Projet", "Project"]);
      return projectName && !projectNames.has(normalizeKey(projectName));
    }).length;
    const unknownParents = projectRows.filter((row) => {
      const parentName = firstValue(row, ["Projet parent", "Parent", "Projet principal"]);
      return parentName && !projectNames.has(normalizeKey(parentName));
    }).length;
    const unknownAthleteSections = athleteRows.filter((row) => {
      const sectionNameValue = firstValue(row, ["Section"]);
      return sectionNameValue && !sectionNames.has(normalizeKey(sectionNameValue));
    }).length;

    if (missingAthletes) warnings.push(`${missingAthletes} ligne(s) Athlètes sans nom complet seront ignorées.`);
    if (missingProjects) warnings.push(`${missingProjects} ligne(s) Projets sans nom de projet seront ignorées.`);
    if (missingTasks) warnings.push(`${missingTasks} ligne(s) Tâches sans titre seront ignorées.`);
    if (missingSections) warnings.push(`${missingSections} ligne(s) Sections sans nom de section seront ignorées.`);
    if (unknownTaskProjects) warnings.push(`${unknownTaskProjects} tâche(s) indiquent un projet associé introuvable. Elles seront importées sans rattachement.`);
    if (unknownParents) warnings.push(`${unknownParents} sous-projet(s) indiquent un projet parent introuvable. Ils resteront au niveau principal.`);
    if (unknownAthleteSections) warnings.push(`${unknownAthleteSections} athlète(s) indiquent une section introuvable. Ils seront importés sans rattachement.`);
    if (!projectRows.length && !taskRows.length && !sectionRows.length && !athleteRows.length) warnings.push("Aucune ligne exploitable détectée dans les onglets Projets, Tâches, Sections ou Athlètes.");

    return warnings;
  };

  const triggerImport = (scope) => {
    importScopeRef.current = scope;
    setImportScope(scope);
    importInputRef.current?.click();
  };

  const handleImportFile = async (event) => {
    const file = event.target.files?.[0];
    event.target.value = "";
    if (!file) return;
    try {
      const scope = importScopeRef.current || importScope;
      const include = (entity) => scope === "all" || scope === entity;
      const buffer = await file.arrayBuffer();
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);
      const projectRows = include("projects") ? readSheetRows(workbook, ["Projets", "Projects"]) : [];
      const taskRows = include("tasks") ? readSheetRows(workbook, ["Tâches", "Taches", "Tasks"]) : [];
      const sectionRows = include("clubs") ? readSheetRows(workbook, ["Sections", "Section", "Clubs"]) : [];
      const athleteRows = include("athletes") ? readSheetRows(workbook, ["Athlètes", "Athletes", "Sportifs"]) : [];
      const activeTab = scope !== "all" ? scope : athleteRows.length ? "athletes" : sectionRows.length ? "clubs" : projectRows.length ? "projects" : "tasks";

      setImportPreview({
        fileName: file.name,
        scope,
        projectRows,
        taskRows,
        sectionRows,
        athleteRows,
        activeTab,
        warnings: buildImportWarnings({ projectRows, taskRows, sectionRows, athleteRows }),
      });
    } catch (error) {
      window.alert(`Import impossible : ${error.message}`);
    }
  };

  const confirmImportPreview = () => {
    if (!importPreview) return;
    const counts = {
      projects: (importPreview.projectRows || []).filter((row) => rowProjectName(row)).length,
      tasks: (importPreview.taskRows || []).filter((row) => rowTaskTitle(row)).length,
      sections: (importPreview.sectionRows || []).filter((row) => rowSectionName(row)).length,
      athletes: (importPreview.athleteRows || []).filter((row) => rowAthleteName(row)).length,
    };
    applyImportRows(importPreview);
    setImportPreview(null);
    window.alert(`Import terminé : ${counts.projects} projet(s), ${counts.tasks} tâche(s), ${counts.sections} section(s), ${counts.athletes} athlète(s).`);
  };

  const addGlobalColumn = ({ name, type }) => {
    const entity = globalColumnModal?.entity;
    if (!entity) return;
    const newCol = { key: "cf_" + uid().slice(0, 8), label: name, type };
    setData((prev) => {
      const current = prev.tableConfig?.[entity] || { customCols: [], hiddenCols: [], renamed: {}, typeOverrides: {}, order: [], groupBy: null, filters: {} };
      return {
        ...prev,
        tableConfig: {
          ...(prev.tableConfig || {}),
          [entity]: {
            ...current,
            customCols: [...(current.customCols || []), newCol],
            order: current.order?.length ? [...current.order, newCol.key] : (current.order || []),
          },
        },
      };
    });
    setGlobalColumnModal(null);
    if (entity === "projects") setProjectViewMode("table");
  };

  // Dashboard stats (workspace-scoped)
  const totalTasks = wsTasks.length;
  const completedTasks = wsTasks.filter((t) => t.status === "Terminé").length;
  const inProgressTasks = wsTasks.filter((t) => t.status === "En cours").length;
  const avgProgress = totalTasks ? Math.round(wsTasks.reduce((s, t) => s + normalizeTaskProgress(t.progress, t.status), 0) / totalTasks) : 0;
  const todayTime = useMemo(() => {
    const date = new Date();
    date.setHours(0, 0, 0, 0);
    return date.getTime();
  }, []);
  const openTasks = wsTasks.filter((task) => task.status !== "Terminé");
  const overdueTasks = openTasks.filter((task) => {
    const days = daysFromToday(task.endDate, todayTime);
    return days !== null && days < 0;
  });
  const upcomingTasks = openTasks
    .map((task) => ({ ...task, daysLeft: daysFromToday(task.endDate, todayTime) }))
    .filter((task) => task.daysLeft !== null && task.daysLeft >= 0 && task.daysLeft <= 14)
    .sort((a, b) => a.daysLeft - b.daysLeft)
    .slice(0, 5);
  const criticalTasks = openTasks.filter((task) => task.priority === "Critique" || task.priority === "Haute");
  const overdueProjects = wsProjects.filter((project) => {
    const days = daysFromToday(project.endDate, todayTime);
    return project.status !== "Terminé" && days !== null && days < 0;
  });
  const incompleteAthletes = wsAthletes.filter((athlete) => !athlete.email || !athlete.phone || !athlete.sectionId || !athlete.subSection || !athlete.category);
  const projectBudgetFields = useMemo(() => {
    const financial = projectCustomFields.filter((field) => field.type === "financier");
    const budgetNamed = financial.filter((field) => normalizeKey(field.name).includes("budget"));
    return budgetNamed.length ? budgetNamed : financial;
  }, [projectCustomFields]);
  const budgetTotals = useMemo(() => projectBudgetFields.map((field) => ({
    field,
    total: wsProjects.reduce((sum, project) => sum + (Number(project.customFieldValues?.[field.id]) || 0), 0),
  })), [projectBudgetFields, wsProjects]);
  const totalBudget = budgetTotals.reduce((sum, row) => sum + row.total, 0);
  const dashboardWidgets = data.dashboardWidgets?.length ? data.dashboardWidgets : DEFAULT_DASHBOARD_WIDGETS;

  const dashboardSourceRows = (source) => {
    if (source === "projects") return wsProjects;
    if (source === "athletes") return wsAthletes;
    if (source === "clubs") return wsClubs;
    return wsTasks;
  };
  const dashboardFilterLabel = (filter) => DASHBOARD_FILTERS.find((item) => item.id === filter)?.label || "Tous";
  const dashboardSourceLabel = (source) => DASHBOARD_SOURCES.find((item) => item.id === source)?.label || "Tâches";
  const dashboardTypeLabel = (type) => DASHBOARD_WIDGET_TYPES.find((item) => item.id === type)?.label || "Nombre";
  const dashboardGroupValue = (row, widget) => {
    if (widget.groupBy === "priority") return row.priority || "Sans priorité";
    if (widget.groupBy === "section") return sectionName(row.sectionId) || row.name || "Sans section";
    if (widget.groupBy === "profile") return row.category || "Sans profil";
    if (widget.groupBy === "assignee") return row.assignee || row.owner || row.coach || row.contactName || "Non assigné";
    return row.status || "Sans statut";
  };
  const dashboardFilteredRows = (widget) => {
    const rows = dashboardSourceRows(widget.source);
    const source = widget.source || "tasks";
    const filter = widget.filter || "all";
    if (filter === "completed") return rows.filter((row) => row.status === "Terminé" || row.status === "Validé");
    if (filter === "open") return rows.filter((row) => row.status !== "Terminé" && row.status !== "Validé");
    if (filter === "overdue") {
      if (source === "projects") return rows.filter((row) => row.status !== "Terminé" && daysFromToday(row.endDate, todayTime) < 0);
      if (source === "tasks") return rows.filter((row) => row.status !== "Terminé" && daysFromToday(row.endDate, todayTime) < 0);
      return [];
    }
    if (filter === "upcoming") {
      if (source !== "tasks" && source !== "projects") return [];
      return rows.filter((row) => {
        const days = daysFromToday(row.endDate, todayTime);
        return row.status !== "Terminé" && days !== null && days >= 0 && days <= 14;
      });
    }
    if (filter === "highPriority") return rows.filter((row) => row.priority === "Haute" || row.priority === "Critique");
    if (filter === "unassigned") return rows.filter((row) => !row.assignee && !row.owner && !row.coach && !row.contactName);
    if (filter === "incomplete") {
      if (source === "athletes") return rows.filter((row) => !row.email || !row.phone || !row.sectionId || !row.subSection || !row.category);
      if (source === "tasks") return rows.filter((row) => !row.assignee || !row.endDate);
      if (source === "projects") return rows.filter((row) => !row.owner || !row.endDate);
      return rows.filter((row) => !row.contactName || !row.email);
    }
    return rows;
  };
  const dashboardNumberValue = (widget) => {
    if (widget.filter === "budget") return formatMoney(totalBudget);
    return dashboardFilteredRows(widget).length;
  };
  const dashboardGroupedData = (widget) => {
    const rows = dashboardFilteredRows(widget);
    const map = new Map();
    rows.forEach((row) => {
      const key = dashboardGroupValue(row, widget);
      map.set(key, (map.get(key) || 0) + 1);
    });
    return Array.from(map.entries()).map(([label, value]) => ({ label, value })).sort((a, b) => b.value - a.value);
  };
  const dashboardRowsForWidget = (widget, groupValue = null) => {
    if (!widget) return [];
    let rows = dashboardFilteredRows(widget);
    if (groupValue) rows = rows.filter((row) => dashboardGroupValue(row, widget) === groupValue);
    return rows;
  };
  const dashboardRowTitle = (row, source) => {
    if (source === "tasks") return row.title || "Tâche sans titre";
    if (source === "projects") return row.name || "Projet sans nom";
    if (source === "athletes") return row.name || "Athlète sans nom";
    return row.name || "Section sans nom";
  };
  const dashboardRowText = (row, source) => {
    const common = [dashboardRowTitle(row, source), row.status, row.priority, row.owner, row.assignee, row.coach, row.contactName, row.email, row.phone, row.endDate, row.city].filter(Boolean);
    if (source === "tasks") common.push(data.projects.find((project) => project.id === row.projectId)?.name);
    if (source === "athletes") common.push(sectionName(row.sectionId), row.subSection, row.category);
    if (source === "clubs") common.push((row.subSections || []).join(" "));
    return common.join(" ").toLowerCase();
  };
  const dashboardBudgetForProject = (project) => projectBudgetFields.reduce((sum, field) => sum + (Number(project.customFieldValues?.[field.id]) || 0), 0);
  const dashboardDrilldownRows = dashboardRowsForWidget(dashboardDrilldown?.widget, dashboardDrilldown?.groupValue);
  const dashboardDrilldownVisibleRows = dashboardDrilldownRows
    .filter((row) => !dashboardDataSearch.trim() || dashboardRowText(row, dashboardDrilldown?.widget?.source).includes(dashboardDataSearch.trim().toLowerCase()))
    .sort((a, b) => {
      const source = dashboardDrilldown?.widget?.source || "tasks";
      if (dashboardDataSort === "date") return String(a.endDate || "").localeCompare(String(b.endDate || ""));
      if (dashboardDataSort === "status") return String(a.status || "").localeCompare(String(b.status || ""));
      if (dashboardDataSort === "priority") return String(a.priority || "").localeCompare(String(b.priority || ""));
      return dashboardRowTitle(a, source).localeCompare(dashboardRowTitle(b, source));
    });
  const openDashboardWidget = (widget, groupValue = null) => {
    setDashboardDrilldown({ widget: { ...widget }, groupValue });
    setDashboardDataSearch("");
    setDashboardDataSort("name");
    setView("dashboardData");
  };
  const closeDashboardDrilldown = () => {
    setDashboardDrilldown(null);
    setView("dashboard");
  };
  const editDashboardRow = (row, source) => {
    if (source === "projects") return handleEditProject(row);
    if (source === "tasks") {
      setForm({ ...row });
      setTaskModal({ open: true, item: row });
      return;
    }
    if (source === "athletes") {
      setForm({ ...row });
      setAthleteModal({ open: true, item: row });
      return;
    }
    handleEditSection(row);
  };
  const dashboardDataColumns = (source) => {
    if (source === "projects") {
      return [
        { key: "name", label: "Projet", render: (row) => <span className="font-semibold text-slate-900">{row.name}</span> },
        { key: "status", label: "Statut", render: (row) => <Badge className={badgeClass("status", row.status)}>{row.status || "—"}</Badge> },
        { key: "priority", label: "Priorité", render: (row) => <Badge className={badgeClass("priority", row.priority)}>{row.priority || "—"}</Badge> },
        { key: "owner", label: "Responsable", render: (row) => row.owner || "—" },
        { key: "endDate", label: "Échéance", render: (row) => row.endDate || "—" },
        { key: "budget", label: "Budget", render: (row) => formatMoney(dashboardBudgetForProject(row)) },
      ];
    }
    if (source === "athletes") {
      return [
        { key: "name", label: "Athlète", render: (row) => <span className="font-semibold text-slate-900">{row.name}</span> },
        { key: "sectionId", label: "Section", render: (row) => sectionName(row.sectionId) || "—" },
        { key: "subSection", label: "Sous-section", render: (row) => row.subSection || "—" },
        { key: "category", label: "Profil", render: (row) => row.category || "—" },
        { key: "coach", label: "Coach", render: (row) => row.coach || "—" },
        { key: "status", label: "Statut", render: (row) => <Badge>{row.status || "—"}</Badge> },
        { key: "contact", label: "Contact", render: (row) => row.email || row.phone || "—" },
      ];
    }
    if (source === "clubs") {
      return [
        { key: "name", label: "Section", render: (row) => <span className="font-semibold text-slate-900">{row.name}</span> },
        { key: "status", label: "Statut", render: (row) => <Badge>{row.status || "—"}</Badge> },
        { key: "subSections", label: "Sous-sections", render: (row) => row.subSections?.length ? row.subSections.join(", ") : "—" },
        { key: "athletes", label: "Athlètes", render: (row) => sectionAthleteCount(row.id) },
        { key: "contactName", label: "Responsable", render: (row) => row.contactName || "—" },
        { key: "phone", label: "Téléphone", render: (row) => row.phone || "—" },
      ];
    }
    return [
      { key: "title", label: "Tâche", render: (row) => <span className="font-semibold text-slate-900">{row.title}</span> },
      { key: "status", label: "Statut", render: (row) => <Badge className={badgeClass("status", row.status)}>{row.status || "—"}</Badge> },
      { key: "priority", label: "Priorité", render: (row) => <Badge className={badgeClass("priority", row.priority)}>{row.priority || "—"}</Badge> },
      { key: "assignee", label: "Responsable", render: (row) => row.assignee || "—" },
      { key: "endDate", label: "Échéance", render: (row) => row.endDate || "—" },
      { key: "project", label: "Projet", render: (row) => data.projects.find((project) => project.id === row.projectId)?.name || "—" },
      { key: "progress", label: "Progression", render: (row) => `${normalizeTaskProgress(row.progress, row.status)}%` },
    ];
  };
  const openWidgetEditor = (widget = null) => {
    setDashboardWidgetForm(widget ? { ...widget } : { type: "number", title: "Nouveau KPI", source: "tasks", filter: "all", groupBy: "status", tone: "slate", size: "small" });
    setDashboardWidgetModal({ open: true, item: widget });
  };
  const saveDashboardWidget = () => {
    const item = { ...dashboardWidgetForm, id: dashboardWidgetForm.id || "dw_" + uid().slice(0, 8) };
    setData((prev) => {
      const widgets = prev.dashboardWidgets?.length ? prev.dashboardWidgets : DEFAULT_DASHBOARD_WIDGETS;
      const exists = widgets.some((widget) => widget.id === item.id);
      return { ...prev, dashboardWidgets: exists ? widgets.map((widget) => widget.id === item.id ? item : widget) : [...widgets, item] };
    });
    setDashboardWidgetModal({ open: false, item: null });
    setDashboardWidgetForm({});
  };
  const deleteDashboardWidget = (widgetId) => {
    setData((prev) => ({ ...prev, dashboardWidgets: (prev.dashboardWidgets || DEFAULT_DASHBOARD_WIDGETS).filter((widget) => widget.id !== widgetId) }));
  };
  const resetDashboardWidgets = () => setData((prev) => ({ ...prev, dashboardWidgets: DEFAULT_DASHBOARD_WIDGETS }));

  const navGroups = [
    {
      label: "Pilotage",
      items: [
        { id: "dashboard", label: "Tableau de bord", Icon: LayoutDashboard, count: overdueTasks.length ? `${overdueTasks.length}` : null, tone: "danger", tooltip: `${overdueTasks.length} tâche${overdueTasks.length > 1 ? "s" : ""} en retard` },
      ],
    },
    {
      label: "Production",
      items: [
        { id: "projects", label: "Projets", Icon: FolderKanban, count: wsProjects.length, alert: overdueProjects.length, tooltip: overdueProjects.length ? `${overdueProjects.length} projet${overdueProjects.length > 1 ? "s" : ""} en retard` : `${wsProjects.length} projet${wsProjects.length > 1 ? "s" : ""} au total` },
        { id: "tasks", label: "Tâches", Icon: ListChecks, count: wsTasks.length, alert: overdueTasks.length, tooltip: overdueTasks.length ? `${overdueTasks.length} tâche${overdueTasks.length > 1 ? "s" : ""} en retard` : `${wsTasks.length} tâche${wsTasks.length > 1 ? "s" : ""} au total` },
        { id: "athletes", label: "Athlètes", Icon: Users, count: wsAthletes.length, alert: incompleteAthletes.length, tooltip: incompleteAthletes.length ? `${incompleteAthletes.length} fiche${incompleteAthletes.length > 1 ? "s" : ""} athlète à compléter` : `${wsAthletes.length} athlète${wsAthletes.length > 1 ? "s" : ""} au total` },
        { id: "clubs", label: "Sections", Icon: Building2, count: wsClubs.length, tooltip: `${wsClubs.length} section${wsClubs.length > 1 ? "s" : ""} au total` },
      ],
    },
    {
      label: "Communication",
      items: [
        { id: "documents", label: "Documents", Icon: FileText, count: wsDocs.length, tooltip: `${wsDocs.length} document${wsDocs.length > 1 ? "s" : ""} au total` },
        { id: "convocations", label: "Convocations", Icon: CalendarDays, count: wsConvocations.length, tooltip: `${wsConvocations.length} convocation${wsConvocations.length > 1 ? "s" : ""} au total` },
        { id: "notifications", label: "Notifications", Icon: Bell, count: wsNotifications.length, tooltip: `${wsNotifications.length} notification${wsNotifications.length > 1 ? "s" : ""} au total` },
      ],
    },
    {
      label: "Système",
      items: [
        { id: "settings", label: "Paramètres", Icon: BarChart3 },
      ],
    },
  ];

  function NavCounter({ value, active, alert, tone, tooltip }) {
    if (value === null || value === undefined) return null;
    const counterClass = active
      ? alert ? "bg-white text-red-600" : "bg-white/15 text-white"
      : alert || tone === "danger" ? "bg-red-50 text-red-600" : "bg-slate-100 text-slate-500";
    return (
      <span className="relative group/counter shrink-0" title={tooltip}>
        <span
          aria-label={tooltip}
          className={`block min-w-6 rounded-full px-1.5 py-0.5 text-[11px] font-bold text-center ${counterClass}`}
        >
          {value}
        </span>
        {tooltip && (
          <span className="pointer-events-none absolute left-full top-1/2 z-50 ml-2 hidden -translate-y-1/2 whitespace-nowrap rounded-xl border border-slate-200 bg-white px-3 py-2 text-xs font-semibold text-slate-700 shadow-xl group-hover/counter:block">
            {tooltip}
          </span>
        )}
      </span>
    );
  }

  // Shared filter bar
  function FilterBar({ withSport = true, withStatus = true, withPriority = false, statusOptions = STATUS_OPTIONS }) {
    const activeCount = [searchQ, withSport && filterSport !== "all", withStatus && filterStatus !== "all", withPriority && filterPriority !== "all"].filter(Boolean).length;
    const clearFilters = () => {
      setSearchQ("");
      setFilterSport("all");
      setFilterStatus("all");
      setFilterPriority("all");
    };
    return (
      <div className="rounded-3xl border border-slate-200 bg-white p-3 mb-5 shadow-sm">
        <div className="flex flex-wrap items-center gap-3">
          <div className="flex items-center gap-2 text-sm font-semibold text-slate-700 px-2">
            <Filter className="w-4 h-4 text-slate-400" />
            Filtres
            {activeCount > 0 && <Badge>{activeCount} actif{activeCount > 1 ? "s" : ""}</Badge>}
          </div>
          <div className="relative flex-1 min-w-[220px]">
            <Search className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
            <input className="w-full pl-9 pr-4 py-2.5 rounded-2xl border border-slate-200 text-sm bg-slate-50 focus:bg-white focus:outline-none focus:ring-2 focus:ring-slate-900" placeholder="Rechercher..." value={searchQ} onChange={(e) => setSearchQ(e.target.value)} />
          </div>
          {withSport && (
            <select className="rounded-2xl border border-slate-200 px-3 py-2.5 text-sm bg-white focus:outline-none focus:ring-2 focus:ring-slate-900" value={filterSport} onChange={(e) => setFilterSport(e.target.value)}>
              <option value="all">Tous les sports</option>
              {allSports.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
            </select>
          )}
          {withStatus && (
            <select className="rounded-2xl border border-slate-200 px-3 py-2.5 text-sm bg-white focus:outline-none focus:ring-2 focus:ring-slate-900" value={filterStatus} onChange={(e) => setFilterStatus(e.target.value)}>
              <option value="all">Tous les statuts</option>
              {statusOptions.map((s) => <option key={s} value={s}>{s}</option>)}
            </select>
          )}
          {withPriority && (
            <select className="rounded-2xl border border-slate-200 px-3 py-2.5 text-sm bg-white focus:outline-none focus:ring-2 focus:ring-slate-900" value={filterPriority} onChange={(e) => setFilterPriority(e.target.value)}>
              <option value="all">Toutes les priorités</option>
              {PRIORITY_OPTIONS.map((p) => <option key={p} value={p}>{p}</option>)}
            </select>
          )}
          {activeCount > 0 && (
            <button onClick={clearFilters} className="flex items-center gap-2 px-3 py-2.5 rounded-2xl text-sm font-semibold text-slate-600 hover:bg-slate-100">
              <X className="w-4 h-4" />
              Réinitialiser
            </button>
          )}
        </div>
      </div>
    );
  }

  function PageHeader({ title, subtitle, onNew, newLabel, actions }) {
    return (
      <div className="flex items-center justify-between mb-6 gap-3 flex-wrap">
        <div>
          <h1 className="text-2xl font-bold text-slate-900">{title}</h1>
          {subtitle && <p className="text-slate-500 mt-1">{subtitle}</p>}
        </div>
        <div className="flex items-center gap-2 flex-wrap">
          {actions}
          {onNew && (
            <button onClick={onNew} className="flex items-center gap-2 px-4 py-2 rounded-2xl bg-slate-900 text-white text-sm">
              <Plus className="w-4 h-4" /> {newLabel}
            </button>
          )}
        </div>
      </div>
    );
  }

  function ActionBtns({ onEdit, onDelete, onSubProject }) {
    return (
      <div className="flex gap-2 shrink-0">
        {onSubProject && (
          <button onClick={onSubProject} className="px-3 py-2 rounded-xl border bg-white hover:bg-violet-50 text-violet-700" title="Créer un sous-projet"><FolderKanban className="w-4 h-4" /></button>
        )}
        <button onClick={onEdit} className="px-3 py-2 rounded-xl border bg-white hover:bg-slate-50"><Edit3 className="w-4 h-4" /></button>
        <button onClick={onDelete} className="px-3 py-2 rounded-xl border bg-white hover:bg-red-50 text-red-600"><Trash2 className="w-4 h-4" /></button>
      </div>
    );
  }

  const tableKanbanOptions = [
    { id: "table", label: "Tableau", Icon: ListChecks },
    { id: "kanban", label: "Kanban", Icon: Layers },
  ];
  const listKanbanOptions = [
    { id: "list", label: "Liste", Icon: ListChecks },
    { id: "kanban", label: "Kanban", Icon: Layers },
  ];

  const projectViewOptions = [
    { id: "tree", label: "Arborescence", Icon: FolderKanban },
    ...tableKanbanOptions,
  ];

  const renderProjectKanbanCard = (project) => (
    <div className="space-y-3">
      <div className="flex items-start justify-between gap-3">
        <div className="min-w-0">
          <h3 className="font-semibold text-sm text-slate-900 truncate">{project.name}</h3>
          <p className="text-xs text-slate-500 truncate">{sportName(project.sport)}</p>
        </div>
        <ActionBtns onSubProject={() => handleNewSubProject(project)} onEdit={() => handleEditProject(project)} onDelete={() => requestDelete("project", project)} />
      </div>
      <div className="flex flex-wrap gap-1.5">
        <Badge className={badgeClass("priority", project.priority)}>{project.priority}</Badge>
        {project.owner && <Badge>{project.owner}</Badge>}
      </div>
      <ProgressBar value={project.computedTotalProgress || 0} label="Avancement" />
      <div className="flex justify-between text-xs text-slate-500">
        <span>{wsTasks.filter((task) => task.projectId === project.id).length} tâche{wsTasks.filter((task) => task.projectId === project.id).length > 1 ? "s" : ""}</span>
        {project.endDate && <span>{project.endDate}</span>}
      </div>
    </div>
  );

  const renderTaskKanbanCard = (task) => {
    const progress = normalizeTaskProgress(task.progress, task.status);
    return (
      <div className="space-y-3">
        <div className="flex items-start justify-between gap-3">
          <div className="min-w-0">
            <h3 className="font-semibold text-sm text-slate-900 truncate">{task.title}</h3>
            <p className="text-xs text-slate-500 truncate">{data.projects.find((project) => project.id === task.projectId)?.name || "Sans projet"}</p>
          </div>
          <ActionBtns onEdit={() => { setForm({ ...task }); setTaskModal({ open: true, item: task }); }} onDelete={() => requestDelete("task", task)} />
        </div>
        <div className="flex flex-wrap gap-1.5">
          <Badge className={badgeClass("priority", task.priority)}>{task.priority}</Badge>
          {task.assignee && <Badge>{task.assignee}</Badge>}
        </div>
        <ProgressBar value={progress} label="Avancement" />
        {(task.startDate || task.endDate) && (
          <p className="text-xs text-slate-500">{task.startDate || "—"} → {task.endDate || "—"}</p>
        )}
      </div>
    );
  };

  const renderAthleteKanbanCard = (athlete) => (
    <div className="space-y-3">
      <div className="flex items-start justify-between gap-3">
        <div className="min-w-0">
          <h3 className="font-semibold text-sm text-slate-900 truncate">{athlete.name}</h3>
          <p className="text-xs text-slate-500 truncate">{sectionName(athlete.sectionId) || sportName(athlete.sport)}</p>
        </div>
        <ActionBtns onEdit={() => { setForm({ ...athlete }); setAthleteModal({ open: true, item: athlete }); }} onDelete={() => requestDelete("athlete", athlete)} />
      </div>
      <div className="flex flex-wrap gap-1.5">
        {athlete.subSection && <Badge>{athlete.subSection}</Badge>}
        {athlete.category && <Badge>{athlete.category}</Badge>}
        {athlete.coach && <Badge>{athlete.coach}</Badge>}
      </div>
      {(athlete.email || athlete.phone) && <p className="text-xs text-slate-500 truncate">{athlete.email || athlete.phone}</p>}
    </div>
  );

  const renderClubKanbanCard = (club) => (
    <div className="space-y-3">
      <div className="flex items-start justify-between gap-3">
        <div className="min-w-0">
          <h3 className="font-semibold text-sm text-slate-900 truncate">{club.name}</h3>
          <p className="text-xs text-slate-500 truncate">{club.city || "Section JAM"}</p>
        </div>
        <ActionBtns onEdit={() => handleEditSection(club)} onDelete={() => requestDelete("club", club)} />
      </div>
      <div className="flex flex-wrap gap-1.5">
        {club.subSections?.length > 0 && <Badge>{club.subSections.length} sous-sections</Badge>}
      </div>
      {(club.contactName || club.email) && <p className="text-xs text-slate-500 truncate">{club.contactName || club.email}</p>}
    </div>
  );

  const renderDocumentKanbanCard = (doc) => (
    <div className="space-y-3">
      <div className="flex items-start justify-between gap-3">
        <div className="min-w-0">
          <h3 className="font-semibold text-sm text-slate-900 truncate">{doc.title}</h3>
          <p className="text-xs text-slate-500 truncate">{doc.url ? "Lien externe" : "Document interne"}</p>
        </div>
        <ActionBtns onEdit={() => { setForm({ ...doc }); setDocModal({ open: true, item: doc }); }} onDelete={() => requestDelete("document", doc)} />
      </div>
      <div className="flex flex-wrap gap-1.5">
        {doc.type && <Badge>{doc.type}</Badge>}
        <Badge>{sportName(doc.sport)}</Badge>
      </div>
      {doc.notes && <p className="text-xs text-slate-500 line-clamp-2">{doc.notes}</p>}
    </div>
  );

  const renderConvocationKanbanCard = (conv) => (
    <div className="space-y-3">
      <div className="flex items-start justify-between gap-3">
        <div className="min-w-0">
          <h3 className="font-semibold text-sm text-slate-900 truncate">{conv.title}</h3>
          <p className="text-xs text-slate-500 truncate">{conv.location || sportName(conv.sport)}</p>
        </div>
        <ActionBtns onEdit={() => { setForm({ ...conv }); setConvModal({ open: true, item: conv }); }} onDelete={() => requestDelete("convocation", conv)} />
      </div>
      <div className="flex flex-wrap gap-1.5">
        {conv.type && <Badge>{conv.type}</Badge>}
        <Badge>{sportName(conv.sport)}</Badge>
      </div>
      {(conv.date || conv.time) && <p className="text-xs text-slate-500">{conv.date || "—"} {conv.time || ""}</p>}
    </div>
  );

  const renderNotificationKanbanCard = (notif) => (
    <div className="space-y-3">
      <div className="flex items-start justify-between gap-3">
        <div className="min-w-0">
          <h3 className="font-semibold text-sm text-slate-900 truncate">{notif.title}</h3>
          <p className="text-xs text-slate-500 truncate">{notif.audience || sportName(notif.sport)}</p>
        </div>
        <ActionBtns onEdit={() => { setForm({ ...notif }); setNotifModal({ open: true, item: notif }); }} onDelete={() => requestDelete("notification", notif)} />
      </div>
      <div className="flex flex-wrap gap-1.5">
        {notif.channel && <Badge>{notif.channel}</Badge>}
        <Badge>{sportName(notif.sport)}</Badge>
      </div>
      {notif.sentAt && <p className="text-xs text-slate-500">Envoi : {notif.sentAt}</p>}
    </div>
  );

  const widgetToneClass = (tone) => ({
    red: "bg-red-50 text-red-700 border-red-100",
    amber: "bg-amber-50 text-amber-700 border-amber-100",
    green: "bg-emerald-50 text-emerald-700 border-emerald-100",
    blue: "bg-blue-50 text-blue-700 border-blue-100",
    violet: "bg-violet-50 text-violet-700 border-violet-100",
    slate: "bg-slate-50 text-slate-700 border-slate-100",
  }[tone || "slate"] || "bg-slate-50 text-slate-700 border-slate-100");

  const widgetAccentClass = (tone) => ({
    red: "bg-red-500",
    amber: "bg-amber-500",
    green: "bg-emerald-500",
    blue: "bg-blue-500",
    violet: "bg-violet-500",
    slate: "bg-slate-900",
  }[tone || "slate"] || "bg-slate-900");

  const dashboardWidgetSpan = (widget) => widget.size === "wide" ? "md:col-span-2" : "";
  const dashboardWidgetKeyDown = (event, widget) => {
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      openDashboardWidget(widget);
    }
  };

  const renderWidgetActions = (widget) => dashboardEditMode && (
    <div className="absolute right-3 top-3 flex gap-1 rounded-xl border bg-white/90 p-1 shadow-sm" onClick={(event) => event.stopPropagation()}>
      <button
        type="button"
        onClick={() => openWidgetEditor(widget)}
        className="rounded-lg p-1.5 text-slate-500 hover:bg-slate-100 hover:text-slate-900"
        title="Modifier le widget"
      >
        <Edit3 className="h-4 w-4" />
      </button>
      <button
        type="button"
        onClick={() => deleteDashboardWidget(widget.id)}
        className="rounded-lg p-1.5 text-slate-500 hover:bg-red-50 hover:text-red-600"
        title="Supprimer le widget"
      >
        <Trash2 className="h-4 w-4" />
      </button>
    </div>
  );

  const renderNumberWidget = (widget) => (
    <Card
      key={widget.id}
      role="button"
      tabIndex={0}
      onClick={() => openDashboardWidget(widget)}
      onKeyDown={(event) => dashboardWidgetKeyDown(event, widget)}
      className={`relative cursor-pointer p-5 transition hover:-translate-y-0.5 hover:border-slate-300 hover:shadow-md focus:outline-none focus:ring-2 focus:ring-slate-900/20 ${dashboardWidgetSpan(widget)}`}
    >
      {renderWidgetActions(widget)}
      <div className="flex min-h-36 flex-col justify-between gap-5">
        <div className="flex items-start justify-between gap-4 pr-16">
          <div className="min-w-0">
            <p className="truncate text-sm font-semibold text-slate-600">{widget.title}</p>
            <p className="mt-1 text-xs text-slate-400">{dashboardSourceLabel(widget.source)} · {dashboardFilterLabel(widget.filter)}</p>
          </div>
          <span className={`inline-flex h-11 w-11 shrink-0 items-center justify-center rounded-2xl border ${widgetToneClass(widget.tone)}`}>
            <BarChart3 className="h-5 w-5" />
          </span>
        </div>
        <div>
          <p className="text-4xl font-bold tracking-normal text-slate-950">{dashboardNumberValue(widget)}</p>
          <p className="mt-3 text-sm font-medium text-slate-500">Ouvrir {dashboardSourceLabel(widget.source).toLowerCase()}</p>
        </div>
      </div>
    </Card>
  );

  const renderBarWidget = (widget) => {
    const rows = dashboardGroupedData(widget).slice(0, 6);
    const max = Math.max(1, ...rows.map((row) => row.value));
    return (
      <Card
        key={widget.id}
        role="button"
        tabIndex={0}
        onClick={() => openDashboardWidget(widget)}
        onKeyDown={(event) => dashboardWidgetKeyDown(event, widget)}
        className={`relative cursor-pointer p-5 transition hover:-translate-y-0.5 hover:border-slate-300 hover:shadow-md focus:outline-none focus:ring-2 focus:ring-slate-900/20 ${dashboardWidgetSpan(widget)}`}
      >
        {renderWidgetActions(widget)}
        <div className="mb-5 pr-20">
          <p className="text-sm font-semibold text-slate-900">{widget.title}</p>
          <p className="mt-1 text-xs text-slate-400">{dashboardSourceLabel(widget.source)} · groupé par {DASHBOARD_GROUPS.find((item) => item.id === widget.groupBy)?.label || "Statut"}</p>
        </div>
        <div className="space-y-3">
          {rows.map((row) => (
            <button
              key={row.label}
              type="button"
              onClick={(event) => {
                event.stopPropagation();
                openDashboardWidget(widget, row.label);
              }}
              className="grid w-full grid-cols-[minmax(90px,150px)_1fr_32px] items-center gap-3 rounded-xl px-2 py-1 text-left text-sm hover:bg-slate-50"
              title={`Voir les données : ${row.label}`}
            >
              <span className="truncate text-slate-600">{row.label}</span>
              <div className="h-3 overflow-hidden rounded-full bg-slate-100">
                <div className={`h-full rounded-full ${widgetAccentClass(widget.tone)}`} style={{ width: `${Math.max(8, (row.value / max) * 100)}%` }} />
              </div>
              <span className="text-right font-semibold text-slate-900">{row.value}</span>
            </button>
          ))}
          {!rows.length && <p className="rounded-2xl border border-dashed p-5 text-center text-sm text-slate-400">Aucune donnée pour ce widget.</p>}
        </div>
      </Card>
    );
  };

  const renderDonutWidget = (widget) => {
    const rows = dashboardGroupedData(widget).slice(0, 5);
    const total = rows.reduce((sum, row) => sum + row.value, 0);
    const colors = ["#7c3aed", "#2563eb", "#059669", "#d97706", "#dc2626"];
    let start = 0;
    const gradient = total
      ? rows.map((row, index) => {
          const end = start + (row.value / total) * 100;
          const segment = `${colors[index % colors.length]} ${start}% ${end}%`;
          start = end;
          return segment;
        }).join(", ")
      : "#e2e8f0 0% 100%";
    return (
      <Card
        key={widget.id}
        role="button"
        tabIndex={0}
        onClick={() => openDashboardWidget(widget)}
        onKeyDown={(event) => dashboardWidgetKeyDown(event, widget)}
        className={`relative cursor-pointer p-5 transition hover:-translate-y-0.5 hover:border-slate-300 hover:shadow-md focus:outline-none focus:ring-2 focus:ring-slate-900/20 ${dashboardWidgetSpan(widget)}`}
      >
        {renderWidgetActions(widget)}
        <div className="mb-5 pr-20">
          <p className="text-sm font-semibold text-slate-900">{widget.title}</p>
          <p className="mt-1 text-xs text-slate-400">{dashboardSourceLabel(widget.source)} · {dashboardFilterLabel(widget.filter)}</p>
        </div>
        <div className="grid items-center gap-6 sm:grid-cols-[190px_1fr]">
          <div className="relative mx-auto h-44 w-44 rounded-full" style={{ background: `conic-gradient(${gradient})` }}>
            <div className="absolute inset-8 flex flex-col items-center justify-center rounded-full bg-white">
              <span className="text-3xl font-bold text-slate-950">{total}</span>
              <span className="text-xs text-slate-400">total</span>
            </div>
          </div>
          <div className="space-y-2">
            {rows.map((row, index) => (
              <button
                key={row.label}
                type="button"
                onClick={(event) => {
                  event.stopPropagation();
                  openDashboardWidget(widget, row.label);
                }}
                className="flex w-full items-center justify-between gap-3 rounded-xl px-2 py-1.5 text-left text-sm hover:bg-slate-50"
                title={`Voir les données : ${row.label}`}
              >
                <span className="flex min-w-0 items-center gap-2">
                  <span className="h-2.5 w-2.5 rounded-full" style={{ backgroundColor: colors[index % colors.length] }} />
                  <span className="truncate text-slate-600">{row.label}</span>
                </span>
                <span className="font-semibold text-slate-900">{row.value}</span>
              </button>
            ))}
            {!rows.length && <p className="text-sm text-slate-400">Aucune donnée.</p>}
          </div>
        </div>
      </Card>
    );
  };

  const renderDashboardWidget = (widget) => {
    if (widget.type === "bar") return renderBarWidget(widget);
    if (widget.type === "donut") return renderDonutWidget(widget);
    return renderNumberWidget(widget);
  };

  return (
    <div className="min-h-screen bg-slate-50 flex">
      <input ref={importInputRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={handleImportFile} />
      {/* Sidebar */}
      <aside className="w-64 bg-white border-r border-slate-200 flex flex-col shrink-0 sticky top-0 h-screen">
        <div className="px-5 py-4 border-b border-slate-200">
          <div className="flex items-center gap-3">
            <div className="w-9 h-9 rounded-2xl bg-slate-900 flex items-center justify-center shrink-0">
              <Trophy className="w-5 h-5 text-white" />
            </div>
            <div>
              <p className="font-bold text-slate-900 leading-none">JAM SaaS</p>
              <p className="text-xs text-slate-500 mt-0.5">v13</p>
            </div>
          </div>
        </div>
        <WorkspaceSelector
          workspaces={data.workspaces || []}
          currentWorkspaceId={currentWsId}
          onChange={switchWorkspace}
          onNew={() => { setWsForm({ color: "#0f172a" }); setWsModal(true); }}
        />
        <nav className="flex-1 px-3 py-4 space-y-5 overflow-auto">
          {navGroups.map((group) => (
            <div key={group.label}>
              <p className="px-3 mb-2 text-[11px] font-bold uppercase tracking-wide text-slate-400">{group.label}</p>
              <div className="space-y-1">
                {group.items.map(({ id, label, Icon, count, alert, tone, tooltip }) => (
                  <button
                    key={id}
                    onClick={() => goToView(id)}
                    className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-2xl text-sm transition-colors ${view === id || (id === "dashboard" && view === "dashboardData") ? "bg-slate-900 text-white shadow-sm" : "text-slate-600 hover:bg-slate-100"}`}
                  >
                    <Icon className="w-4 h-4 shrink-0" />
                    <span className="flex-1 text-left truncate">{label}</span>
                    <NavCounter
                      value={alert > 0 ? alert : count}
                      active={view === id || (id === "dashboard" && view === "dashboardData")}
                      alert={alert > 0}
                      tone={tone}
                      tooltip={tooltip}
                    />
                  </button>
                ))}
              </div>
            </div>
          ))}
        </nav>
        <div className="px-4 py-4 border-t border-slate-200">
          <div className="rounded-2xl border border-slate-200 bg-slate-50 p-3">
            <div className="flex items-center justify-between text-xs font-semibold text-slate-600 mb-2">
              <span>Avancement global</span>
              <span>{avgProgress}%</span>
            </div>
            <div className="h-2 rounded-full bg-white border border-slate-200 overflow-hidden">
              <div className={`h-full rounded-full ${progressTone(avgProgress)}`} style={{ width: `${avgProgress}%` }} />
            </div>
            <div className="grid grid-cols-2 gap-2 mt-3 text-xs">
              <button
                type="button"
                onClick={() => goToView("tasks")}
                className="rounded-xl bg-white border border-slate-200 p-2 text-left transition hover:border-red-200 hover:bg-red-50 focus:outline-none focus:ring-2 focus:ring-red-200"
                title="Voir les tâches en retard"
              >
                <p className="font-bold text-slate-900">{overdueTasks.length}</p>
                <p className="text-slate-500">retard</p>
              </button>
              <button
                type="button"
                onClick={() => goToView("tasks")}
                className="rounded-xl bg-white border border-slate-200 p-2 text-left transition hover:border-blue-200 hover:bg-blue-50 focus:outline-none focus:ring-2 focus:ring-blue-200"
                title="Voir les tâches à venir"
              >
                <p className="font-bold text-slate-900">{upcomingTasks.length}</p>
                <p className="text-slate-500">à venir</p>
              </button>
            </div>
          </div>
        </div>
      </aside>

      {/* Main */}
      <main className="flex-1 overflow-auto">

        {/* ── DASHBOARD ── */}
        {view === "dashboard" && (
          <div className="p-6 max-w-7xl mx-auto">
            <div className="mb-6 flex flex-wrap items-start justify-between gap-4">
              <div>
                <h1 className="text-3xl font-bold text-slate-950">Tableau de bord</h1>
                <p className="mt-1 text-slate-500">Une vue pilotable de la JAM, avec des indicateurs que tu peux adapter.</p>
              </div>
              <div className="flex flex-wrap gap-2">
                <button
                  type="button"
                  onClick={() => openWidgetEditor()}
                  className="inline-flex items-center gap-2 rounded-2xl border bg-white px-4 py-2 text-sm font-semibold text-slate-700 shadow-sm hover:bg-slate-50"
                >
                  <Plus className="h-4 w-4" /> Ajouter un widget
                </button>
                <button
                  type="button"
                  onClick={() => setDashboardEditMode((value) => !value)}
                  className={`inline-flex items-center gap-2 rounded-2xl px-4 py-2 text-sm font-semibold shadow-sm ${dashboardEditMode ? "bg-slate-900 text-white" : "border bg-white text-slate-700 hover:bg-slate-50"}`}
                >
                  <BarChart3 className="h-4 w-4" /> {dashboardEditMode ? "Terminer" : "Personnaliser"}
                </button>
                {dashboardEditMode && (
                  <button
                    type="button"
                    onClick={resetDashboardWidgets}
                    className="inline-flex items-center gap-2 rounded-2xl border bg-white px-4 py-2 text-sm font-semibold text-slate-500 shadow-sm hover:bg-slate-50"
                  >
                    Réinitialiser
                  </button>
                )}
              </div>
            </div>

            <div className="mb-5 grid gap-4 md:grid-cols-4">
              {dashboardWidgets.map((widget) => renderDashboardWidget(widget))}
            </div>

            <div className="grid gap-4 lg:grid-cols-[1.2fr_0.8fr]">
              <Card className="p-5">
                <div className="mb-5 flex items-start justify-between gap-3">
                  <div>
                    <h2 className="text-lg font-semibold text-slate-950">Priorités opérationnelles</h2>
                    <p className="mt-1 text-sm text-slate-500">Les signaux qui doivent guider la journée.</p>
                  </div>
                  <Badge className={overdueTasks.length ? "bg-red-50 text-red-700 border-red-200" : "bg-green-50 text-green-700 border-green-200"}>
                    {overdueTasks.length ? `${overdueTasks.length} retard${overdueTasks.length > 1 ? "s" : ""}` : "À jour"}
                  </Badge>
                </div>
                <div className="grid gap-3 md:grid-cols-3">
                  <button type="button" onClick={() => goToView("tasks")} className="rounded-2xl border border-red-100 bg-red-50 p-4 text-left transition hover:-translate-y-0.5 hover:shadow-sm">
                    <p className="text-xs font-bold uppercase tracking-wide text-red-700">Retards</p>
                    <p className="mt-2 text-3xl font-bold text-red-700">{overdueTasks.length + overdueProjects.length}</p>
                    <p className="mt-1 text-sm text-red-700/70">tâches + projets</p>
                  </button>
                  <button type="button" onClick={() => goToView("tasks")} className="rounded-2xl border border-amber-100 bg-amber-50 p-4 text-left transition hover:-translate-y-0.5 hover:shadow-sm">
                    <p className="text-xs font-bold uppercase tracking-wide text-amber-700">Critiques</p>
                    <p className="mt-2 text-3xl font-bold text-amber-700">{criticalTasks.length}</p>
                    <p className="mt-1 text-sm text-amber-700/70">haute priorité</p>
                  </button>
                  <button type="button" onClick={() => goToView("tasks")} className="rounded-2xl border border-blue-100 bg-blue-50 p-4 text-left transition hover:-translate-y-0.5 hover:shadow-sm">
                    <p className="text-xs font-bold uppercase tracking-wide text-blue-700">14 jours</p>
                    <p className="mt-2 text-3xl font-bold text-blue-700">{upcomingTasks.length}</p>
                    <p className="mt-1 text-sm text-blue-700/70">échéances proches</p>
                  </button>
                </div>
                <div className="mt-5">
                  <ProgressBar value={avgProgress} label="Progression moyenne des tâches" />
                </div>
              </Card>

              <Card className="p-5">
                <h2 className="mb-4 text-lg font-semibold text-slate-950">Données à compléter</h2>
                <div className="space-y-2">
                  <button type="button" onClick={() => goToView("athletes")} className="flex w-full justify-between rounded-xl px-2 py-2 text-left text-sm hover:bg-slate-50"><span>Athlètes sans email</span><Badge>{wsAthletes.filter((a) => !a.email).length}</Badge></button>
                  <button type="button" onClick={() => goToView("athletes")} className="flex w-full justify-between rounded-xl px-2 py-2 text-left text-sm hover:bg-slate-50"><span>Athlètes sans téléphone</span><Badge>{wsAthletes.filter((a) => !a.phone).length}</Badge></button>
                  <button type="button" onClick={() => goToView("athletes")} className="flex w-full justify-between rounded-xl px-2 py-2 text-left text-sm hover:bg-slate-50"><span>Athlètes sans section</span><Badge>{wsAthletes.filter((a) => !a.sectionId).length}</Badge></button>
                  <button type="button" onClick={() => goToView("athletes")} className="flex w-full justify-between rounded-xl px-2 py-2 text-left text-sm hover:bg-slate-50"><span>Athlètes sans profil</span><Badge>{wsAthletes.filter((a) => !a.category).length}</Badge></button>
                  <button type="button" onClick={() => goToView("tasks")} className="flex w-full justify-between rounded-xl px-2 py-2 text-left text-sm hover:bg-slate-50"><span>Tâches non assignées</span><Badge>{wsTasks.filter((task) => !task.assignee).length}</Badge></button>
                </div>
              </Card>
            </div>
          </div>
        )}

        {/* ── DASHBOARD DATA ── */}
        {view === "dashboardData" && (
          <div className="p-6 max-w-7xl mx-auto">
            {dashboardDrilldown?.widget ? (
              <>
                <div className="mb-5 flex flex-wrap items-start justify-between gap-4">
                  <div>
                    <button
                      type="button"
                      onClick={closeDashboardDrilldown}
                      className="mb-3 inline-flex items-center gap-2 text-sm font-semibold text-slate-500 hover:text-slate-900"
                    >
                      ← Retour au tableau de bord
                    </button>
                    <h1 className="text-3xl font-bold text-slate-950">Données du graphique</h1>
                    <p className="mt-1 text-slate-500">
                      {dashboardDrilldown.widget.title}
                      {dashboardDrilldown.groupValue ? ` · ${dashboardDrilldown.groupValue}` : ""}
                    </p>
                  </div>
                  <div className="flex flex-wrap gap-2">
                    <button
                      type="button"
                      onClick={() => openWidgetEditor(dashboardDrilldown.widget)}
                      className="inline-flex items-center gap-2 rounded-2xl border bg-white px-4 py-2 text-sm font-semibold text-slate-700 shadow-sm hover:bg-slate-50"
                    >
                      <Edit3 className="h-4 w-4" /> Modifier le widget
                    </button>
                    <button
                      type="button"
                      onClick={() => goToView(dashboardDrilldown.widget.source === "clubs" ? "clubs" : dashboardDrilldown.widget.source || "tasks")}
                      className="inline-flex items-center gap-2 rounded-2xl bg-slate-900 px-4 py-2 text-sm font-semibold text-white shadow-sm hover:bg-slate-800"
                    >
                      Ouvrir la vue complète
                    </button>
                  </div>
                </div>

                <div className="mb-4 grid gap-3 md:grid-cols-4">
                  <Card className="p-4">
                    <p className="text-xs font-bold uppercase tracking-wide text-slate-400">Source</p>
                    <p className="mt-1 font-semibold text-slate-900">{dashboardSourceLabel(dashboardDrilldown.widget.source)}</p>
                  </Card>
                  <Card className="p-4">
                    <p className="text-xs font-bold uppercase tracking-wide text-slate-400">Type</p>
                    <p className="mt-1 font-semibold text-slate-900">{dashboardTypeLabel(dashboardDrilldown.widget.type)}</p>
                  </Card>
                  <Card className="p-4">
                    <p className="text-xs font-bold uppercase tracking-wide text-slate-400">Filtre</p>
                    <p className="mt-1 font-semibold text-slate-900">{dashboardFilterLabel(dashboardDrilldown.widget.filter)}</p>
                  </Card>
                  <Card className="p-4">
                    <p className="text-xs font-bold uppercase tracking-wide text-slate-400">Résultat</p>
                    <p className="mt-1 font-semibold text-slate-900">{dashboardDrilldownVisibleRows.length} ligne{dashboardDrilldownVisibleRows.length > 1 ? "s" : ""}</p>
                  </Card>
                </div>

                <Card className="mb-4 p-4">
                  <div className="flex flex-wrap items-center gap-3">
                    <div className="relative min-w-[260px] flex-1">
                      <Search className="absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-slate-400" />
                      <input
                        className="w-full rounded-2xl border border-slate-200 bg-white py-2.5 pl-9 pr-4 text-sm focus:outline-none focus:ring-2 focus:ring-slate-900/10"
                        placeholder="Rechercher dans les données..."
                        value={dashboardDataSearch}
                        onChange={(event) => setDashboardDataSearch(event.target.value)}
                      />
                    </div>
                    <select
                      className="rounded-2xl border border-slate-200 bg-white px-4 py-2.5 text-sm"
                      value={dashboardDataSort}
                      onChange={(event) => setDashboardDataSort(event.target.value)}
                    >
                      <option value="name">Tri : nom</option>
                      <option value="status">Tri : statut</option>
                      <option value="priority">Tri : priorité</option>
                      <option value="date">Tri : échéance</option>
                    </select>
                    {dashboardDrilldown.groupValue && (
                      <Badge className="bg-violet-50 text-violet-700 border-violet-200">
                        Segment : {dashboardDrilldown.groupValue}
                      </Badge>
                    )}
                  </div>
                </Card>

                <Card className="overflow-hidden">
                  <div className="overflow-x-auto">
                    <table className="w-full border-collapse text-sm">
                      <thead className="bg-slate-50">
                        <tr>
                          {dashboardDataColumns(dashboardDrilldown.widget.source).map((column) => (
                            <th key={column.key} className="border-b border-r border-slate-200 px-4 py-3 text-left text-xs font-bold uppercase tracking-wide text-slate-500">
                              {column.label}
                            </th>
                          ))}
                          <th className="border-b border-slate-200 px-4 py-3 text-right text-xs font-bold uppercase tracking-wide text-slate-500">Action</th>
                        </tr>
                      </thead>
                      <tbody>
                        {dashboardDrilldownVisibleRows.map((row) => (
                          <tr key={row.id} className="border-b border-slate-100 bg-white hover:bg-slate-50">
                            {dashboardDataColumns(dashboardDrilldown.widget.source).map((column) => (
                              <td key={column.key} className="border-r border-slate-100 px-4 py-3 text-slate-700">
                                {column.render(row)}
                              </td>
                            ))}
                            <td className="px-4 py-3 text-right">
                              <button
                                type="button"
                                onClick={() => editDashboardRow(row, dashboardDrilldown.widget.source)}
                                className="inline-flex items-center gap-1.5 rounded-xl border bg-white px-3 py-1.5 text-xs font-semibold text-slate-600 hover:bg-slate-50"
                              >
                                <Edit3 className="h-3.5 w-3.5" /> Modifier
                              </button>
                            </td>
                          </tr>
                        ))}
                        {!dashboardDrilldownVisibleRows.length && (
                          <tr>
                            <td colSpan={dashboardDataColumns(dashboardDrilldown.widget.source).length + 1} className="px-4 py-10 text-center text-slate-500">
                              Aucune ligne ne correspond à ce graphique.
                            </td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                </Card>
              </>
            ) : (
              <Card className="p-8 text-center">
                <h1 className="text-2xl font-bold text-slate-950">Aucune donnée sélectionnée</h1>
                <p className="mt-2 text-slate-500">Retourne au tableau de bord et clique sur un widget.</p>
                <button onClick={closeDashboardDrilldown} className="mt-5 rounded-2xl bg-slate-900 px-4 py-2 text-sm font-semibold text-white">
                  Tableau de bord
                </button>
              </Card>
            )}
          </div>
        )}

        {/* ── PROJECTS ── */}
        {view === "projects" && (
          <div className="p-6 max-w-7xl mx-auto">
            <div className="mb-5 rounded-3xl border border-slate-200 bg-white p-5 shadow-sm">
              <div className="flex items-start justify-between flex-wrap gap-4">
                <div className="min-w-0">
                  <div className="flex items-center gap-3">
                    <span className="inline-flex h-11 w-11 items-center justify-center rounded-2xl bg-slate-900 text-white shadow-sm">
                      <FolderKanban className="w-5 h-5" />
                    </span>
                    <div>
                      <h1 className="text-2xl font-bold text-slate-900">Projets</h1>
                      <p className="text-slate-500 mt-0.5">{filteredProjects.length} projet{filteredProjects.length > 1 ? "s" : ""} affiché{filteredProjects.length > 1 ? "s" : ""}</p>
                    </div>
                  </div>
                  <div className="mt-4 flex flex-wrap gap-2">
                    <Badge className="bg-slate-50 text-slate-700 border-slate-200">{wsProjects.length} total</Badge>
                    <Badge className="bg-blue-50 text-blue-700 border-blue-200">{wsProjects.filter((project) => project.status === "En cours").length} en cours</Badge>
                    <Badge className="bg-emerald-50 text-emerald-700 border-emerald-200">{wsProjects.filter((project) => project.status === "Terminé").length} terminés</Badge>
                    {overdueProjects.length > 0 && <Badge className="bg-red-50 text-red-700 border-red-200">{overdueProjects.length} en retard</Badge>}
                  </div>
                </div>
                <div className="flex w-full items-center gap-2 flex-wrap justify-start xl:w-auto xl:justify-end">
                  <ViewToggle options={projectViewOptions} value={projectViewMode} onChange={setProjectViewMode} />
                  {fileActions("projects", "projets")}
                  <button onClick={() => setGlobalColumnModal({ entity: "projects", label: "Projets" })}
                    className="flex items-center gap-2 px-4 py-2 rounded-2xl border bg-white text-slate-700 text-sm font-medium hover:bg-slate-50">
                    <Plus className="w-4 h-4" /> Ajouter une colonne
                  </button>
                  <button onClick={() => setGuidedOpen(true)}
                    className="flex items-center gap-2 px-4 py-2 rounded-2xl border border-slate-900 text-slate-900 text-sm font-medium hover:bg-slate-50">
                    <Send className="w-4 h-4" /> Saisie guidée
                  </button>
                  <button onClick={openNewProjectModal}
                    className="flex items-center gap-2 px-4 py-2 rounded-2xl bg-slate-900 text-white text-sm font-medium">
                    <Plus className="w-4 h-4" /> Nouveau projet
                  </button>
                </div>
              </div>
            </div>
            <FilterBar withPriority={false} />
            {projectViewMode === "tree" ? (
              <>
                <ProjectTree
                  nodes={projectTree}
                  tasks={wsTasks}
                  projects={enrichedProjects}
                  expanded={expanded}
                  setExpanded={setExpanded}
                  onQuickTask={handleQuickTask}
                  onNewSubProject={handleNewSubProject}
                  onEditProject={handleEditProject}
                  requestDelete={requestDelete}
                  onDuplicate={handleDuplicate}
                  projectCustomFields={projectCustomFields}
                />
                {!filteredProjects.length && <p className="text-center text-slate-500 py-12">Aucun projet trouvé.</p>}
              </>
            ) : projectViewMode === "table" ? (
              <ProjectTable
                projects={filteredProjects}
                tasks={wsTasks}
                onEdit={handleEditProject}
                requestDelete={requestDelete}
                onQuickTask={handleQuickTask}
                onNewSubProject={handleNewSubProject}
                config={data.tableConfig?.projects}
                onConfigChange={(next) => setData((prev) => ({ ...prev, tableConfig: { ...(prev.tableConfig || {}), projects: next } }))}
                customFields={legacyProjectCustomFields}
              />
            ) : (
              <KanbanBoard
                rows={filteredProjects}
                statuses={STATUS_OPTIONS}
                onStatusChange={updateProjectStatus}
                renderCard={renderProjectKanbanCard}
                emptyLabel="Aucun projet trouvé."
                celebrateStatus="Terminé"
              />
            )}
          </div>
        )}

        {/* ── TASKS ── */}
        {view === "tasks" && (
          <div className="p-6 max-w-7xl mx-auto">
            <PageHeader
              title="Tâches"
              subtitle={`${filteredTasks.length} tâche${filteredTasks.length > 1 ? "s" : ""}`}
              actions={fileActions("tasks", "tâches")}
              onNew={() => { setForm({ status: "À faire", priority: "Moyenne", progress: 0, customFieldValues: {} }); setTaskModal({ open: true, item: null }); }}
              newLabel="Nouvelle tâche"
            />
            <div className="-mt-4 mb-5 flex justify-end">
              <ViewToggle options={tableKanbanOptions} value={taskViewMode} onChange={setTaskViewMode} />
            </div>
            <FilterBar withSport={false} withPriority />
            {taskViewMode === "table" ? (
              <EntityTable
                rows={filteredTasks}
                nameKey="title"
                emptyLabel="Aucune tâche trouvée."
                builtinCols={[
                  { key: "title", label: "Tâche", type: "texte", sortable: true },
                  { key: "status", label: "Statut", type: "statut", sortable: true, options: STATUS_OPTIONS,
                    render: (t) => <Badge className={badgeClass("status", t.status)}>{t.status}</Badge> },
                  { key: "priority", label: "Priorité", type: "statut", sortable: true, options: PRIORITY_OPTIONS,
                    render: (t) => <Badge className={badgeClass("priority", t.priority)}>{t.priority}</Badge> },
                  { key: "projectId", label: "Projet", type: "texte", sortable: true,
                    accessor: (t) => data.projects.find((p) => p.id === t.projectId)?.name || "—",
                    render: (t) => <span className="text-slate-700">{data.projects.find((p) => p.id === t.projectId)?.name || "—"}</span> },
                  { key: "assignee", label: "Assigné", type: "personne", sortable: true },
                  { key: "progress", label: "Avancement", type: "nombre", sortable: true,
                    accessor: (t) => normalizeTaskProgress(t.progress, t.status),
                    render: (t) => {
                      const p = normalizeTaskProgress(t.progress, t.status);
                      return (
                        <div className="flex items-center gap-2 min-w-32">
                          <div className="flex-1 h-2 rounded-full bg-slate-200 overflow-hidden">
                            <div className={`h-full rounded-full ${progressTone(p)}`} style={{ width: `${p}%` }} />
                          </div>
                          <span className="text-xs font-bold text-slate-700 w-9 text-right">{p}%</span>
                        </div>
                      );
                    } },
                  { key: "startDate", label: "Début", type: "date", sortable: true },
                  { key: "endDate", label: "Échéance", type: "date", sortable: true },
                ]}
                config={data.tableConfig?.tasks}
                onConfigChange={(next) => setData((prev) => ({ ...prev, tableConfig: { ...(prev.tableConfig || {}), tasks: next } }))}
                onEdit={(task) => { setForm({ ...task }); setTaskModal({ open: true, item: task }); }}
                onDelete={(task) => requestDelete("task", task)}
                customFields={legacyTaskCustomFields}
              />
            ) : (
              <KanbanBoard
                rows={filteredTasks}
                statuses={STATUS_OPTIONS}
                onStatusChange={updateTaskStatus}
                renderCard={renderTaskKanbanCard}
                emptyLabel="Aucune tâche trouvée."
              />
            )}
          </div>
        )}

        {/* ── ATHLETES ── */}
        {view === "athletes" && (
          <div className="p-6 max-w-7xl mx-auto">
            <PageHeader
              title="Athlètes"
              subtitle={`${filteredAthletes.length} athlète${filteredAthletes.length > 1 ? "s" : ""}`}
              actions={fileActions("athletes", "athlètes")}
              onNew={() => { setForm({ status: "Actif" }); setAthleteModal({ open: true, item: null }); }}
              newLabel="Nouvel athlète"
            />
            <div className="-mt-4 mb-5 flex justify-end">
              <ViewToggle options={tableKanbanOptions} value={athleteViewMode} onChange={setAthleteViewMode} />
            </div>
            <FilterBar withSport={false} statusOptions={ATHLETE_STATUS} />
            {athleteViewMode === "table" ? (
              <EntityTable
                rows={filteredAthletes}
                nameKey="name"
                emptyLabel="Aucun athlète trouvé."
                builtinCols={[
                  { key: "name", label: "Athlète", type: "texte", sortable: true },
                  { key: "status", label: "Statut", type: "statut", sortable: true, options: ATHLETE_STATUS,
                    render: (a) => <Badge className={badgeClass("status", a.status === "Actif" ? "En cours" : a.status === "Blessé" ? "En attente" : "À faire")}>{a.status}</Badge> },
                  { key: "sectionId", label: "Section", type: "texte", sortable: true,
                    accessor: (a) => sectionName(a.sectionId),
                    render: (a) => <span className="text-slate-600">{sectionName(a.sectionId) || "—"}</span> },
                  { key: "subSection", label: "Sous-section", type: "texte", sortable: true },
                  { key: "category", label: "Profil", type: "statut", sortable: true, options: ATHLETE_PROFILE_OPTIONS },
                  { key: "coach", label: "Coach", type: "personne", sortable: true },
                  { key: "email", label: "Email", type: "texte", sortable: true },
                  { key: "phone", label: "Téléphone", type: "texte", sortable: true },
                ]}
                config={data.tableConfig?.athletes}
                onConfigChange={(next) => setData((prev) => ({ ...prev, tableConfig: { ...(prev.tableConfig || {}), athletes: next } }))}
                onEdit={(athlete) => { setForm({ ...athlete }); setAthleteModal({ open: true, item: athlete }); }}
                onDelete={(athlete) => requestDelete("athlete", athlete)}
                customFields={athleteCustomFields}
              />
            ) : (
              <KanbanBoard
                rows={filteredAthletes}
                statuses={ATHLETE_STATUS}
                onStatusChange={(id, status) => updateEntityStatus("athletes", id, status)}
                renderCard={renderAthleteKanbanCard}
                emptyLabel="Aucun athlète trouvé."
              />
            )}
          </div>
        )}

        {/* ── SECTIONS ── */}
        {view === "clubs" && (
          <div className="p-6 max-w-7xl mx-auto">
            <PageHeader
              title="Sections"
              subtitle={`${filteredClubs.length} section${filteredClubs.length > 1 ? "s" : ""}`}
              actions={fileActions("clubs", "sections")}
              onNew={() => { setForm({ status: "À contacter", subSections: [] }); setClubModal({ open: true, item: null }); }}
              newLabel="Nouvelle section"
            />
            <div className="-mt-4 mb-5 flex justify-end">
              <ViewToggle options={tableKanbanOptions} value={clubViewMode} onChange={setClubViewMode} />
            </div>
            <FilterBar withSport={false} statusOptions={CLUB_STATUS} />
            {clubViewMode === "table" ? (
              <EntityTable
                rows={filteredClubs}
                nameKey="name"
                emptyLabel="Aucune section trouvée."
                builtinCols={[
                  { key: "name", label: "Section", type: "texte", sortable: true },
                  { key: "status", label: "Statut", type: "statut", sortable: true, options: CLUB_STATUS,
                    render: (c) => {
                      const tone = c.status === "Validé" ? "Terminé" : c.status === "Inscrit" ? "En cours" : c.status === "Invité" ? "En attente" : "À faire";
                      return <Badge className={badgeClass("status", tone)}>{c.status}</Badge>;
                    } },
                  { key: "city", label: "Ville", type: "texte", sortable: true },
                  { key: "subSections", label: "Sous-sections", type: "texte", sortable: false,
                    accessor: (c) => (c.subSections || []).join(", "),
                    render: (c) => <span className="text-slate-600">{c.subSections?.length ? c.subSections.join(", ") : "—"}</span> },
                  { key: "contactName", label: "Responsable", type: "personne", sortable: true },
                  { key: "email", label: "Email", type: "texte", sortable: true },
                  { key: "phone", label: "Téléphone", type: "texte", sortable: true },
                  { key: "athleteCount", label: "Athlètes", type: "nombre", sortable: true,
                    accessor: (section) => sectionAthleteCount(section.id),
                    render: (section) => <span className="font-medium text-slate-700">{sectionAthleteCount(section.id)} athlète{sectionAthleteCount(section.id) > 1 ? "s" : ""}</span> },
                ]}
                config={data.tableConfig?.clubs}
                onConfigChange={(next) => setData((prev) => ({ ...prev, tableConfig: { ...(prev.tableConfig || {}), clubs: next } }))}
                onEdit={handleEditSection}
                onDelete={(club) => requestDelete("club", club)}
                customFields={clubCustomFields}
              />
            ) : (
              <KanbanBoard
                rows={filteredClubs}
                statuses={CLUB_STATUS}
                onStatusChange={(id, status) => updateEntityStatus("clubs", id, status)}
                renderCard={renderClubKanbanCard}
                emptyLabel="Aucune section trouvée."
              />
            )}
          </div>
        )}

        {/* ── DOCUMENTS ── */}
        {view === "documents" && (
          <div className="p-6 max-w-7xl mx-auto">
            <PageHeader
              title="Documents"
              subtitle={`${filteredDocs.length} document${filteredDocs.length > 1 ? "s" : ""}`}
              onNew={() => { setForm({ sport: "developpe_couche", type: "Suivi sportif", status: "Brouillon", url: "", content: "" }); setDocModal({ open: true, item: null }); }}
              newLabel="Nouveau document"
            />
            <div className="-mt-4 mb-5 flex justify-end">
              <ViewToggle options={tableKanbanOptions} value={documentViewMode} onChange={setDocumentViewMode} />
            </div>
            <div className="flex flex-wrap gap-3 mb-5">
              <div className="relative">
                <Search className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
                <input className="pl-9 pr-4 py-2 rounded-2xl border text-sm bg-white" placeholder="Rechercher..." value={searchQ} onChange={(e) => setSearchQ(e.target.value)} />
              </div>
              <select className="rounded-2xl border px-3 py-2 text-sm bg-white" value={filterSport} onChange={(e) => setFilterSport(e.target.value)}>
                <option value="all">Tous les sports</option>
                {allSports.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
              </select>
              <select className="rounded-2xl border px-3 py-2 text-sm bg-white" value={filterStatus} onChange={(e) => setFilterStatus(e.target.value)}>
                <option value="all">Tous les statuts</option>
                {DOC_STATUS.map((s) => <option key={s} value={s}>{s}</option>)}
              </select>
              {DOC_MODES.map((m) => (
                <button key={m} onClick={() => setDocMode(m)}
                  className={`px-3 py-2 rounded-2xl text-sm border ${docMode === m ? "bg-slate-900 text-white border-slate-900" : "bg-white"}`}>
                  {m === "all" ? "Tous" : m === "internal" ? "Internes" : "Liens externes"}
                </button>
              ))}
            </div>
            {documentViewMode === "table" ? (
              <EntityTable
                rows={filteredDocs}
                nameKey="title"
                emptyLabel="Aucun document trouvé."
                builtinCols={[
                  { key: "title", label: "Document", type: "texte", sortable: true },
                  { key: "status", label: "Statut", type: "statut", sortable: true, options: DOC_STATUS,
                    render: (d) => <Badge className={badgeClass("status", d.status === "Publié" || d.status === "Validé" ? "Terminé" : d.status === "À valider" ? "En attente" : d.status === "Archivé" ? "À faire" : "En cours")}>{d.status}</Badge> },
                  { key: "type", label: "Type", type: "statut", sortable: true,
                    render: (d) => d.type ? <Badge>{d.type}</Badge> : <span className="text-slate-400">—</span> },
                  { key: "sport", label: "Sport", type: "statut", sortable: true,
                    accessor: (d) => sportName(d.sport),
                    render: (d) => <Badge>{sportName(d.sport)}</Badge> },
                  { key: "url", label: "Source", type: "texte", sortable: true,
                    accessor: (d) => d.url ? "Externe" : "Interne",
                    render: (d) => d.url
                      ? <a href={d.url} target="_blank" rel="noopener noreferrer" className="text-xs text-blue-600 underline truncate max-w-xs inline-block">{d.url}</a>
                      : <Badge className="bg-slate-50 text-slate-600 border-slate-200">Interne</Badge> },
                  { key: "notes", label: "Notes", type: "texte", sortable: true },
                ]}
                config={data.tableConfig?.documents}
                onConfigChange={(next) => setData((prev) => ({ ...prev, tableConfig: { ...(prev.tableConfig || {}), documents: next } }))}
                onEdit={(doc) => { setForm({ ...doc }); setDocModal({ open: true, item: doc }); }}
                onDelete={(doc) => requestDelete("document", doc)}
                customFields={documentCustomFields}
              />
            ) : (
              <KanbanBoard
                rows={filteredDocs}
                statuses={DOC_STATUS}
                onStatusChange={(id, status) => updateEntityStatus("documents", id, status)}
                renderCard={renderDocumentKanbanCard}
                emptyLabel="Aucun document trouvé."
              />
            )}
          </div>
        )}

        {/* ── CONVOCATIONS ── */}
        {view === "convocations" && (
          <div className="p-6 max-w-7xl mx-auto">
            <PageHeader
              title="Convocations"
              subtitle={`${filteredConvocations.length} convocation${filteredConvocations.length > 1 ? "s" : ""}`}
              onNew={() => { setForm({ sport: "football", type: "Match", status: "Brouillon", workspaceId: currentWsId }); setConvModal({ open: true, item: null }); }}
              newLabel="Nouvelle convocation"
            />
            <div className="-mt-4 mb-5 flex justify-end">
              <ViewToggle options={listKanbanOptions} value={convocationViewMode} onChange={setConvocationViewMode} />
            </div>
            <FilterBar statusOptions={COMM_STATUS} />
            {convocationViewMode === "list" ? (
              <div className="space-y-3">
                {filteredConvocations.map((conv) => (
                  <Card key={conv.id} className="p-4">
                    <div className="flex items-start justify-between gap-4">
                      <div className="flex-1 min-w-0">
                        <div className="flex flex-wrap items-center gap-2 mb-1">
                          <h3 className="font-semibold">{conv.title}</h3>
                          <Badge className={badgeClass("status", conv.status === "Envoyée" ? "Terminé" : conv.status === "Prêt à envoyer" ? "En cours" : "À faire")}>{conv.status}</Badge>
                          {conv.type && <Badge>{conv.type}</Badge>}
                          <Badge>{sportName(conv.sport)}</Badge>
                        </div>
                        <div className="flex flex-wrap gap-3 text-xs text-slate-500 mt-1">
                          {conv.date && <span>Date : {conv.date}</span>}
                          {conv.time && <span>Heure : {conv.time}</span>}
                          {conv.location && <span>Lieu : {conv.location}</span>}
                          {conv.attachmentTitle && <span>PJ : {conv.attachmentTitle}</span>}
                        </div>
                        {conv.message && <p className="text-sm text-slate-600 mt-2 line-clamp-2">{conv.message}</p>}
                      </div>
                      <ActionBtns
                        onEdit={() => { setForm({ ...conv }); setConvModal({ open: true, item: conv }); }}
                        onDelete={() => requestDelete("convocation", conv)}
                      />
                    </div>
                  </Card>
                ))}
                {!filteredConvocations.length && <p className="text-center text-slate-500 py-12">Aucune convocation.</p>}
              </div>
            ) : (
              <KanbanBoard
                rows={filteredConvocations}
                statuses={COMM_STATUS}
                onStatusChange={(id, status) => updateEntityStatus("convocations", id, status)}
                renderCard={renderConvocationKanbanCard}
                emptyLabel="Aucune convocation."
              />
            )}
          </div>
        )}

        {/* ── NOTIFICATIONS ── */}
        {view === "notifications" && (
          <div className="p-6 max-w-7xl mx-auto">
            <PageHeader
              title="Notifications"
              subtitle={`${filteredNotifications.length} notification${filteredNotifications.length > 1 ? "s" : ""}`}
              onNew={() => { setForm({ sport: "developpe_couche", channel: "Email", status: "Brouillon", workspaceId: currentWsId }); setNotifModal({ open: true, item: null }); }}
              newLabel="Nouvelle notification"
            />
            <div className="-mt-4 mb-5 flex justify-end">
              <ViewToggle options={listKanbanOptions} value={notificationViewMode} onChange={setNotificationViewMode} />
            </div>
            <FilterBar statusOptions={COMM_STATUS} />
            {notificationViewMode === "list" ? (
              <div className="space-y-3">
                {filteredNotifications.map((notif) => (
                  <Card key={notif.id} className="p-4">
                    <div className="flex items-start justify-between gap-4">
                      <div className="flex-1 min-w-0">
                        <div className="flex flex-wrap items-center gap-2 mb-1">
                          <h3 className="font-semibold">{notif.title}</h3>
                          <Badge className={badgeClass("status", notif.status === "Envoyée" ? "Terminé" : notif.status === "Prêt à envoyer" ? "En cours" : "À faire")}>{notif.status}</Badge>
                          {notif.channel && <Badge>{notif.channel}</Badge>}
                          <Badge>{sportName(notif.sport)}</Badge>
                        </div>
                        <div className="flex flex-wrap gap-3 text-xs text-slate-500 mt-1">
                          {notif.audience && <span>Audience : {notif.audience}</span>}
                          {notif.sentAt && <span>Envoyée le : {notif.sentAt}</span>}
                          {notif.attachmentTitle && <span>PJ : {notif.attachmentTitle}</span>}
                        </div>
                        {notif.message && <p className="text-sm text-slate-600 mt-2 line-clamp-2">{notif.message}</p>}
                      </div>
                      <ActionBtns
                        onEdit={() => { setForm({ ...notif }); setNotifModal({ open: true, item: notif }); }}
                        onDelete={() => requestDelete("notification", notif)}
                      />
                    </div>
                  </Card>
                ))}
                {!filteredNotifications.length && <p className="text-center text-slate-500 py-12">Aucune notification.</p>}
              </div>
            ) : (
              <KanbanBoard
                rows={filteredNotifications}
                statuses={COMM_STATUS}
                onStatusChange={(id, status) => updateEntityStatus("notifications", id, status)}
                renderCard={renderNotificationKanbanCard}
                emptyLabel="Aucune notification."
              />
            )}
          </div>
        )}

        {/* ── SETTINGS ── */}
        {view === "settings" && (
          <div className="p-6 max-w-4xl mx-auto">
            <PageHeader title="Paramètres" />

            <Card className="p-6 mb-6 bg-slate-50/60">
              <h2 className="font-semibold text-base mb-2">Colonnes personnalisées</h2>
              <p className="text-sm text-slate-600">
                Les colonnes personnalisées se gèrent maintenant directement depuis chaque vue tableau (Projets, Tâches, Athlètes, Sections, Documents).
                Clique sur l'icône <span className="inline-flex items-center justify-center w-5 h-5 rounded bg-white border align-middle">+</span> en bout de rangée d'en-têtes, ou sur le menu d'une colonne existante (« Ajouter à gauche / à droite »).
              </p>
            </Card>

            <Card className="p-6">
              <h2 className="font-semibold text-lg mb-4">Sports</h2>
              <div className="space-y-2 mb-4">
                {allSports.map((sport) => (
                  <div key={sport.id} className="flex items-center justify-between p-3 border rounded-2xl">
                    <span className="text-sm font-medium">{sport.name}</span>
                    {data.customSports.find((s) => s.id === sport.id) ? (
                      <button onClick={() => upd("customSports", (arr) => arr.filter((s) => s.id !== sport.id))}
                        className="px-2 py-1 rounded-xl border hover:bg-red-50 text-red-600"><Trash2 className="w-3 h-3" /></button>
                    ) : (
                      <Badge className="bg-slate-50 text-slate-500 border-slate-200">Intégré</Badge>
                    )}
                  </div>
                ))}
              </div>
              <div className="flex gap-2">
                <input className="flex-1 rounded-2xl border px-4 py-2 text-sm" placeholder="Nom du nouveau sport..."
                  value={newSportName} onChange={(e) => setNewSportName(e.target.value)}
                  onKeyDown={(e) => e.key === "Enter" && addSport()} />
                <button onClick={addSport} className="px-4 py-2 rounded-2xl bg-slate-900 text-white text-sm">Ajouter</button>
              </div>
            </Card>
          </div>
        )}
      </main>

      {/* ===== MODALS ===== */}

      {/* Project */}
      <Modal open={projectModal.open} title={projectModal.item ? "Modifier le projet" : "Nouveau projet"}
        onClose={() => { setProjectModal({ open: false, item: null }); setForm({}); }} onSave={saveProject}>
        <div className="grid md:grid-cols-2 gap-4">
          <input className="w-full rounded-2xl border px-4 py-3" placeholder="Nom du projet *" value={form.name || ""} onChange={(e) => f("name", e.target.value)} />
          <input className="w-full rounded-2xl border px-4 py-3" placeholder="Responsable" value={form.owner || ""} onChange={(e) => f("owner", e.target.value)} />
          <textarea className="w-full rounded-2xl border px-4 py-3 md:col-span-2 resize-none h-20" placeholder="Description" value={form.description || ""} onChange={(e) => f("description", e.target.value)} />
          <select className="w-full rounded-2xl border px-4 py-3" value={form.sport || ""} onChange={(e) => f("sport", e.target.value)}>
            <option value="">Sport</option>
            {allSports.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
          </select>
          <div className="space-y-2 md:col-span-2 rounded-2xl border bg-slate-50 p-3">
            <p className="text-xs font-bold uppercase tracking-wide text-slate-500">Organisation du projet</p>
            <div className="grid grid-cols-2 gap-2">
              <button
                type="button"
                onClick={() => f("parentId", null)}
                className={`rounded-2xl border px-3 py-3 text-left text-sm font-semibold ${!form.parentId ? "border-slate-900 bg-slate-900 text-white" : "bg-white hover:bg-slate-50"}`}
              >
                <span className="block">Projet principal</span>
                <span className={`mt-1 block text-xs font-normal ${!form.parentId ? "text-slate-200" : "text-slate-500"}`}>Niveau racine</span>
              </button>
              <button
                type="button"
                disabled={!hasAvailableProjectParent}
                onClick={() => f("parentId", form.parentId || projectParentChoices[0]?.id || null)}
                className={`rounded-2xl border px-3 py-3 text-left text-sm font-semibold disabled:cursor-not-allowed disabled:opacity-40 ${form.parentId ? "border-slate-900 bg-slate-900 text-white" : "bg-white hover:bg-slate-50"}`}
              >
                <span className="block">Sous-projet</span>
                <span className={`mt-1 block text-xs font-normal ${form.parentId ? "text-slate-200" : "text-slate-500"}`}>Rattaché à un projet</span>
              </button>
            </div>
            {form.parentId && (
              <select className="w-full rounded-2xl border px-4 py-3" value={form.parentId || ""} onChange={(e) => f("parentId", e.target.value || null)}>
                {projectParentChoices.map((p) => <option key={p.id} value={p.id}>Rattaché à : {p.name}</option>)}
              </select>
            )}
            {!hasAvailableProjectParent && <p className="text-xs text-slate-500">Crée d'abord un projet principal pour pouvoir créer un sous-projet.</p>}
          </div>
          <select className="w-full rounded-2xl border px-4 py-3" value={form.status || "À faire"} onChange={(e) => f("status", e.target.value)}>
            {STATUS_OPTIONS.map((s) => <option key={s} value={s}>{s}</option>)}
          </select>
          <select className="w-full rounded-2xl border px-4 py-3" value={form.priority || "Moyenne"} onChange={(e) => f("priority", e.target.value)}>
            {PRIORITY_OPTIONS.map((p) => <option key={p} value={p}>{p}</option>)}
          </select>
          <input type="date" className="w-full rounded-2xl border px-4 py-3" value={form.startDate || ""} onChange={(e) => f("startDate", e.target.value)} />
          <input type="date" className="w-full rounded-2xl border px-4 py-3" value={form.endDate || ""} onChange={(e) => f("endDate", e.target.value)} />
        </div>
        {projectCustomFields.length > 0 && (
          <div className="mt-4">
            <p className="text-sm font-medium text-slate-700 mb-3">Colonnes ajoutées au tableau Projets</p>
            <CustomFieldsEditor fields={projectCustomFields} values={form.customFieldValues} onChange={fcf} />
          </div>
        )}
      </Modal>

      {/* Task */}
      <Modal open={taskModal.open} title={taskModal.item ? "Modifier la tâche" : "Nouvelle tâche"}
        onClose={() => { setTaskModal({ open: false, item: null }); setForm({}); }} onSave={saveTask}>
        <div className="grid md:grid-cols-2 gap-4">
          <input className="w-full rounded-2xl border px-4 py-3 md:col-span-2" placeholder="Titre *" value={form.title || ""} onChange={(e) => f("title", e.target.value)} />
          <textarea className="w-full rounded-2xl border px-4 py-3 md:col-span-2 resize-none h-16" placeholder="Description" value={form.description || ""} onChange={(e) => f("description", e.target.value)} />
          <input className="w-full rounded-2xl border px-4 py-3" placeholder="Assigné à" value={form.assignee || ""} onChange={(e) => f("assignee", e.target.value)} />
          <select className="w-full rounded-2xl border px-4 py-3" value={form.projectId || ""} onChange={(e) => f("projectId", e.target.value)}>
            <option value="">Projet</option>
            {data.projects.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
          </select>
          <select className="w-full rounded-2xl border px-4 py-3" value={form.status || "À faire"} onChange={(e) => f("status", e.target.value)}>
            {STATUS_OPTIONS.map((s) => <option key={s} value={s}>{s}</option>)}
          </select>
          <select className="w-full rounded-2xl border px-4 py-3" value={form.priority || "Moyenne"} onChange={(e) => f("priority", e.target.value)}>
            {PRIORITY_OPTIONS.map((p) => <option key={p} value={p}>{p}</option>)}
          </select>
          <input type="date" className="w-full rounded-2xl border px-4 py-3" value={form.startDate || ""} onChange={(e) => f("startDate", e.target.value)} />
          <input type="date" className="w-full rounded-2xl border px-4 py-3" value={form.endDate || ""} onChange={(e) => f("endDate", e.target.value)} />
          <div className="md:col-span-2">
            <label className="block text-sm text-slate-700 mb-2">Progression : <span className="font-semibold">{form.progress || 0}%</span></label>
            <input type="range" min="0" max="100" className="w-full accent-slate-900" value={form.progress || 0} onChange={(e) => f("progress", Number(e.target.value))} />
          </div>
        </div>
        {taskCustomFields.length > 0 && (
          <div className="mt-4">
            <p className="text-sm font-medium text-slate-700 mb-3">Colonnes ajoutées au tableau Tâches</p>
            <CustomFieldsEditor fields={taskCustomFields} values={form.customFieldValues} onChange={fcf} />
          </div>
        )}
      </Modal>

      {/* Athlete */}
      <Modal open={athleteModal.open} title={athleteModal.item ? "Modifier l'athlète" : "Nouvel athlète"}
        onClose={() => { setAthleteModal({ open: false, item: null }); setForm({}); }} onSave={saveAthlete}>
        <div className="grid md:grid-cols-2 gap-4">
          <input className="w-full rounded-2xl border px-4 py-3 md:col-span-2" placeholder="Nom complet *" value={form.name || ""} onChange={(e) => f("name", e.target.value)} />
          <select
            className="w-full rounded-2xl border px-4 py-3"
            value={form.sectionId || ""}
            onChange={(e) => {
              const section = wsClubs.find((item) => item.id === e.target.value);
              setForm((prev) => ({ ...prev, sectionId: e.target.value, subSection: "", sport: section?.sport || prev.sport }));
            }}
          >
            <option value="">Section JAM</option>
            {wsClubs.map((section) => <option key={section.id} value={section.id}>{section.name}</option>)}
          </select>
          {subSectionsFor(form.sectionId).length > 0 ? (
            <select className="w-full rounded-2xl border px-4 py-3" value={form.subSection || ""} onChange={(e) => f("subSection", e.target.value)}>
              <option value="">Sous-section</option>
              {subSectionsFor(form.sectionId).map((sub) => <option key={sub} value={sub}>{sub}</option>)}
            </select>
          ) : (
            <input className="w-full rounded-2xl border px-4 py-3" placeholder="Sous-section (ex : U13, -88 kg…)" value={form.subSection || ""} onChange={(e) => f("subSection", e.target.value)} />
          )}
          <select className="w-full rounded-2xl border px-4 py-3" value={form.category || ""} onChange={(e) => f("category", e.target.value)}>
            <option value="">Profil</option>
            {ATHLETE_PROFILE_OPTIONS.map((profile) => <option key={profile} value={profile}>{profile}</option>)}
          </select>
          <input className="w-full rounded-2xl border px-4 py-3" placeholder="Coach" value={form.coach || ""} onChange={(e) => f("coach", e.target.value)} />
          <select className="w-full rounded-2xl border px-4 py-3" value={form.status || "Actif"} onChange={(e) => f("status", e.target.value)}>
            {ATHLETE_STATUS.map((s) => <option key={s} value={s}>{s}</option>)}
          </select>
          <input type="email" className="w-full rounded-2xl border px-4 py-3" placeholder="Email" value={form.email || ""} onChange={(e) => f("email", e.target.value)} />
          <input className="w-full rounded-2xl border px-4 py-3" placeholder="Téléphone" value={form.phone || ""} onChange={(e) => f("phone", e.target.value)} />
          <textarea className="w-full rounded-2xl border px-4 py-3 md:col-span-2 resize-none h-16" placeholder="Notes" value={form.notes || ""} onChange={(e) => f("notes", e.target.value)} />
        </div>
        {athleteCustomFields.length > 0 && (
          <div className="mt-4">
            <p className="text-sm font-medium text-slate-700 mb-3">Colonnes ajoutées au tableau Athlètes</p>
            <CustomFieldsEditor fields={athleteCustomFields} values={form.customFieldValues} onChange={fcf} />
          </div>
        )}
      </Modal>

      {/* Section */}
      <Modal open={clubModal.open} title={clubModal.item ? "Modifier la section" : "Nouvelle section"}
        onClose={() => { setClubModal({ open: false, item: null }); setForm({}); }} onSave={saveClub}>
        <div className="grid md:grid-cols-2 gap-4">
          <input className="w-full rounded-2xl border px-4 py-3 md:col-span-2" placeholder="Nom de la section *" value={form.name || ""} onChange={(e) => f("name", e.target.value)} />
          <input className="w-full rounded-2xl border px-4 py-3" placeholder="Lieu / ville" value={form.city || ""} onChange={(e) => f("city", e.target.value)} />
          <input className="w-full rounded-2xl border px-4 py-3" placeholder="Responsable de section" value={form.contactName || ""} onChange={(e) => f("contactName", e.target.value)} />
          <select className="w-full rounded-2xl border px-4 py-3" value={form.status || "À contacter"} onChange={(e) => f("status", e.target.value)}>
            {CLUB_STATUS.map((s) => <option key={s} value={s}>{s}</option>)}
          </select>
          <div className="rounded-2xl border px-4 py-3 bg-slate-50">
            <p className="text-xs font-semibold text-slate-500 uppercase">Nombre d'athlètes</p>
            <p className="text-lg font-bold text-slate-900">{sectionAthleteCount(form.id || clubModal.item?.id)} athlète{sectionAthleteCount(form.id || clubModal.item?.id) > 1 ? "s" : ""}</p>
          </div>
          <input type="email" className="w-full rounded-2xl border px-4 py-3" placeholder="Email" value={form.email || ""} onChange={(e) => f("email", e.target.value)} />
          <input className="w-full rounded-2xl border px-4 py-3" placeholder="Téléphone" value={form.phone || ""} onChange={(e) => f("phone", e.target.value)} />
          <textarea
            className="w-full rounded-2xl border px-4 py-3 md:col-span-2 resize-none h-24"
            placeholder={"Sous-sections, une par ligne\nEx : U7\nU9\nU11"}
            value={form.subSectionsText ?? (form.subSections || []).join("\n")}
            onChange={(e) => f("subSectionsText", e.target.value)}
          />
          <textarea className="w-full rounded-2xl border px-4 py-3 md:col-span-2 resize-none h-16" placeholder="Notes" value={form.notes || ""} onChange={(e) => f("notes", e.target.value)} />
        </div>
        {clubCustomFields.length > 0 && (
          <div className="mt-4">
            <p className="text-sm font-medium text-slate-700 mb-3">Colonnes ajoutées au tableau Sections</p>
            <CustomFieldsEditor fields={clubCustomFields} values={form.customFieldValues} onChange={fcf} />
          </div>
        )}
      </Modal>

      {/* Document */}
      <Modal open={docModal.open} title={docModal.item ? "Modifier le document" : "Nouveau document"}
        onClose={() => { setDocModal({ open: false, item: null }); setForm({}); }} onSave={saveDoc}>
        <div className="grid md:grid-cols-2 gap-4">
          <input className="w-full rounded-2xl border px-4 py-3 md:col-span-2" placeholder="Titre *" value={form.title || ""} onChange={(e) => f("title", e.target.value)} />
          <select className="w-full rounded-2xl border px-4 py-3" value={form.sport || ""} onChange={(e) => f("sport", e.target.value)}>
            <option value="">Sport</option>
            {allSports.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
          </select>
          <input className="w-full rounded-2xl border px-4 py-3" placeholder="Type (Match, Suivi sportif…)" value={form.type || ""} onChange={(e) => f("type", e.target.value)} />
          <select className="w-full rounded-2xl border px-4 py-3" value={form.status || "Brouillon"} onChange={(e) => f("status", e.target.value)}>
            {DOC_STATUS.map((s) => <option key={s} value={s}>{s}</option>)}
          </select>
          <input className="w-full rounded-2xl border px-4 py-3 md:col-span-2" placeholder="URL (laisser vide = document interne)" value={form.url || ""} onChange={(e) => f("url", e.target.value)} />
          <textarea className="w-full rounded-2xl border px-4 py-3 md:col-span-2 resize-none h-20" placeholder="Contenu (si document interne)" value={form.content || ""} onChange={(e) => f("content", e.target.value)} />
          <textarea className="w-full rounded-2xl border px-4 py-3 md:col-span-2 resize-none h-16" placeholder="Notes" value={form.notes || ""} onChange={(e) => f("notes", e.target.value)} />
        </div>
        {documentCustomFields.length > 0 && (
          <div className="mt-4">
            <p className="text-sm font-medium text-slate-700 mb-3">Colonnes ajoutées au tableau Documents</p>
            <CustomFieldsEditor fields={documentCustomFields} values={form.customFieldValues} onChange={fcf} />
          </div>
        )}
      </Modal>

      {/* Convocation */}
      <Modal open={convModal.open} title={convModal.item ? "Modifier la convocation" : "Nouvelle convocation"}
        onClose={() => { setConvModal({ open: false, item: null }); setForm({}); }} onSave={saveConv}>
        <div className="grid md:grid-cols-2 gap-4">
          <input className="w-full rounded-2xl border px-4 py-3 md:col-span-2" placeholder="Titre *" value={form.title || ""} onChange={(e) => f("title", e.target.value)} />
          <select className="w-full rounded-2xl border px-4 py-3" value={form.sport || ""} onChange={(e) => f("sport", e.target.value)}>
            <option value="">Sport</option>
            {allSports.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
          </select>
          <input className="w-full rounded-2xl border px-4 py-3" placeholder="Type (Match, Entraînement…)" value={form.type || ""} onChange={(e) => f("type", e.target.value)} />
          <input type="date" className="w-full rounded-2xl border px-4 py-3" value={form.date || ""} onChange={(e) => f("date", e.target.value)} />
          <input type="time" className="w-full rounded-2xl border px-4 py-3" value={form.time || ""} onChange={(e) => f("time", e.target.value)} />
          <input className="w-full rounded-2xl border px-4 py-3 md:col-span-2" placeholder="Lieu" value={form.location || ""} onChange={(e) => f("location", e.target.value)} />
          <select className="w-full rounded-2xl border px-4 py-3" value={form.status || "Brouillon"} onChange={(e) => f("status", e.target.value)}>
            {COMM_STATUS.map((s) => <option key={s} value={s}>{s}</option>)}
          </select>
          <input className="w-full rounded-2xl border px-4 py-3" placeholder="Pièce jointe (titre)" value={form.attachmentTitle || ""} onChange={(e) => f("attachmentTitle", e.target.value)} />
          <textarea className="w-full rounded-2xl border px-4 py-3 md:col-span-2 resize-none h-20" placeholder="Message" value={form.message || ""} onChange={(e) => f("message", e.target.value)} />
        </div>
      </Modal>

      {/* Notification */}
      <Modal open={notifModal.open} title={notifModal.item ? "Modifier la notification" : "Nouvelle notification"}
        onClose={() => { setNotifModal({ open: false, item: null }); setForm({}); }} onSave={saveNotif}>
        <div className="grid md:grid-cols-2 gap-4">
          <input className="w-full rounded-2xl border px-4 py-3 md:col-span-2" placeholder="Titre *" value={form.title || ""} onChange={(e) => f("title", e.target.value)} />
          <select className="w-full rounded-2xl border px-4 py-3" value={form.sport || ""} onChange={(e) => f("sport", e.target.value)}>
            <option value="">Sport</option>
            {allSports.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
          </select>
          <select className="w-full rounded-2xl border px-4 py-3" value={form.channel || "Email"} onChange={(e) => f("channel", e.target.value)}>
            <option value="Email">Email</option>
            <option value="SMS">SMS</option>
            <option value="Push">Push</option>
            <option value="WhatsApp">WhatsApp</option>
          </select>
          <input className="w-full rounded-2xl border px-4 py-3" placeholder="Audience cible" value={form.audience || ""} onChange={(e) => f("audience", e.target.value)} />
          <select className="w-full rounded-2xl border px-4 py-3" value={form.status || "Brouillon"} onChange={(e) => f("status", e.target.value)}>
            {COMM_STATUS.map((s) => <option key={s} value={s}>{s}</option>)}
          </select>
          <input className="w-full rounded-2xl border px-4 py-3" placeholder="Pièce jointe (titre)" value={form.attachmentTitle || ""} onChange={(e) => f("attachmentTitle", e.target.value)} />
          <input type="date" className="w-full rounded-2xl border px-4 py-3" placeholder="Date d'envoi" value={form.sentAt || ""} onChange={(e) => f("sentAt", e.target.value)} />
          <textarea className="w-full rounded-2xl border px-4 py-3 md:col-span-2 resize-none h-20" placeholder="Message" value={form.message || ""} onChange={(e) => f("message", e.target.value)} />
        </div>
      </Modal>

      {/* Custom Field */}
      <Modal open={cfModal.open} title={cfModal.item ? "Modifier le champ" : "Nouveau champ personnalisé"}
        onClose={() => { setCfModal({ open: false, item: null }); setForm({}); }} onSave={saveCF}>
        <div className="grid md:grid-cols-2 gap-4">
          <input className="w-full rounded-2xl border px-4 py-3 md:col-span-2" placeholder="Nom du champ *" value={form.name || ""} onChange={(e) => f("name", e.target.value)} />
          <select className="w-full rounded-2xl border px-4 py-3" value={form.type || "texte"} onChange={(e) => f("type", e.target.value)}>
            {CUSTOM_FIELD_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}
          </select>
          <div className="flex flex-col gap-2 justify-center">
            <p className="text-sm text-slate-700 font-medium">Appliquer à :</p>
            {["projects", "tasks"].map((target) => (
              <label key={target} className="flex items-center gap-2 text-sm cursor-pointer">
                <input type="checkbox"
                  checked={(form.appliesTo || []).includes(target)}
                  onChange={(e) => {
                    const cur = form.appliesTo || [];
                    f("appliesTo", e.target.checked ? [...cur, target] : cur.filter((x) => x !== target));
                  }} />
                {target === "projects" ? "Projets" : "Tâches"}
              </label>
            ))}
          </div>
        </div>
      </Modal>

      {/* Dashboard widget */}
      <Modal
        open={dashboardWidgetModal.open}
        title={dashboardWidgetModal.item ? "Modifier le widget" : "Nouveau widget"}
        onClose={() => { setDashboardWidgetModal({ open: false, item: null }); setDashboardWidgetForm({}); }}
        onSave={saveDashboardWidget}
        saveLabel="Enregistrer le widget"
      >
        <div className="grid gap-4 md:grid-cols-2">
          <label className="block md:col-span-2">
            <span className="mb-1.5 block text-xs font-bold uppercase tracking-wide text-slate-500">Titre</span>
            <input
              className="w-full rounded-2xl border px-4 py-3"
              value={dashboardWidgetForm.title || ""}
              onChange={(e) => setDashboardWidgetForm((prev) => ({ ...prev, title: e.target.value }))}
              placeholder="ex : Tâches terminées"
            />
          </label>
          <label className="block">
            <span className="mb-1.5 block text-xs font-bold uppercase tracking-wide text-slate-500">Type de widget</span>
            <select
              className="w-full rounded-2xl border px-4 py-3"
              value={dashboardWidgetForm.type || "number"}
              onChange={(e) => setDashboardWidgetForm((prev) => ({ ...prev, type: e.target.value }))}
            >
              {DASHBOARD_WIDGET_TYPES.map((item) => <option key={item.id} value={item.id}>{item.label}</option>)}
            </select>
          </label>
          <label className="block">
            <span className="mb-1.5 block text-xs font-bold uppercase tracking-wide text-slate-500">Donnée</span>
            <select
              className="w-full rounded-2xl border px-4 py-3"
              value={dashboardWidgetForm.source || "tasks"}
              onChange={(e) => setDashboardWidgetForm((prev) => ({ ...prev, source: e.target.value }))}
            >
              {DASHBOARD_SOURCES.map((item) => <option key={item.id} value={item.id}>{item.label}</option>)}
            </select>
          </label>
          <label className="block">
            <span className="mb-1.5 block text-xs font-bold uppercase tracking-wide text-slate-500">Filtre</span>
            <select
              className="w-full rounded-2xl border px-4 py-3"
              value={dashboardWidgetForm.filter || "all"}
              onChange={(e) => setDashboardWidgetForm((prev) => ({ ...prev, filter: e.target.value }))}
            >
              {DASHBOARD_FILTERS.map((item) => <option key={item.id} value={item.id}>{item.label}</option>)}
            </select>
          </label>
          {(dashboardWidgetForm.type || "number") !== "number" && (
            <label className="block">
              <span className="mb-1.5 block text-xs font-bold uppercase tracking-wide text-slate-500">Grouper par</span>
              <select
                className="w-full rounded-2xl border px-4 py-3"
                value={dashboardWidgetForm.groupBy || "status"}
                onChange={(e) => setDashboardWidgetForm((prev) => ({ ...prev, groupBy: e.target.value }))}
              >
                {DASHBOARD_GROUPS.map((item) => <option key={item.id} value={item.id}>{item.label}</option>)}
              </select>
            </label>
          )}
          <label className="block">
            <span className="mb-1.5 block text-xs font-bold uppercase tracking-wide text-slate-500">Couleur</span>
            <select
              className="w-full rounded-2xl border px-4 py-3"
              value={dashboardWidgetForm.tone || "slate"}
              onChange={(e) => setDashboardWidgetForm((prev) => ({ ...prev, tone: e.target.value }))}
            >
              <option value="slate">JAM sobre</option>
              <option value="blue">Bleu</option>
              <option value="green">Vert</option>
              <option value="amber">Orange</option>
              <option value="red">Rouge</option>
              <option value="violet">Violet</option>
            </select>
          </label>
          <label className="block">
            <span className="mb-1.5 block text-xs font-bold uppercase tracking-wide text-slate-500">Format</span>
            <select
              className="w-full rounded-2xl border px-4 py-3"
              value={dashboardWidgetForm.size || "small"}
              onChange={(e) => setDashboardWidgetForm((prev) => ({ ...prev, size: e.target.value }))}
            >
              <option value="small">Carte simple</option>
              <option value="wide">Carte large</option>
            </select>
          </label>
        </div>
        <div className="mt-4 rounded-2xl border border-slate-200 bg-slate-50 p-4 text-sm text-slate-500">
          Les widgets se mettent à jour automatiquement avec les données actuelles. Pour un KPI budget, choisis “Projets” + “Budget”.
        </div>
      </Modal>

      {/* Guided input panel */}
      <GuidedInputPanel
        open={guidedOpen}
        onClose={() => setGuidedOpen(false)}
        allSports={allSports}
        projects={wsProjects}
        sections={wsClubs}
        onSave={handleGuidedSave}
      />

      <AddColumnModal
        open={!!globalColumnModal}
        onClose={() => setGlobalColumnModal(null)}
        onCreate={addGlobalColumn}
        position="right"
      />

      <ImportPreviewModal
        preview={importPreview}
        onClose={() => setImportPreview(null)}
        onConfirm={confirmImportPreview}
        onTabChange={(activeTab) => setImportPreview((prev) => prev ? { ...prev, activeTab } : prev)}
      />

      {/* Workspace creation modal */}
      {wsModal && (
        <div className="fixed inset-0 z-50 bg-black/40 flex items-center justify-center p-4">
          <div className="bg-white rounded-3xl shadow-2xl w-full max-w-sm p-6">
            <div className="flex items-center justify-between mb-4">
              <h3 className="text-lg font-semibold">Nouvel espace de travail</h3>
              <button onClick={() => { setWsModal(false); setWsForm({}); }} className="text-slate-400 hover:text-slate-700"><X className="w-5 h-5" /></button>
            </div>
            <div className="space-y-4">
              <div>
                <label className="block text-sm font-bold text-slate-800 mb-2">Nom *</label>
                <input autoFocus className="w-full rounded-2xl border px-4 py-3 text-sm focus:outline-none focus:ring-2 focus:ring-slate-900"
                  placeholder="ex : Mon équipe" value={wsForm.name || ""} onChange={(e) => setWsForm((p) => ({ ...p, name: e.target.value }))} />
              </div>
              <div>
                <label className="block text-sm font-bold text-slate-800 mb-2">Couleur</label>
                <div className="flex gap-2 flex-wrap">
                  {["#0f172a","#7c3aed","#059669","#dc2626","#d97706","#2563eb","#db2777"].map((c) => (
                    <button key={c} onClick={() => setWsForm((p) => ({ ...p, color: c }))}
                      className={`w-8 h-8 rounded-lg border-2 ${wsForm.color === c ? "border-slate-900 scale-110" : "border-transparent"} transition-all`}
                      style={{ backgroundColor: c }} />
                  ))}
                </div>
              </div>
            </div>
            <div className="flex justify-end gap-2 mt-6">
              <button onClick={() => { setWsModal(false); setWsForm({}); }} className="px-4 py-2 rounded-2xl border text-sm">Annuler</button>
              <button onClick={saveWorkspace} className="px-4 py-2 rounded-2xl bg-slate-900 text-white text-sm font-semibold">Créer</button>
            </div>
          </div>
        </div>
      )}

      {/* Delete confirm */}
      {deleteConfirm.open && (
        <div className="fixed inset-0 z-50 bg-black/40 flex items-center justify-center p-4">
          <div className="bg-white rounded-3xl shadow-2xl w-full max-w-sm p-6">
            <div className="flex items-center gap-3 mb-4">
              <div className="p-2 rounded-2xl bg-red-100"><AlertTriangle className="w-5 h-5 text-red-600" /></div>
              <h3 className="text-lg font-semibold">Confirmer la suppression</h3>
            </div>
            <p className="text-slate-600 mb-6">
              Supprimer <span className="font-medium">"{deleteConfirm.item?.name || deleteConfirm.item?.title}"</span> ? Cette action est irréversible.
            </p>
            <div className="flex justify-end gap-2">
              <button onClick={() => setDeleteConfirm({ open: false, type: null, item: null })} className="px-4 py-2 rounded-2xl border">Annuler</button>
              <button onClick={confirmDelete} className="px-4 py-2 rounded-2xl bg-red-600 text-white">Supprimer</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
