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
  Menu,
  X,
  EyeOff,
  Layers,
  Paperclip,
  FileImage,
  FileAudio2,
  Link2,
  ExternalLink,
  MessageCircle,
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
const CUSTOM_FIELD_TARGETS = [
  { id: "projects", label: "Projets" },
  { id: "tasks", label: "Tâches" },
  { id: "athletes", label: "Athlètes" },
  { id: "clubs", label: "Sections" },
];
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
  { id: "missingEmail", label: "Sans email" },
  { id: "missingPhone", label: "Sans téléphone" },
  { id: "missingSection", label: "Sans section" },
  { id: "missingProfile", label: "Sans profil" },
  { id: "budget", label: "Budget" },
  { id: "withBudget", label: "Avec budget" },
  { id: "withoutBudget", label: "Sans budget" },
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
const ATTACHMENT_FILE_LIMIT = 4 * 1024 * 1024;

const uid = () => {
  if (globalThis.crypto?.randomUUID) return globalThis.crypto.randomUUID();
  if (globalThis.crypto?.getRandomValues) {
    const bytes = globalThis.crypto.getRandomValues(new Uint8Array(16));
    bytes[6] = (bytes[6] & 0x0f) | 0x40;
    bytes[8] = (bytes[8] & 0x3f) | 0x80;
    const hex = Array.from(bytes, (byte) => byte.toString(16).padStart(2, "0")).join("");
    return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20)}`;
  }
  return `jam-${Date.now()}-${Math.random().toString(16).slice(2, 10)}`;
};

const normalizeAttachments = (attachments) => Array.isArray(attachments) ? attachments.filter(Boolean) : [];

const attachmentKindFromFile = (file) => {
  const type = file?.type || "";
  const name = (file?.name || "").toLowerCase();
  if (type.startsWith("image/")) return "image";
  if (type.startsWith("audio/")) return "audio";
  if (type === "application/pdf" || name.endsWith(".pdf")) return "pdf";
  if (
    type.includes("spreadsheet") ||
    type.includes("excel") ||
    name.endsWith(".xls") ||
    name.endsWith(".xlsx") ||
    name.endsWith(".csv")
  ) return "spreadsheet";
  return "file";
};

const attachmentKindLabel = (attachment) => {
  if (attachment.kind === "whatsapp") return "WhatsApp";
  if (attachment.kind === "image") return "Image";
  if (attachment.kind === "audio") return "Mémo vocal";
  if (attachment.kind === "pdf") return "PDF";
  if (attachment.kind === "spreadsheet") return "Tableur";
  return "Fichier";
};

const attachmentIcon = (attachment) => {
  if (attachment.kind === "whatsapp") return MessageCircle;
  if (attachment.kind === "image") return FileImage;
  if (attachment.kind === "audio") return FileAudio2;
  if (attachment.kind === "spreadsheet") return FileSpreadsheet;
  return FileText;
};

const formatFileSize = (size) => {
  if (!size) return "";
  if (size < 1024) return `${size} o`;
  if (size < 1024 * 1024) return `${(size / 1024).toFixed(1)} Ko`;
  return `${(size / (1024 * 1024)).toFixed(1)} Mo`;
};

const readFileAsDataUrl = (file) => new Promise((resolve, reject) => {
  const reader = new FileReader();
  reader.onload = () => resolve(reader.result);
  reader.onerror = () => reject(reader.error || new Error("Lecture impossible"));
  reader.readAsDataURL(file);
});

const openAttachmentFile = (attachment) => {
  if (!attachment?.url) return;
  if (!attachment.url.startsWith("data:")) {
    window.open(attachment.url, "_blank", "noopener,noreferrer");
    return;
  }
  const popup = window.open("", "_blank", "noopener,noreferrer");
  if (popup) {
    popup.location.href = attachment.url;
    return;
  }
  const link = document.createElement("a");
  link.href = attachment.url;
  link.target = "_blank";
  link.rel = "noopener noreferrer";
  link.download = attachment.name || "piece-jointe";
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
};

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
    visibility: "Oui",
    allowedUserIds: [],
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
    visibility: "Oui",
    allowedUserIds: [],
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
    visibility: "Oui",
    allowedUserIds: [],
    workspaceId: "ws_jam",
  },
];

const ROOT_ID = seedProjects[0].id;
const LOG_ID = seedProjects[1].id;
const PORTAL_ID = seedProjects[2].id;
const USER_ROLE_OPTIONS = ["Administrateur", "Manager", "Utilisateur"];
const defaultUsers = [
  { id: "user_samy", name: "Samy", email: "samy@jam.local", role: "Administrateur", workspaceIds: ["ws_jam", "ws_propcept"] },
  { id: "user_sonia", name: "Sonia", email: "sonia@jam.local", role: "Manager", workspaceIds: ["ws_jam"] },
  { id: "user_thomas", name: "Thomas", email: "thomas@jam.local", role: "Utilisateur", workspaceIds: ["ws_jam"] },
  { id: "user_jeremy", name: "Jérémy", email: "jeremy@jam.local", role: "Utilisateur", workspaceIds: ["ws_jam"] },
];

function normalizeProjectAccess(project, fallbackUserId = "user_samy") {
  const allowedUserIds = Array.isArray(project?.allowedUserIds) ? project.allowedUserIds.filter(Boolean) : [];
  if ((project?.visibility || "Oui") !== "Non") return allowedUserIds;
  return Array.from(new Set(allowedUserIds.length ? allowedUserIds : [fallbackUserId]));
}

const defaultData = {
  workspaces: [
    { id: "ws_jam", name: "JAM", color: "#0f172a", initial: "J" },
    { id: "ws_propcept", name: "Propcept Event", color: "#7c3aed", initial: "P" },
  ],
  currentWorkspaceId: "ws_jam",
  users: defaultUsers,
  currentUserId: "user_samy",
  customFields: [
    { id: uid(), name: "Budget prévisionnel", type: "financier", appliesTo: ["projects"], position: "right" },
    { id: uid(), name: "Date de clôture", type: "date", appliesTo: ["projects", "tasks"], position: "right" },
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
      visibility: "Oui",
      allowedUserIds: [],
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
      visibility: "Oui",
      allowedUserIds: [],
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
      visibility: "Oui",
      allowedUserIds: [],
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
      visibility: "Oui",
      allowedUserIds: [],
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
    const users = (parsed.users || defaultUsers).map((user) => {
      const workspaceIds = Array.isArray(user.workspaceIds) && user.workspaceIds.length ? user.workspaceIds : ["ws_jam"];
      const role = USER_ROLE_OPTIONS.includes(user.role) ? user.role : "Utilisateur";
      return { ...user, workspaceIds, role };
    });
    const currentUserId = parsed.currentUserId || users[0]?.id || "user_samy";
    return {
      ...defaultData,
      ...parsed,
      users,
      currentUserId,
      customFields: (parsed.customFields || defaultData.customFields).map((field) =>
        field.name === "Date comité" ? { ...field, name: "Date de clôture" } : field
      ),
      workspaces: parsed.workspaces || defaultData.workspaces,
      currentWorkspaceId: parsed.currentWorkspaceId || "ws_jam",
      projects: addWs(parsed.projects || defaultData.projects).map((project) => ({
        ...project,
        visibility: project.visibility || "Oui",
        allowedUserIds: normalizeProjectAccess(project, currentUserId),
      })),
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

function priorityCellClass(value) {
  if (value === "Critique") return "bg-red-500 text-white";
  if (value === "Haute") return "bg-orange-500 text-white";
  if (value === "Moyenne") return "bg-violet-500 text-white";
  return "bg-slate-200 text-slate-700";
}

function nextOptionValue(options, currentValue) {
  const currentIndex = Math.max(0, options.indexOf(currentValue));
  return options[(currentIndex + 1) % options.length];
}

function statusVisualValue(value) {
  if (value === "Actif" || value === "Inscrit" || value === "Brouillon") return "En cours";
  if (value === "Blessé" || value === "Invité" || value === "À valider") return "En attente";
  if (value === "Publié" || value === "Validé" || value === "Envoyée") return "Terminé";
  if (value === "Archivé" || value === "Absent" || value === "À relancer" || value === "À contacter") return "À faire";
  return value;
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

function statusProgressValue(status) {
  if (status === "Terminé") return 100;
  if (status === "En cours") return 50;
  if (status === "En attente") return 5;
  return 0;
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

function ownProgress(projectId, tasks, projects = []) {
  const rows = tasks.filter((t) => t.projectId === projectId);
  if (rows.length) {
    return Math.round(rows.reduce((sum, row) => sum + normalizeTaskProgress(row.progress, row.status), 0) / rows.length);
  }

  const directChildren = projects.filter((project) => project.parentId === projectId);
  if (directChildren.length) {
    return Math.round(
      directChildren.reduce((sum, child) => sum + consolidatedProgress(child.id, tasks, projects), 0) / directChildren.length
    );
  }

  const project = projects.find((item) => item.id === projectId);
  return statusProgressValue(project?.status);
}

function consolidatedProgress(projectId, tasks, projects) {
  const ids = descendants(projectId, projects);
  const rows = tasks.filter((t) => ids.has(t.projectId));
  if (!rows.length) {
    const scopedProjects = projects.filter((project) => ids.has(project.id));
    if (!scopedProjects.length) return 0;
    return Math.round(scopedProjects.reduce((sum, project) => sum + statusProgressValue(project.status), 0) / scopedProjects.length);
  }
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

function defaultFinancialValues(fields = []) {
  return fields.reduce((acc, field) => {
    if (field.type === "financier") acc[field.id] = 0;
    return acc;
  }, {});
}

function Card({ children, className = "", ...props }) {
  return <div className={`rounded-[28px] border border-slate-200/80 bg-white shadow-[0_12px_30px_-18px_rgba(15,23,42,0.22)] ${className}`} {...props}>{children}</div>;
}

function Badge({ children, className = "" }) {
  return <span className={`inline-flex items-center rounded-full border px-2.5 py-1 text-xs font-semibold ${className}`}>{children}</span>;
}

function ProgressBar({ value, label }) {
  return (
    <div className="space-y-1.5">
      <div className="flex justify-between text-[11px] font-semibold uppercase tracking-wide text-slate-400">
        <span>{label}</span>
        <span className="text-slate-600">{value}%</span>
      </div>
      <div className="h-2.5 w-full overflow-hidden rounded-full bg-slate-100 ring-1 ring-inset ring-slate-200">
        <div className={`h-full rounded-full ${progressTone(value)}`} style={{ width: `${value}%` }} />
      </div>
    </div>
  );
}

function ViewToggle({ options, value, onChange }) {
  return (
    <div className="flex w-full flex-wrap gap-1 rounded-2xl border border-slate-200 bg-white p-1.5 shadow-sm sm:inline-flex sm:w-auto">
      {options.map(({ id, label, Icon }) => (
        <button
          key={id}
          type="button"
          onClick={() => onChange(id)}
          className={`flex flex-1 items-center justify-center gap-1.5 rounded-xl px-3 py-2 text-sm font-semibold transition-colors sm:flex-none ${value === id ? "bg-slate-900 text-white shadow-sm" : "text-slate-500 hover:bg-slate-100 hover:text-slate-900"}`}
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
      <div className="flex items-start justify-between gap-4">
        <div className="min-w-0">
          <p className="text-sm font-semibold text-slate-500 sm:text-base">{title}</p>
          <p className="mt-2 truncate text-3xl font-bold tracking-tight text-slate-950 sm:mt-3 sm:text-4xl">{value}</p>
          <p className="mt-2 text-sm text-slate-500">{subtitle}</p>
        </div>
        <div className={`rounded-2xl p-3 ${tones[tone] || tones.slate} sm:p-3.5`}><Icon className="h-5 w-5" /></div>
      </div>
      {onClick && <p className="mt-4 text-[10px] font-bold uppercase tracking-[0.14em] text-slate-400 sm:text-[11px] sm:tracking-[0.16em]">Ouvrir le tableau</p>}
    </>
  );
  if (onClick) {
    return (
      <button
        type="button"
        onClick={onClick}
        className="group w-full rounded-[28px] border border-slate-200/80 bg-white p-4 text-left shadow-[0_12px_30px_-18px_rgba(15,23,42,0.22)] transition hover:-translate-y-0.5 hover:border-slate-300 hover:shadow-md focus:outline-none focus:ring-2 focus:ring-slate-900/20 sm:p-5"
      >
        {content}
      </button>
    );
  }
  return (
    <Card className="p-5">
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
              placeholder={field.type === "date" ? "" : field.type === "personne" ? (field.placeholder || `Ex : ${field.name}`) : field.name}
              value={value}
              onChange={(e) => onChange(field.id, e.target.value)}
            />
          </label>
        );
      })}
    </div>
  );
}

function ProjectTreeNode({ node, tasks, projects, expanded, setExpanded, onQuickTask, onNewSubProject, onEditProject, onEditTask, onArchiveProject, requestDelete, onDuplicate, projectCustomFields, onProjectStatusChange, onProjectPriorityChange, onTaskStatusChange, onTaskPriorityChange, onTaskProgressChange, onOpenAttachments }) {
  const [menuOpen, setMenuOpen] = useState(false);
  const menuRef = useRef(null);

  useEffect(() => {
    if (!menuOpen) return;
    const handler = (e) => { if (menuRef.current && !menuRef.current.contains(e.target)) setMenuOpen(false); };
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, [menuOpen]);

  const own = node.computedOwnProgress ?? ownProgress(node.id, tasks, projects);
  const total = node.computedTotalProgress ?? consolidatedProgress(node.id, tasks, projects);
  const isOpen = expanded[node.id] ?? true;
  const childCount = node.children.length;
  const projectScope = descendants(node.id, projects);
  const directTasks = tasks.filter((t) => t.projectId === node.id);
  const inheritedTasks = tasks.filter((t) => projectScope.has(t.projectId) && t.projectId !== node.id);
  const scopeTasks = [...directTasks, ...inheritedTasks];

  const renderTaskItem = (task, showProjectLabel = false) => {
    const prog = normalizeTaskProgress(task.progress, task.status);
    const sourceProject = projects.find((project) => project.id === task.projectId);
    return (
      <div key={task.id} className="rounded-xl border bg-white p-3">
        <div className="flex items-start justify-between gap-3">
          <div className="flex items-center gap-2 min-w-0">
            <CheckCircle2 className={`w-4 h-4 shrink-0 ${task.status === "Terminé" ? "text-green-500" : "text-slate-300"}`} />
            <div className="min-w-0">
              <button type="button" onClick={() => onEditTask?.(task)} className="truncate text-left text-sm font-medium text-slate-900 hover:underline">{task.title}</button>
              {showProjectLabel && sourceProject && (
                <p className="mt-0.5 text-[11px] font-semibold uppercase tracking-wide text-slate-400">{sourceProject.name}</p>
              )}
            </div>
          </div>
          <div className="flex items-center gap-1.5 shrink-0">
            <button
              type="button"
              onClick={() => onTaskStatusChange?.(task.id, nextOptionValue(STATUS_OPTIONS, task.status))}
              className={`rounded-full border px-2 py-1 text-xs font-semibold transition hover:brightness-95 ${badgeClass("status", task.status)}`}
              title="Cliquer pour changer le statut"
            >
              {task.status}
            </button>
            <button
              type="button"
              onClick={() => onTaskPriorityChange?.(task.id, nextOptionValue(PRIORITY_OPTIONS, task.priority))}
              className={`rounded-full border px-2 py-1 text-xs font-semibold transition hover:brightness-95 ${badgeClass("priority", task.priority)}`}
              title="Cliquer pour changer la priorité"
            >
              {task.priority}
            </button>
          </div>
        </div>
        <div className="mt-2 flex items-center gap-3 pl-6">
          <span className="text-xs text-slate-500">{task.assignee || "Non assigné"}</span>
          {task.endDate && <span className="text-xs text-slate-400">· {task.endDate}</span>}
        </div>
        <div className="mt-2 pl-6">
          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={() => onTaskProgressChange?.(task.id, prog - 10)}
              className="inline-flex h-7 w-7 items-center justify-center rounded-full border border-slate-200 text-sm font-bold text-slate-500 hover:bg-slate-100"
              title="Diminuer la progression"
            >
              -
            </button>
            <div className="flex-1 h-1.5 rounded-full bg-slate-200 overflow-hidden">
              <div className={`h-full rounded-full ${progressTone(prog)}`} style={{ width: `${prog}%` }} />
            </div>
            <button
              type="button"
              onClick={() => onTaskProgressChange?.(task.id, prog + 10)}
              className="inline-flex h-7 w-7 items-center justify-center rounded-full border border-slate-200 text-sm font-bold text-slate-500 hover:bg-slate-100"
              title="Augmenter la progression"
            >
              +
            </button>
            <span className="text-xs font-bold text-slate-600 w-7 text-right">{prog}%</span>
          </div>
        </div>
      </div>
    );
  };

  return (
    <div className="border rounded-2xl p-4 bg-white shadow-sm">
      <div className="flex items-start justify-between gap-4">
        <div className="flex gap-3 flex-1 min-w-0">
          <button onClick={() => setExpanded((prev) => ({ ...prev, [node.id]: !isOpen }))} className="mt-1 text-slate-500 shrink-0">
            {childCount ? (isOpen ? <ChevronDown className="w-4 h-4" /> : <ChevronRight className="w-4 h-4" />) : <div className="w-4" />}
          </button>
          <div className="flex-1 min-w-0">
            <div className="flex gap-3">
              {node.coverImage ? (
                <button
                  type="button"
                  onClick={() => onEditProject?.(node)}
                  className="shrink-0 overflow-hidden rounded-2xl border border-slate-200 bg-white shadow-sm"
                  title="Modifier le projet"
                >
                  <img src={node.coverImage} alt="" className="h-20 w-28 object-cover" />
                </button>
              ) : null}
              <div className="min-w-0 flex-1">
                <div className="flex flex-wrap items-center gap-2">
                  <h3 className="font-semibold truncate">{node.name}</h3>
                  <button
                    type="button"
                    onClick={() => onProjectStatusChange?.(node.id, nextOptionValue(STATUS_OPTIONS, node.status))}
                    className={`rounded-full border px-2 py-1 text-xs font-semibold transition hover:brightness-95 ${badgeClass("status", node.status)}`}
                    title="Cliquer pour changer le statut"
                  >
                    {node.status}
                  </button>
                  <button
                    type="button"
                    onClick={() => onProjectPriorityChange?.(node.id, nextOptionValue(PRIORITY_OPTIONS, node.priority))}
                    className={`rounded-full border px-2 py-1 text-xs font-semibold transition hover:brightness-95 ${badgeClass("priority", node.priority)}`}
                    title="Cliquer pour changer la priorité"
                  >
                    {node.priority}
                  </button>
                  {node.sportName && <Badge>{node.sportName}</Badge>}
                  {normalizeAttachments(node.attachments).length > 0 && (
                    <button
                      type="button"
                      onClick={() => onOpenAttachments?.(`Pièces jointes · ${node.name}`, node.attachments)}
                      className="inline-flex items-center gap-1 rounded-full border border-slate-200 bg-white px-2 py-1 text-[11px] font-semibold text-slate-600 hover:bg-slate-100"
                      title="Voir les pièces jointes"
                    >
                      <Paperclip className="h-3 w-3" /> {normalizeAttachments(node.attachments).length}
                    </button>
                  )}
                  <Badge className={node.visibility === "Non" ? "bg-slate-100 text-slate-600 border-slate-200" : "bg-emerald-50 text-emerald-700 border-emerald-200"}>
                    Visible par tous : {node.visibility || "Oui"}
                  </Badge>
                </div>
                {node.description && <p className="text-sm text-slate-500 mt-1">{node.description}</p>}
              </div>
            </div>
            <div className="grid md:grid-cols-2 gap-4 mt-3 max-w-3xl">
              <ProgressBar value={own} label="Progression directe" />
              <ProgressBar value={total} label="Progression consolidée" />
            </div>
            <div className="flex flex-wrap gap-3 text-xs text-slate-500 mt-2">
              {node.owner && <span>Responsable : {node.owner}</span>}
              {node.startDate && <span>Début : {node.startDate}</span>}
              {node.endDate && <span>Fin : {node.endDate}</span>}
              <span>Axes analytiques : {childCount}</span>
              <span>Tâches du périmètre : {scopeTasks.length}</span>
            </div>

            {projectCustomFields.length > 0 && (
              <div className="flex flex-wrap gap-2 mt-2">
                {projectCustomFields.map((field) => (
                  <Badge key={field.id}>{field.name} : {formatCustomFieldValue(field.type, node.customFieldValues?.[field.id])}</Badge>
                ))}
              </div>
            )}

            {scopeTasks.length > 0 && (
              <div className="mt-4 rounded-2xl bg-slate-50 border border-slate-200 p-3">
                {directTasks.length > 0 && (
                  <>
                    <p className="mb-2 text-xs font-semibold uppercase tracking-wide text-slate-500">Tâches directes ({directTasks.length})</p>
                    <div className="space-y-2">
                      {directTasks.map((task) => renderTaskItem(task))}
                    </div>
                  </>
                )}
                {inheritedTasks.length > 0 && (
                  <>
                    <p className={`text-xs font-semibold uppercase tracking-wide text-slate-500 ${directTasks.length > 0 ? "mt-4 mb-2" : "mb-2"}`}>Tâches des axes analytiques ({inheritedTasks.length})</p>
                    <div className="space-y-2">
                      {inheritedTasks.map((task) => renderTaskItem(task, true))}
                    </div>
                  </>
                )}
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
                <FolderKanban className="w-4 h-4 text-slate-400" /> Créer un axe analytique
              </button>
              <button onClick={() => { onDuplicate(node); setMenuOpen(false); }} className="w-full text-left px-3 py-2 text-sm hover:bg-slate-50 flex items-center gap-2">
                <Copy className="w-4 h-4 text-slate-400" /> Dupliquer
              </button>
              <button onClick={() => { onArchiveProject?.(node); setMenuOpen(false); }} className="w-full text-left px-3 py-2 text-sm hover:bg-amber-50 text-amber-700 flex items-center gap-2">
                <Library className="w-4 h-4" /> Archiver
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
                <ProjectTree nodes={node.children} tasks={tasks} projects={projects} expanded={expanded} setExpanded={setExpanded} onQuickTask={onQuickTask} onNewSubProject={onNewSubProject} onEditProject={onEditProject} onEditTask={onEditTask} onArchiveProject={onArchiveProject} requestDelete={requestDelete} onDuplicate={onDuplicate} projectCustomFields={projectCustomFields} onProjectStatusChange={onProjectStatusChange} onProjectPriorityChange={onProjectPriorityChange} onTaskStatusChange={onTaskStatusChange} onTaskPriorityChange={onTaskPriorityChange} onTaskProgressChange={onTaskProgressChange} onOpenAttachments={onOpenAttachments} />
        </div>
      )}
    </div>
  );
}

function ProjectTree({ nodes, tasks, projects, expanded, setExpanded, onQuickTask, onNewSubProject, onEditProject, onEditTask, onArchiveProject, requestDelete, onDuplicate, projectCustomFields, onProjectStatusChange, onProjectPriorityChange, onTaskStatusChange, onTaskPriorityChange, onTaskProgressChange, onOpenAttachments }) {
  return (
    <div className="space-y-3">
      {nodes.map((node) => (
        <ProjectTreeNode key={node.id} node={node} tasks={tasks} projects={projects} expanded={expanded} setExpanded={setExpanded} onQuickTask={onQuickTask} onNewSubProject={onNewSubProject} onEditProject={onEditProject} onEditTask={onEditTask} onArchiveProject={onArchiveProject} requestDelete={requestDelete} onDuplicate={onDuplicate} projectCustomFields={projectCustomFields} onProjectStatusChange={onProjectStatusChange} onProjectPriorityChange={onProjectPriorityChange} onTaskStatusChange={onTaskStatusChange} onTaskPriorityChange={onTaskPriorityChange} onTaskProgressChange={onTaskProgressChange} onOpenAttachments={onOpenAttachments} />
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

const ADD_COLUMN_HINTS = {
  texte: {
    placeholder: "Ex : Commentaire interne",
    helper: "Choisis librement le titre de la colonne.",
  },
  nombre: {
    placeholder: "Ex : Volume prévisionnel",
    helper: "Utile pour des quantités, compteurs ou charges.",
  },
  pourcentage: {
    placeholder: "Ex : Taux d'avancement",
    helper: "Parfait pour un suivi en pourcentage de 0 à 100.",
  },
  date: {
    placeholder: "Ex : Date de validation",
    helper: "Pratique pour les échéances, relances et dates clés.",
  },
  financier: {
    placeholder: "Ex : Budget complémentaire",
    helper: "Le montant sera affiché automatiquement en euros.",
  },
  statut: {
    placeholder: "Ex : Validation comité",
    helper: "Ajoute une colonne de statut complémentaire au suivi principal.",
  },
  case_a_cocher: {
    placeholder: "Ex : Validé par le bureau",
    helper: "Tu peux personnaliser le texte affiché si la case est cochée ou non.",
  },
  personne: {
    placeholder: "Ex : Responsable logistique",
    helper: "Idéal pour un référent, un contact ou un propriétaire secondaire.",
  },
};

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
  const [checkedLabel, setCheckedLabel] = useState("Oui");
  const [uncheckedLabel, setUncheckedLabel] = useState("Non");
  const [personPlaceholder, setPersonPlaceholder] = useState("");
  const inputRef = useRef(null);
  useEffect(() => {
    if (!open) return;
    setName("");
    setType("texte");
    setCheckedLabel("Oui");
    setUncheckedLabel("Non");
    setPersonPlaceholder("");
  }, [open]);
  if (!open) return null;
  const activeHint = ADD_COLUMN_HINTS[type] || ADD_COLUMN_HINTS.texte;
  const handleTypeSelect = (nextType) => {
    setType(nextType);
    window.requestAnimationFrame(() => inputRef.current?.focus());
  };
  const handleCreate = () => {
    const payload = { name: name.trim(), type };
    if (type === "case_a_cocher") {
      payload.checkedLabel = checkedLabel.trim() || "Oui";
      payload.uncheckedLabel = uncheckedLabel.trim() || "Non";
    }
    if (type === "personne" && personPlaceholder.trim()) {
      payload.placeholder = personPlaceholder.trim();
    }
    onCreate(payload);
    setName("");
    setType("texte");
    setCheckedLabel("Oui");
    setUncheckedLabel("Non");
    setPersonPlaceholder("");
  };
  return (
    <div className="fixed inset-0 z-[60] bg-black/40 flex items-center justify-center" onClick={onClose}>
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-md" onClick={(e) => e.stopPropagation()}>
        <div className="px-5 py-4 border-b flex items-center justify-between">
          <h3 className="font-bold">Ajout de colonne {position === "left" ? "à gauche" : position === "right" ? "à droite" : ""}</h3>
          <button type="button" onClick={onClose}><X className="w-5 h-5" /></button>
        </div>
        <div className="p-5 space-y-4">
          <div>
            <label className="block text-xs font-bold text-slate-700 mb-1.5 uppercase tracking-wide">Nom</label>
            <input ref={inputRef} autoFocus type="text" value={name} onChange={(e) => setName(e.target.value)} placeholder={activeHint.placeholder}
              className="w-full rounded-xl border px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-slate-900" />
            <p className="mt-1.5 text-xs text-slate-500">{activeHint.helper}</p>
          </div>
          <div>
            <label className="block text-xs font-bold text-slate-700 mb-1.5 uppercase tracking-wide">Type</label>
            <div className="grid grid-cols-2 gap-2">
              {COL_TYPES.map((t) => (
                <button key={t.id} type="button" onClick={() => handleTypeSelect(t.id)}
                  className={`flex items-center gap-2 px-3 py-2 rounded-xl border text-sm text-left ${type === t.id ? "bg-slate-900 text-white border-slate-900" : "hover:bg-slate-50"}`}>
                  <ColumnTypeMark type={t.id} />
                  <span>{t.label}</span>
                </button>
              ))}
            </div>
          </div>
          {type === "case_a_cocher" && (
            <div className="grid grid-cols-2 gap-3">
              <label className="block">
                <span className="mb-1.5 block text-xs font-bold uppercase tracking-wide text-slate-500">Texte si coché</span>
                <input
                  type="text"
                  value={checkedLabel}
                  onChange={(e) => setCheckedLabel(e.target.value)}
                  placeholder="Ex : Oui"
                  className="w-full rounded-xl border px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-slate-900"
                />
              </label>
              <label className="block">
                <span className="mb-1.5 block text-xs font-bold uppercase tracking-wide text-slate-500">Texte si non coché</span>
                <input
                  type="text"
                  value={uncheckedLabel}
                  onChange={(e) => setUncheckedLabel(e.target.value)}
                  placeholder="Ex : Non"
                  className="w-full rounded-xl border px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-slate-900"
                />
              </label>
            </div>
          )}
          {type === "personne" && (
            <label className="block">
              <span className="mb-1.5 block text-xs font-bold uppercase tracking-wide text-slate-500">Texte d'aide dans la cellule</span>
              <input
                type="text"
                value={personPlaceholder}
                onChange={(e) => setPersonPlaceholder(e.target.value)}
                placeholder="Ex : Choisir un référent"
                className="w-full rounded-xl border px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-slate-900"
              />
            </label>
          )}
        </div>
        <div className="px-5 py-4 border-t flex justify-end gap-2">
          <button type="button" onClick={onClose} className="px-4 py-2 rounded-xl text-sm font-semibold hover:bg-slate-100">Annuler</button>
          <button type="button" disabled={!name.trim()} onClick={handleCreate}
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
function EntityTable({ rows, builtinCols, config, onConfigChange, onEdit, onDelete, onArchive, onUnarchive, nameKey = "name", emptyLabel = "Aucun élément.", customFields = [], onStatusChange, onPriorityChange, onProgressChange, onFieldChange, onOpenAttachments, allowRootColumnEditing = true }) {
  const [openColMenu, setOpenColMenu] = useState(null);
  const [submenu, setSubmenu] = useState(null);
  const [addColModal, setAddColModal] = useState(null);

  const cfg = config || { customCols: [], hiddenCols: [], renamed: {}, typeOverrides: {}, order: [], groupBy: null, filters: {}, sortCol: nameKey, sortDir: "asc" };
  const setCfg = (patch) => onConfigChange({ ...cfg, ...(typeof patch === "function" ? patch(cfg) : patch) });
  const sortCol = cfg.sortCol || nameKey;
  const sortDir = cfg.sortDir || "asc";
  const closeMenu = () => { setOpenColMenu(null); setSubmenu(null); };
  const cycleStatus = (row) => {
    if (!onStatusChange) return;
    const statusCol = allCols.find((col) => col.key === "status");
    const options = statusCol?.options || STATUS_OPTIONS;
    onStatusChange(row.id, nextOptionValue(options, row.status));
  };
  const cyclePriority = (row) => {
    if (!onPriorityChange) return;
    const priorityCol = allCols.find((col) => col.key === "priority");
    const options = priorityCol?.options || PRIORITY_OPTIONS;
    onPriorityChange(row.id, nextOptionValue(options, row.priority));
  };

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
  const addColumn = ({ name, type, ...meta }) => {
    const newKey = "cf_" + uid().slice(0, 8);
    const anchor = addColModal?.anchorKey;
    setCfg((c) => {
      const existing = [...(c.customCols || []), { key: newKey, label: name, type, ...meta }];
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
    if (col.key === "priority" && onPriorityChange) {
      const v = valueFor(row, col);
      return (
        <button
          type="button"
          onClick={() => cyclePriority(row)}
          className={`rounded-full border px-2 py-1 text-xs font-semibold transition hover:brightness-95 ${badgeClass("priority", v)}`}
          title="Cliquer pour changer la priorité"
        >
          {v || "—"}
        </button>
      );
    }
    const v = valueFor(row, col);
    const rawValue = col.builtin ? row[col.key] : row.customFieldValues?.[col.key];
    if (col.key === "progress" && onProgressChange) {
      const progressValue = Number(v) || 0;
      return (
        <div className="flex min-w-52 items-center gap-2">
          <button
            type="button"
            onClick={() => onProgressChange(row.id, progressValue - 10)}
            className="inline-flex h-7 w-7 items-center justify-center rounded-full border border-slate-200 text-sm font-bold text-slate-500 hover:bg-slate-100"
            title="Diminuer la progression"
          >
            -
          </button>
          <div className="flex-1 h-2 overflow-hidden rounded-full bg-slate-200">
            <div className={`h-full rounded-full ${progressTone(progressValue)}`} style={{ width: `${progressValue}%` }} />
          </div>
          <button
            type="button"
            onClick={() => onProgressChange(row.id, progressValue + 10)}
            className="inline-flex h-7 w-7 items-center justify-center rounded-full border border-slate-200 text-sm font-bold text-slate-500 hover:bg-slate-100"
            title="Augmenter la progression"
          >
            +
          </button>
          <input
            type="number"
            min="0"
            max="100"
            value={progressValue}
            onChange={(event) => onProgressChange(row.id, Number(event.target.value))}
            className="w-16 rounded-xl border border-slate-200 bg-white px-2 py-1 text-right text-xs font-bold text-slate-700 focus:outline-none focus:ring-2 focus:ring-slate-900/10"
            title="Saisir une progression"
          />
          <span className="text-xs font-bold text-slate-500">%</span>
        </div>
      );
    }
    if (col.key === "attachmentsCount") {
      const count = normalizeAttachments(row.attachments).length;
      return count > 0 ? (
        <button
          type="button"
          onClick={() => onOpenAttachments?.(`Pièces jointes · ${row[nameKey] || row.title || row.name || "élément"}`, row.attachments)}
          className="inline-flex items-center gap-2 rounded-full border border-slate-200 bg-slate-50 px-2.5 py-1 text-xs font-semibold text-slate-700 transition hover:bg-slate-100"
          title="Voir les pièces jointes"
        >
          <Paperclip className="h-3.5 w-3.5 text-slate-500" />
          {count}
        </button>
      ) : <span className="text-slate-300">—</span>;
    }
    if (col.quickEdit && onFieldChange) {
      if (col.quickEdit === "text") {
        return (
          <input
            list={col.type === "personne" ? "people-options" : undefined}
            type="text"
            value={rawValue || ""}
            onChange={(event) => onFieldChange(row.id, col.key, event.target.value)}
            className="w-full rounded-xl border border-slate-200 bg-white px-3 py-2 text-sm text-slate-700 focus:outline-none focus:ring-2 focus:ring-slate-900/10"
            placeholder={col.placeholder || "—"}
          />
        );
      }
      if (col.quickEdit === "date") {
        return (
          <input
            type="date"
            value={rawValue || ""}
            onChange={(event) => onFieldChange(row.id, col.key, event.target.value)}
            className="w-full rounded-xl border border-slate-200 bg-white px-3 py-2 text-sm text-slate-700 focus:outline-none focus:ring-2 focus:ring-slate-900/10"
          />
        );
      }
      if (col.quickEdit === "select") {
        const options = col.editOptions?.(row) || [];
        return (
          <select
            value={rawValue || ""}
            onChange={(event) => onFieldChange(row.id, col.key, event.target.value)}
            className="w-full rounded-xl border border-slate-200 bg-white px-3 py-2 text-sm text-slate-700 focus:outline-none focus:ring-2 focus:ring-slate-900/10"
          >
            <option value="">{col.placeholder || "—"}</option>
            {options.map((option) => <option key={option.value} value={option.value}>{option.label}</option>)}
          </select>
        );
      }
    }
    if (col.key === "status" && v && col.options) {
      const toneValue = col.badgeTone ? col.badgeTone(row) : statusVisualValue(v);
      return onStatusChange ? (
        <button
          type="button"
          onClick={() => cycleStatus(row)}
          className={`rounded-full border px-2 py-1 text-xs font-semibold transition hover:brightness-95 ${badgeClass("status", toneValue)}`}
          title="Cliquer pour changer le statut"
        >
          {v}
        </button>
      ) : <Badge className={badgeClass("status", toneValue)}>{v}</Badge>;
    }
    if (col.render) return col.render(row);
    if (col.type === "case_a_cocher") return <span>{v === true ? (col.checkedLabel || "Oui") : v === false ? (col.uncheckedLabel || "Non") : "—"}</span>;
    if (col.type === "financier" || col.type === "nombre" || col.type === "pourcentage") return <span className="text-slate-700">{formatCustomFieldValue(col.type, v)}</span>;
    if (col.key === "priority" && v && col.options) {
      return onPriorityChange ? (
        <button
          type="button"
          onClick={() => cyclePriority(row)}
          className={`rounded-full border px-2 py-1 text-xs font-semibold transition hover:brightness-95 ${badgeClass("priority", v)}`}
          title="Cliquer pour changer la priorité"
        >
          {v}
        </button>
      ) : <Badge className={badgeClass("priority", v)}>{v}</Badge>;
    }
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
              <button onClick={() => {
                if (!allowRootColumnEditing) return;
                setAddColModal({ position: "left", anchorKey: col.key });
                closeMenu();
              }} className={`w-full text-left px-3 py-2 text-sm flex items-center gap-2 ${allowRootColumnEditing ? "hover:bg-slate-50" : "text-slate-300 cursor-not-allowed"}`}>
                <Plus className="w-4 h-4 text-slate-400" /> Ajouter une colonne à gauche
              </button>
              <button onClick={() => {
                if (!allowRootColumnEditing) return;
                setAddColModal({ position: "right", anchorKey: col.key });
                closeMenu();
              }} className={`w-full text-left px-3 py-2 text-sm flex items-center gap-2 ${allowRootColumnEditing ? "hover:bg-slate-50" : "text-slate-300 cursor-not-allowed"}`}>
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
      {allowRootColumnEditing && (
        <div className="mb-3 flex flex-col gap-2 sm:flex-row sm:justify-end">
          <button
            type="button"
            onClick={() => setAddColModal({ position: "right", anchorKey: visibleCols[visibleCols.length - 1]?.key })}
            className="inline-flex w-full items-center justify-center gap-2 rounded-2xl border bg-white px-4 py-3 text-sm font-semibold text-slate-700 shadow-sm hover:bg-slate-50 sm:w-auto sm:py-2"
          >
            <Plus className="w-4 h-4" /> Ajouter une colonne
          </button>
        </div>
      )}

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
                {allowRootColumnEditing ? (
                  <button onClick={(e) => { e.stopPropagation(); setAddColModal({ position: "right", anchorKey: visibleCols[visibleCols.length - 1]?.key }); }}
                    className="p-1.5 rounded-lg hover:bg-slate-200 text-slate-500" title="Ajouter une colonne">
                    <Plus className="w-4 h-4" />
                  </button>
                ) : null}
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
                          onEdit ? (
                            <button
                              type="button"
                              onClick={() => onEdit(row)}
                              className="text-left font-semibold text-slate-900 hover:underline"
                              title="Ouvrir la fiche"
                            >
                              {row[nameKey] || row.title || "—"}
                            </button>
                          ) : (
                            <p className="font-semibold text-slate-900">{row[nameKey] || row.title || "—"}</p>
                          )
                        ) : renderCell(row, col)}
                      </td>
                    ))}
                    <td />
                    <td className="py-3 px-3" onClick={(e) => e.stopPropagation()}>
                      <div className="flex gap-1.5 justify-end">
                        {onEdit && <button onClick={() => onEdit(row)} className="p-1.5 rounded-lg hover:bg-slate-100 text-slate-600" title="Modifier"><Edit3 className="w-3.5 h-3.5" /></button>}
                        {onArchive && <button onClick={() => onArchive(row)} className="p-1.5 rounded-lg hover:bg-amber-100 text-amber-700" title="Archiver"><Library className="w-3.5 h-3.5" /></button>}
                        {onUnarchive && <button onClick={() => onUnarchive(row)} className="p-1.5 rounded-lg hover:bg-emerald-100 text-emerald-700" title="Désarchiver"><Library className="w-3.5 h-3.5" /></button>}
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

function ProjectTable({ projects, tasks, allProjects = projects, onEdit, onEditTask, onArchiveProject, requestDelete, onQuickTask, onNewSubProject, config, onConfigChange, customFields = [], onStatusChange, onPriorityChange, onTaskStatusChange, onTaskPriorityChange, onTaskProgressChange, onProjectFieldChange, onOpenAttachments }) {
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
  const cycleProjectStatus = (project) => {
    if (!onStatusChange) return;
    onStatusChange(project.id, nextOptionValue(STATUS_OPTIONS, project.status));
  };
  const cycleProjectPriority = (project) => {
    if (!onPriorityChange) return;
    onPriorityChange(project.id, nextOptionValue(PRIORITY_OPTIONS, project.priority));
  };
  const cycleTaskStatus = (task) => {
    if (!onTaskStatusChange) return;
    onTaskStatusChange(task.id, nextOptionValue(STATUS_OPTIONS, task.status));
  };
  const cycleTaskPriority = (task) => {
    if (!onTaskPriorityChange) return;
    onTaskPriorityChange(task.id, nextOptionValue(PRIORITY_OPTIONS, task.priority));
  };

  const handleSort = (key, dir) => { setCfg({ sortCol: key, sortDir: dir }); closeMenu(); };

  // ===== Column model =====
  const builtinCols = [
    { key: "name", label: "Projet", type: "texte", builtin: true, sortable: true },
    { key: "sportName", label: "Sport", type: "texte", builtin: true, sortable: true },
    { key: "status", label: "Statut", type: "statut", builtin: true, sortable: true, options: STATUS_OPTIONS },
    { key: "priority", label: "Priorité", type: "statut", builtin: true, sortable: true, options: PRIORITY_OPTIONS },
    { key: "attachmentsCount", label: "PJ", type: "nombre", builtin: true, sortable: true, accessor: (project) => normalizeAttachments(project.attachments).length },
    { key: "computedTotalProgress", label: "Avancement", type: "nombre", builtin: true, sortable: true },
    { key: "tasksCount", label: "Tâches", type: "nombre", builtin: true, sortable: false },
    { key: "owner", label: "Responsable", type: "personne", builtin: true, sortable: true },
    { key: "visibility", label: "Visible", type: "statut", builtin: true, sortable: true, options: ["Oui", "Non"] },
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
  const scopedTaskCountFor = (projectId) => tasks.filter((task) => descendants(projectId, allProjects).has(task.projectId)).length;
  const directTaskCountFor = (projectId) => tasks.filter((task) => task.projectId === projectId).length;

  // ===== Filtering =====
  const valueFor = (project, col) => {
    if (col.builtin) {
      if (col.key === "tasksCount") return scopedTaskCountFor(project.id);
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
  const visibleTaskCount = filtered.reduce((sum, project) => sum + scopedTaskCountFor(project.id), 0);
  const selectedCount = selectedIds.size;

  // ===== Cell renderer =====
  const renderCell = (project, col) => {
    const v = valueFor(project, col);
    if (col.builtin && col.key === "name") return null; // handled in special name cell
    if (col.builtin && col.key === "status") return (
      <button
        type="button"
        onClick={() => cycleProjectStatus(project)}
        className={`block w-full rounded-xl px-3 py-2 text-center text-xs font-bold transition hover:brightness-95 ${statusCellClass(project.status)}`}
        title="Cliquer pour changer le statut"
      >
        {project.status}
      </button>
    );
    if (col.builtin && col.key === "priority") return (
      <button
        type="button"
        onClick={() => cycleProjectPriority(project)}
        className={`inline-flex items-center gap-2 rounded-xl px-3 py-2 text-sm font-semibold transition hover:brightness-95 ${priorityCellClass(project.priority)}`}
        title="Cliquer pour changer la priorité"
      >
        <span className={`h-2.5 w-2.5 rounded-full ${priorityDotClass(project.priority)}`} />
        <span>{project.priority || "—"}</span>
      </button>
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
      const directCount = directTaskCountFor(project.id);
      const scopedCount = scopedTaskCountFor(project.id);
      return (
        <div className="space-y-1">
          <span className="block text-slate-600"><span className="font-medium text-slate-800">{scopedCount}</span> tâche{scopedCount > 1 ? "s" : ""}</span>
          <span className="block text-[11px] text-slate-400">{directCount} directe{directCount > 1 ? "s" : ""}</span>
        </div>
      );
    }
    if (col.builtin && col.key === "sportName") return <span className="text-slate-700">{project.sportName || "—"}</span>;
    if (col.builtin && col.key === "visibility") {
      return (
        <button
          type="button"
          onClick={() => onProjectFieldChange?.(project.id, "visibility", project.visibility === "Non" ? "Oui" : "Non")}
          className={`rounded-full border px-2 py-1 text-xs font-semibold transition hover:brightness-95 ${project.visibility === "Non" ? "border-slate-200 bg-slate-100 text-slate-600" : "border-emerald-200 bg-emerald-50 text-emerald-700"}`}
          title="Cliquer pour changer la visibilité"
        >
          {project.visibility || "Oui"}
        </button>
      );
    }
    if (col.builtin && col.key === "owner") {
      return (
        <input
          list="people-options"
          type="text"
          value={project.owner || ""}
          onChange={(event) => onProjectFieldChange?.(project.id, "owner", event.target.value)}
          className="w-full rounded-xl border border-slate-200 bg-white px-3 py-2 text-sm text-slate-700 focus:outline-none focus:ring-2 focus:ring-slate-900/10"
          placeholder="Responsable"
        />
      );
    }
    if (col.builtin && col.key === "endDate") {
      return (
        <input
          type="date"
          value={project.endDate || ""}
          onChange={(event) => onProjectFieldChange?.(project.id, "endDate", event.target.value)}
          className="w-full rounded-xl border border-slate-200 bg-white px-3 py-2 text-sm text-slate-700 focus:outline-none focus:ring-2 focus:ring-slate-900/10"
        />
      );
    }
    if (col.builtin && col.key === "attachmentsCount") {
      const count = normalizeAttachments(project.attachments).length;
      return count > 0 ? (
        <button
          type="button"
          onClick={() => onOpenAttachments?.(`Pièces jointes · ${project.name}`, project.attachments)}
          className="inline-flex items-center gap-2 rounded-full border border-slate-200 bg-slate-50 px-2.5 py-1 text-xs font-semibold text-slate-700 transition hover:bg-slate-100"
          title="Voir les pièces jointes"
        >
          <Paperclip className="h-3.5 w-3.5 text-slate-500" />
          {count}
        </button>
      ) : <span className="text-slate-300">—</span>;
    }
    if (col.type === "case_a_cocher") return <span>{v === true ? (col.checkedLabel || "Oui") : v === false ? (col.uncheckedLabel || "Non") : "—"}</span>;
    if (col.type === "financier" || col.type === "nombre" || col.type === "pourcentage") return <span className="text-slate-700">{formatCustomFieldValue(col.type, v)}</span>;
    if (col.key === "status" && v) return (
      <button
        type="button"
        onClick={() => cycleProjectStatus(project)}
        className={`block w-full rounded-xl px-3 py-2 text-center text-xs font-bold transition hover:brightness-95 ${statusCellClass(v)}`}
        title="Cliquer pour changer le statut"
      >
        {v}
      </button>
    );
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
  const addColumn = ({ name, type, ...meta }) => {
    const newKey = "cf_" + uid().slice(0, 8);
    const newCol = { key: newKey, label: name, type, ...meta };
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
    const directTaskCount = projectTasks.length;
    const scopedTaskCount = scopedTaskCountFor(project.id);
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
                {project.coverImage && <img src={project.coverImage} alt="" className="h-10 w-10 rounded-xl object-cover border border-slate-200" />}
                {depth > 0 && <span className="text-slate-300 text-xs">↳</span>}
                <p className="font-semibold text-slate-900">{project.name}</p>
                {normalizeAttachments(project.attachments).length > 0 && (
                  <button
                    type="button"
                    onClick={(event) => { event.stopPropagation(); openAttachmentViewer(`Pièces jointes · ${project.name}`, project.attachments); }}
                    className="inline-flex h-7 w-7 items-center justify-center rounded-full border border-slate-200 text-slate-500 hover:bg-slate-100"
                    title="Voir les pièces jointes"
                  >
                    <Paperclip className="h-4 w-4" />
                  </button>
                )}
                <Badge className={depth === 0 ? "bg-blue-50 text-blue-700 border-blue-200" : "bg-violet-50 text-violet-700 border-violet-200"}>
                  {depth === 0 ? "Projet" : "Axe analytique"}
                </Badge>
                <Badge className={project.visibility === "Non" ? "bg-slate-100 text-slate-600 border-slate-200" : "bg-emerald-50 text-emerald-700 border-emerald-200"}>
                  {project.visibility || "Oui"}
                </Badge>
              </div>
              <div className="mt-1 flex flex-wrap items-center gap-2 text-xs text-slate-400">
                <span>{project.sportName || "Sport non renseigné"}</span>
                {childProjects.length > 0 && <span>{childProjects.length} axe{childProjects.length > 1 ? "s" : ""} analytique{childProjects.length > 1 ? "s" : ""}</span>}
                <span>{directTaskCount} tâche{directTaskCount > 1 ? "s" : ""} directe{directTaskCount > 1 ? "s" : ""}</span>
                <span>{scopedTaskCount} sur le périmètre</span>
                {normalizeAttachments(project.attachments).length > 0 && <span>{normalizeAttachments(project.attachments).length} pièce{normalizeAttachments(project.attachments).length > 1 ? "s jointes" : " jointe"}</span>}
              </div>
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
            <button onClick={() => onNewSubProject(project)} className="p-1.5 rounded-lg hover:bg-violet-100 text-violet-700" title="Créer un axe analytique"><FolderKanban className="w-3.5 h-3.5" /></button>
            {onArchiveProject && <button onClick={() => onArchiveProject(project)} className="p-1.5 rounded-lg hover:bg-amber-100 text-amber-700" title="Archiver"><Library className="w-3.5 h-3.5" /></button>}
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
                    <button type="button" onClick={() => onEditTask?.(task)} className="text-left text-slate-800 font-medium hover:underline">{task.title}</button>
                    {normalizeAttachments(task.attachments).length > 0 && (
                      <button
                        type="button"
                        onClick={() => openAttachmentViewer(`Pièces jointes · ${task.title}`, task.attachments)}
                        className="inline-flex items-center gap-1 rounded-full border border-slate-200 bg-white px-2 py-1 text-[11px] font-semibold text-slate-600 hover:bg-slate-100"
                        title="Voir les pièces jointes"
                      >
                        <Paperclip className="h-3 w-3" /> {normalizeAttachments(task.attachments).length}
                      </button>
                    )}
                  </div>
                  <p className="text-xs text-slate-400 mt-0.5 pl-5">{task.assignee || "Non assigné"} · {task.endDate || "Sans échéance"}</p>
                </div>
              </td>
            )}
            {otherCols.map((col) => {
              if (col.key === "status") return (
                <td key={col.key} className="py-2.5 px-3 border-r border-slate-200">
                  <button
                    type="button"
                    onClick={() => cycleTaskStatus(task)}
                    className={`block w-full rounded-xl px-3 py-2 text-center text-xs font-bold transition hover:brightness-95 ${statusCellClass(task.status)}`}
                    title="Cliquer pour changer le statut"
                  >
                    {task.status}
                  </button>
                </td>
              );
              if (col.key === "priority") return (
                <td key={col.key} className="py-2.5 px-3 border-r border-slate-200">
                  <button
                    type="button"
                    onClick={() => cycleTaskPriority(task)}
                    className={`inline-flex items-center gap-2 rounded-xl px-3 py-2 text-sm font-semibold transition hover:brightness-95 ${priorityCellClass(task.priority)}`}
                    title="Cliquer pour changer la priorité"
                  >
                    <span className={`h-2.5 w-2.5 rounded-full ${priorityDotClass(task.priority)}`} />
                    <span>{task.priority || "—"}</span>
                  </button>
                </td>
              );
              if (col.key === "computedTotalProgress") return (
                <td key={col.key} className="py-2.5 px-3 min-w-32 border-r border-slate-200">
                  <div className="flex items-center gap-2">
                    <button
                      type="button"
                      onClick={() => onTaskProgressChange?.(task.id, prog - 10)}
                      className="inline-flex h-7 w-7 items-center justify-center rounded-full border border-slate-200 text-sm font-bold text-slate-500 hover:bg-slate-100"
                      title="Diminuer la progression"
                    >
                      -
                    </button>
                    <div className="flex-1 h-1.5 rounded-full bg-slate-200 overflow-hidden">
                      <div className={`h-full rounded-full ${progressTone(prog)}`} style={{ width: `${prog}%` }} />
                    </div>
                    <button
                      type="button"
                      onClick={() => onTaskProgressChange?.(task.id, prog + 10)}
                      className="inline-flex h-7 w-7 items-center justify-center rounded-full border border-slate-200 text-sm font-bold text-slate-500 hover:bg-slate-100"
                      title="Augmenter la progression"
                    >
                      +
                    </button>
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
          <div className="flex w-full flex-col gap-2 sm:w-auto sm:flex-row sm:flex-wrap sm:items-center">
            {selectedCount > 0 && (
              <button onClick={clearSelection} className="inline-flex w-full items-center justify-center gap-2 rounded-2xl border border-slate-200 bg-slate-50 px-3 py-3 text-sm font-semibold text-slate-700 hover:bg-slate-100 sm:w-auto sm:py-2">
                {selectedCount} sélectionné{selectedCount > 1 ? "s" : ""}
                <X className="w-4 h-4" />
              </button>
            )}
            <button
              type="button"
              onClick={() => setCfg({ hiddenCols: [] })}
              className="inline-flex w-full items-center justify-center gap-2 rounded-2xl border bg-white px-3 py-3 text-sm font-semibold text-slate-700 hover:bg-slate-50 sm:w-auto sm:py-2"
            >
              <EyeOff className="w-4 h-4" /> Afficher colonnes
            </button>
            <button
              type="button"
              onClick={() => setAddColModal({ position: "right", anchorKey: visibleCols[visibleCols.length - 1]?.key })}
              className="inline-flex w-full items-center justify-center gap-2 rounded-2xl bg-slate-900 px-4 py-3 text-sm font-semibold text-white shadow-sm hover:bg-slate-800 sm:w-auto sm:py-2"
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
            <GuidedField label={form.parentId ? "Sports concernés par l'axe analytique" : "Sport du projet"}>
              {form.parentId ? (
                <div className="space-y-3">
                  <div className="flex flex-wrap gap-2">
                    <button type="button" onClick={() => ff("sports", allSports.map((sport) => sport.id))} className="rounded-full border px-3 py-1 text-xs font-medium text-slate-700 hover:bg-slate-50">Tous</button>
                    <button type="button" onClick={() => ff("sports", [])} className="rounded-full border px-3 py-1 text-xs font-medium text-slate-700 hover:bg-slate-50">Aucun</button>
                  </div>
                  <div className="grid gap-2 sm:grid-cols-2">
                    {allSports.map((sport) => {
                      const checked = (form.sports || []).includes(sport.id);
                      return (
                        <label key={sport.id} className={`flex items-center gap-3 rounded-2xl border px-3 py-3 text-sm font-medium cursor-pointer transition ${checked ? "border-slate-900 bg-white text-slate-900" : "bg-white text-slate-600 hover:bg-slate-100"}`}>
                          <input
                            type="checkbox"
                            checked={checked}
                            onChange={() => ff("sports", checked ? (form.sports || []).filter((id) => id !== sport.id) : [...(form.sports || []), sport.id])}
                            className="h-4 w-4 rounded border-slate-300"
                          />
                          <span>{sport.name}</span>
                        </label>
                      );
                    })}
                  </div>
                </div>
              ) : (
                <select className={inp} value={form.sport || ""} onChange={(e) => ff("sport", e.target.value)}>
                  {allSports.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
                </select>
              )}
            </GuidedField>
            <GuidedField label="Priorité">
              <StatusToggle options={PRIORITY_OPTIONS} value={form.priority} onChange={(v) => ff("priority", v)} />
            </GuidedField>
            <div className="grid grid-cols-2 gap-4">
              <GuidedField label="Date de début"><input type="date" className={inp} value={form.startDate || ""} onChange={(e) => ff("startDate", e.target.value)} /></GuidedField>
              <GuidedField label="Date de clôture"><input type="date" className={inp} value={form.endDate || ""} onChange={(e) => ff("endDate", e.target.value)} /></GuidedField>
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
                  Axe analytique
                </button>
              </div>
              {form.parentId && (
                <select className={`${inp} mt-2`} value={form.parentId || ""} onChange={(e) => ff("parentId", e.target.value || null)}>
                  {projects.map((p) => <option key={p.id} value={p.id}>Rattaché au projet : {p.name}</option>)}
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

function AttachmentManager({ attachments = [], onAddFiles, onAddWhatsApp, onRemove, whatsAppDraft, onWhatsAppDraftChange }) {
  const fileInputRef = useRef(null);

  return (
    <div className="space-y-3 rounded-2xl border bg-slate-50 p-4 md:col-span-2">
      <div className="flex items-center justify-between gap-3 flex-wrap">
        <div>
          <p className="text-xs font-bold uppercase tracking-wide text-slate-500">Pièces jointes</p>
          <p className="mt-1 text-xs text-slate-500">PDF, JPEG, XLS, audio et messages WhatsApp. Limite conseillée : 4 Mo par fichier.</p>
        </div>
        <button
          type="button"
          onClick={() => fileInputRef.current?.click()}
          className="inline-flex items-center gap-2 rounded-2xl border bg-white px-3 py-2 text-sm font-semibold text-slate-700 hover:bg-slate-100"
        >
          <Paperclip className="w-4 h-4" />
          Ajouter un fichier
        </button>
      </div>

      <input
        ref={fileInputRef}
        type="file"
        multiple
        accept=".pdf,.jpg,.jpeg,.png,.xls,.xlsx,.csv,.mp3,.wav,.m4a,.ogg,.webm,image/*,audio/*,application/pdf,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,text/csv"
        className="hidden"
        onChange={(event) => {
          onAddFiles?.(Array.from(event.target.files || []));
          event.target.value = "";
        }}
      />

      <div className="grid gap-2 md:grid-cols-[1fr_1fr_auto]">
        <input
          className="rounded-2xl border bg-white px-4 py-3 text-sm"
          placeholder="Titre WhatsApp"
          value={whatsAppDraft?.title || ""}
          onChange={(event) => onWhatsAppDraftChange?.("title", event.target.value)}
        />
        <input
          className="rounded-2xl border bg-white px-4 py-3 text-sm"
          placeholder="Lien WhatsApp (optionnel)"
          value={whatsAppDraft?.url || ""}
          onChange={(event) => onWhatsAppDraftChange?.("url", event.target.value)}
        />
        <button
          type="button"
          onClick={onAddWhatsApp}
          className="rounded-2xl bg-slate-900 px-4 py-3 text-sm font-semibold text-white"
        >
          Ajouter WhatsApp
        </button>
      </div>
      <textarea
        className="w-full rounded-2xl border bg-white px-4 py-3 text-sm resize-none h-20"
        placeholder="Contenu du message WhatsApp"
        value={whatsAppDraft?.message || ""}
        onChange={(event) => onWhatsAppDraftChange?.("message", event.target.value)}
      />

      {attachments.length ? (
        <div className="grid gap-2">
          {attachments.map((attachment) => {
            const Icon = attachmentIcon(attachment);
            return (
              <div key={attachment.id} className="flex items-start justify-between gap-3 rounded-2xl border bg-white px-3 py-3">
                <div className="flex min-w-0 items-start gap-3">
                  <div className="mt-0.5 inline-flex h-9 w-9 shrink-0 items-center justify-center rounded-xl bg-slate-100 text-slate-600">
                    <Icon className="h-4 w-4" />
                  </div>
                  <div className="min-w-0">
                    <p className="truncate text-sm font-semibold text-slate-900">{attachment.name || "Pièce jointe"}</p>
                    <div className="mt-1 flex flex-wrap gap-2 text-xs text-slate-500">
                      <span>{attachmentKindLabel(attachment)}</span>
                      {attachment.size ? <span>{formatFileSize(attachment.size)}</span> : null}
                    </div>
                    {attachment.message ? <p className="mt-1 text-xs text-slate-600 line-clamp-2">{attachment.message}</p> : null}
                  </div>
                </div>
                <div className="flex shrink-0 items-center gap-1">
                  {attachment.url ? (
                    <a
                      href={attachment.url}
                      target="_blank"
                      rel="noopener noreferrer"
                      className="inline-flex h-8 w-8 items-center justify-center rounded-lg border text-slate-600 hover:bg-slate-100"
                      title="Ouvrir"
                    >
                      <ExternalLink className="h-4 w-4" />
                    </a>
                  ) : null}
                  <button
                    type="button"
                    onClick={() => onRemove?.(attachment.id)}
                    className="inline-flex h-8 w-8 items-center justify-center rounded-lg border text-red-600 hover:bg-red-50"
                    title="Supprimer"
                  >
                    <Trash2 className="h-4 w-4" />
                  </button>
                </div>
              </div>
            );
          })}
        </div>
      ) : (
        <div className="rounded-2xl border border-dashed bg-white px-4 py-5 text-sm text-slate-500">
          Aucune pièce jointe pour le moment.
        </div>
      )}
    </div>
  );
}

function ProjectCoverManager({ coverImage, onAddImage, onRemoveImage }) {
  const inputRef = useRef(null);

  return (
    <div className="space-y-3 rounded-2xl border bg-slate-50 p-4 md:col-span-2">
      <div className="flex items-center justify-between gap-3 flex-wrap">
        <div>
          <p className="text-xs font-bold uppercase tracking-wide text-slate-500">Miniature du projet</p>
          <p className="mt-1 text-xs text-slate-500">Importe une image pour identifier visuellement le projet dans les listes et les cartes.</p>
        </div>
        <div className="flex gap-2">
          <button
            type="button"
            onClick={() => inputRef.current?.click()}
            className="inline-flex items-center gap-2 rounded-2xl border bg-white px-3 py-2 text-sm font-semibold text-slate-700 hover:bg-slate-100"
          >
            <FileImage className="w-4 h-4" />
            Importer une image
          </button>
          {coverImage ? (
            <button
              type="button"
              onClick={onRemoveImage}
              className="inline-flex items-center gap-2 rounded-2xl border bg-white px-3 py-2 text-sm font-semibold text-red-600 hover:bg-red-50"
            >
              <Trash2 className="w-4 h-4" />
              Retirer
            </button>
          ) : null}
        </div>
      </div>

      <input
        ref={inputRef}
        type="file"
        accept="image/*"
        className="hidden"
        onChange={(event) => {
          onAddImage?.(Array.from(event.target.files || []));
          event.target.value = "";
        }}
      />

      {coverImage ? (
        <div className="overflow-hidden rounded-2xl border bg-white">
          <img src={coverImage} alt="Miniature du projet" className="h-48 w-full object-cover" />
        </div>
      ) : (
        <div className="rounded-2xl border border-dashed bg-white px-4 py-8 text-center text-sm text-slate-500">
          Aucune miniature pour le moment.
        </div>
      )}
    </div>
  );
}

function AttachmentViewerModal({ open, title, attachments = [], onClose }) {
  if (!open) return null;
  return (
    <div className="fixed inset-0 z-[80] bg-black/40 flex items-center justify-center p-4">
      <div className="bg-white rounded-3xl shadow-2xl w-full max-w-3xl max-h-[92vh] overflow-auto p-6">
        <div className="mb-4 flex items-center justify-between gap-4">
          <div>
            <h3 className="text-xl font-semibold text-slate-900">{title}</h3>
            <p className="mt-1 text-sm text-slate-500">{attachments.length} pièce{attachments.length > 1 ? "s jointes" : " jointe"}</p>
          </div>
          <button onClick={onClose} className="text-slate-500 text-xl">✕</button>
        </div>
        <div className="space-y-3">
          {attachments.length ? attachments.map((attachment) => {
            const Icon = attachmentIcon(attachment);
            return (
              <div key={attachment.id} className="rounded-2xl border p-4">
                <div className="flex items-start justify-between gap-3">
                  <div className="flex min-w-0 items-start gap-3">
                    <div className="mt-0.5 inline-flex h-10 w-10 shrink-0 items-center justify-center rounded-xl bg-slate-100 text-slate-600">
                      <Icon className="h-5 w-5" />
                    </div>
                    <div className="min-w-0">
                      <p className="truncate text-sm font-semibold text-slate-900">{attachment.name || "Pièce jointe"}</p>
                      <div className="mt-1 flex flex-wrap gap-2 text-xs text-slate-500">
                        <span>{attachmentKindLabel(attachment)}</span>
                        {attachment.size ? <span>{formatFileSize(attachment.size)}</span> : null}
                      </div>
                      {attachment.message ? <p className="mt-2 text-sm text-slate-600 whitespace-pre-wrap">{attachment.message}</p> : null}
                    </div>
                  </div>
                  {attachment.url ? (
                    <button
                      type="button"
                      onClick={() => openAttachmentFile(attachment)}
                      className="inline-flex items-center gap-2 rounded-2xl border px-3 py-2 text-sm font-semibold text-slate-700 hover:bg-slate-50"
                    >
                      <ExternalLink className="h-4 w-4" /> Ouvrir
                    </button>
                  ) : null}
                </div>
                {attachment.kind === "image" && attachment.url ? (
                  <img src={attachment.url} alt={attachment.name || ""} className="mt-3 max-h-72 w-full rounded-2xl border object-cover" />
                ) : null}
                {attachment.kind === "audio" && attachment.url ? (
                  <audio controls className="mt-3 w-full">
                    <source src={attachment.url} type={attachment.mimeType || "audio/mpeg"} />
                  </audio>
                ) : null}
              </div>
            );
          }) : (
            <div className="rounded-2xl border border-dashed p-6 text-center text-sm text-slate-500">Aucune pièce jointe.</div>
          )}
        </div>
        <div className="mt-5 flex justify-end">
          <button onClick={onClose} className="px-4 py-2 rounded-2xl border bg-white text-sm font-semibold hover:bg-slate-100">Fermer</button>
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
  const [mobileNavOpen, setMobileNavOpen] = useState(false);
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
  const [attachmentViewer, setAttachmentViewer] = useState({ open: false, title: "", attachments: [] });
  const [athleteModal, setAthleteModal] = useState({ open: false, item: null });
  const [clubModal, setClubModal] = useState({ open: false, item: null });
  const [docModal, setDocModal] = useState({ open: false, item: null });
  const [convModal, setConvModal] = useState({ open: false, item: null });
  const [notifModal, setNotifModal] = useState({ open: false, item: null });
  const [cfModal, setCfModal] = useState({ open: false, item: null });
  const [deleteConfirm, setDeleteConfirm] = useState({ open: false, type: null, item: null });
  const [projectViewMode, setProjectViewMode] = useState("tree");
  const [archiveEntity, setArchiveEntity] = useState("projects");
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
  const [cfForm, setCfForm] = useState({});
  const [newSportName, setNewSportName] = useState("");
  const [globalColumnModal, setGlobalColumnModal] = useState(null);
  const [importPreview, setImportPreview] = useState(null);
  const [importScope, setImportScope] = useState("all");
  const importScopeRef = useRef("all");
  const importInputRef = useRef(null);

  const allSports = useMemo(() => [...BASE_SPORTS, ...data.customSports], [data.customSports]);
  const sportName = (value) => {
    if (Array.isArray(value)) return value.map((id) => allSports.find((s) => s.id === id)?.name || id).filter(Boolean).join(", ");
    return allSports.find((s) => s.id === value)?.name || value;
  };
  const projectSportIds = (project) => Array.isArray(project?.sports) && project.sports.length ? project.sports : (project?.sport ? [project.sport] : []);
  const projectMatchesSport = (project, sportId) => sportId === "all" || projectSportIds(project).includes(sportId);
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
  const updateWhatsAppDraft = (key, value) => {
    setForm((prev) => ({ ...prev, whatsAppDraft: { ...(prev.whatsAppDraft || {}), [key]: value } }));
  };
  const toggleFormAllowedUser = (userId) => {
    setForm((prev) => {
      const nextAllowed = new Set(normalizeProjectAccess(prev, currentUser?.id));
      if (nextAllowed.has(userId)) nextAllowed.delete(userId);
      else nextAllowed.add(userId);
      return { ...prev, allowedUserIds: Array.from(nextAllowed) };
    });
  };
  const openAttachmentViewer = (title, attachments) => {
    const safeAttachments = normalizeAttachments(attachments);
    if (!safeAttachments.length) return;
    setAttachmentViewer({ open: true, title, attachments: safeAttachments });
  };
  const removeAttachment = (attachmentId) => {
    setForm((prev) => ({ ...prev, attachments: normalizeAttachments(prev.attachments).filter((attachment) => attachment.id !== attachmentId) }));
  };
  const addWhatsAppAttachment = () => {
    const draft = form.whatsAppDraft || {};
    const title = (draft.title || "").trim();
    const message = (draft.message || "").trim();
    const url = (draft.url || "").trim();
    if (!title && !message && !url) return;
    const attachment = {
      id: uid(),
      kind: "whatsapp",
      name: title || "Message WhatsApp",
      message,
      url,
      createdAt: new Date().toISOString(),
    };
    setForm((prev) => ({
      ...prev,
      attachments: [...normalizeAttachments(prev.attachments), attachment],
      whatsAppDraft: { title: "", message: "", url: "" },
    }));
  };
  const addAttachmentFiles = async (files) => {
    if (!files?.length) return;
    const nextAttachments = [];
    for (const file of files) {
      if (file.size > ATTACHMENT_FILE_LIMIT) {
        window.alert(`Le fichier "${file.name}" dépasse la limite conseillée de 4 Mo.`);
        continue;
      }
      try {
        const url = await readFileAsDataUrl(file);
        nextAttachments.push({
          id: uid(),
          kind: attachmentKindFromFile(file),
          name: file.name,
          size: file.size,
          mimeType: file.type,
          url,
          createdAt: new Date().toISOString(),
        });
      } catch (error) {
        window.alert(`Impossible d'ajouter le fichier "${file.name}".`);
      }
    }
    if (!nextAttachments.length) return;
    setForm((prev) => ({ ...prev, attachments: [...normalizeAttachments(prev.attachments), ...nextAttachments] }));
  };
  const addProjectCoverImage = async (files) => {
    const file = files?.[0];
    if (!file) return;
    if (!file.type.startsWith("image/")) {
      window.alert("Merci de choisir une image pour la miniature du projet.");
      return;
    }
    if (file.size > ATTACHMENT_FILE_LIMIT) {
      window.alert(`L'image "${file.name}" dépasse la limite conseillée de 4 Mo.`);
      return;
    }
    try {
      const url = await readFileAsDataUrl(file);
      setForm((prev) => ({ ...prev, coverImage: url }));
    } catch (error) {
      window.alert(`Impossible d'ajouter l'image "${file.name}".`);
    }
  };
  const toggleFormSport = (sportId) => setForm((prev) => {
    const current = Array.isArray(prev.sports) ? prev.sports : [];
    return {
      ...prev,
      sports: current.includes(sportId) ? current.filter((id) => id !== sportId) : [...current, sportId],
    };
  });
  const addInlineSport = (target = "sport", multi = false) => {
    const rawName = String(form.newSportName || "").trim();
    if (!rawName) return;
    const id = makeSportId(rawName);
    if (!id) return;
    if (!allSports.some((sport) => sport.id === id)) {
      upd("customSports", (arr) => [...arr, { id, name: rawName, icon: Trophy }]);
    }
    setForm((prev) => ({
      ...prev,
      [target]: multi ? undefined : id,
      sports: multi ? Array.from(new Set([...(prev.sports || []), id])) : prev.sports,
      newSportName: "",
    }));
  };
  const addPendingAxis = () => {
    const rawName = String(form.newAxisName || "").trim();
    if (!rawName) return;
    setForm((prev) => {
      const current = Array.isArray(prev.pendingAxes) ? prev.pendingAxes : [];
      if (current.some((axisName) => normalizeKey(axisName) === normalizeKey(rawName))) {
        return { ...prev, newAxisName: "" };
      }
      return { ...prev, pendingAxes: [...current, rawName], newAxisName: "" };
    });
  };
  const removePendingAxis = (axisName) => {
    setForm((prev) => ({
      ...prev,
      pendingAxes: (prev.pendingAxes || []).filter((item) => item !== axisName),
    }));
  };
  const setTaskFormProgress = (value) => {
    const progress = clampProgress(value);
    setForm((prev) => ({ ...prev, progress, status: statusFromProgress(progress, prev.status || "À faire") }));
  };

  const resetFilters = () => {
    setSearchQ(""); setFilterSport("all"); setFilterStatus("all");
    setFilterPriority("all"); setDocMode("all");
  };
  const goToView = (nextView) => {
    setView(nextView);
    setMobileNavOpen(false);
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
    setForm({ projectId, title: "", assignee: "", status: "À faire", priority: "Moyenne", progress: 0, startDate: "", endDate: "", description: "", customFieldValues: {}, attachments: [], whatsAppDraft: { title: "", message: "", url: "" } });
    setTaskModal({ open: true, item: null });
  };

  const handleEditProject = (item) => {
    setForm({
      ...item,
      visibility: item.visibility || "Oui",
      allowedUserIds: normalizeProjectAccess(item, currentUser?.id),
      coverImage: item.coverImage || "",
      sports: Array.isArray(item.sports) ? item.sports : (item.parentId && item.sport ? [item.sport] : []),
      pendingAxes: [],
      newAxisName: "",
      customFieldValues: { ...defaultFinancialValues(projectCustomFields), ...(item.customFieldValues || {}) },
      attachments: normalizeAttachments(item.attachments),
      whatsAppDraft: { title: "", message: "", url: "" },
    });
    setProjectModal({ open: true, item });
  };

  const handleEditSection = (section) => {
    setForm({ ...section, subSectionsText: (section.subSections || []).join("\n") });
    setClubModal({ open: true, item: section });
  };
  const openTaskEditor = (task) => {
    setForm({ ...task, attachments: normalizeAttachments(task.attachments), whatsAppDraft: { title: "", message: "", url: "" } });
    setTaskModal({ open: true, item: task });
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
    const selectedSports = parentId ? Array.from(new Set((form.sports || []).filter(Boolean))) : [];
    const projectId = form.id || uid();
    const item = {
      workspaceId: currentWsId,
      ...form,
      parentId,
      id: projectId,
      sport: parentId ? (selectedSports[0] || form.sport || allSports[0]?.id || "developpe_couche") : (form.sport || allSports[0]?.id || "developpe_couche"),
      sports: selectedSports,
      visibility: form.visibility || "Oui",
      allowedUserIds: normalizeProjectAccess({ visibility: form.visibility || "Oui", allowedUserIds: form.allowedUserIds }, currentUser?.id),
      archived: Boolean(form.archived),
      coverImage: form.coverImage || "",
      attachments: normalizeAttachments(form.attachments),
      durationDays: daysBetween(form.startDate, form.endDate),
      customFieldValues: { ...defaultFinancialValues(projectCustomFields), ...(form.customFieldValues || {}) },
    };
    const pendingAxes = !parentId ? (form.pendingAxes || []).filter(Boolean) : [];
    upd("projects", (arr) => {
      const nextProjects = form.id ? arr.map((p) => p.id === form.id ? item : p) : [...arr, item];
      if (!pendingAxes.length) return nextProjects;
      const existingByParent = nextProjects.filter((project) => project.parentId === projectId);
      const createdAxes = pendingAxes
        .filter((axisName) => !existingByParent.some((project) => normalizeKey(project.name) === normalizeKey(axisName)))
        .map((axisName) => ({
          workspaceId: currentWsId,
          id: uid(),
          name: axisName,
          description: "",
          owner: item.owner || "",
          parentId: projectId,
          sport: item.sport,
          sports: item.sports?.length ? item.sports : [item.sport].filter(Boolean),
          status: "À faire",
          priority: "Moyenne",
          visibility: item.visibility || "Oui",
          allowedUserIds: normalizeProjectAccess(item, currentUser?.id),
          archived: false,
          coverImage: "",
          attachments: [],
          startDate: item.startDate || "",
          endDate: item.endDate || "",
          durationDays: daysBetween(item.startDate, item.endDate),
          customFieldValues: { ...defaultFinancialValues(projectCustomFields) },
        }));
      return [...nextProjects, ...createdAxes];
    });
    setProjectModal({ open: false, item: null }); setForm({});
  };

  const saveTask = () => {
    const progress = normalizeTaskProgress(form.progress, form.status);
    const status = form.status === "Terminé" ? "Terminé" : statusFromProgress(progress, form.status);
    const linkedProject = workspaceProjects.find((project) => project.id === form.projectId);
    const item = {
      workspaceId: currentWsId,
      ...form,
      id: form.id || uid(),
      sport: form.sport || linkedProject?.sport || allSports[0]?.id || "developpe_couche",
      archived: Boolean(form.archived),
      attachments: normalizeAttachments(form.attachments),
      progress,
      status,
      durationDays: daysBetween(form.startDate, form.endDate),
      customFieldValues: form.customFieldValues || {},
    };
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

  const openScopedCustomFieldModal = (target) => {
    setCfForm({
      name: "",
      type: "texte",
      appliesTo: [target],
      checkedLabel: "Oui",
      uncheckedLabel: "Non",
      placeholder: "",
    });
    setCfModal({ open: true, item: null });
  };

  const saveCF = () => {
    const item = {
      ...cfForm,
      id: cfForm.id || uid(),
      appliesTo: cfForm.appliesTo || ["projects"],
      checkedLabel: (cfForm.type || "texte") === "case_a_cocher" ? (cfForm.checkedLabel?.trim() || "Oui") : cfForm.checkedLabel,
      uncheckedLabel: (cfForm.type || "texte") === "case_a_cocher" ? (cfForm.uncheckedLabel?.trim() || "Non") : cfForm.uncheckedLabel,
      placeholder: (cfForm.type || "texte") === "personne" ? (cfForm.placeholder?.trim() || "Choisir un référent") : cfForm.placeholder,
    };
    upd("customFields", (arr) => cfForm.id ? arr.map((x) => x.id === cfForm.id ? item : x) : [...arr, item]);
    setForm((prev) => ({
      ...prev,
      customFieldValues: { ...(prev.customFieldValues || {}), [item.id]: prev.customFieldValues?.[item.id] ?? "" },
    }));
    setCfModal({ open: false, item: null });
    setCfForm({});
  };

  const addSport = () => {
    if (!newSportName.trim()) return;
    const id = makeSportId(newSportName);
    if (allSports.find((s) => s.id === id)) return;
    upd("customSports", (arr) => [...arr, { id, name: newSportName.trim(), icon: Trophy }]);
    setNewSportName("");
  };

  const currentWsId = data.currentWorkspaceId || "ws_jam";
  const currentWs = (data.workspaces || []).find((ws) => ws.id === currentWsId) || data.workspaces?.[0] || null;
  const currentUser = (data.users || []).find((user) => user.id === data.currentUserId) || data.users?.[0] || defaultUsers[0];
  const isAdmin = currentUser?.role === "Administrateur";
  const workspaceUsers = useMemo(
    () => (data.users || []).filter((user) => (user.workspaceIds || []).includes(currentWsId)),
    [data.users, currentWsId]
  );
  const switchWorkspace = (wsId) => setData((prev) => ({ ...prev, currentWorkspaceId: wsId }));
  const switchCurrentUser = (userId) => setData((prev) => ({ ...prev, currentUserId: userId }));
  const updateUserRole = (userId, role) => {
    setData((prev) => ({
      ...prev,
      users: (prev.users || []).map((user) => user.id === userId ? { ...user, role } : user),
    }));
  };
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
    const projectScope = descendants(id, data.projects);
    upd("projects", (arr) => arr.map((project) => projectScope.has(project.id) ? { ...project, status } : project));
    upd("tasks", (arr) => arr.map((task) => {
      if (!projectScope.has(task.projectId)) return task;
      const progress = normalizeTaskProgress(task.progress, status);
      return { ...task, status: statusFromProgress(progress, status), progress };
    }));
  };

  const updateTaskStatus = (id, status) => {
    upd("tasks", (arr) => arr.map((task) => {
      if (task.id !== id) return task;
      const progress = normalizeTaskProgress(task.progress, status);
      return { ...task, status: statusFromProgress(progress, status), progress };
    }));
  };

  const updateTaskProgress = (id, nextProgress) => {
    upd("tasks", (arr) => arr.map((task) => {
      if (task.id !== id) return task;
      const progress = clampProgress(nextProgress);
      const status = statusFromProgress(progress, task.status);
      return { ...task, progress, status };
    }));
  };

  const updateProjectPriority = (id, priority) => {
    upd("projects", (arr) => arr.map((project) => project.id === id ? { ...project, priority } : project));
  };

  const updateTaskPriority = (id, priority) => {
    upd("tasks", (arr) => arr.map((task) => task.id === id ? { ...task, priority } : task));
  };

  const toggleProjectArchived = (id, archived = true) => {
    upd("projects", (arr) => arr.map((project) => project.id === id ? { ...project, archived, archivedAt: archived ? new Date().toISOString() : "" } : project));
  };

  const toggleTaskArchived = (id, archived = true) => {
    upd("tasks", (arr) => arr.map((task) => task.id === id ? { ...task, archived, archivedAt: archived ? new Date().toISOString() : "" } : task));
  };

  const updateProjectField = (id, key, value) => {
    upd("projects", (arr) => arr.map((project) => {
      if (project.id !== id) return project;
      const next = { ...project, [key]: value };
      if (key === "visibility") {
        next.allowedUserIds = normalizeProjectAccess({ ...project, visibility: value }, currentUser?.id);
      }
      if (key === "startDate" || key === "endDate") {
        next.durationDays = daysBetween(key === "startDate" ? value : next.startDate, key === "endDate" ? value : next.endDate);
      }
      return next;
    }));
  };

  const updateTaskField = (id, key, value) => {
    upd("tasks", (arr) => arr.map((task) => {
      if (task.id !== id) return task;
      const next = { ...task, [key]: value };
      if (key === "startDate" || key === "endDate") {
        next.durationDays = daysBetween(key === "startDate" ? value : next.startDate, key === "endDate" ? value : next.endDate);
      }
      return next;
    }));
  };

  const updateAthleteField = (id, key, value) => {
    upd("athletes", (arr) => arr.map((athlete) => {
      if (athlete.id !== id) return athlete;
      if (key === "sectionId") {
        const section = wsClubs.find((item) => item.id === value);
        return { ...athlete, sectionId: value, subSection: "", sport: section?.sport || athlete.sport };
      }
      return { ...athlete, [key]: value };
    }));
  };

  const updateEntityStatus = (key, id, status) => {
    upd(key, (arr) => arr.map((item) => item.id === id ? { ...item, status } : item));
  };

  const canAccessProject = (project) => {
    if (!project || project.workspaceId !== currentWsId) return false;
    if (isAdmin) return true;
    if ((project.visibility || "Oui") !== "Non") return true;
    if (normalizeKey(project.owner || "") === normalizeKey(currentUser?.name || "")) return true;
    return normalizeProjectAccess(project, currentUser?.id).includes(currentUser?.id);
  };

  // Enriched + filtered data (workspace-scoped)
  const rawWorkspaceProjects = useMemo(() => data.projects.filter((p) => p.workspaceId === currentWsId), [data.projects, currentWsId]);
  const workspaceProjects = useMemo(() => rawWorkspaceProjects.filter((project) => canAccessProject(project)), [rawWorkspaceProjects, currentWsId, currentUser?.id, currentUser?.name, isAdmin]);
  const visibleProjectIds = useMemo(() => new Set(workspaceProjects.map((project) => project.id)), [workspaceProjects]);
  const rawWorkspaceTasks = useMemo(() => data.tasks.filter((t) => t.workspaceId === currentWsId), [data.tasks, currentWsId]);
  const workspaceTasks = useMemo(
    () => rawWorkspaceTasks.filter((task) => !task.projectId || visibleProjectIds.has(task.projectId)),
    [rawWorkspaceTasks, visibleProjectIds]
  );
  const wsProjects = useMemo(() => workspaceProjects.filter((project) => !project.archived), [workspaceProjects]);
  const wsTasks = useMemo(() => workspaceTasks.filter((task) => !task.archived), [workspaceTasks]);
  const archivedProjects = useMemo(() => workspaceProjects.filter((project) => project.archived), [workspaceProjects]);
  const archivedTasks = useMemo(() => workspaceTasks.filter((task) => task.archived), [workspaceTasks]);
  const wsAthletes = useMemo(() => data.athletes.filter((a) => a.workspaceId === currentWsId), [data.athletes, currentWsId]);
  const wsClubs = useMemo(() => data.clubs.filter((c) => c.workspaceId === currentWsId), [data.clubs, currentWsId]);

  const enrichedProjects = useMemo(() =>
    wsProjects.map((p) => {
      const projectScope = descendants(p.id, wsProjects);
      const hasLinkedTasks = wsTasks.some((task) => projectScope.has(task.projectId));
      const computedOwnProgress = ownProgress(p.id, wsTasks, wsProjects);
      const computedTotalProgress = consolidatedProgress(p.id, wsTasks, wsProjects);
      return {
        ...p,
        status: hasLinkedTasks ? deriveProjectStatus(computedTotalProgress, p.status) : p.status,
        sportName: sportName(projectSportIds(p)),
        computedOwnProgress,
        computedTotalProgress,
      };
    }), [wsProjects, wsTasks, allSports]);

  const filteredProjects = useMemo(() => enrichedProjects.filter((p) => {
    if (!projectMatchesSport(p, filterSport)) return false;
    if (filterStatus !== "all" && p.status !== filterStatus) return false;
    if (searchQ && !p.name.toLowerCase().includes(searchQ.toLowerCase())) return false;
    return true;
  }), [enrichedProjects, filterSport, filterStatus, searchQ]);
  const filteredRootProjects = useMemo(() => filteredProjects.filter((project) => !project.parentId), [filteredProjects]);
  const filteredSubProjects = useMemo(() => filteredProjects.filter((project) => project.parentId), [filteredProjects]);
  const wsRootProjects = useMemo(() => wsProjects.filter((project) => !project.parentId), [wsProjects]);
  const wsSubProjects = useMemo(() => wsProjects.filter((project) => project.parentId), [wsProjects]);
  const filteredArchivedProjects = useMemo(() => archivedProjects.filter((project) => {
    if (!projectMatchesSport(project, filterSport)) return false;
    if (filterStatus !== "all" && project.status !== filterStatus) return false;
    if (searchQ && !project.name.toLowerCase().includes(searchQ.toLowerCase())) return false;
    return true;
  }), [archivedProjects, filterSport, filterStatus, searchQ]);

  const filteredTasks = useMemo(() => wsTasks.filter((t) => {
    if (filterStatus !== "all" && t.status !== filterStatus) return false;
    if (filterPriority !== "all" && t.priority !== filterPriority) return false;
    if (searchQ && !t.title.toLowerCase().includes(searchQ.toLowerCase())) return false;
    return true;
  }), [wsTasks, filterStatus, filterPriority, searchQ]);
  const filteredArchivedTasks = useMemo(() => archivedTasks.filter((t) => {
    if (filterStatus !== "all" && t.status !== filterStatus) return false;
    if (filterPriority !== "all" && t.priority !== filterPriority) return false;
    if (searchQ && !t.title.toLowerCase().includes(searchQ.toLowerCase())) return false;
    return true;
  }), [archivedTasks, filterStatus, filterPriority, searchQ]);

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
  const peopleOptions = useMemo(() => Array.from(new Set([
    ...(data.users || []).map((user) => user.name),
    ...workspaceProjects.map((project) => project.owner),
    ...workspaceTasks.map((task) => task.assignee),
    ...wsClubs.map((club) => club.contactName),
    ...wsAthletes.map((athlete) => athlete.coach),
  ].filter(Boolean))).sort((a, b) => a.localeCompare(b)), [data.users, workspaceProjects, workspaceTasks, wsClubs, wsAthletes]);

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

  const projectTree = useMemo(() => buildTree(filteredRootProjects), [filteredRootProjects]);
  const subProjectTree = useMemo(() => buildTree(filteredSubProjects), [filteredSubProjects]);

  const projectParentChoices = parentOptionsForProject(form.id, wsProjects);
  const hasAvailableProjectParent = projectParentChoices.length > 0;
  const openNewProjectModal = () => {
    setForm({ status: "À faire", priority: "Moyenne", visibility: "Oui", allowedUserIds: [currentUser?.id].filter(Boolean), sport: allSports[0]?.id || "developpe_couche", sports: [], parentId: null, coverImage: "", pendingAxes: [], newAxisName: "", customFieldValues: defaultFinancialValues(projectCustomFields), attachments: [], whatsAppDraft: { title: "", message: "", url: "" } });
    setProjectModal({ open: true, item: null });
  };
  const openNewStandaloneSubProjectModal = () => {
    const defaultParentId = projectParentChoices[0]?.id || null;
    setForm({
      status: "À faire",
      priority: "Moyenne",
      visibility: "Oui",
      allowedUserIds: [currentUser?.id].filter(Boolean),
      sport: allSports[0]?.id || "developpe_couche",
      sports: [],
      parentId: defaultParentId,
      coverImage: "",
      pendingAxes: [],
      newAxisName: "",
      customFieldValues: defaultFinancialValues(projectCustomFields),
      attachments: [],
      whatsAppDraft: { title: "", message: "", url: "" },
    });
    setProjectModal({ open: true, item: null });
  };
  const handleNewSubProject = (parentProject) => {
    setForm({
      status: "À faire",
      priority: parentProject.priority || "Moyenne",
      visibility: parentProject.visibility || "Oui",
      allowedUserIds: normalizeProjectAccess(parentProject, currentUser?.id),
      sport: parentProject.sport || allSports[0]?.id || "developpe_couche",
      sports: projectSportIds(parentProject),
      parentId: parentProject.id,
      coverImage: "",
      pendingAxes: [],
      newAxisName: "",
      customFieldValues: defaultFinancialValues(projectCustomFields),
      attachments: [],
      whatsAppDraft: { title: "", message: "", url: "" },
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
      appendSheet(wb, "Notifications", mergedNotifications.map((notif) => ({ Titre: notif.title, Sport: sportName(notif.sport), Type: notif.type || "", Statut: notif.status || "", "Date envoi": notif.sentAt || "", Message: notif.message || "", Origine: notif.systemGenerated ? "Alerte automatique" : "Manuelle" })));
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
    if (unknownParents) warnings.push(`${unknownParents} axe(s) analytique(s) indiquent un projet parent introuvable. Ils resteront au niveau principal.`);
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

  const addGlobalColumn = ({ name, type, ...meta }) => {
    const entity = globalColumnModal?.entity;
    if (!entity) return;
    const newCol = { key: "cf_" + uid().slice(0, 8), label: name, type, ...meta };
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
  const rootProjects = wsProjects.filter((project) => !project.parentId);
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
  const upcomingProjects = rootProjects
    .map((project) => ({ ...project, daysLeft: daysFromToday(project.endDate, todayTime) }))
    .filter((project) => project.status !== "Terminé" && project.daysLeft !== null && project.daysLeft >= 0 && project.daysLeft <= 14)
    .sort((a, b) => a.daysLeft - b.daysLeft)
    .slice(0, 5);
  const criticalTasks = openTasks.filter((task) => task.priority === "Critique" || task.priority === "Haute");
  const overdueProjects = rootProjects.filter((project) => {
    const days = daysFromToday(project.endDate, todayTime);
    return project.status !== "Terminé" && days !== null && days < 0;
  });
  const systemDeadlineNotifications = useMemo(() => [
    ...upcomingProjects.map((project) => ({
      id: `system-project-upcoming-${project.id}`,
      title: `Rappel clôture projet : ${project.name}`,
      sport: project.sport,
      channel: "Système",
      audience: project.owner || "Pilotage",
      message: project.daysLeft === 0
        ? `Le projet ${project.name} arrive à sa date de clôture aujourd'hui${project.endDate ? ` (${project.endDate})` : ""}.`
        : `Le projet ${project.name} arrive à échéance dans ${project.daysLeft} jour${project.daysLeft > 1 ? "s" : ""}${project.endDate ? ` (${project.endDate})` : ""}.`,
      attachmentTitle: "",
      status: "Prêt à envoyer",
      sentAt: "",
      workspaceId: currentWsId,
      systemGenerated: true,
      sourceType: "project",
      sourceId: project.id,
    })),
    ...overdueProjects.map((project) => ({
      id: `system-project-overdue-${project.id}`,
      title: `Projet en retard : ${project.name}`,
      sport: project.sport,
      channel: "Système",
      audience: project.owner || "Pilotage",
      message: `La date de clôture du projet ${project.name}${project.endDate ? ` (${project.endDate})` : ""} est dépassée. Une relance est nécessaire.`,
      attachmentTitle: "",
      status: "Prêt à envoyer",
      sentAt: "",
      workspaceId: currentWsId,
      systemGenerated: true,
      sourceType: "project",
      sourceId: project.id,
    })),
    ...upcomingTasks.map((task) => ({
      id: `system-task-upcoming-${task.id}`,
      title: `Rappel clôture tâche : ${task.title}`,
      sport: task.sport || workspaceProjects.find((project) => project.id === task.projectId)?.sport || "",
      channel: "Système",
      audience: task.assignee || "Pilotage",
      message: task.daysLeft === 0
        ? `La tâche ${task.title} arrive à sa date de clôture aujourd'hui${task.endDate ? ` (${task.endDate})` : ""}.`
        : `La tâche ${task.title} arrive à échéance dans ${task.daysLeft} jour${task.daysLeft > 1 ? "s" : ""}${task.endDate ? ` (${task.endDate})` : ""}.`,
      attachmentTitle: "",
      status: "Prêt à envoyer",
      sentAt: "",
      workspaceId: currentWsId,
      systemGenerated: true,
      sourceType: "task",
      sourceId: task.id,
    })),
    ...overdueTasks.map((task) => ({
      id: `system-task-overdue-${task.id}`,
      title: `Tâche en retard : ${task.title}`,
      sport: task.sport || workspaceProjects.find((project) => project.id === task.projectId)?.sport || "",
      channel: "Système",
      audience: task.assignee || "Pilotage",
      message: `La date de clôture de la tâche ${task.title}${task.endDate ? ` (${task.endDate})` : ""} est dépassée. Une relance est nécessaire.`,
      attachmentTitle: "",
      status: "Prêt à envoyer",
      sentAt: "",
      workspaceId: currentWsId,
      systemGenerated: true,
      sourceType: "task",
      sourceId: task.id,
    })),
  ], [upcomingProjects, overdueProjects, upcomingTasks, overdueTasks, currentWsId, workspaceProjects]);
  const mergedNotifications = useMemo(() => [...systemDeadlineNotifications, ...wsNotifications], [systemDeadlineNotifications, wsNotifications]);
  const filteredNotifications = useMemo(() => mergedNotifications.filter((notif) => {
    if (filterSport !== "all" && notif.sport !== filterSport) return false;
    if (filterStatus !== "all" && notif.status !== filterStatus) return false;
    if (searchQ && !notif.title.toLowerCase().includes(searchQ.toLowerCase())) return false;
    return true;
  }), [mergedNotifications, filterSport, filterStatus, searchQ]);
  const incompleteAthletes = wsAthletes.filter((athlete) => !athlete.email || !athlete.phone || !athlete.sectionId || !athlete.subSection || !athlete.category);
  const projectBudgetFields = useMemo(() => {
    const financial = projectCustomFields.filter((field) => field.type === "financier");
    const budgetNamed = financial.filter((field) => normalizeKey(field.name).includes("budget"));
    return budgetNamed.length ? budgetNamed : financial;
  }, [projectCustomFields]);
  const dashboardBudgetForProject = (project) => projectBudgetFields.reduce((sum, field) => sum + (Number(project.customFieldValues?.[field.id]) || 0), 0);
  const budgetTotals = useMemo(() => projectBudgetFields.map((field) => ({
    field,
    total: rootProjects.reduce((sum, project) => sum + (Number(project.customFieldValues?.[field.id]) || 0), 0),
  })), [projectBudgetFields, rootProjects]);
  const totalBudget = budgetTotals.reduce((sum, row) => sum + row.total, 0);
  const budgetedRootProjects = rootProjects.filter((project) => dashboardBudgetForProject(project) > 0);
  const unbudgetedRootProjects = rootProjects.filter((project) => dashboardBudgetForProject(project) <= 0);
  const averageBudget = budgetedRootProjects.length ? totalBudget / budgetedRootProjects.length : 0;
  const topBudgetProjects = [...budgetedRootProjects]
    .sort((a, b) => dashboardBudgetForProject(b) - dashboardBudgetForProject(a))
    .slice(0, 4);
  const dashboardWidgets = data.dashboardWidgets?.length ? data.dashboardWidgets : DEFAULT_DASHBOARD_WIDGETS;

  const dashboardSourceRows = (source) => {
    if (source === "projects") return rootProjects;
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
    if (filter === "budget") {
      if (source !== "projects") return [];
      return rows.filter((row) => dashboardBudgetForProject(row) > 0);
    }
    if (filter === "withBudget") {
      if (source !== "projects") return [];
      return rows.filter((row) => dashboardBudgetForProject(row) > 0);
    }
    if (filter === "withoutBudget") {
      if (source !== "projects") return [];
      return rows.filter((row) => dashboardBudgetForProject(row) <= 0);
    }
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
    if (filter === "missingEmail") {
      if (source === "athletes") return rows.filter((row) => !row.email);
      if (source === "clubs") return rows.filter((row) => !row.email);
      return [];
    }
    if (filter === "missingPhone") {
      if (source === "athletes") return rows.filter((row) => !row.phone);
      if (source === "clubs") return rows.filter((row) => !row.phone);
      return [];
    }
    if (filter === "missingSection") {
      if (source === "athletes") return rows.filter((row) => !row.sectionId);
      return [];
    }
    if (filter === "missingProfile") {
      if (source === "athletes") return rows.filter((row) => !row.category);
      return [];
    }
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
  const renderDashboardEditableCell = (row, column, source) => {
    if (source === "projects") {
      if (column.key === "status") {
        return (
          <button
            type="button"
            onClick={() => updateProjectStatus(row.id, nextOptionValue(STATUS_OPTIONS, row.status))}
            className={`rounded-full border px-2.5 py-1 text-xs font-semibold transition hover:brightness-95 ${badgeClass("status", row.status)}`}
            title="Cliquer pour changer le statut"
          >
            {row.status || "—"}
          </button>
        );
      }
      if (column.key === "priority") {
        return (
          <button
            type="button"
            onClick={() => updateProjectPriority(row.id, nextOptionValue(PRIORITY_OPTIONS, row.priority))}
            className={`rounded-full border px-2.5 py-1 text-xs font-semibold transition hover:brightness-95 ${badgeClass("priority", row.priority)}`}
            title="Cliquer pour changer la priorité"
          >
            {row.priority || "—"}
          </button>
        );
      }
      if (column.key === "owner") {
        return (
          <input
            type="text"
            value={row.owner || ""}
            onChange={(event) => updateProjectField(row.id, "owner", event.target.value)}
            className="w-full rounded-xl border border-slate-200 bg-white px-3 py-2 text-sm text-slate-700 focus:outline-none focus:ring-2 focus:ring-slate-900/10"
            placeholder="Responsable"
          />
        );
      }
      if (column.key === "endDate") {
        return (
          <input
            type="date"
            value={row.endDate || ""}
            onChange={(event) => updateProjectField(row.id, "endDate", event.target.value)}
            className="w-full rounded-xl border border-slate-200 bg-white px-3 py-2 text-sm text-slate-700 focus:outline-none focus:ring-2 focus:ring-slate-900/10"
          />
        );
      }
    }
    if (source === "tasks") {
      if (column.key === "status") {
        return (
          <button
            type="button"
            onClick={() => updateTaskStatus(row.id, nextOptionValue(STATUS_OPTIONS, row.status))}
            className={`rounded-full border px-2.5 py-1 text-xs font-semibold transition hover:brightness-95 ${badgeClass("status", row.status)}`}
            title="Cliquer pour changer le statut"
          >
            {row.status || "—"}
          </button>
        );
      }
      if (column.key === "priority") {
        return (
          <button
            type="button"
            onClick={() => updateTaskPriority(row.id, nextOptionValue(PRIORITY_OPTIONS, row.priority))}
            className={`rounded-full border px-2.5 py-1 text-xs font-semibold transition hover:brightness-95 ${badgeClass("priority", row.priority)}`}
            title="Cliquer pour changer la priorité"
          >
            {row.priority || "—"}
          </button>
        );
      }
      if (column.key === "assignee") {
        return (
          <input
            type="text"
            value={row.assignee || ""}
            onChange={(event) => updateTaskField(row.id, "assignee", event.target.value)}
            className="w-full rounded-xl border border-slate-200 bg-white px-3 py-2 text-sm text-slate-700 focus:outline-none focus:ring-2 focus:ring-slate-900/10"
            placeholder="Responsable"
          />
        );
      }
      if (column.key === "endDate") {
        return (
          <input
            type="date"
            value={row.endDate || ""}
            onChange={(event) => updateTaskField(row.id, "endDate", event.target.value)}
            className="w-full rounded-xl border border-slate-200 bg-white px-3 py-2 text-sm text-slate-700 focus:outline-none focus:ring-2 focus:ring-slate-900/10"
          />
        );
      }
      if (column.key === "progress") {
        const progress = normalizeTaskProgress(row.progress, row.status);
        return (
          <div className="flex min-w-[172px] items-center gap-2">
            <button
              type="button"
              onClick={() => updateTaskProgress(row.id, progress - 10)}
              className="inline-flex h-8 w-8 items-center justify-center rounded-full border border-slate-200 bg-white text-slate-600 transition hover:border-slate-300 hover:bg-slate-50"
              title="Réduire de 10%"
            >
              -
            </button>
            <div className="min-w-0 flex-1">
              <div className="h-2 overflow-hidden rounded-full bg-slate-200">
                <div className={`h-full rounded-full ${progressTone(progress)}`} style={{ width: `${progress}%` }} />
              </div>
            </div>
            <button
              type="button"
              onClick={() => updateTaskProgress(row.id, progress + 10)}
              className="inline-flex h-8 w-8 items-center justify-center rounded-full border border-slate-200 bg-white text-slate-600 transition hover:border-slate-300 hover:bg-slate-50"
              title="Augmenter de 10%"
            >
              +
            </button>
            <span className="w-10 text-right text-xs font-bold text-slate-700">{progress}%</span>
          </div>
        );
      }
    }
    if (source === "athletes") {
      if (column.key === "status") {
        return (
          <button
            type="button"
            onClick={() => updateEntityStatus("athletes", row.id, nextOptionValue(ATHLETE_STATUS, row.status))}
            className={`rounded-full border px-2.5 py-1 text-xs font-semibold transition hover:brightness-95 ${badgeClass("status", row.status)}`}
            title="Cliquer pour changer le statut"
          >
            {row.status || "—"}
          </button>
        );
      }
      if (column.key === "coach") {
        return (
          <input
            type="text"
            value={row.coach || ""}
            onChange={(event) => updateAthleteField(row.id, "coach", event.target.value)}
            className="w-full rounded-xl border border-slate-200 bg-white px-3 py-2 text-sm text-slate-700 focus:outline-none focus:ring-2 focus:ring-slate-900/10"
            placeholder="Coach"
          />
        );
      }
    }
    if (source === "clubs" && column.key === "status") {
      return (
        <button
          type="button"
          onClick={() => updateEntityStatus("clubs", row.id, nextOptionValue(CLUB_STATUS, row.status))}
          className={`rounded-full border px-2.5 py-1 text-xs font-semibold transition hover:brightness-95 ${badgeClass("status", row.status)}`}
          title="Cliquer pour changer le statut"
        >
          {row.status || "—"}
        </button>
      );
    }
    return column.render(row);
  };
  const dashboardDataColumns = (source) => {
    if (source === "projects") {
      return [
        { key: "name", label: "Projet", render: (row) => <span className="font-semibold text-slate-900">{row.name}</span> },
        { key: "status", label: "Statut", render: (row) => <Badge className={badgeClass("status", row.status)}>{row.status || "—"}</Badge>, editable: true },
        { key: "priority", label: "Priorité", render: (row) => <Badge className={badgeClass("priority", row.priority)}>{row.priority || "—"}</Badge>, editable: true },
        { key: "owner", label: "Responsable", render: (row) => row.owner || "—", editable: true },
        { key: "endDate", label: "Échéance", render: (row) => row.endDate || "—", editable: true },
        { key: "budget", label: "Budget", render: (row) => formatMoney(dashboardBudgetForProject(row)) },
      ];
    }
    if (source === "athletes") {
      return [
        { key: "name", label: "Athlète", render: (row) => <span className="font-semibold text-slate-900">{row.name}</span> },
        { key: "sectionId", label: "Section", render: (row) => sectionName(row.sectionId) || "—" },
        { key: "subSection", label: "Sous-section", render: (row) => row.subSection || "—" },
        { key: "category", label: "Profil", render: (row) => row.category || "—" },
        { key: "coach", label: "Coach", render: (row) => row.coach || "—", editable: true },
        { key: "status", label: "Statut", render: (row) => <Badge>{row.status || "—"}</Badge>, editable: true },
        { key: "contact", label: "Contact", render: (row) => row.email || row.phone || "—" },
      ];
    }
    if (source === "clubs") {
      return [
        { key: "name", label: "Section", render: (row) => <span className="font-semibold text-slate-900">{row.name}</span> },
        { key: "status", label: "Statut", render: (row) => <Badge>{row.status || "—"}</Badge>, editable: true },
        { key: "subSections", label: "Sous-sections", render: (row) => row.subSections?.length ? row.subSections.join(", ") : "—" },
        { key: "athletes", label: "Athlètes", render: (row) => sectionAthleteCount(row.id) },
        { key: "contactName", label: "Responsable", render: (row) => row.contactName || "—" },
        { key: "phone", label: "Téléphone", render: (row) => row.phone || "—" },
      ];
    }
    return [
      { key: "title", label: "Tâche", render: (row) => <span className="font-semibold text-slate-900">{row.title}</span> },
      { key: "status", label: "Statut", render: (row) => <Badge className={badgeClass("status", row.status)}>{row.status || "—"}</Badge>, editable: true },
      { key: "priority", label: "Priorité", render: (row) => <Badge className={badgeClass("priority", row.priority)}>{row.priority || "—"}</Badge>, editable: true },
      { key: "assignee", label: "Responsable", render: (row) => row.assignee || "—", editable: true },
      { key: "endDate", label: "Échéance", render: (row) => row.endDate || "—", editable: true },
      { key: "project", label: "Projet", render: (row) => data.projects.find((project) => project.id === row.projectId)?.name || "—" },
      { key: "progress", label: "Progression", render: (row) => `${normalizeTaskProgress(row.progress, row.status)}%`, editable: true },
    ];
  };
  const exportDashboardDrilldown = async () => {
    if (!dashboardDrilldown?.widget) return;
    const source = dashboardDrilldown.widget.source || "tasks";
    const wb = new ExcelJS.Workbook();
    wb.creator = "JAM SaaS";
    wb.created = new Date();
    const rows = dashboardDrilldownVisibleRows.map((row) => {
      if (source === "projects") {
        return {
          Projet: row.name || "",
          Statut: row.status || "",
          Priorité: row.priority || "",
          Responsable: row.owner || "",
          "Date fin": row.endDate || "",
          Budget: dashboardBudgetForProject(row),
        };
      }
      if (source === "athletes") {
        return {
          Athlète: row.name || "",
          Section: sectionName(row.sectionId) || "",
          "Sous-section": row.subSection || "",
          Profil: row.category || "",
          Statut: row.status || "",
          Coach: row.coach || "",
          Email: row.email || "",
          Téléphone: row.phone || "",
        };
      }
      if (source === "clubs") {
        return {
          Section: row.name || "",
          Statut: row.status || "",
          "Sous-sections": row.subSections?.join(", ") || "",
          Responsable: row.contactName || "",
          Email: row.email || "",
          Téléphone: row.phone || "",
          Athlètes: sectionAthleteCount(row.id),
        };
      }
      return {
        Tâche: row.title || "",
        Statut: row.status || "",
        Priorité: row.priority || "",
        Responsable: row.assignee || "",
        Projet: data.projects.find((project) => project.id === row.projectId)?.name || "",
        "Date fin": row.endDate || "",
        Progression: normalizeTaskProgress(row.progress, row.status),
      };
    });
    appendSheet(wb, "Données dashboard", rows.length ? rows : [{ Résultat: "Aucune donnée" }]);
    const slug = normalizeKey(dashboardDrilldown.widget.title || "dashboard").replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "") || "dashboard";
    await downloadWorkbook(wb, `export-dashboard-${slug}.xlsx`);
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
        { id: "projects", label: "Projets", Icon: FolderKanban, count: wsRootProjects.length, alert: overdueProjects.filter((project) => !project.parentId).length, tooltip: overdueProjects.filter((project) => !project.parentId).length ? `${overdueProjects.filter((project) => !project.parentId).length} projet${overdueProjects.filter((project) => !project.parentId).length > 1 ? "s" : ""} en retard` : `${wsRootProjects.length} projet${wsRootProjects.length > 1 ? "s" : ""}` },
        { id: "subprojects", label: "Axes analytiques", Icon: Layers, count: wsSubProjects.length, tooltip: `${wsSubProjects.length} axe${wsSubProjects.length > 1 ? "s" : ""} analytique${wsSubProjects.length > 1 ? "s" : ""} au total` },
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
        { id: "notifications", label: "Notifications", Icon: Bell, count: mergedNotifications.length, tooltip: `${mergedNotifications.length} notification${mergedNotifications.length > 1 ? "s" : ""} au total` },
      ],
    },
    {
      label: "Système",
      items: [
        { id: "archives", label: "Archives", Icon: Library, count: archivedProjects.length + archivedTasks.length, tooltip: `${archivedProjects.length} projet${archivedProjects.length > 1 ? "s" : ""} et ${archivedTasks.length} tâche${archivedTasks.length > 1 ? "s" : ""} archivés` },
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
      <div className="mb-5 rounded-[28px] border border-slate-200/80 bg-white p-3.5 shadow-[0_12px_30px_-18px_rgba(15,23,42,0.22)]">
        <div className="flex flex-wrap items-stretch gap-3">
          <div className="flex min-h-[48px] items-center gap-2 px-2 text-sm font-semibold text-slate-700">
            <Filter className="w-4 h-4 text-slate-400" />
            Filtres
            {activeCount > 0 && <Badge>{activeCount} actif{activeCount > 1 ? "s" : ""}</Badge>}
          </div>
          <div className="relative w-full min-w-0 sm:min-w-[220px] sm:flex-1">
            <Search className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
            <input className="w-full rounded-2xl border border-slate-200 bg-slate-50/80 py-3 pl-9 pr-4 text-sm text-slate-700 focus:bg-white focus:outline-none focus:ring-2 focus:ring-slate-900" placeholder="Rechercher..." value={searchQ} onChange={(e) => setSearchQ(e.target.value)} />
          </div>
          {withSport && (
            <select className="w-full rounded-2xl border border-slate-200 bg-white px-3 py-3 text-sm text-slate-700 focus:outline-none focus:ring-2 focus:ring-slate-900 sm:w-auto" value={filterSport} onChange={(e) => setFilterSport(e.target.value)}>
              <option value="all">Tous les sports</option>
              {allSports.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
            </select>
          )}
          {withStatus && (
            <select className="w-full rounded-2xl border border-slate-200 bg-white px-3 py-3 text-sm text-slate-700 focus:outline-none focus:ring-2 focus:ring-slate-900 sm:w-auto" value={filterStatus} onChange={(e) => setFilterStatus(e.target.value)}>
              <option value="all">Tous les statuts</option>
              {statusOptions.map((s) => <option key={s} value={s}>{s}</option>)}
            </select>
          )}
          {withPriority && (
            <select className="w-full rounded-2xl border border-slate-200 bg-white px-3 py-3 text-sm text-slate-700 focus:outline-none focus:ring-2 focus:ring-slate-900 sm:w-auto" value={filterPriority} onChange={(e) => setFilterPriority(e.target.value)}>
              <option value="all">Toutes les priorités</option>
              {PRIORITY_OPTIONS.map((p) => <option key={p} value={p}>{p}</option>)}
            </select>
          )}
          {activeCount > 0 && (
            <button onClick={clearFilters} className="flex w-full items-center justify-center gap-2 rounded-2xl px-3 py-3 text-sm font-semibold text-slate-600 hover:bg-slate-100 sm:w-auto">
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
      <div className="mb-6 flex flex-wrap items-end justify-between gap-4">
        <div className="space-y-1">
          <h1 className="text-3xl font-bold tracking-tight text-slate-950 sm:text-[2rem]">{title}</h1>
          {subtitle && <p className="max-w-2xl text-sm text-slate-500 md:text-base">{subtitle}</p>}
        </div>
        <div className="flex w-full items-center gap-2 flex-wrap sm:w-auto">
          {actions}
          {onNew && (
            <button onClick={onNew} className="flex w-full items-center justify-center gap-2 rounded-2xl bg-slate-900 px-4 py-2.5 text-sm font-semibold text-white shadow-sm transition hover:bg-slate-800 sm:w-auto">
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
          <button onClick={onSubProject} className="px-3 py-2 rounded-xl border bg-white hover:bg-violet-50 text-violet-700" title="Créer un axe analytique"><FolderKanban className="w-4 h-4" /></button>
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
          {project.coverImage && <img src={project.coverImage} alt="" className="mb-2 h-24 w-full rounded-2xl object-cover border border-slate-200" />}
          <h3 className="font-semibold text-sm text-slate-900 truncate">{project.name}</h3>
          <p className="text-xs text-slate-500 truncate">{sportName(project.sport)}</p>
        </div>
        <div className="flex gap-2 shrink-0">
          <button onClick={() => toggleProjectArchived(project.id, true)} className="px-3 py-2 rounded-xl border bg-white hover:bg-amber-50 text-amber-700" title="Archiver"><Library className="w-4 h-4" /></button>
          <ActionBtns onSubProject={() => handleNewSubProject(project)} onEdit={() => handleEditProject(project)} onDelete={() => requestDelete("project", project)} />
        </div>
      </div>
      <div className="flex flex-wrap gap-1.5">
        <button
          type="button"
          onClick={() => updateProjectPriority(project.id, nextOptionValue(PRIORITY_OPTIONS, project.priority))}
          className={`rounded-full border px-2 py-1 text-xs font-semibold transition hover:brightness-95 ${badgeClass("priority", project.priority)}`}
          title="Cliquer pour changer la priorité"
        >
          {project.priority}
        </button>
        {project.owner && <Badge>{project.owner}</Badge>}
        <Badge className={project.visibility === "Non" ? "bg-slate-100 text-slate-600 border-slate-200" : "bg-emerald-50 text-emerald-700 border-emerald-200"}>{project.visibility || "Oui"}</Badge>
        {normalizeAttachments(project.attachments).length > 0 && (
          <button
            type="button"
            onClick={() => openAttachmentViewer(`Pièces jointes · ${project.name}`, project.attachments)}
            className="inline-flex items-center gap-1 rounded-full border border-slate-200 bg-white px-2 py-1 text-[11px] font-semibold text-slate-600 hover:bg-slate-100"
            title="Voir les pièces jointes"
          >
            <Paperclip className="h-3 w-3" /> {normalizeAttachments(project.attachments).length}
          </button>
        )}
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
          <ActionBtns onEdit={() => openTaskEditor(task)} onDelete={() => requestDelete("task", task)} />
        </div>
      <div className="flex flex-wrap gap-1.5">
        {normalizeAttachments(task.attachments).length > 0 && (
          <button
            type="button"
            onClick={() => openAttachmentViewer(`Pièces jointes · ${task.title}`, task.attachments)}
            className="inline-flex items-center gap-1 rounded-full border border-slate-200 bg-white px-2 py-1 text-[11px] font-semibold text-slate-600 hover:bg-slate-100"
            title="Voir les pièces jointes"
          >
            <Paperclip className="h-3 w-3" /> {normalizeAttachments(task.attachments).length}
          </button>
        )}
        <button
            type="button"
            onClick={() => updateTaskPriority(task.id, nextOptionValue(PRIORITY_OPTIONS, task.priority))}
            className={`rounded-full border px-2 py-1 text-xs font-semibold transition hover:brightness-95 ${badgeClass("priority", task.priority)}`}
            title="Cliquer pour changer la priorité"
          >
            {task.priority}
          </button>
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
        {!notif.systemGenerated && (
          <ActionBtns onEdit={() => { setForm({ ...notif }); setNotifModal({ open: true, item: notif }); }} onDelete={() => requestDelete("notification", notif)} />
        )}
      </div>
      <div className="flex flex-wrap gap-1.5">
        {notif.systemGenerated && <Badge className="bg-amber-50 text-amber-700 border-amber-200">Alerte automatique</Badge>}
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

  const currentViewLabel = {
    dashboard: "Tableau de bord",
    dashboardData: "Données du tableau de bord",
    projects: "Projets",
    subprojects: "Axes analytiques",
    tasks: "Tâches",
    athletes: "Athlètes",
    clubs: "Sections",
    documents: "Documents",
    convocations: "Convocations",
    notifications: "Notifications",
    settings: "Paramètres",
    archives: "Archives",
  }[view] || "JAM SaaS";

  return (
    <div className="min-h-screen bg-slate-50 lg:flex">
      <input ref={importInputRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={handleImportFile} />
      {mobileNavOpen && (
        <button
          type="button"
          aria-label="Fermer le menu"
          onClick={() => setMobileNavOpen(false)}
          className="fixed inset-0 z-30 bg-slate-950/35 lg:hidden"
        />
      )}
      {/* Sidebar */}
      <aside className={`fixed inset-y-0 left-0 z-40 flex h-screen w-72 max-w-[88vw] flex-col border-r border-slate-200 bg-white shadow-2xl transition-transform duration-200 lg:sticky lg:top-0 lg:z-auto lg:w-64 lg:max-w-none lg:translate-x-0 lg:shadow-none ${mobileNavOpen ? "translate-x-0" : "-translate-x-full"}`}>
        <div className="px-5 py-4 border-b border-slate-200">
          <div className="flex items-center justify-between gap-3">
            <div>
              <p className="font-bold text-slate-900 leading-none">JAM SaaS</p>
              <p className="text-xs text-slate-500 mt-0.5">v13</p>
            </div>
            <button
              type="button"
              onClick={() => setMobileNavOpen(false)}
              className="inline-flex h-10 w-10 items-center justify-center rounded-2xl border border-slate-200 text-slate-500 lg:hidden"
              aria-label="Fermer le menu"
            >
              <X className="h-4 w-4" />
            </button>
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
                onClick={() => openDashboardWidget({ id: "sidebar_overdue_tasks", type: "number", title: "Tâches en retard", source: "tasks", filter: "overdue", tone: "red" })}
                className="rounded-xl bg-white border border-slate-200 p-2 text-left transition hover:border-red-200 hover:bg-red-50 focus:outline-none focus:ring-2 focus:ring-red-200"
                title="Voir les tâches en retard"
              >
                <p className="font-bold text-slate-900">{overdueTasks.length}</p>
                <p className="text-slate-500">retard</p>
              </button>
              <button
                type="button"
                onClick={() => openDashboardWidget({ id: "sidebar_upcoming_tasks", type: "number", title: "Échéances à 14 jours", source: "tasks", filter: "upcoming", tone: "blue" })}
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
      <main className="min-w-0 flex-1 overflow-x-hidden">
        <div className="sticky top-0 z-20 border-b border-slate-200 bg-white/90 px-4 py-3 backdrop-blur lg:hidden">
          <div className="flex items-center gap-3">
            <button
              type="button"
              onClick={() => setMobileNavOpen(true)}
              className="inline-flex h-10 w-10 items-center justify-center rounded-2xl border border-slate-200 bg-white text-slate-700 shadow-sm"
              aria-label="Ouvrir le menu"
            >
              <Menu className="h-4.5 w-4.5" />
            </button>
            <div className="min-w-0">
              <p className="truncate text-[15px] font-semibold text-slate-900">{currentViewLabel}</p>
              <p className="truncate text-[11px] text-slate-500">{currentWs?.name || "JAM"}</p>
            </div>
          </div>
        </div>

        {/* ── DASHBOARD ── */}
        {view === "dashboard" && (
          <div className="mx-auto max-w-7xl px-4 pb-28 pt-4 sm:p-6">
            <div className="mb-6 flex flex-wrap items-start justify-between gap-4">
              <div className="max-w-2xl">
                <h1 className="text-[2.2rem] font-bold leading-none text-slate-950 sm:text-3xl">Tableau de bord</h1>
                <p className="mt-2 text-base text-slate-500 sm:mt-1">Une vue pilotable de la JAM, avec des indicateurs que tu peux adapter.</p>
              </div>
              <div className="grid w-full grid-cols-2 gap-2 sm:flex sm:w-auto sm:flex-wrap">
                <button
                  type="button"
                  onClick={() => openWidgetEditor()}
                  className="inline-flex items-center justify-center gap-2 rounded-2xl border bg-white px-3 py-2.5 text-sm font-semibold text-slate-700 shadow-sm hover:bg-slate-50 sm:px-4 sm:py-2"
                >
                  <Plus className="h-4 w-4" /> Ajouter un widget
                </button>
                <button
                  type="button"
                  onClick={() => setDashboardEditMode((value) => !value)}
                  className={`inline-flex items-center justify-center gap-2 rounded-2xl px-3 py-2.5 text-sm font-semibold shadow-sm sm:px-4 sm:py-2 ${dashboardEditMode ? "bg-slate-900 text-white" : "border bg-white text-slate-700 hover:bg-slate-50"}`}
                >
                  <BarChart3 className="h-4 w-4" /> {dashboardEditMode ? "Terminer" : "Personnaliser"}
                </button>
                {dashboardEditMode && (
                  <button
                    type="button"
                    onClick={resetDashboardWidgets}
                    className="col-span-2 inline-flex items-center justify-center gap-2 rounded-2xl border bg-white px-4 py-2.5 text-sm font-semibold text-slate-500 shadow-sm hover:bg-slate-50 sm:col-auto sm:px-4 sm:py-2"
                  >
                    Réinitialiser
                  </button>
                )}
              </div>
            </div>

            <div className="mb-6 grid gap-3 sm:gap-4 md:grid-cols-2 xl:grid-cols-4">
              <StatCard
                title="Projets principaux"
                value={rootProjects.length}
                subtitle={`${overdueProjects.length} en retard sur les projets principaux`}
                Icon={FolderKanban}
                tone="red"
                onClick={() => openDashboardWidget({ id: "overview_root_projects", type: "number", title: "Projets principaux", source: "projects", filter: "all", tone: "red" })}
              />
              <StatCard
                title="Tâches totales"
                value={totalTasks}
                subtitle={`${openTasks.length} ouvertes, ${completedTasks} terminées`}
                Icon={ListChecks}
                tone="blue"
                onClick={() => openDashboardWidget({ id: "overview_all_tasks", type: "number", title: "Toutes les tâches", source: "tasks", filter: "all", tone: "blue" })}
              />
              <StatCard
                title="Athlètes suivis"
                value={wsAthletes.length}
                subtitle={`${incompleteAthletes.length} fiches à compléter`}
                Icon={Users}
                tone="amber"
                onClick={() => openDashboardWidget({ id: "overview_all_athletes", type: "number", title: "Tous les athlètes", source: "athletes", filter: "all", tone: "amber" })}
              />
              <StatCard
                title="Budget consolidé"
                value={formatMoney(totalBudget)}
                subtitle={`${projectBudgetFields.length} colonne${projectBudgetFields.length > 1 ? "s" : ""} financière${projectBudgetFields.length > 1 ? "s" : ""} suivie${projectBudgetFields.length > 1 ? "s" : ""}`}
                Icon={BarChart3}
                tone="green"
                onClick={() => openDashboardWidget({ id: "overview_budget_projects", type: "number", title: "Budget consolidé", source: "projects", filter: "budget", tone: "green" })}
              />
            </div>

            <Card className="mb-5 overflow-hidden p-5">
              <div className="flex flex-wrap items-start justify-between gap-4">
                <div className="max-w-2xl">
                  <p className="text-[11px] font-bold uppercase tracking-[0.16em] text-slate-400">Cockpit</p>
                  <h2 className="mt-2 text-xl font-bold text-slate-950">Vue synthétique des opérations</h2>
                  <p className="mt-1 text-sm text-slate-500">Les widgets restent personnalisables, mais l’ensemble devient plus lisible et plus orienté décision.</p>
                </div>
                <div className="grid w-full gap-3 sm:grid-cols-3 lg:min-w-[260px] lg:w-auto">
                  <div className="rounded-2xl border border-slate-200 bg-slate-50/80 px-4 py-3">
                    <p className="text-[11px] font-bold uppercase tracking-[0.16em] text-slate-400">Progression moyenne</p>
                    <p className="mt-2 text-2xl font-bold text-slate-950">{avgProgress}%</p>
                  </div>
                  <div className="rounded-2xl border border-red-100 bg-red-50 px-4 py-3">
                    <p className="text-[11px] font-bold uppercase tracking-[0.16em] text-red-500">Retards</p>
                    <p className="mt-2 text-2xl font-bold text-red-700">{overdueTasks.length + overdueProjects.length}</p>
                  </div>
                  <div className="rounded-2xl border border-blue-100 bg-blue-50 px-4 py-3">
                    <p className="text-[11px] font-bold uppercase tracking-[0.16em] text-blue-500">À venir</p>
                    <p className="mt-2 text-2xl font-bold text-blue-700">{upcomingTasks.length}</p>
                  </div>
                </div>
              </div>
            </Card>

            <Card className="mb-5 p-5">
              <div className="mb-5 flex flex-wrap items-start justify-between gap-4">
                <div className="max-w-2xl">
                  <p className="text-[11px] font-bold uppercase tracking-[0.16em] text-emerald-500">Finance</p>
                  <h2 className="mt-2 text-xl font-bold text-slate-950">Pilotage financier</h2>
                  <p className="mt-1 text-sm text-slate-500">Une lecture claire des budgets saisis sur les projets principaux, avec accès direct au détail sans quitter le tableau de bord.</p>
                </div>
                <Badge className={budgetedRootProjects.length ? "bg-emerald-50 text-emerald-700 border-emerald-200" : "bg-amber-50 text-amber-700 border-amber-200"}>
                  {budgetedRootProjects.length ? `${budgetedRootProjects.length} projet${budgetedRootProjects.length > 1 ? "s" : ""} budgété${budgetedRootProjects.length > 1 ? "s" : ""}` : "Aucun budget saisi"}
                </Badge>
              </div>

              <div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-4">
                <button
                  type="button"
                  onClick={() => openDashboardWidget({ id: "finance_total_budget", type: "number", title: "Budget total consolidé", source: "projects", filter: "budget", tone: "green" })}
                  className="rounded-2xl border border-emerald-100 bg-emerald-50 p-4 text-left transition hover:-translate-y-0.5 hover:shadow-sm"
                >
                  <p className="text-xs font-bold uppercase tracking-wide text-emerald-700">Budget total</p>
                  <p className="mt-2 text-3xl font-bold text-emerald-700">{formatMoney(totalBudget)}</p>
                  <p className="mt-1 text-sm text-emerald-700/70">ouvrir les projets budgétés</p>
                </button>
                <button
                  type="button"
                  onClick={() => openDashboardWidget({ id: "finance_budgeted_projects", type: "number", title: "Projets avec budget", source: "projects", filter: "withBudget", tone: "green" })}
                  className="rounded-2xl border border-slate-200 bg-white p-4 text-left transition hover:-translate-y-0.5 hover:shadow-sm"
                >
                  <p className="text-xs font-bold uppercase tracking-wide text-slate-500">Projets budgétés</p>
                  <p className="mt-2 text-3xl font-bold text-slate-950">{budgetedRootProjects.length}</p>
                  <p className="mt-1 text-sm text-slate-500">sur {rootProjects.length} projets principaux</p>
                </button>
                <button
                  type="button"
                  onClick={() => openDashboardWidget({ id: "finance_unbudgeted_projects", type: "number", title: "Projets sans budget", source: "projects", filter: "withoutBudget", tone: "amber" })}
                  className="rounded-2xl border border-amber-100 bg-amber-50 p-4 text-left transition hover:-translate-y-0.5 hover:shadow-sm"
                >
                  <p className="text-xs font-bold uppercase tracking-wide text-amber-700">Sans budget</p>
                  <p className="mt-2 text-3xl font-bold text-amber-700">{unbudgetedRootProjects.length}</p>
                  <p className="mt-1 text-sm text-amber-700/70">projets à compléter</p>
                </button>
                <button
                  type="button"
                  onClick={() => openDashboardWidget({ id: "finance_avg_budget", type: "number", title: "Budget moyen par projet", source: "projects", filter: "withBudget", tone: "blue" })}
                  className="rounded-2xl border border-blue-100 bg-blue-50 p-4 text-left transition hover:-translate-y-0.5 hover:shadow-sm"
                >
                  <p className="text-xs font-bold uppercase tracking-wide text-blue-700">Budget moyen</p>
                  <p className="mt-2 text-3xl font-bold text-blue-700">{formatMoney(averageBudget)}</p>
                  <p className="mt-1 text-sm text-blue-700/70">sur les projets budgétés</p>
                </button>
              </div>

              <div className="mt-5 grid gap-4 xl:grid-cols-[1.2fr_0.8fr]">
                <div className="rounded-2xl border border-slate-200 bg-slate-50/70 p-4">
                  <div className="mb-3 flex items-center justify-between gap-3">
                    <div>
                      <p className="text-sm font-semibold text-slate-900">Top budgets</p>
                      <p className="text-xs text-slate-500">Les plus gros montants saisis actuellement.</p>
                    </div>
                    <button
                      type="button"
                      onClick={() => openDashboardWidget({ id: "finance_budget_ranking", type: "bar", title: "Budgets par projet", source: "projects", filter: "withBudget", groupBy: "status", tone: "green" })}
                      className="text-xs font-semibold text-slate-500 hover:text-slate-900"
                    >
                      Voir le détail
                    </button>
                  </div>
                  <div className="space-y-2">
                    {topBudgetProjects.length ? topBudgetProjects.map((project, index) => (
                      <button
                        key={project.id}
                        type="button"
                        onClick={() => editDashboardRow(project, "projects")}
                        className="flex w-full items-center justify-between gap-3 rounded-2xl bg-white px-3 py-3 text-left transition hover:-translate-y-0.5 hover:shadow-sm"
                      >
                        <div className="min-w-0">
                          <p className="text-xs font-bold uppercase tracking-wide text-slate-400">#{index + 1}</p>
                          <p className="truncate text-sm font-semibold text-slate-900">{project.name}</p>
                          <p className="text-xs text-slate-500">{project.owner || "Responsable non renseigné"}</p>
                        </div>
                        <div className="shrink-0 text-right">
                          <p className="text-sm font-bold text-slate-950">{formatMoney(dashboardBudgetForProject(project))}</p>
                          <p className="text-xs text-slate-500">{project.endDate || "Sans échéance"}</p>
                        </div>
                      </button>
                    )) : (
                      <div className="rounded-2xl bg-white px-3 py-6 text-center text-sm text-slate-500">
                        Aucun budget n’a encore été saisi sur les projets principaux.
                      </div>
                    )}
                  </div>
                </div>

                <div className="rounded-2xl border border-slate-200 bg-white p-4">
                  <p className="text-sm font-semibold text-slate-900">Couverture budgétaire</p>
                  <p className="mt-1 text-xs text-slate-500">Part des projets principaux qui ont déjà un budget renseigné.</p>
                  <div className="mt-4">
                    <ProgressBar value={rootProjects.length ? Math.round((budgetedRootProjects.length / rootProjects.length) * 100) : 0} label="Projets principaux couverts" />
                  </div>
                  <div className="mt-4 grid gap-2">
                    <button type="button" onClick={() => openDashboardWidget({ id: "finance_budgeted_cover", type: "number", title: "Projets avec budget", source: "projects", filter: "withBudget", tone: "green" })} className="flex items-center justify-between rounded-xl bg-emerald-50 px-3 py-2.5 text-left text-sm hover:bg-emerald-100/80">
                      <span className="text-emerald-800">Avec budget</span>
                      <span className="font-bold text-emerald-700">{budgetedRootProjects.length}</span>
                    </button>
                    <button type="button" onClick={() => openDashboardWidget({ id: "finance_budget_missing", type: "number", title: "Projets sans budget", source: "projects", filter: "withoutBudget", tone: "amber" })} className="flex items-center justify-between rounded-xl bg-amber-50 px-3 py-2.5 text-left text-sm hover:bg-amber-100/80">
                      <span className="text-amber-800">Sans budget</span>
                      <span className="font-bold text-amber-700">{unbudgetedRootProjects.length}</span>
                    </button>
                    <button type="button" onClick={() => openDashboardWidget({ id: "finance_all_root_projects", type: "number", title: "Tous les projets principaux", source: "projects", filter: "all", tone: "slate" })} className="flex items-center justify-between rounded-xl bg-slate-50 px-3 py-2.5 text-left text-sm hover:bg-slate-100">
                      <span className="text-slate-700">Ouvrir les projets</span>
                      <span className="font-bold text-slate-900">{rootProjects.length}</span>
                    </button>
                  </div>
                </div>
              </div>
            </Card>

            <div className="mb-3 flex items-center justify-between gap-3">
              <div>
                <p className="text-[11px] font-bold uppercase tracking-[0.16em] text-slate-400">Widgets</p>
                <h2 className="mt-1 text-lg font-bold text-slate-950">Indicateurs personnalisables</h2>
              </div>
              <p className="text-sm text-slate-500">{dashboardWidgets.length} widget{dashboardWidgets.length > 1 ? "s" : ""}</p>
            </div>

            <div className="mb-5 grid gap-4 sm:grid-cols-2 xl:grid-cols-4">
              {dashboardWidgets.map((widget) => renderDashboardWidget(widget))}
            </div>

            <div className="grid gap-4 xl:grid-cols-[1.2fr_0.8fr]">
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
                <div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-4">
                  <button type="button" onClick={() => openDashboardWidget({ id: "operations_overdue_tasks", type: "number", title: "Tâches en retard", source: "tasks", filter: "overdue", tone: "red" })} className="rounded-2xl border border-red-100 bg-red-50 p-4 text-left transition hover:-translate-y-0.5 hover:shadow-sm">
                    <p className="text-xs font-bold uppercase tracking-wide text-red-700">Tâches en retard</p>
                    <p className="mt-2 text-3xl font-bold text-red-700">{overdueTasks.length}</p>
                    <p className="mt-1 text-sm text-red-700/70">ouvrir le tableau</p>
                  </button>
                  <button type="button" onClick={() => openDashboardWidget({ id: "operations_overdue_projects", type: "number", title: "Projets en retard", source: "projects", filter: "overdue", tone: "red" })} className="rounded-2xl border border-rose-100 bg-rose-50 p-4 text-left transition hover:-translate-y-0.5 hover:shadow-sm">
                    <p className="text-xs font-bold uppercase tracking-wide text-rose-700">Projets en retard</p>
                    <p className="mt-2 text-3xl font-bold text-rose-700">{overdueProjects.length}</p>
                    <p className="mt-1 text-sm text-rose-700/70">ouvrir le tableau</p>
                  </button>
                  <button type="button" onClick={() => openDashboardWidget({ id: "operations_critical_tasks", type: "number", title: "Tâches critiques", source: "tasks", filter: "highPriority", tone: "amber" })} className="rounded-2xl border border-amber-100 bg-amber-50 p-4 text-left transition hover:-translate-y-0.5 hover:shadow-sm">
                    <p className="text-xs font-bold uppercase tracking-wide text-amber-700">Critiques</p>
                    <p className="mt-2 text-3xl font-bold text-amber-700">{criticalTasks.length}</p>
                    <p className="mt-1 text-sm text-amber-700/70">haute priorité</p>
                  </button>
                  <button type="button" onClick={() => openDashboardWidget({ id: "operations_upcoming_tasks", type: "number", title: "Échéances à 14 jours", source: "tasks", filter: "upcoming", tone: "blue" })} className="rounded-2xl border border-blue-100 bg-blue-50 p-4 text-left transition hover:-translate-y-0.5 hover:shadow-sm">
                    <p className="text-xs font-bold uppercase tracking-wide text-blue-700">14 jours</p>
                    <p className="mt-2 text-3xl font-bold text-blue-700">{upcomingTasks.length}</p>
                    <p className="mt-1 text-sm text-blue-700/70">échéances proches</p>
                  </button>
                </div>
                <div className="mt-5">
                  <ProgressBar value={avgProgress} label="Progression moyenne des tâches" />
                </div>
                {upcomingTasks.length > 0 && (
                  <div className="mt-5 rounded-2xl border border-slate-200 bg-slate-50/80 p-4">
                    <div className="mb-3 flex items-center justify-between gap-3">
                      <p className="text-sm font-semibold text-slate-900">Échéances à surveiller</p>
                      <button type="button" onClick={() => openDashboardWidget({ id: "operations_upcoming_tasks_link", type: "number", title: "Échéances à 14 jours", source: "tasks", filter: "upcoming", tone: "blue" })} className="text-xs font-semibold text-slate-500 hover:text-slate-900">
                        Ouvrir les tâches
                      </button>
                    </div>
                    <div className="space-y-2">
                      {upcomingTasks.slice(0, 3).map((task) => (
                        <div key={task.id} className="flex items-center justify-between gap-3 rounded-2xl bg-white px-3 py-2.5">
                          <div className="min-w-0">
                            <p className="truncate text-sm font-semibold text-slate-900">{task.title}</p>
                            <p className="text-xs text-slate-500">{data.projects.find((project) => project.id === task.projectId)?.name || "Sans projet"}</p>
                          </div>
                          <div className="shrink-0 text-right">
                            <p className="text-sm font-semibold text-slate-900">{task.endDate || "—"}</p>
                            <p className="text-xs text-slate-500">{task.daysLeft === 0 ? "Aujourd’hui" : `${task.daysLeft} jour${task.daysLeft > 1 ? "s" : ""}`}</p>
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                )}
              </Card>

              <Card className="p-5">
                <h2 className="mb-4 text-lg font-semibold text-slate-950">Données à compléter</h2>
                <div className="space-y-2">
                  <button type="button" onClick={() => openDashboardWidget({ id: "data_missing_email", type: "number", title: "Athlètes sans email", source: "athletes", filter: "missingEmail", tone: "amber" })} className="flex w-full justify-between rounded-xl px-2 py-2 text-left text-sm hover:bg-slate-50"><span>Athlètes sans email</span><Badge>{wsAthletes.filter((a) => !a.email).length}</Badge></button>
                  <button type="button" onClick={() => openDashboardWidget({ id: "data_missing_phone", type: "number", title: "Athlètes sans téléphone", source: "athletes", filter: "missingPhone", tone: "amber" })} className="flex w-full justify-between rounded-xl px-2 py-2 text-left text-sm hover:bg-slate-50"><span>Athlètes sans téléphone</span><Badge>{wsAthletes.filter((a) => !a.phone).length}</Badge></button>
                  <button type="button" onClick={() => openDashboardWidget({ id: "data_missing_section", type: "number", title: "Athlètes sans section", source: "athletes", filter: "missingSection", tone: "amber" })} className="flex w-full justify-between rounded-xl px-2 py-2 text-left text-sm hover:bg-slate-50"><span>Athlètes sans section</span><Badge>{wsAthletes.filter((a) => !a.sectionId).length}</Badge></button>
                  <button type="button" onClick={() => openDashboardWidget({ id: "data_missing_profile", type: "number", title: "Athlètes sans profil", source: "athletes", filter: "missingProfile", tone: "amber" })} className="flex w-full justify-between rounded-xl px-2 py-2 text-left text-sm hover:bg-slate-50"><span>Athlètes sans profil</span><Badge>{wsAthletes.filter((a) => !a.category).length}</Badge></button>
                  <button type="button" onClick={() => openDashboardWidget({ id: "data_unassigned_tasks", type: "number", title: "Tâches non assignées", source: "tasks", filter: "unassigned", tone: "blue" })} className="flex w-full justify-between rounded-xl px-2 py-2 text-left text-sm hover:bg-slate-50"><span>Tâches non assignées</span><Badge>{wsTasks.filter((task) => !task.assignee).length}</Badge></button>
                </div>
              </Card>
            </div>
          </div>
        )}

        {/* ── DASHBOARD DATA ── */}
        {view === "dashboardData" && (
          <div className="mx-auto max-w-7xl p-4 sm:p-6">
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
                      onClick={exportDashboardDrilldown}
                      className="inline-flex items-center gap-2 rounded-2xl bg-slate-900 px-4 py-2 text-sm font-semibold text-white shadow-sm hover:bg-slate-800"
                    >
                      <Download className="h-4 w-4" /> Exporter Excel
                    </button>
                  </div>
                </div>

                <div className="mb-4 grid gap-3 sm:grid-cols-2 xl:grid-cols-4">
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
                    <div className="relative w-full min-w-0 sm:min-w-[260px] sm:flex-1">
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
                                {renderDashboardEditableCell(row, column, dashboardDrilldown.widget.source)}
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
          <div className="mx-auto max-w-7xl p-4 sm:p-6">
            <div className="mb-5 rounded-3xl border border-slate-200 bg-white p-5 shadow-sm">
              <div className="flex items-start justify-between flex-wrap gap-4">
                <div className="min-w-0">
                  <div className="flex items-center gap-3">
                    <span className="inline-flex h-11 w-11 items-center justify-center rounded-2xl bg-slate-900 text-white shadow-sm">
                      <FolderKanban className="w-5 h-5" />
                    </span>
                    <div>
                      <h1 className="text-2xl font-bold text-slate-900">Projets</h1>
                      <p className="text-slate-500 mt-0.5">{filteredRootProjects.length} projet{filteredRootProjects.length > 1 ? "s" : ""} principal{filteredRootProjects.length > 1 ? "aux" : ""} affiché{filteredRootProjects.length > 1 ? "s" : ""}</p>
                    </div>
                  </div>
                  <div className="mt-4 flex flex-wrap gap-2">
                    <Badge className="bg-slate-50 text-slate-700 border-slate-200">{wsRootProjects.length} principaux</Badge>
                    <Badge className="bg-blue-50 text-blue-700 border-blue-200">{wsRootProjects.filter((project) => project.status === "En cours").length} en cours</Badge>
                    <Badge className="bg-emerald-50 text-emerald-700 border-emerald-200">{wsRootProjects.filter((project) => project.status === "Terminé").length} terminés</Badge>
                    {overdueProjects.filter((project) => !project.parentId).length > 0 && <Badge className="bg-red-50 text-red-700 border-red-200">{overdueProjects.filter((project) => !project.parentId).length} en retard</Badge>}
                    {wsSubProjects.length > 0 && <Badge className="bg-violet-50 text-violet-700 border-violet-200">{wsSubProjects.length} axes analytiques rangés à part</Badge>}
                  </div>
                </div>
                <div className="flex w-full items-center gap-2 flex-wrap justify-start xl:w-auto xl:justify-end">
                  <ViewToggle options={projectViewOptions} value={projectViewMode} onChange={setProjectViewMode} />
                  {fileActions("projects", "projets")}
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
            {FilterBar({ withPriority: false })}
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
                  onEditTask={(task) => openTaskEditor(task)}
                  onArchiveProject={(project) => toggleProjectArchived(project.id, true)}
                  requestDelete={requestDelete}
                  onDuplicate={handleDuplicate}
                  projectCustomFields={projectCustomFields}
                  onProjectStatusChange={updateProjectStatus}
                  onProjectPriorityChange={updateProjectPriority}
                  onTaskStatusChange={updateTaskStatus}
                  onTaskPriorityChange={updateTaskPriority}
                  onTaskProgressChange={updateTaskProgress}
                  onOpenAttachments={openAttachmentViewer}
                />
                {!filteredRootProjects.length && <p className="text-center text-slate-500 py-12">Aucun projet principal trouvé.</p>}
              </>
            ) : projectViewMode === "table" ? (
              <ProjectTable
                projects={filteredRootProjects}
                tasks={wsTasks}
                allProjects={enrichedProjects}
                onEdit={handleEditProject}
                onEditTask={(task) => openTaskEditor(task)}
                onArchiveProject={(project) => toggleProjectArchived(project.id, true)}
                requestDelete={requestDelete}
                onQuickTask={handleQuickTask}
                onNewSubProject={handleNewSubProject}
                config={data.tableConfig?.projects}
                onConfigChange={(next) => setData((prev) => ({ ...prev, tableConfig: { ...(prev.tableConfig || {}), projects: next } }))}
                customFields={legacyProjectCustomFields}
                onStatusChange={updateProjectStatus}
                onPriorityChange={updateProjectPriority}
                onTaskStatusChange={updateTaskStatus}
                onTaskPriorityChange={updateTaskPriority}
                onTaskProgressChange={updateTaskProgress}
                onProjectFieldChange={updateProjectField}
                onOpenAttachments={openAttachmentViewer}
              />
            ) : (
              <KanbanBoard
                rows={filteredRootProjects}
                statuses={STATUS_OPTIONS}
                onStatusChange={updateProjectStatus}
                renderCard={renderProjectKanbanCard}
                emptyLabel="Aucun projet trouvé."
                celebrateStatus="Terminé"
              />
            )}
          </div>
        )}

        {view === "subprojects" && (
          <div className="mx-auto max-w-7xl p-4 sm:p-6">
            <div className="mb-5 rounded-3xl border border-slate-200 bg-white p-5 shadow-sm">
              <div className="flex items-start justify-between flex-wrap gap-4">
                <div className="min-w-0">
                  <div className="flex items-center gap-3">
                    <span className="inline-flex h-11 w-11 items-center justify-center rounded-2xl bg-violet-600 text-white shadow-sm">
                      <Layers className="w-5 h-5" />
                    </span>
                    <div>
                      <h1 className="text-2xl font-bold text-slate-900">Axes analytiques</h1>
                      <p className="text-slate-500 mt-0.5">{filteredSubProjects.length} axe{filteredSubProjects.length > 1 ? "s transverses" : " transverse"} affiché{filteredSubProjects.length > 1 ? "s" : ""}</p>
                    </div>
                  </div>
                  <div className="mt-4 flex flex-wrap gap-2">
                    <Badge className="bg-violet-50 text-violet-700 border-violet-200">{wsSubProjects.length} total</Badge>
                    <Badge className="bg-blue-50 text-blue-700 border-blue-200">{wsSubProjects.filter((project) => project.status === "En cours").length} en cours</Badge>
                    <Badge className="bg-emerald-50 text-emerald-700 border-emerald-200">{wsSubProjects.filter((project) => project.status === "Terminé").length} terminés</Badge>
                  </div>
                </div>
                <div className="flex w-full items-center gap-2 flex-wrap justify-start xl:w-auto xl:justify-end">
                  <ViewToggle options={projectViewOptions} value={projectViewMode} onChange={setProjectViewMode} />
                  <button onClick={openNewStandaloneSubProjectModal}
                    className="flex items-center gap-2 px-4 py-2 rounded-2xl bg-slate-900 text-white text-sm font-medium">
                    <Plus className="w-4 h-4" /> Nouvel axe analytique
                  </button>
                </div>
              </div>
            </div>
            {FilterBar({ withPriority: false })}
            {projectViewMode === "tree" ? (
              <>
                <ProjectTree
                  nodes={subProjectTree}
                  tasks={wsTasks}
                  projects={enrichedProjects}
                  expanded={expanded}
                  setExpanded={setExpanded}
                  onQuickTask={handleQuickTask}
                  onNewSubProject={handleNewSubProject}
                  onEditProject={handleEditProject}
                  onEditTask={(task) => openTaskEditor(task)}
                  onArchiveProject={(project) => toggleProjectArchived(project.id, true)}
                  requestDelete={requestDelete}
                  onDuplicate={handleDuplicate}
                  projectCustomFields={projectCustomFields}
                  onProjectStatusChange={updateProjectStatus}
                  onProjectPriorityChange={updateProjectPriority}
                  onTaskStatusChange={updateTaskStatus}
                  onTaskPriorityChange={updateTaskPriority}
                  onTaskProgressChange={updateTaskProgress}
                  onOpenAttachments={openAttachmentViewer}
                />
                {!filteredSubProjects.length && <p className="text-center text-slate-500 py-12">Aucun axe analytique trouvé.</p>}
              </>
            ) : projectViewMode === "table" ? (
              <ProjectTable
                projects={filteredSubProjects}
                tasks={wsTasks}
                allProjects={enrichedProjects}
                onEdit={handleEditProject}
                onEditTask={(task) => openTaskEditor(task)}
                onArchiveProject={(project) => toggleProjectArchived(project.id, true)}
                requestDelete={requestDelete}
                onQuickTask={handleQuickTask}
                onNewSubProject={handleNewSubProject}
                config={data.tableConfig?.projects}
                onConfigChange={(next) => setData((prev) => ({ ...prev, tableConfig: { ...(prev.tableConfig || {}), projects: next } }))}
                customFields={legacyProjectCustomFields}
                onStatusChange={updateProjectStatus}
                onPriorityChange={updateProjectPriority}
                onTaskStatusChange={updateTaskStatus}
                onTaskPriorityChange={updateTaskPriority}
                onTaskProgressChange={updateTaskProgress}
                onProjectFieldChange={updateProjectField}
                onOpenAttachments={openAttachmentViewer}
              />
            ) : (
              <KanbanBoard
                rows={filteredSubProjects}
                statuses={STATUS_OPTIONS}
                onStatusChange={updateProjectStatus}
                renderCard={renderProjectKanbanCard}
                emptyLabel="Aucun axe analytique trouvé."
                celebrateStatus="Terminé"
              />
            )}
          </div>
        )}

        {/* ── TASKS ── */}
        {view === "tasks" && (
          <div className="mx-auto max-w-7xl p-4 sm:p-6">
            <PageHeader
              title="Tâches"
              subtitle={`${filteredTasks.length} tâche${filteredTasks.length > 1 ? "s" : ""}`}
              actions={fileActions("tasks", "tâches")}
              onNew={() => { setForm({ status: "À faire", priority: "Moyenne", progress: 0, customFieldValues: {}, attachments: [], whatsAppDraft: { title: "", message: "", url: "" } }); setTaskModal({ open: true, item: null }); }}
              newLabel="Nouvelle tâche"
            />
            <div className="-mt-4 mb-5 flex justify-end">
              <ViewToggle options={tableKanbanOptions} value={taskViewMode} onChange={setTaskViewMode} />
            </div>
            {FilterBar({ withSport: false, withPriority: true })}
            {taskViewMode === "table" ? (
              <EntityTable
                rows={filteredTasks}
                nameKey="title"
                emptyLabel="Aucune tâche trouvée."
                allowRootColumnEditing={false}
                builtinCols={[
                  { key: "title", label: "Tâche", type: "texte", sortable: true,
                    render: (t) => (
                      <button type="button" onClick={() => openTaskEditor(t)} className="flex items-center gap-2 text-left hover:underline">
                        <span className="font-semibold text-slate-900">{t.title}</span>
                        {normalizeAttachments(t.attachments).length > 0 && (
                          <button
                            type="button"
                            onClick={(event) => { event.stopPropagation(); openAttachmentViewer(`Pièces jointes · ${t.title}`, t.attachments); }}
                            className="inline-flex items-center gap-1 rounded-full border border-slate-200 bg-white px-2 py-1 text-[11px] font-semibold text-slate-600 hover:bg-slate-100"
                            title="Voir les pièces jointes"
                          >
                            <Paperclip className="h-3 w-3" /> {normalizeAttachments(t.attachments).length}
                          </button>
                        )}
                      </button>
                    ) },
                  { key: "status", label: "Statut", type: "statut", sortable: true, options: STATUS_OPTIONS },
                  { key: "priority", label: "Priorité", type: "statut", sortable: true, options: PRIORITY_OPTIONS,
                    render: (t) => <Badge className={badgeClass("priority", t.priority)}>{t.priority}</Badge> },
                  { key: "attachmentsCount", label: "PJ", type: "nombre", sortable: true, accessor: (t) => normalizeAttachments(t.attachments).length },
                  { key: "projectId", label: "Projet", type: "texte", sortable: true,
                    accessor: (t) => data.projects.find((p) => p.id === t.projectId)?.name || "—",
                    render: (t) => <span className="text-slate-700">{data.projects.find((p) => p.id === t.projectId)?.name || "—"}</span> },
                  { key: "assignee", label: "Assigné", type: "personne", sortable: true, quickEdit: "text", placeholder: "Assigné à" },
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
                  { key: "startDate", label: "Début", type: "date", sortable: true, quickEdit: "date" },
                  { key: "endDate", label: "Échéance", type: "date", sortable: true, quickEdit: "date" },
                ]}
                config={data.tableConfig?.tasks}
                onConfigChange={(next) => setData((prev) => ({ ...prev, tableConfig: { ...(prev.tableConfig || {}), tasks: next } }))}
                onEdit={(task) => openTaskEditor(task)}
                onArchive={(task) => toggleTaskArchived(task.id, true)}
                onDelete={(task) => requestDelete("task", task)}
                customFields={legacyTaskCustomFields}
                onStatusChange={updateTaskStatus}
                onPriorityChange={updateTaskPriority}
                onProgressChange={updateTaskProgress}
                onFieldChange={updateTaskField}
                onOpenAttachments={openAttachmentViewer}
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
          <div className="mx-auto max-w-7xl p-4 sm:p-6">
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
            {FilterBar({ withSport: false, statusOptions: ATHLETE_STATUS })}
            {athleteViewMode === "table" ? (
              <EntityTable
                rows={filteredAthletes}
                nameKey="name"
                emptyLabel="Aucun athlète trouvé."
                allowRootColumnEditing={false}
                builtinCols={[
                  { key: "name", label: "Athlète", type: "texte", sortable: true },
                  { key: "status", label: "Statut", type: "statut", sortable: true, options: ATHLETE_STATUS,
                    badgeTone: (a) => a.status === "Actif" ? "En cours" : a.status === "Blessé" ? "En attente" : "À faire" },
                  { key: "sectionId", label: "Section", type: "texte", sortable: true,
                    accessor: (a) => sectionName(a.sectionId),
                    quickEdit: "select",
                    editOptions: () => wsClubs.map((section) => ({ value: section.id, label: section.name })) },
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
                onStatusChange={(id, status) => updateEntityStatus("athletes", id, status)}
                onFieldChange={updateAthleteField}
                onOpenAttachments={openAttachmentViewer}
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
          <div className="mx-auto max-w-7xl p-4 sm:p-6">
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
            {FilterBar({ withSport: false, statusOptions: CLUB_STATUS })}
            {clubViewMode === "table" ? (
              <EntityTable
                rows={filteredClubs}
                nameKey="name"
                emptyLabel="Aucune section trouvée."
                allowRootColumnEditing={false}
                builtinCols={[
                  { key: "name", label: "Section", type: "texte", sortable: true },
                  { key: "status", label: "Statut", type: "statut", sortable: true, options: CLUB_STATUS,
                    badgeTone: (c) => c.status === "Validé" ? "Terminé" : c.status === "Inscrit" ? "En cours" : c.status === "Invité" ? "En attente" : "À faire" },
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
                onStatusChange={(id, status) => updateEntityStatus("clubs", id, status)}
                onOpenAttachments={openAttachmentViewer}
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
          <div className="mx-auto max-w-7xl p-4 sm:p-6">
            <PageHeader
              title="Documents"
              subtitle={`${filteredDocs.length} document${filteredDocs.length > 1 ? "s" : ""}`}
              onNew={() => { setForm({ sport: "developpe_couche", type: "Suivi sportif", status: "Brouillon", url: "", content: "" }); setDocModal({ open: true, item: null }); }}
              newLabel="Nouveau document"
            />
            <div className="-mt-4 mb-5 flex justify-end">
              <ViewToggle options={tableKanbanOptions} value={documentViewMode} onChange={setDocumentViewMode} />
            </div>
            <div className="mb-5 flex flex-wrap gap-3">
              <div className="relative w-full sm:w-auto sm:min-w-[260px]">
                <Search className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
                <input className="w-full rounded-2xl border bg-white py-2 pl-9 pr-4 text-sm sm:min-w-[260px]" placeholder="Rechercher..." value={searchQ} onChange={(e) => setSearchQ(e.target.value)} />
              </div>
              <select className="w-full rounded-2xl border px-3 py-2 text-sm bg-white sm:w-auto" value={filterSport} onChange={(e) => setFilterSport(e.target.value)}>
                <option value="all">Tous les sports</option>
                {allSports.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
              </select>
              <select className="w-full rounded-2xl border px-3 py-2 text-sm bg-white sm:w-auto" value={filterStatus} onChange={(e) => setFilterStatus(e.target.value)}>
                <option value="all">Tous les statuts</option>
                {DOC_STATUS.map((s) => <option key={s} value={s}>{s}</option>)}
              </select>
              {DOC_MODES.map((m) => (
                <button key={m} onClick={() => setDocMode(m)}
                  className={`w-full rounded-2xl border px-3 py-2 text-sm sm:w-auto ${docMode === m ? "bg-slate-900 text-white border-slate-900" : "bg-white"}`}>
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
                    badgeTone: (d) => d.status === "Publié" || d.status === "Validé" ? "Terminé" : d.status === "À valider" ? "En attente" : d.status === "Archivé" ? "À faire" : "En cours" },
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
                onStatusChange={(id, status) => updateEntityStatus("documents", id, status)}
                onOpenAttachments={openAttachmentViewer}
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
          <div className="mx-auto max-w-7xl p-4 sm:p-6">
            <PageHeader
              title="Convocations"
              subtitle={`${filteredConvocations.length} convocation${filteredConvocations.length > 1 ? "s" : ""}`}
              onNew={() => { setForm({ sport: "football", type: "Match", status: "Brouillon", workspaceId: currentWsId }); setConvModal({ open: true, item: null }); }}
              newLabel="Nouvelle convocation"
            />
            <div className="-mt-4 mb-5 flex justify-end">
              <ViewToggle options={listKanbanOptions} value={convocationViewMode} onChange={setConvocationViewMode} />
            </div>
            {FilterBar({ statusOptions: COMM_STATUS })}
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
          <div className="mx-auto max-w-7xl p-4 sm:p-6">
            <PageHeader
              title="Notifications"
              subtitle={`${filteredNotifications.length} notification${filteredNotifications.length > 1 ? "s" : ""}`}
              onNew={() => { setForm({ sport: "developpe_couche", channel: "Email", status: "Brouillon", workspaceId: currentWsId }); setNotifModal({ open: true, item: null }); }}
              newLabel="Nouvelle notification"
            />
            <div className="-mt-4 mb-5 flex justify-end">
              <ViewToggle options={listKanbanOptions} value={notificationViewMode} onChange={setNotificationViewMode} />
            </div>
            {FilterBar({ statusOptions: COMM_STATUS })}
            {notificationViewMode === "list" ? (
              <div className="space-y-3">
                {filteredNotifications.map((notif) => (
                  <Card key={notif.id} className="p-4">
                    <div className="flex items-start justify-between gap-4">
                      <div className="flex-1 min-w-0">
                        <div className="flex flex-wrap items-center gap-2 mb-1">
                          <h3 className="font-semibold">{notif.title}</h3>
                          <Badge className={badgeClass("status", notif.status === "Envoyée" ? "Terminé" : notif.status === "Prêt à envoyer" ? "En cours" : "À faire")}>{notif.status}</Badge>
                          {notif.systemGenerated && <Badge className="bg-amber-50 text-amber-700 border-amber-200">Alerte automatique</Badge>}
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
                      {!notif.systemGenerated && (
                        <ActionBtns
                          onEdit={() => { setForm({ ...notif }); setNotifModal({ open: true, item: notif }); }}
                          onDelete={() => requestDelete("notification", notif)}
                        />
                      )}
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

        {view === "archives" && (
          <div className="mx-auto max-w-7xl p-4 sm:p-6">
            <PageHeader
              title="Archives"
              subtitle="Consulte, modifie et réactive les projets et tâches sortis du flux principal."
            />
            <div className="mb-5 flex flex-wrap items-center gap-2">
              {[
                { id: "projects", label: `Projets archivés (${archivedProjects.length})` },
                { id: "tasks", label: `Tâches archivées (${archivedTasks.length})` },
              ].map((option) => (
                <button
                  key={option.id}
                  type="button"
                  onClick={() => setArchiveEntity(option.id)}
                  className={`rounded-2xl border px-4 py-2 text-sm font-semibold ${archiveEntity === option.id ? "border-slate-900 bg-slate-900 text-white" : "bg-white text-slate-700 hover:bg-slate-50"}`}
                >
                  {option.label}
                </button>
              ))}
            </div>

            {archiveEntity === "projects" ? (
              <EntityTable
                rows={archivedProjects}
                nameKey="name"
                emptyLabel="Aucun projet archivé."
                builtinCols={[
                  { key: "name", label: "Projet", type: "texte", sortable: true },
                  { key: "sportName", label: "Sport", type: "texte", sortable: true, accessor: (project) => sportName(project.sport) },
                  { key: "status", label: "Statut", type: "statut", sortable: true, options: STATUS_OPTIONS },
                  { key: "owner", label: "Responsable", type: "personne", sortable: true },
                  { key: "visibility", label: "Visible", type: "statut", sortable: true, options: ["Oui", "Non"] },
                  { key: "endDate", label: "Échéance", type: "date", sortable: true },
                  { key: "archivedAt", label: "Archivé le", type: "date", sortable: true },
                ]}
                config={data.tableConfig?.archivedProjects}
                onConfigChange={(next) => setData((prev) => ({ ...prev, tableConfig: { ...(prev.tableConfig || {}), archivedProjects: next } }))}
                onEdit={handleEditProject}
                onUnarchive={(project) => toggleProjectArchived(project.id, false)}
                onDelete={(project) => requestDelete("project", project)}
                onStatusChange={updateProjectStatus}
                onFieldChange={updateProjectField}
                onOpenAttachments={openAttachmentViewer}
              />
            ) : (
              <EntityTable
                rows={archivedTasks}
                nameKey="title"
                emptyLabel="Aucune tâche archivée."
                builtinCols={[
                  { key: "title", label: "Tâche", type: "texte", sortable: true,
                    render: (task) => (
                      <button type="button" onClick={() => openTaskEditor(task)} className="flex items-center gap-2 text-left hover:underline">
                        <span className="font-semibold text-slate-900">{task.title}</span>
                        {normalizeAttachments(task.attachments).length > 0 && (
                          <button
                            type="button"
                            onClick={(event) => { event.stopPropagation(); openAttachmentViewer(`Pièces jointes · ${task.title}`, task.attachments); }}
                            className="inline-flex items-center gap-1 rounded-full border border-slate-200 bg-white px-2 py-1 text-[11px] font-semibold text-slate-600 hover:bg-slate-100"
                            title="Voir les pièces jointes"
                          >
                            <Paperclip className="h-3 w-3" /> {normalizeAttachments(task.attachments).length}
                          </button>
                        )}
                      </button>
                    ) },
                  { key: "status", label: "Statut", type: "statut", sortable: true, options: STATUS_OPTIONS },
                  { key: "priority", label: "Priorité", type: "statut", sortable: true, options: PRIORITY_OPTIONS },
                  { key: "attachmentsCount", label: "PJ", type: "nombre", sortable: true, accessor: (task) => normalizeAttachments(task.attachments).length },
                  { key: "projectId", label: "Projet", type: "texte", sortable: true, accessor: (task) => workspaceProjects.find((project) => project.id === task.projectId)?.name || "—" },
                  { key: "assignee", label: "Assigné", type: "personne", sortable: true },
                  { key: "progress", label: "Avancement", type: "nombre", sortable: true, accessor: (task) => normalizeTaskProgress(task.progress, task.status) },
                  { key: "archivedAt", label: "Archivé le", type: "date", sortable: true },
                ]}
                config={data.tableConfig?.archivedTasks}
                onConfigChange={(next) => setData((prev) => ({ ...prev, tableConfig: { ...(prev.tableConfig || {}), archivedTasks: next } }))}
                onEdit={(task) => openTaskEditor(task)}
                onUnarchive={(task) => toggleTaskArchived(task.id, false)}
                onDelete={(task) => requestDelete("task", task)}
                onStatusChange={updateTaskStatus}
                onPriorityChange={updateTaskPriority}
                onProgressChange={updateTaskProgress}
                onFieldChange={updateTaskField}
                onOpenAttachments={openAttachmentViewer}
              />
            )}
          </div>
        )}

        {/* ── SETTINGS ── */}
        {view === "settings" && (
          <div className="p-6 max-w-4xl mx-auto">
            <PageHeader title="Paramètres" />

            <Card className="p-6 mb-6">
              <div className="flex flex-col gap-4 md:flex-row md:items-start md:justify-between">
                <div>
                  <h2 className="font-semibold text-lg">Utilisateurs et rôles</h2>
                  <p className="mt-1 text-sm text-slate-600">
                    Cette première version permet de simuler qui est connecté dans l'application et de rendre un projet confidentiel.
                    Un projet avec <span className="font-semibold">Visible par tous : Non</span> est accessible seulement aux administrateurs, au responsable du projet et aux utilisateurs autorisés.
                  </p>
                </div>
                <div className="rounded-2xl border bg-slate-50 px-4 py-3 text-sm">
                  <p className="font-semibold text-slate-900">Session active</p>
                  <p className="mt-1 text-slate-600">{currentUser?.name} · {currentUser?.role}</p>
                </div>
              </div>

              <div className="mt-5 grid gap-3">
                {workspaceUsers.map((user) => {
                  const isCurrent = user.id === currentUser?.id;
                  return (
                    <div key={user.id} className={`rounded-2xl border px-4 py-4 ${isCurrent ? "border-slate-900 bg-slate-50" : "border-slate-200 bg-white"}`}>
                      <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
                        <div>
                          <p className="font-semibold text-slate-900">{user.name}</p>
                          <p className="text-sm text-slate-500">{user.email || "Aucun email renseigné"}</p>
                        </div>
                        <div className="flex flex-col gap-2 md:flex-row md:items-center">
                          <button
                            type="button"
                            onClick={() => switchCurrentUser(user.id)}
                            className={`rounded-2xl px-4 py-2 text-sm font-semibold ${isCurrent ? "bg-slate-900 text-white" : "border border-slate-200 bg-white text-slate-700 hover:bg-slate-100"}`}
                          >
                            {isCurrent ? "Utilisateur connecté" : "Se connecter en test"}
                          </button>
                          <select
                            className="rounded-2xl border px-4 py-2 text-sm"
                            value={user.role || "Utilisateur"}
                            disabled={!isAdmin}
                            onChange={(event) => updateUserRole(user.id, event.target.value)}
                          >
                            {USER_ROLE_OPTIONS.map((role) => <option key={role} value={role}>{role}</option>)}
                          </select>
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>

              {!isAdmin && (
                <p className="mt-4 text-xs text-slate-500">
                  Seul un administrateur peut changer les rôles. En mode test, tu peux quand même basculer d'un utilisateur à l'autre pour vérifier la confidentialité des projets.
                </p>
              )}
            </Card>

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

      <datalist id="people-options">
        {peopleOptions.map((person) => <option key={person} value={person} />)}
      </datalist>

      <AttachmentViewerModal
        open={attachmentViewer.open}
        title={attachmentViewer.title}
        attachments={attachmentViewer.attachments}
        onClose={() => setAttachmentViewer({ open: false, title: "", attachments: [] })}
      />

      {/* Project */}
      <Modal open={projectModal.open} title={projectModal.item ? "Modifier le projet" : "Nouveau projet"}
        onClose={() => { setProjectModal({ open: false, item: null }); setForm({}); }} onSave={saveProject}>
        <div className="grid md:grid-cols-2 gap-4">
          <input className="w-full rounded-2xl border px-4 py-3" placeholder="Nom du projet *" value={form.name || ""} onChange={(e) => f("name", e.target.value)} />
          <input list="people-options" className="w-full rounded-2xl border px-4 py-3" placeholder="Responsable" value={form.owner || ""} onChange={(e) => f("owner", e.target.value)} />
          <textarea className="w-full rounded-2xl border px-4 py-3 md:col-span-2 resize-none h-20" placeholder="Description" value={form.description || ""} onChange={(e) => f("description", e.target.value)} />
          <div className="md:col-span-2 rounded-2xl border bg-slate-50 p-4 space-y-3">
            <div className="flex items-center justify-between gap-3 flex-wrap">
              <p className="text-xs font-bold uppercase tracking-wide text-slate-500">
                {form.parentId ? "Sports concernés par l'axe analytique" : "Sport du projet"}
              </p>
              {form.parentId && (
                <div className="flex gap-2">
                  <button type="button" onClick={() => setForm((prev) => ({ ...prev, sports: allSports.map((sport) => sport.id) }))} className="rounded-xl border bg-white px-3 py-1.5 text-xs font-semibold text-slate-600 hover:bg-slate-100">Tous</button>
                  <button type="button" onClick={() => setForm((prev) => ({ ...prev, sports: [] }))} className="rounded-xl border bg-white px-3 py-1.5 text-xs font-semibold text-slate-600 hover:bg-slate-100">Aucun</button>
                </div>
              )}
            </div>

            {form.parentId ? (
              <div className="grid gap-2 md:grid-cols-2 xl:grid-cols-3">
                {allSports.map((sport) => {
                  const checked = (form.sports || []).includes(sport.id);
                  return (
                    <label key={sport.id} className={`flex items-center gap-3 rounded-2xl border px-3 py-3 text-sm font-medium cursor-pointer transition ${checked ? "border-slate-900 bg-white text-slate-900" : "bg-white text-slate-600 hover:bg-slate-100"}`}>
                      <input
                        type="checkbox"
                        checked={checked}
                        onChange={() => toggleFormSport(sport.id)}
                        className="h-4 w-4 rounded border-slate-300"
                      />
                      <span>{sport.name}</span>
                    </label>
                  );
                })}
              </div>
            ) : (
              <select className="w-full rounded-2xl border bg-white px-4 py-3" value={form.sport || ""} onChange={(e) => f("sport", e.target.value)}>
                <option value="">Sport</option>
                {allSports.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
              </select>
            )}

            <div className="flex gap-2">
              <input
                className="flex-1 rounded-2xl border bg-white px-4 py-3 text-sm"
                placeholder="Créer un nouveau sport"
                value={form.newSportName || ""}
                onChange={(e) => f("newSportName", e.target.value)}
                onKeyDown={(e) => e.key === "Enter" && (e.preventDefault(), addInlineSport(form.parentId ? "sports" : "sport", Boolean(form.parentId)))}
              />
              <button
                type="button"
                onClick={() => addInlineSport(form.parentId ? "sports" : "sport", Boolean(form.parentId))}
                className="rounded-2xl bg-slate-900 px-4 py-3 text-sm font-semibold text-white"
              >
                Ajouter sport
              </button>
            </div>
          </div>
          <ProjectCoverManager
            coverImage={form.coverImage || ""}
            onAddImage={addProjectCoverImage}
            onRemoveImage={() => f("coverImage", "")}
          />
          <div className="space-y-2 md:col-span-2 rounded-2xl border bg-slate-50 p-3">
            <p className="text-xs font-bold uppercase tracking-wide text-slate-500">Organisation du projet</p>
            <div className="grid grid-cols-2 gap-2">
              <button
                type="button"
                onClick={() => f("parentId", null)}
                className={`rounded-2xl border px-3 py-3 text-left text-sm font-semibold ${!form.parentId ? "border-slate-900 bg-slate-900 text-white" : "bg-white hover:bg-slate-50"}`}
              >
                <span className="block">Projet principal</span>
                <span className={`mt-1 block text-xs font-normal ${!form.parentId ? "text-slate-200" : "text-slate-500"}`}>Projet autonome</span>
              </button>
              <button
                type="button"
                disabled={!hasAvailableProjectParent}
                onClick={() => f("parentId", form.parentId || projectParentChoices[0]?.id || null)}
                className={`rounded-2xl border px-3 py-3 text-left text-sm font-semibold disabled:cursor-not-allowed disabled:opacity-40 ${form.parentId ? "border-slate-900 bg-slate-900 text-white" : "bg-white hover:bg-slate-50"}`}
              >
                <span className="block">Axe analytique</span>
                <span className={`mt-1 block text-xs font-normal ${form.parentId ? "text-slate-200" : "text-slate-500"}`}>Axe transverse de pilotage</span>
              </button>
            </div>
            {form.parentId && (
              <select className="w-full rounded-2xl border px-4 py-3" value={form.parentId || ""} onChange={(e) => f("parentId", e.target.value || null)}>
                {projectParentChoices.map((p) => <option key={p.id} value={p.id}>Rattaché au projet : {p.name}</option>)}
              </select>
            )}
            {!hasAvailableProjectParent && <p className="text-xs text-slate-500">Crée d'abord un projet principal pour pouvoir créer un axe analytique.</p>}
          </div>
          {!form.parentId && (
            <div className="space-y-3 md:col-span-2 rounded-2xl border bg-violet-50/60 p-4">
              <div>
                <p className="text-xs font-bold uppercase tracking-wide text-violet-700">Axes analytiques à rattacher</p>
                <p className="mt-1 text-sm text-slate-600">Tu peux préparer ici les axes transverses de pilotage qui seront créés automatiquement avec ce projet.</p>
              </div>
              <div className="flex gap-2">
                <input
                  className="flex-1 rounded-2xl border bg-white px-4 py-3 text-sm"
                  placeholder="Ex : Comptabilité, RH, Logistique..."
                  value={form.newAxisName || ""}
                  onChange={(e) => f("newAxisName", e.target.value)}
                  onKeyDown={(e) => e.key === "Enter" && (e.preventDefault(), addPendingAxis())}
                />
                <button
                  type="button"
                  onClick={addPendingAxis}
                  className="rounded-2xl bg-violet-700 px-4 py-3 text-sm font-semibold text-white"
                >
                  Ajouter axe
                </button>
              </div>
              {(form.pendingAxes || []).length > 0 && (
                <div className="flex flex-wrap gap-2">
                  {form.pendingAxes.map((axisName) => (
                    <span key={axisName} className="inline-flex items-center gap-2 rounded-full border border-violet-200 bg-white px-3 py-1.5 text-sm font-medium text-violet-700">
                      {axisName}
                      <button
                        type="button"
                        onClick={() => removePendingAxis(axisName)}
                        className="rounded-full p-0.5 text-violet-500 hover:bg-violet-100 hover:text-violet-700"
                        title={`Retirer ${axisName}`}
                      >
                        <X className="h-3.5 w-3.5" />
                      </button>
                    </span>
                  ))}
                </div>
              )}
            </div>
          )}
          <select className="w-full rounded-2xl border px-4 py-3" value={form.status || "À faire"} onChange={(e) => f("status", e.target.value)}>
            {STATUS_OPTIONS.map((s) => <option key={s} value={s}>{s}</option>)}
          </select>
          <select className="w-full rounded-2xl border px-4 py-3" value={form.priority || "Moyenne"} onChange={(e) => f("priority", e.target.value)}>
            {PRIORITY_OPTIONS.map((p) => <option key={p} value={p}>{p}</option>)}
          </select>
          <select className="w-full rounded-2xl border px-4 py-3" value={form.visibility || "Oui"} onChange={(e) => f("visibility", e.target.value)}>
            <option value="Oui">Visible par tous : Oui</option>
            <option value="Non">Projet confidentiel : Non</option>
          </select>
          {form.visibility === "Non" && (
            <div className="md:col-span-2 rounded-2xl border border-slate-200 bg-slate-50 p-4">
              <p className="text-sm font-semibold text-slate-900">Utilisateurs autorisés</p>
              <p className="mt-1 text-sm text-slate-500">
                Les administrateurs voient toujours le projet. Le responsable du projet garde aussi l'accès automatiquement.
              </p>
              <div className="mt-3 grid gap-2 md:grid-cols-2">
                {workspaceUsers.map((user) => {
                  const checked = normalizeProjectAccess(form, currentUser?.id).includes(user.id);
                  return (
                    <label key={user.id} className={`flex items-center justify-between gap-3 rounded-2xl border px-3 py-3 text-sm font-medium transition ${checked ? "border-slate-900 bg-white text-slate-900" : "border-slate-200 bg-white text-slate-600 hover:bg-slate-100"}`}>
                      <span>
                        {user.name}
                        <span className="ml-2 text-xs text-slate-400">{user.role}</span>
                      </span>
                      <input
                        type="checkbox"
                        checked={checked}
                        onChange={() => toggleFormAllowedUser(user.id)}
                        className="h-4 w-4 rounded border-slate-300"
                      />
                    </label>
                  );
                })}
              </div>
            </div>
          )}
          <input type="date" className="w-full rounded-2xl border px-4 py-3" value={form.startDate || ""} onChange={(e) => f("startDate", e.target.value)} />
          <input type="date" className="w-full rounded-2xl border px-4 py-3" value={form.endDate || ""} onChange={(e) => f("endDate", e.target.value)} />
          <AttachmentManager
            attachments={normalizeAttachments(form.attachments)}
            onAddFiles={addAttachmentFiles}
            onAddWhatsApp={addWhatsAppAttachment}
            onRemove={removeAttachment}
            whatsAppDraft={form.whatsAppDraft}
            onWhatsAppDraftChange={updateWhatsAppDraft}
          />
        </div>
        <div className="mt-4 rounded-2xl border border-slate-200 bg-slate-50 p-4">
          <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
            <div>
              <p className="text-sm font-semibold text-slate-900">Structure du tableau Projets</p>
              <p className="mt-1 text-sm text-slate-500">Ajoute ici une colonne qui sera disponible sur tous les projets de l'application.</p>
            </div>
            <button
              type="button"
              onClick={() => openScopedCustomFieldModal("projects")}
              className="inline-flex w-full items-center justify-center gap-2 rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm font-semibold text-slate-700 hover:bg-slate-100 md:w-auto"
            >
              <Plus className="h-4 w-4" />
              Ajouter une colonne au tableau Projets
            </button>
          </div>
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
          <input list="people-options" className="w-full rounded-2xl border px-4 py-3" placeholder="Assigné à" value={form.assignee || ""} onChange={(e) => f("assignee", e.target.value)} />
          <select className="w-full rounded-2xl border px-4 py-3" value={form.projectId || ""} onChange={(e) => {
            const projectId = e.target.value;
            const project = workspaceProjects.find((item) => item.id === projectId);
            setForm((prev) => ({ ...prev, projectId, sport: project?.sport || prev.sport || "" }));
          }}>
            <option value="">Projet</option>
            {wsProjects.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
          </select>
          <div className="md:col-span-2 rounded-2xl border bg-slate-50 p-4 space-y-3">
            <p className="text-xs font-bold uppercase tracking-wide text-slate-500">Sport lié à la tâche</p>
            <div className="grid gap-2 md:grid-cols-[1fr_auto]">
              <select className="w-full rounded-2xl border bg-white px-4 py-3" value={form.sport || ""} onChange={(e) => f("sport", e.target.value)}>
                <option value="">Sport</option>
                {allSports.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
              </select>
              <div className="flex gap-2 md:min-w-[320px]">
                <input
                  className="flex-1 rounded-2xl border bg-white px-4 py-3 text-sm"
                  placeholder="Créer un nouveau sport"
                  value={form.newSportName || ""}
                  onChange={(e) => f("newSportName", e.target.value)}
                  onKeyDown={(e) => e.key === "Enter" && (e.preventDefault(), addInlineSport("sport", false))}
                />
                <button
                  type="button"
                  onClick={() => addInlineSport("sport", false)}
                  className="rounded-2xl bg-slate-900 px-4 py-3 text-sm font-semibold text-white"
                >
                  Ajouter sport
                </button>
              </div>
            </div>
          </div>
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
            <div className="flex items-center gap-3">
              <input type="range" min="0" max="100" className="w-full accent-slate-900" value={form.progress || 0} onChange={(e) => setTaskFormProgress(Number(e.target.value))} />
              <div className="flex items-center gap-2">
                <input
                  type="number"
                  min="0"
                  max="100"
                  className="w-20 rounded-2xl border px-3 py-2 text-right text-sm"
                  value={form.progress || 0}
                  onChange={(e) => setTaskFormProgress(Number(e.target.value))}
                />
                <span className="text-sm font-semibold text-slate-500">%</span>
              </div>
            </div>
          </div>
          <AttachmentManager
            attachments={normalizeAttachments(form.attachments)}
            onAddFiles={addAttachmentFiles}
            onAddWhatsApp={addWhatsAppAttachment}
            onRemove={removeAttachment}
            whatsAppDraft={form.whatsAppDraft}
            onWhatsAppDraftChange={updateWhatsAppDraft}
          />
        </div>
        <div className="mt-4 rounded-2xl border border-slate-200 bg-slate-50 p-4">
          <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
            <div>
              <p className="text-sm font-semibold text-slate-900">Structure du tableau Tâches</p>
              <p className="mt-1 text-sm text-slate-500">Ajoute ici une colonne qui sera disponible sur toutes les tâches de l'application.</p>
            </div>
            <button
              type="button"
              onClick={() => openScopedCustomFieldModal("tasks")}
              className="inline-flex w-full items-center justify-center gap-2 rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm font-semibold text-slate-700 hover:bg-slate-100 md:w-auto"
            >
              <Plus className="h-4 w-4" />
              Ajouter une colonne au tableau Tâches
            </button>
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
        <div className="mt-4 rounded-2xl border border-slate-200 bg-slate-50 p-4">
          <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
            <div>
              <p className="text-sm font-semibold text-slate-900">Structure du tableau Athlètes</p>
              <p className="mt-1 text-sm text-slate-500">Ajoute ici une colonne qui sera disponible sur tous les athlètes de l'application.</p>
            </div>
            <button
              type="button"
              onClick={() => openScopedCustomFieldModal("athletes")}
              className="inline-flex w-full items-center justify-center gap-2 rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm font-semibold text-slate-700 hover:bg-slate-100 md:w-auto"
            >
              <Plus className="h-4 w-4" />
              Ajouter une colonne au tableau Athlètes
            </button>
          </div>
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
          <input list="people-options" className="w-full rounded-2xl border px-4 py-3" placeholder="Responsable de section" value={form.contactName || ""} onChange={(e) => f("contactName", e.target.value)} />
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
        <div className="mt-4 rounded-2xl border border-slate-200 bg-slate-50 p-4">
          <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
            <div>
              <p className="text-sm font-semibold text-slate-900">Structure du tableau Sections</p>
              <p className="mt-1 text-sm text-slate-500">Ajoute ici une colonne qui sera disponible sur toutes les sections de l'application.</p>
            </div>
            <button
              type="button"
              onClick={() => openScopedCustomFieldModal("clubs")}
              className="inline-flex w-full items-center justify-center gap-2 rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm font-semibold text-slate-700 hover:bg-slate-100 md:w-auto"
            >
              <Plus className="h-4 w-4" />
              Ajouter une colonne au tableau Sections
            </button>
          </div>
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
        onClose={() => { setCfModal({ open: false, item: null }); setCfForm({}); }} onSave={saveCF}>
        <div className="grid md:grid-cols-2 gap-4">
          <div className="md:col-span-2">
            <input
              className="w-full rounded-2xl border px-4 py-3"
              placeholder={(ADD_COLUMN_HINTS[cfForm.type || "texte"] || ADD_COLUMN_HINTS.texte).placeholder}
              value={cfForm.name || ""}
              onChange={(e) => setCfForm((prev) => ({ ...prev, name: e.target.value }))}
            />
            <p className="mt-1.5 text-xs text-slate-500">{(ADD_COLUMN_HINTS[cfForm.type || "texte"] || ADD_COLUMN_HINTS.texte).helper}</p>
          </div>
          <select className="w-full rounded-2xl border px-4 py-3" value={cfForm.type || "texte"} onChange={(e) => setCfForm((prev) => ({ ...prev, type: e.target.value }))}>
            {CUSTOM_FIELD_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}
          </select>
          <div className="flex flex-col gap-2 justify-center md:col-span-2">
            <p className="text-sm text-slate-700 font-medium">Appliquer à :</p>
            <div className="grid gap-2 sm:grid-cols-2">
            {CUSTOM_FIELD_TARGETS.map((target) => (
              <label key={target.id} className="flex items-center gap-2 text-sm cursor-pointer">
                <input type="checkbox"
                  checked={(cfForm.appliesTo || []).includes(target.id)}
                  onChange={(e) => {
                    const cur = cfForm.appliesTo || [];
                    setCfForm((prev) => ({
                      ...prev,
                      appliesTo: e.target.checked ? [...cur, target.id] : cur.filter((x) => x !== target.id),
                    }));
                  }} />
                {target.label}
              </label>
            ))}
            </div>
          </div>
          {(cfForm.type || "texte") === "case_a_cocher" && (
            <>
              <label className="block">
                <span className="mb-1.5 block text-xs font-bold uppercase tracking-wide text-slate-500">Texte si coché</span>
                <input
                  className="w-full rounded-2xl border px-4 py-3"
                  placeholder="Ex : Oui"
                  value={cfForm.checkedLabel || "Oui"}
                  onChange={(e) => setCfForm((prev) => ({ ...prev, checkedLabel: e.target.value }))}
                />
              </label>
              <label className="block">
                <span className="mb-1.5 block text-xs font-bold uppercase tracking-wide text-slate-500">Texte si non coché</span>
                <input
                  className="w-full rounded-2xl border px-4 py-3"
                  placeholder="Ex : Non"
                  value={cfForm.uncheckedLabel || "Non"}
                  onChange={(e) => setCfForm((prev) => ({ ...prev, uncheckedLabel: e.target.value }))}
                />
              </label>
            </>
          )}
          {(cfForm.type || "texte") === "personne" && (
            <label className="block md:col-span-2">
              <span className="mb-1.5 block text-xs font-bold uppercase tracking-wide text-slate-500">Texte d'aide dans la cellule</span>
              <input
                className="w-full rounded-2xl border px-4 py-3"
                placeholder="Ex : Choisir un référent"
                value={cfForm.placeholder || ""}
                onChange={(e) => setCfForm((prev) => ({ ...prev, placeholder: e.target.value }))}
              />
            </label>
          )}
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
