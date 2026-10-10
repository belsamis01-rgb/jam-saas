// Synchronisation JAM SaaS <-> Supabase (connexion email / mot de passe)
// Sans dépendance externe : appels REST directs vers Supabase Auth et PostgREST.

const SUPABASE_URL = import.meta.env.VITE_SUPABASE_URL || "https://agcoulfqoodzllvxrjfg.supabase.co";
// Clé publiable : prévue pour le navigateur, l'accès aux données est protégé par les règles RLS.
const SUPABASE_KEY = import.meta.env.VITE_SUPABASE_PUBLISHABLE_KEY || "sb_publishable_Fvsb2ALt7jkzf8jjdNmD9A_--6MoKfH";

export const LOCAL_DATA_KEY = "jam_saas_v19_clean";
const SESSION_KEY = "jam_saas_cloud_session";
const BACKUP_KEY = "jam_saas_local_backup";
const UNSYNCED_KEY = "jam_saas_unsynced";
const UNSYNCED_CONFLICT_KEY = "jam_saas_unsynced_conflict";
const ROW_ID = "jam-test";
const TABLE_URL = `${SUPABASE_URL}/rest/v1/jam_shared_state`;
// Champs propres à chaque appareil : jamais envoyés dans la base partagée.
const DEVICE_KEYS = ["currentUserId", "currentWorkspaceId"];
const SAVE_DELAY_MS = 1200;
// Migration unique validée par Samy le 10/10/2026 : tant que la base contient encore la version de juillet
// (version 33 enregistrée par Thomas), le premier navigateur connecté qui a des projets envoie ses données,
// qui deviennent la version partagée. La version passe alors à 34 et cette règle ne s'applique plus.
const SEED_FROM_LOCAL = { version: 33, updatedBy: "Thomas" };

/* ---------- Stockage local sûr ---------- */
function readLocal(key) {
  try {
    const raw = localStorage.getItem(key);
    return raw ? JSON.parse(raw) : null;
  } catch {
    return null;
  }
}
function writeLocal(key, value) {
  try {
    localStorage.setItem(key, JSON.stringify(value));
  } catch {
    /* stockage indisponible : on continue sans planter */
  }
}
function removeLocal(key) {
  try {
    localStorage.removeItem(key);
  } catch {
    /* rien */
  }
}

/* ---------- Authentification ---------- */
function frenchAuthError(payload, status) {
  const code = payload?.error_code || payload?.error || "";
  const msg = String(payload?.msg || payload?.error_description || payload?.message || "");
  if (code === "invalid_credentials" || /invalid login credentials/i.test(msg)) return "Email ou mot de passe incorrect.";
  if (code === "email_not_confirmed" || /not confirmed/i.test(msg)) return "Ce compte n'est pas encore confirmé.";
  if (status === 429) return "Trop de tentatives. Réessaie dans quelques minutes.";
  if (status >= 500) return "Le serveur Supabase ne répond pas. Réessaie dans un instant.";
  return msg || "Connexion impossible.";
}

function storeSession(payload) {
  const expiresAt = payload.expires_at || Math.floor(Date.now() / 1000) + (Number(payload.expires_in) || 3600);
  const session = {
    access_token: payload.access_token,
    refresh_token: payload.refresh_token,
    expires_at: expiresAt,
    email: payload.user?.email || "",
  };
  writeLocal(SESSION_KEY, session);
  return session;
}

export async function signIn(email, password) {
  const cleanEmail = String(email || "").trim();
  if (!cleanEmail || !password) throw new Error("Renseigne ton email et ton mot de passe.");
  let res;
  try {
    res = await fetch(`${SUPABASE_URL}/auth/v1/token?grant_type=password`, {
      method: "POST",
      headers: { apikey: SUPABASE_KEY, "Content-Type": "application/json" },
      body: JSON.stringify({ email: cleanEmail, password }),
    });
  } catch {
    throw new Error("Pas de connexion internet ou serveur injoignable.");
  }
  const payload = await res.json().catch(() => ({}));
  if (!res.ok || !payload.access_token) throw new Error(frenchAuthError(payload, res.status));
  return storeSession(payload);
}

/* ---------- Mot de passe oublié ---------- */
export async function requestPasswordReset(email) {
  const cleanEmail = String(email || "").trim();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(cleanEmail)) throw new Error("Renseigne une adresse email valide.");
  const redirectTo = `${window.location.origin}/`;
  let res;
  try {
    res = await fetch(`${SUPABASE_URL}/auth/v1/recover?redirect_to=${encodeURIComponent(redirectTo)}`, {
      method: "POST",
      headers: { apikey: SUPABASE_KEY, "Content-Type": "application/json" },
      body: JSON.stringify({ email: cleanEmail }),
    });
  } catch {
    throw new Error("Pas de connexion internet ou serveur injoignable.");
  }
  if (res.status === 429) throw new Error("Trop de demandes. Réessaie dans une heure.");
  if (!res.ok && res.status >= 500) throw new Error("Le serveur Supabase ne répond pas. Réessaie dans un instant.");
  // Pour ne pas révéler quels emails existent, on affiche toujours le même message de confirmation.
  return true;
}

// Lit le lien reçu par email (#access_token=...&type=recovery) et nettoie l'adresse.
export function readRecoveryFromUrl() {
  const hash = window.location.hash.startsWith("#") ? window.location.hash.slice(1) : "";
  if (!hash) return null;
  const params = new URLSearchParams(hash);
  const clean = () => window.history.replaceState(null, "", window.location.pathname + window.location.search);
  if (params.get("error") || params.get("error_code")) {
    clean();
    const expired = params.get("error_code") === "otp_expired";
    return { error: expired ? "Ce lien a expiré ou a déjà été utilisé. Fais une nouvelle demande." : "Ce lien n'est pas valide. Fais une nouvelle demande." };
  }
  if (params.get("type") !== "recovery" || !params.get("access_token")) return null;
  clean();
  return {
    access_token: params.get("access_token"),
    refresh_token: params.get("refresh_token"),
    expires_at: Number(params.get("expires_at")) || undefined,
    expires_in: Number(params.get("expires_in")) || 3600,
  };
}

function frenchPasswordError(payload, status) {
  const code = payload?.error_code || "";
  const msg = String(payload?.msg || payload?.message || "");
  if (code === "same_password" || /different from the old/i.test(msg)) return "Choisis un mot de passe différent de l'ancien.";
  if (code === "weak_password" || /at least/i.test(msg)) return "Mot de passe trop court ou trop simple (6 caractères minimum).";
  if (status === 401 || status === 403) return "Ce lien a expiré. Fais une nouvelle demande de mot de passe.";
  return msg || "Impossible de changer le mot de passe.";
}

// Enregistre le nouveau mot de passe à partir du lien reçu, puis ouvre la session.
export async function setNewPassword(recovery, password, confirm) {
  if (!password || password.length < 6) throw new Error("Le mot de passe doit contenir au moins 6 caractères.");
  if (password !== confirm) throw new Error("Les deux mots de passe ne sont pas identiques.");
  let res;
  try {
    res = await fetch(`${SUPABASE_URL}/auth/v1/user`, {
      method: "PUT",
      headers: { apikey: SUPABASE_KEY, Authorization: `Bearer ${recovery.access_token}`, "Content-Type": "application/json" },
      body: JSON.stringify({ password }),
    });
  } catch {
    throw new Error("Pas de connexion internet ou serveur injoignable.");
  }
  const payload = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(frenchPasswordError(payload, res.status));
  return storeSession({ ...recovery, user: { email: payload.email || payload.user?.email || "" } });
}

async function refreshSession(session) {
  const res = await fetch(`${SUPABASE_URL}/auth/v1/token?grant_type=refresh_token`, {
    method: "POST",
    headers: { apikey: SUPABASE_KEY, "Content-Type": "application/json" },
    body: JSON.stringify({ refresh_token: session.refresh_token }),
  });
  const payload = await res.json().catch(() => ({}));
  if (!res.ok || !payload.access_token) {
    removeLocal(SESSION_KEY);
    return null;
  }
  return storeSession({ ...payload, user: payload.user || { email: session.email } });
}

let refreshing = null;
export async function getSession() {
  const session = readLocal(SESSION_KEY);
  if (!session?.access_token || !session?.refresh_token) return null;
  const now = Math.floor(Date.now() / 1000);
  if (session.expires_at - 60 > now) return session;
  if (!refreshing) refreshing = refreshSession(session).finally(() => { refreshing = null; });
  return refreshing;
}

export async function signOut() {
  const session = readLocal(SESSION_KEY);
  stopSync();
  removeLocal(SESSION_KEY);
  if (session?.access_token) {
    try {
      await fetch(`${SUPABASE_URL}/auth/v1/logout`, {
        method: "POST",
        headers: { apikey: SUPABASE_KEY, Authorization: `Bearer ${session.access_token}` },
      });
    } catch {
      /* déconnexion locale suffisante */
    }
  }
}

/* ---------- Accès aux données ---------- */
class AuthExpiredError extends Error {}
export function isAuthExpired(error) {
  return error instanceof AuthExpiredError;
}

async function authHeaders(extra = {}) {
  const session = await getSession();
  if (!session) throw new AuthExpiredError("Session expirée");
  return { apikey: SUPABASE_KEY, Authorization: `Bearer ${session.access_token}`, ...extra };
}

async function checkResponse(res) {
  if (res.status === 401) {
    removeLocal(SESSION_KEY);
    throw new AuthExpiredError("Session expirée");
  }
  if (!res.ok) {
    const detail = await res.text().catch(() => "");
    throw new Error(`Erreur Supabase (${res.status}) ${detail}`.trim());
  }
  return res.json();
}

export async function fetchRemote() {
  const res = await fetch(`${TABLE_URL}?id=eq.${ROW_ID}&select=state,version,updated_by,updated_at`, {
    headers: await authHeaders(),
  });
  const rows = await checkResponse(res);
  return Array.isArray(rows) && rows.length ? rows[0] : null;
}

function cloudState(data) {
  const copy = { ...data };
  DEVICE_KEYS.forEach((key) => delete copy[key]);
  return copy;
}

/* ---------- Synchronisation ---------- */
const sync = {
  active: false,
  email: "",
  version: null,
  exists: false,
  lastJson: null,
  pending: null,
  timer: null,
  saving: false,
  status: "idle", // idle | saving | saved | error | conflict | remote-newer
  remoteBy: "",
  listeners: new Set(),
};

function setStatus(status, extra = {}) {
  Object.assign(sync, extra, { status });
  sync.listeners.forEach((fn) => fn(getSyncInfo()));
}

export function getSyncInfo() {
  return { status: sync.status, email: sync.email, version: sync.version, remoteBy: sync.remoteBy };
}

export function subscribeSync(fn) {
  sync.listeners.add(fn);
  return () => sync.listeners.delete(fn);
}

// Copie de secours des données déjà présentes dans ce navigateur, faite une seule fois.
function backupLocalOnce() {
  if (readLocal(BACKUP_KEY)) return;
  const local = readLocal(LOCAL_DATA_KEY);
  if (local) writeLocal(BACKUP_KEY, { savedAt: new Date().toISOString(), data: local });
}

// Écrit l'état distant dans le navigateur avant le démarrage de l'application.
function applyRemoteToLocal(remoteState) {
  const local = readLocal(LOCAL_DATA_KEY) || {};
  const merged = { ...remoteState };
  DEVICE_KEYS.forEach((key) => {
    if (local[key] !== undefined) merged[key] = local[key];
  });
  writeLocal(LOCAL_DATA_KEY, merged);
}

// Appelé après connexion, avant d'afficher l'application.
export async function startSync(session) {
  backupLocalOnce();
  const remote = await fetchRemote();
  sync.email = session.email;
  const unsynced = readLocal(UNSYNCED_KEY);
  removeLocal(UNSYNCED_KEY);
  if (remote && remote.state && typeof remote.state === "object") {
    sync.version = Number(remote.version) || 0;
    sync.exists = true;
    sync.lastJson = JSON.stringify(cloudState(remote.state));
    const local = readLocal(LOCAL_DATA_KEY);
    const seedFromLocal = sync.version === SEED_FROM_LOCAL.version
      && remote.updated_by === SEED_FROM_LOCAL.updatedBy
      && Array.isArray(local?.projects) && local.projects.length > 0;
    if (seedFromLocal) {
      // On garde les données du navigateur : elles seront envoyées au premier enregistrement.
      sync.lastJson = null;
    } else if (unsynced?.state && unsynced.baseVersion === sync.version) {
      // Modifications faites ici mais pas encore envoyées : on les reprend.
      applyRemoteToLocal(unsynced.state);
    } else {
      if (unsynced?.state) writeLocal(UNSYNCED_CONFLICT_KEY, { savedAt: new Date().toISOString(), ...unsynced });
      applyRemoteToLocal(remote.state);
    }
  } else {
    // Aucune donnée en ligne : les données locales seront envoyées au premier enregistrement.
    sync.version = 0;
    sync.exists = false;
    sync.lastJson = null;
  }
  sync.active = true;
  setStatus("saved");
  window.addEventListener("visibilitychange", onVisibility);
  window.addEventListener("pagehide", onPageHide);
}

export function stopSync() {
  clearTimeout(sync.timer);
  Object.assign(sync, { active: false, pending: null, timer: null });
  window.removeEventListener("visibilitychange", onVisibility);
  window.removeEventListener("pagehide", onPageHide);
}

// Appelé par l'application à chaque modification des données.
export function queueCloudSave(data) {
  if (!sync.active || sync.status === "conflict") return;
  const state = cloudState(data);
  const json = JSON.stringify(state);
  if (json === sync.lastJson) return;
  sync.pending = { state, json };
  writeLocal(UNSYNCED_KEY, { baseVersion: sync.version, state });
  clearTimeout(sync.timer);
  sync.timer = setTimeout(flush, SAVE_DELAY_MS);
}

async function flush({ keepalive = false } = {}) {
  if (!sync.active || !sync.pending) return;
  if (sync.saving) {
    clearTimeout(sync.timer);
    sync.timer = setTimeout(flush, SAVE_DELAY_MS);
    return;
  }
  const { state, json } = sync.pending;
  sync.pending = null;
  sync.saving = true;
  setStatus("saving");
  const nextVersion = (sync.version || 0) + 1;
  const body = { state, version: nextVersion, updated_by: sync.email, updated_at: new Date().toISOString() };
  try {
    let res;
    if (sync.exists) {
      res = await fetch(`${TABLE_URL}?id=eq.${ROW_ID}&version=eq.${sync.version}`, {
        method: "PATCH",
        keepalive,
        headers: await authHeaders({ "Content-Type": "application/json", Prefer: "return=representation" }),
        body: JSON.stringify(body),
      });
    } else {
      res = await fetch(TABLE_URL, {
        method: "POST",
        keepalive,
        headers: await authHeaders({ "Content-Type": "application/json", Prefer: "return=representation" }),
        body: JSON.stringify({ id: ROW_ID, ...body }),
      });
    }
    if (!sync.exists && res.status === 409) {
      // Quelqu'un a créé la ligne entre-temps.
      setStatus("conflict", { remoteBy: "" });
      return;
    }
    const rows = await checkResponse(res);
    if (!Array.isArray(rows) || rows.length === 0) {
      // La version en ligne a changé : une autre personne a enregistré avant nous.
      const remote = await fetchRemote().catch(() => null);
      setStatus("conflict", { remoteBy: remote?.updated_by || "" });
      return;
    }
    Object.assign(sync, { version: nextVersion, exists: true, lastJson: json });
    if (sync.pending) writeLocal(UNSYNCED_KEY, { baseVersion: sync.version, state: sync.pending.state });
    else removeLocal(UNSYNCED_KEY);
    setStatus(sync.pending ? "saving" : "saved");
  } catch (error) {
    if (isAuthExpired(error)) {
      // Les modifications restent dans le navigateur (clé jam_saas_unsynced) et seront reprises après reconnexion.
      stopSync();
      window.location.reload();
      return;
    }
    // On garde la modification pour réessayer.
    sync.pending = sync.pending || { state, json };
    setStatus("error");
    clearTimeout(sync.timer);
    sync.timer = setTimeout(flush, 10000);
  } finally {
    sync.saving = false;
  }
}

function onPageHide() {
  if (sync.pending) {
    clearTimeout(sync.timer);
    flush({ keepalive: true });
  }
}

async function onVisibility() {
  if (document.visibilityState !== "visible" || !sync.active || sync.pending || sync.saving) return;
  if (sync.status === "conflict") return;
  try {
    const remote = await fetchRemote();
    if (remote && Number(remote.version) > (sync.version || 0)) {
      setStatus("remote-newer", { remoteBy: remote.updated_by || "" });
    }
  } catch {
    /* on réessaiera au prochain retour sur l'onglet */
  }
}

// Recharge l'application avec la dernière version en ligne (abandonne les modifications non enregistrées).
export async function reloadFromCloud() {
  const remote = await fetchRemote();
  stopSync();
  removeLocal(UNSYNCED_KEY);
  if (remote?.state) applyRemoteToLocal(remote.state);
  window.location.reload();
}
