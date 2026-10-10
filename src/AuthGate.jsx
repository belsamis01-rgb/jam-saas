import React, { useEffect, useState } from "react";
import { getSession, signIn, signOut, startSync, subscribeSync, getSyncInfo, reloadFromCloud, isAuthExpired, requestPasswordReset, readRecoveryFromUrl, setNewPassword } from "./cloud.js";

// Lien « mot de passe oublié » reçu par email : lu une seule fois au chargement de la page.
const initialRecovery = readRecoveryFromUrl();

const font = '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';

const styles = {
  page: { minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: "#f4f5f7", fontFamily: font, padding: 16, boxSizing: "border-box" },
  card: { width: "100%", maxWidth: 380, background: "#fff", borderRadius: 16, padding: 28, boxShadow: "0 8px 30px rgba(15,23,42,0.08)", boxSizing: "border-box" },
  title: { margin: 0, fontSize: 22, fontWeight: 700, color: "#0f172a" },
  subtitle: { margin: "6px 0 22px", fontSize: 14, color: "#64748b" },
  label: { display: "block", fontSize: 13, fontWeight: 600, color: "#334155", marginBottom: 6 },
  input: { width: "100%", minHeight: 44, padding: "10px 12px", fontSize: 15, borderRadius: 10, border: "1px solid #cbd5e1", boxSizing: "border-box", marginBottom: 14, fontFamily: font },
  button: { width: "100%", minHeight: 44, border: 0, borderRadius: 10, background: "#0f172a", color: "#fff", fontSize: 15, fontWeight: 600, cursor: "pointer", fontFamily: font },
  error: { background: "#fef2f2", color: "#b91c1c", border: "1px solid #fecaca", borderRadius: 10, padding: "10px 12px", fontSize: 13, marginBottom: 14 },
  center: { minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: font, color: "#64748b", fontSize: 15, padding: 16, textAlign: "center" },
};

const linkButton = { background: "none", border: 0, padding: "10px 0", minHeight: 44, width: "100%", color: "#2563eb", fontSize: 14, cursor: "pointer", fontFamily: font };

function LoginForm({ onLogged, initialError = "" }) {
  const [mode, setMode] = useState("login"); // login | forgot | sent
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState(initialError);
  const [busy, setBusy] = useState(false);

  const switchMode = (next) => {
    setError("");
    setMode(next);
  };

  const submit = async (event) => {
    event.preventDefault();
    setError("");
    setBusy(true);
    try {
      if (mode === "forgot") {
        await requestPasswordReset(email);
        setMode("sent");
        setBusy(false);
        return;
      }
      const session = await signIn(email, password);
      await onLogged(session);
    } catch (err) {
      setError(err.message || "Connexion impossible.");
      setBusy(false);
    }
  };

  if (mode === "sent") {
    return (
      <div style={styles.page}>
        <div style={styles.card}>
          <h1 style={styles.title}>Email envoyé</h1>
          <p style={styles.subtitle}>Si un compte existe pour {email.trim()}, tu vas recevoir un email avec un lien pour choisir un nouveau mot de passe. Pense à regarder dans les spams.</p>
          <button style={styles.button} type="button" onClick={() => switchMode("login")}>Retour à la connexion</button>
        </div>
      </div>
    );
  }

  const forgot = mode === "forgot";
  return (
    <div style={styles.page}>
      <form style={styles.card} onSubmit={submit} noValidate>
        <h1 style={styles.title}>{forgot ? "Mot de passe oublié" : "JAM SaaS"}</h1>
        <p style={styles.subtitle}>{forgot ? "Saisis ton email : tu recevras un lien pour choisir un nouveau mot de passe." : "Connecte-toi pour accéder aux données partagées."}</p>
        {error && <div style={styles.error} role="alert">{error}</div>}
        <label style={styles.label} htmlFor="jam-email">Email</label>
        <input id="jam-email" style={styles.input} type="email" autoComplete="username" value={email} onChange={(e) => setEmail(e.target.value)} disabled={busy} />
        {!forgot && (
          <>
            <label style={styles.label} htmlFor="jam-password">Mot de passe</label>
            <input id="jam-password" style={styles.input} type="password" autoComplete="current-password" value={password} onChange={(e) => setPassword(e.target.value)} disabled={busy} />
          </>
        )}
        <button style={{ ...styles.button, opacity: busy ? 0.6 : 1 }} type="submit" disabled={busy}>
          {busy ? (forgot ? "Envoi…" : "Connexion…") : (forgot ? "Recevoir le lien" : "Se connecter")}
        </button>
        <button type="button" style={linkButton} onClick={() => switchMode(forgot ? "login" : "forgot")} disabled={busy}>
          {forgot ? "Retour à la connexion" : "Mot de passe oublié ?"}
        </button>
      </form>
    </div>
  );
}

function NewPasswordForm({ recovery, onDone }) {
  const [password, setPassword] = useState("");
  const [confirm, setConfirm] = useState("");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);

  const submit = async (event) => {
    event.preventDefault();
    setError("");
    setBusy(true);
    try {
      const session = await setNewPassword(recovery, password, confirm);
      await onDone(session);
    } catch (err) {
      setError(err.message || "Impossible de changer le mot de passe.");
      setBusy(false);
    }
  };

  return (
    <div style={styles.page}>
      <form style={styles.card} onSubmit={submit} noValidate>
        <h1 style={styles.title}>Nouveau mot de passe</h1>
        <p style={styles.subtitle}>Choisis ton nouveau mot de passe (6 caractères minimum).</p>
        {error && <div style={styles.error} role="alert">{error}</div>}
        <label style={styles.label} htmlFor="jam-new-password">Nouveau mot de passe</label>
        <input id="jam-new-password" style={styles.input} type="password" autoComplete="new-password" value={password} onChange={(e) => setPassword(e.target.value)} disabled={busy} />
        <label style={styles.label} htmlFor="jam-confirm-password">Confirme le mot de passe</label>
        <input id="jam-confirm-password" style={styles.input} type="password" autoComplete="new-password" value={confirm} onChange={(e) => setConfirm(e.target.value)} disabled={busy} />
        <button style={{ ...styles.button, opacity: busy ? 0.6 : 1 }} type="submit" disabled={busy}>
          {busy ? "Enregistrement…" : "Enregistrer et me connecter"}
        </button>
      </form>
    </div>
  );
}

const STATUS_LABELS = {
  idle: "Connexion…",
  saving: "Enregistrement…",
  saved: "Synchronisé",
  error: "Hors ligne – nouvel essai automatique",
  conflict: "Modification non enregistrée",
  "remote-newer": "Nouvelle version disponible",
};

function SyncBar() {
  const [info, setInfo] = useState(getSyncInfo());
  const [busy, setBusy] = useState(false);
  useEffect(() => subscribeSync(setInfo), []);

  const reload = async () => {
    setBusy(true);
    try {
      await reloadFromCloud();
    } catch {
      setBusy(false);
    }
  };
  const logout = async () => {
    await signOut();
    window.location.reload();
  };

  const alert = info.status === "conflict" || info.status === "remote-newer";
  const who = info.remoteBy ? ` par ${info.remoteBy}` : "";
  const alertText = info.status === "conflict"
    ? `Des modifications ont été enregistrées${who} pendant que tu travaillais. Ta dernière modification n'a pas été enregistrée.`
    : `Une version plus récente a été enregistrée${who}.`;

  const pill = { position: "fixed", right: 12, bottom: 12, zIndex: 9999, display: "flex", alignItems: "center", gap: 8, background: "#fff", border: "1px solid #e2e8f0", borderRadius: 999, padding: "4px 6px 4px 12px", boxShadow: "0 4px 14px rgba(15,23,42,0.08)", fontFamily: font, fontSize: 12, color: "#334155", maxWidth: "calc(100vw - 24px)", boxSizing: "border-box" };
  const dotColor = info.status === "saved" ? "#16a34a" : info.status === "error" || alert ? "#d97706" : "#64748b";
  const smallBtn = { minHeight: 32, minWidth: 44, border: "1px solid #e2e8f0", background: "#f8fafc", borderRadius: 999, padding: "0 10px", fontSize: 12, cursor: "pointer", fontFamily: font, color: "#0f172a" };

  return (
    <>
      {alert && (
        <div role="alert" style={{ position: "fixed", left: 12, right: 12, top: 12, zIndex: 10000, background: "#fffbeb", border: "1px solid #fcd34d", color: "#78350f", borderRadius: 12, padding: "12px 14px", fontFamily: font, fontSize: 14, display: "flex", flexWrap: "wrap", gap: 10, alignItems: "center", justifyContent: "space-between", boxShadow: "0 6px 20px rgba(15,23,42,0.12)" }}>
          <span>{alertText}</span>
          <button style={{ ...smallBtn, minHeight: 44, background: "#0f172a", color: "#fff", border: 0, padding: "0 16px" }} onClick={reload} disabled={busy}>
            {busy ? "Chargement…" : "Charger la dernière version"}
          </button>
        </div>
      )}
      <div style={pill}>
        <span style={{ width: 8, height: 8, borderRadius: 8, background: dotColor, flex: "0 0 auto" }} />
        <span style={{ whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }} title={info.email}>
          {STATUS_LABELS[info.status] || "Synchronisé"}{info.email ? ` · ${info.email}` : ""}
        </span>
        <button style={smallBtn} onClick={logout}>Déconnexion</button>
      </div>
    </>
  );
}

export default function AuthGate({ children }) {
  const [phase, setPhase] = useState("checking"); // checking | login | new-password | loading | ready | error
  const [message, setMessage] = useState("");

  const boot = async (session) => {
    setPhase("loading");
    try {
      await startSync(session);
      setPhase("ready");
    } catch (err) {
      if (isAuthExpired(err)) {
        setPhase("login");
        return;
      }
      setMessage("Impossible de charger les données en ligne. Vérifie ta connexion puis recharge la page.");
      setPhase("error");
    }
  };

  useEffect(() => {
    let cancelled = false;
    if (initialRecovery?.access_token) {
      setPhase("new-password");
      return () => { cancelled = true; };
    }
    if (initialRecovery?.error) {
      setPhase("login");
      return () => { cancelled = true; };
    }
    getSession()
      .then((session) => {
        if (cancelled) return;
        if (session) boot(session);
        else setPhase("login");
      })
      .catch(() => !cancelled && setPhase("login"));
    return () => { cancelled = true; };
  }, []);

  if (phase === "new-password") return <NewPasswordForm recovery={initialRecovery} onDone={boot} />;
  if (phase === "login") return <LoginForm onLogged={boot} initialError={initialRecovery?.error || ""} />;
  if (phase === "ready") return (<>{children}<SyncBar /></>);
  if (phase === "error") {
    return (
      <div style={styles.center}>
        <div>
          <p>{message}</p>
          <button style={{ ...styles.button, width: "auto", padding: "0 20px", marginTop: 8 }} onClick={() => window.location.reload()}>Recharger</button>
        </div>
      </div>
    );
  }
  return <div style={styles.center}>Chargement des données…</div>;
}
