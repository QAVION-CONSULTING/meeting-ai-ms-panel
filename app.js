(() => {
  "use strict";

  const cfg = window.OTTOMEET_CONFIG || {};
  const labels = cfg.labels || {};
  const teams = window.microsoftTeams;
  const el = (id) => document.getElementById(id);

  const POLL_MS = Math.max(5, Number(cfg.statusPollSeconds) || 10) * 1000;
  const FAST_POLL_MS = 3000;
  const FAST_POLL_DURATION_MS = 2 * 60 * 1000;

  const ctx = { inTeams: false, chatId: "", userOid: "", upn: "" };
  let pollTimer = null;
  let fastPollUntil = 0;
  let timerInterval = null;

  const setMessage = (text) => { el("message").textContent = text || ""; };

  function isHttps(url) {
    try {
      return new URL(url).protocol === "https:";
    } catch {
      return false;
    }
  }

  async function openExternal(url) {
    if (!isHttps(url)) throw new Error("Es sind nur HTTPS-Links erlaubt.");
    if (teams?.app?.openLink) {
      try {
        await teams.app.openLink(url);
        return;
      } catch (_) {
        // Manche Clients lehnen openLink für externe URLs ab – dann window.open.
      }
    }
    const w = window.open(url, "_blank", "noopener,noreferrer");
    if (!w) throw new Error("Der Link konnte nicht geöffnet werden (Popup blockiert?).");
  }

  function backendUrl(path, params) {
    const url = new URL(path, cfg.backendUrl + "/");
    for (const [key, value] of Object.entries(params)) {
      if (value) url.searchParams.set(key, value);
    }
    return url.href;
  }

  // ---------- Darstellung ----------

  function renderCard(state, icon, title, text) {
    el("statusCard").className = `status-card state-${state}`;
    el("statusIcon").textContent = icon;
    el("statusTitle").textContent = title;
    el("statusText").textContent = text;
  }

  function renderButtons({ showConsent, consentEnabled = true, showWithdraw }) {
    el("consentButton").hidden = !showConsent;
    el("consentButton").disabled = !consentEnabled;
    el("withdrawButton").hidden = !showWithdraw;
  }

  function renderTranscript(active, startedAt) {
    el("transcriptText").textContent = active ? "Transkription aktiv" : "Transkription noch nicht aktiv";
    el("transcriptDot").parentElement.classList.toggle("active", active);
    clearInterval(timerInterval);
    el("transcriptTimer").textContent = "";
    const start = active && startedAt ? Date.parse(startedAt) : NaN;
    if (Number.isNaN(start)) return;
    const tick = () => {
      const s = Math.max(0, Math.floor((Date.now() - start) / 1000));
      const h = Math.floor(s / 3600);
      const mm = String(Math.floor((s % 3600) / 60)).padStart(2, "0");
      const ss = String(s % 60).padStart(2, "0");
      el("transcriptTimer").textContent = h ? `${h}:${mm}:${ss}` : `${mm}:${ss}`;
    };
    tick();
    timerInterval = setInterval(tick, 1000);
  }

  function renderStatus(status) {
    if (!status.session_active) {
      renderCard("pending", "i", "OttoMeet ist nicht aktiv",
        "In diesem Meeting läuft gerade keine OttoMeet-Transkription. Eine Einwilligung ist möglich, sobald OttoMeet gestartet wurde.");
      renderButtons({ showConsent: true, consentEnabled: false, showWithdraw: false });
      renderTranscript(false);
      return;
    }

    if (status.decision === "decline") {
      renderCard("bad", "✕", "Consent abgelehnt",
        "Du hast der Transkription widersprochen. Dein Mikrofon bleibt gesperrt, solange OttoMeet läuft.");
      renderButtons({ showConsent: true, showWithdraw: false });
    } else if (status.consented) {
      let reason = "OttoMeet ist aktiv und verarbeitet das Meeting.";
      if (status.decision !== "accept" && status.recording_mode === "tagged") {
        reason = "In diesem Meeting gilt die Einwilligung standardmäßig. Wenn du nicht einverstanden bist, lehne ab.";
      } else if (status.decision !== "accept" && status.persistent_consent) {
        reason = "Du hast die automatische Einwilligung für alle Meetings aktiviert.";
      }
      renderCard("ok", "✓", "Consent erteilt", reason);
      renderButtons({ showConsent: false, showWithdraw: true });
    } else {
      renderCard("warn", "!", "Consent erforderlich",
        "Damit OttoMeet das Meeting verarbeiten kann, benötigen wir deine Zustimmung.");
      renderButtons({ showConsent: true, showWithdraw: false });
    }
    renderTranscript(true, status.started_at);
  }

  function renderNoStatus(title, text) {
    renderCard("warn", "!", title, text);
    renderButtons({ showConsent: true, showWithdraw: false });
    renderTranscript(false);
  }

  // ---------- Status vom OttoMeet-Backend ----------

  async function fetchStatus() {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 8000);
    try {
      const res = await fetch(
        backendUrl("teams/app-status", { chat_id: ctx.chatId, user_oid: ctx.userOid, upn: ctx.upn }),
        { headers: { Accept: "application/json" }, credentials: "omit", signal: controller.signal }
      );
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      return await res.json();
    } finally {
      clearTimeout(timeout);
    }
  }

  async function refresh() {
    if (!ctx.inTeams) {
      renderNoStatus("Nur in einer Teams-Besprechung verfügbar",
        "Öffne OttoMeet über das App-Symbol in der Meeting-Leiste.");
      return;
    }
    if (!isHttps(cfg.backendUrl)) {
      renderNoStatus("Consent erforderlich",
        "Damit OttoMeet das Meeting verarbeiten kann, benötigen wir deine Zustimmung. (Status-Anzeige: OttoMeet-Backend noch nicht verbunden.)");
      return;
    }
    try {
      renderStatus(await fetchStatus());
    } catch {
      renderNoStatus("Status derzeit nicht verfügbar",
        "Das OttoMeet-Backend ist nicht erreichbar. Du kannst trotzdem einwilligen – der Status erscheint, sobald die Verbindung wieder steht.");
    }
  }

  function schedulePoll() {
    clearTimeout(pollTimer);
    if (!ctx.inTeams || !isHttps(cfg.backendUrl)) return;
    const delay = Date.now() < fastPollUntil ? FAST_POLL_MS : POLL_MS;
    pollTimer = setTimeout(async () => {
      if (!document.hidden) await refresh();
      schedulePoll();
    }, delay);
  }

  function pollFastForAWhile() {
    fastPollUntil = Date.now() + FAST_POLL_DURATION_MS;
    schedulePoll();
  }

  // ---------- Aktionen ----------

  async function openConsentPage(act) {
    setMessage("");
    if (!isHttps(cfg.backendUrl)) {
      setMessage("Das OttoMeet-Backend ist noch nicht konfiguriert (backendUrl in settings.json).");
      return;
    }
    if (!ctx.chatId) {
      setMessage("Das ist nur innerhalb einer Teams-Besprechung möglich.");
      return;
    }
    try {
      await openExternal(backendUrl("teams/consent-entry", { chat_id: ctx.chatId, act }));
      setMessage(act === "accept"
        ? "Bitte die Zustimmung im geöffneten Browserfenster bestätigen."
        : "Bitte die Ablehnung im geöffneten Browserfenster bestätigen.");
      pollFastForAWhile();
    } catch (err) {
      setMessage(err?.message || "Die Seite konnte nicht geöffnet werden.");
    }
  }

  async function openOttoMeet() {
    setMessage("");
    try {
      await openExternal(cfg.ottomeetUrl);
    } catch (err) {
      setMessage(err?.message || "OttoMeet konnte nicht geöffnet werden.");
    }
  }

  function bindUi() {
    el("consentButton").textContent = labels.consent || "Consent geben";
    el("withdrawButton").textContent = labels.withdraw || "Consent ablehnen";
    el("ottomeetButton").textContent = labels.ottomeet || "Transkript starten (OttoMeet öffnen)";
    el("ottomeetButton").hidden = !cfg.ottomeetUrl;
    el("settingsVersion").textContent = cfg.version || "–";

    el("consentButton").addEventListener("click", () => openConsentPage("accept"));
    el("withdrawButton").addEventListener("click", () => openConsentPage("decline"));
    el("ottomeetButton").addEventListener("click", openOttoMeet);

    for (const tab of document.querySelectorAll(".tab")) {
      tab.addEventListener("click", () => {
        for (const t of document.querySelectorAll(".tab")) {
          const active = t === tab;
          t.classList.toggle("active", active);
          t.setAttribute("aria-selected", String(active));
          el(`tab-${t.dataset.tab}`).hidden = !active;
        }
      });
    }

    const onReturn = () => {
      if (!document.hidden && ctx.inTeams) refresh();
    };
    document.addEventListener("visibilitychange", onReturn);
    window.addEventListener("focus", onReturn);
  }

  // ---------- Start ----------

  async function initTeams() {
    if (!teams?.app) return;
    try {
      await teams.app.initialize();
    } catch {
      return;
    }
    const context = await teams.app.getContext();
    ctx.inTeams = true;
    ctx.chatId = context?.chat?.id || "";
    ctx.userOid = context?.user?.id || "";
    ctx.upn = context?.user?.userPrincipalName || "";
    el("settingsUser").textContent = ctx.upn || "–";

    document.body.dataset.theme = context?.app?.theme || "default";
    teams.app.registerOnThemeChangeHandler?.((theme) => { document.body.dataset.theme = theme; });
    teams.app.notifySuccess?.();
  }

  document.addEventListener("DOMContentLoaded", async () => {
    bindUi();
    await initTeams();
    await refresh();
    schedulePoll();
  });
})();
