(() => {
  "use strict";

  const cfg = window.OTTOMEET_CONFIG || {};
  const labels = cfg.labels || {};
  const teams = window.microsoftTeams;
  const el = (id) => document.getElementById(id);

  const DEMO_CHAT_ID = "demo-meeting";
  const DEMO_USER = "demo-user";

  const ctx = { inTeams: false, chatId: DEMO_CHAT_ID, userOid: DEMO_USER, upn: "" };
  let timerInterval = null;
  let memoryMock = null;

  const setMessage = (text) => { el("message").textContent = text || ""; };

  // ---------- Mock-Speicher (ersetzt das OttoMeet-Backend) ----------

  const storageKey = () => `ottomeet:mock:${ctx.chatId}:${ctx.userOid}`;

  function loadMock() {
    try {
      return JSON.parse(localStorage.getItem(storageKey())) || {};
    } catch {
      return {};
    }
  }

  function saveMock(data) {
    try {
      localStorage.setItem(storageKey(), JSON.stringify(data));
    } catch {
      // Ohne localStorage gilt der Status nur bis zum Schließen des Panels.
    }
    memoryMock = data;
  }

  function mockStatus() {
    const data = memoryMock || loadMock();
    if (!data.startedAt) {
      data.startedAt = new Date().toISOString();
      saveMock(data);
    }
    return {
      decision: data.decision || "none",
      consented: data.decision === "accept",
      started_at: data.startedAt,
    };
  }

  function recordDecision(act) {
    const data = memoryMock || loadMock();
    saveMock({ ...data, decision: act, decidedAt: new Date().toISOString() });
  }

  // ---------- Darstellung ----------

  function renderCard(state, icon, title, text) {
    el("statusCard").className = `status-card state-${state}`;
    el("statusIcon").textContent = icon;
    el("statusTitle").textContent = title;
    el("statusText").textContent = text;
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

  function render() {
    const status = mockStatus();
    if (status.decision === "decline") {
      renderCard("bad", "✕", "Consent abgelehnt",
        "Du hast der Transkription widersprochen. Dein Mikrofon bleibt gesperrt, solange OttoMeet läuft.");
    } else if (status.consented) {
      renderCard("ok", "✓", "Consent erteilt", "OttoMeet ist aktiv und verarbeitet das Meeting.");
    } else {
      renderCard("warn", "!", "Consent erforderlich",
        "Damit OttoMeet das Meeting verarbeiten kann, benötigen wir deine Zustimmung.");
    }
    el("consentButton").hidden = status.consented;
    el("withdrawButton").hidden = !status.consented;
    renderTranscript(status.consented, status.started_at);
  }

  // ---------- Aktionen ----------

  function consentPageUrl(act) {
    const url = new URL("mock/consent.html", window.location.href);
    url.searchParams.set("act", act);
    url.searchParams.set("chat_id", ctx.chatId);
    url.searchParams.set("user", ctx.userOid);
    return url.href;
  }

  function dialogSupported() {
    try {
      return ctx.inTeams && teams?.dialog?.url?.isSupported?.() === true;
    } catch {
      return false;
    }
  }

  function openConsentPage(act) {
    setMessage("");
    if (dialogSupported()) {
      teams.dialog.url.open(
        {
          url: consentPageUrl(act),
          title: act === "accept" ? "OttoMeet – Zustimmung" : "OttoMeet – Ablehnung",
          size: { height: 420, width: 480 },
        },
        ({ err, result }) => {
          if (err || !result?.act) return;
          recordDecision(result.act);
          render();
        }
      );
      return;
    }
    const w = window.open(consentPageUrl(act), "_blank");
    if (!w) setMessage("Die Bestätigungsseite konnte nicht geöffnet werden (Popup blockiert?).");
  }

  async function openOttoMeet() {
    setMessage("");
    const url = cfg.ottomeetUrl;
    if (teams?.app?.openLink && ctx.inTeams) {
      try {
        await teams.app.openLink(url);
        return;
      } catch (_) {
        // Manche Clients lehnen openLink für externe URLs ab – dann window.open.
      }
    }
    if (!window.open(url, "_blank", "noopener,noreferrer")) {
      setMessage("OttoMeet konnte nicht geöffnet werden (Popup blockiert?).");
    }
  }

  function resetMock() {
    try {
      localStorage.removeItem(storageKey());
    } catch {
      // ignorieren
    }
    memoryMock = null;
    render();
    setMessage("Mock-Status zurückgesetzt.");
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
    el("resetMockButton").addEventListener("click", resetMock);

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

    // Bestätigung in einem anderen Fenster (Browser-Demo) aktualisiert das Panel.
    window.addEventListener("storage", (event) => {
      if (event.key === storageKey()) {
        memoryMock = null;
        render();
      }
    });
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
    ctx.chatId = context?.chat?.id || DEMO_CHAT_ID;
    ctx.userOid = context?.user?.id || DEMO_USER;
    ctx.upn = context?.user?.userPrincipalName || "";

    document.body.dataset.theme = context?.app?.theme || "default";
    teams.app.registerOnThemeChangeHandler?.((theme) => { document.body.dataset.theme = theme; });
    teams.app.notifySuccess?.();
  }

  document.addEventListener("DOMContentLoaded", async () => {
    bindUi();
    await initTeams();
    el("settingsUser").textContent = ctx.inTeams ? (ctx.upn || "–") : "Browser-Demo (außerhalb von Teams)";
    render();
  });
})();
