(() => {
  "use strict";

  const cfg = window.MEETING_LINKS_CONFIG || {};
  const teams = window.microsoftTeams;
  const el = (id) => document.getElementById(id);
  const setStatus = (text) => { el("status").textContent = text || ""; };
  let chatId = "";

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

  function consentUrl(act) {
    const url = new URL("teams/consent-entry", cfg.consent.backendUrl + "/");
    url.searchParams.set("chat_id", chatId);
    url.searchParams.set("act", act);
    return url.href;
  }

  async function giveConsent() {
    setStatus("");
    if (!isHttps(cfg.consent?.backendUrl)) {
      setStatus("Das OttoMeet-Backend ist noch nicht konfiguriert (consent.backendUrl).");
      return;
    }
    if (!chatId) {
      setStatus("Die Einwilligung ist nur innerhalb einer Teams-Besprechung möglich.");
      return;
    }
    try {
      await openExternal(consentUrl("accept"));
      setStatus("Bitte die Einwilligung im geöffneten Browserfenster bestätigen.");
    } catch (err) {
      setStatus(err?.message || "Die Einwilligungsseite konnte nicht geöffnet werden.");
    }
  }

  function bindButtons() {
    const consentButton = el("consentButton");
    consentButton.textContent = cfg.consent?.label || "I Consent";
    consentButton.addEventListener("click", giveConsent);

    const ottomeetButton = el("ottomeetButton");
    if (!cfg.ottomeet?.url) {
      ottomeetButton.hidden = true;
      return;
    }
    ottomeetButton.textContent = cfg.ottomeet.label || "Open OttoMeet";
    ottomeetButton.addEventListener("click", async () => {
      setStatus("");
      try {
        await openExternal(cfg.ottomeet.url);
      } catch (err) {
        setStatus(err?.message || "OttoMeet konnte nicht geöffnet werden.");
      }
    });
  }

  async function initTeams() {
    if (!teams?.app) {
      setStatus("TeamsJS konnte nicht geladen werden.");
      return;
    }
    try {
      await teams.app.initialize();
    } catch {
      setStatus("Außerhalb von Teams geöffnet – „I Consent“ funktioniert nur in einer Besprechung.");
      return;
    }

    const context = await teams.app.getContext();
    chatId = context?.chat?.id || "";
    document.body.dataset.frame = context?.page?.frameContext || "";
    document.body.dataset.theme = context?.app?.theme || "default";
    teams.app.registerOnThemeChangeHandler?.((theme) => { document.body.dataset.theme = theme; });
    teams.app.notifySuccess?.();
  }

  document.addEventListener("DOMContentLoaded", () => {
    el("title").textContent = cfg.title || "OttoMeet";
    el("subtitle").textContent = cfg.subtitle || "";
    bindButtons();
    initTeams();
  });
})();
