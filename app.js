(() => {
  "use strict";

  const cfg = window.MEETING_LINKS_CONFIG || {};
  const teams = window.microsoftTeams;
  const el = (id) => document.getElementById(id);
  const setStatus = (text) => { el("status").textContent = text || ""; };

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

  function bindLinkButton(buttonId, item) {
    const button = el(buttonId);
    if (!item?.url) {
      button.hidden = true;
      return;
    }
    button.textContent = item.label || item.url;
    button.addEventListener("click", async () => {
      setStatus("");
      try {
        await openExternal(item.url);
      } catch (err) {
        setStatus(err?.message || "Link konnte nicht geöffnet werden.");
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
      setStatus("Außerhalb von Teams geöffnet – Links funktionieren trotzdem.");
      return;
    }

    const context = await teams.app.getContext();
    document.body.dataset.frame = context?.page?.frameContext || "";
    document.body.dataset.theme = context?.app?.theme || "default";
    teams.app.registerOnThemeChangeHandler?.((theme) => { document.body.dataset.theme = theme; });
    teams.app.notifySuccess?.();
  }

  document.addEventListener("DOMContentLoaded", () => {
    el("title").textContent = cfg.title || "Meeting Links";
    el("subtitle").textContent = cfg.subtitle || "";
    bindLinkButton("button1", cfg.button1);
    bindLinkButton("button2", cfg.button2);
    initTeams();
  });
})();
