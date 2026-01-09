/* global CustomFunctions, msal */

(() => {
  const CONFIG = {
    tenantId: "37f72dd3-3cb6-4a3c-9f03-a0e9ad17c700",
    clientId: "a59c788a-3f0e-4fa0-a450-746734ef6fcd",

    siteId: "vinhuys.sharepoint.com,9ed18a60-56d2-4f3a-8d2f-eeea9517f7a1,039a6441-1114-4ac1-b5c2-93e74de479c2",
    listId: "882d9fb5-7971-45d9-a219-de43b6545661",

    fieldIdentifier: "Identifier",
    fieldDate: "Date",
    fieldValue: "Value",

    // IMPORTANT: This MUST be listed as a Redirect URI in your Entra App registration.
    // Put your actual GitHub Pages functions.html URL here:
    redirectUri: "https://dndent.github.io/Database/functions.html"
  };

  const GRAPH_SCOPE = ["https://graph.microsoft.com/.default"];

  // ---- safety helpers (avoid #VALUE) ----
  function safeAssociate(name, fn) {
    try { CustomFunctions.associate(name, fn); } catch (_) {}
  }
  function shortErr(e) {
    const msg = (e && (e.message || e.errorMessage || String(e))) || "Unknown error";
    return msg.length > 220 ? msg.slice(0, 220) + "…" : msg;
  }

  function loadScript(url) {
    return new Promise((resolve, reject) => {
      const s = document.createElement("script");
      s.src = url;
      s.async = true;
      s.onload = resolve;
      s.onerror = () => reject(new Error(`Failed to load script: ${url}`));
      document.head.appendChild(s);
    });
  }

  // ---- MSAL (lazy, so it won't crash at load) ----
  let msalApp = null;
  let tokenCache = null;

  async function ensureMsal() {
    if (typeof msal !== "undefined" && msal?.PublicClientApplication) return;

    // Load MSAL if not present (prevents your earlier top-level crash)
    await loadScript("https://alcdn.msauth.net/browser/2.38.0/js/msal-browser.min.js");

    if (typeof msal === "undefined" || !msal?.PublicClientApplication) {
      throw new Error("MSAL not available (failed to load msal-browser).");
    }
  }

  async function ensureMsalApp() {
    if (msalApp) return msalApp;

    await ensureMsal();

    msalApp = new msal.PublicClientApplication({
      auth: {
        clientId: CONFIG.clientId,
        authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
        redirectUri: CONFIG.redirectUri
      },
      cache: { cacheLocation: "sessionStorage" }
    });

    if (typeof msalApp.initialize === "function") {
      await msalApp.initialize();
    }

    return msalApp;
  }

  async function getToken() {
    if (tokenCache) return tokenCache;

    const app = await ensureMsalApp();
    const accounts = app.getAllAccounts();

    // Try silent first
    if (accounts.length) {
      try {
        const r = await app.acquireTokenSilent({ scopes: GRAPH_SCOPE, account: accounts[0] });
        tokenCache = r.accessToken;
        return tokenCache;
      } catch (_) {}
    }

    // Interactive fallback (if popup blocked, you'll see ERR: ...)
    await app.loginPopup({ scopes: GRAPH_SCOPE });

    const acct = app.getAllAccounts()[0];
    const r2 = await app.acquireTokenSilent({ scopes: GRAPH_SCOPE, account: acct });
    tokenCache = r2.accessToken;
    return tokenCache;
  }

  async function graphGet(url) {
    const token = await getToken();
    const r = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
    if (!r.ok) {
      const txt = await r.text().catch(() => "");
      throw new Error(`Graph ${r.status}: ${txt || r.statusText}`);
    }
    return r.json();
  }

  // ---- Excel matrix helpers ----
  function to2D(v) { return Array.isArray(v) ? v : [[v]]; }
  function shape2D(a) { return { rows: a.length, cols: a[0]?.length ?? 0 }; }
  function escapeODataString(s) { return String(s).replace(/'/g, "''"); }

  // ---- Date parsing (supports Excel serial, Date, ISO, and NL dd-mm-yyyy) ----
  function parseExcelDateLocal(excelVal) {
    if (excelVal === null || excelVal === undefined || excelVal === "") return null;

    if (typeof excelVal === "number" && isFinite(excelVal)) {
      const epoch = new Date(1899, 11, 30); // local epoch
      return new Date(epoch.getTime() + excelVal * 86400000);
    }

    if (excelVal instanceof Date) {
      return isNaN(excelVal.getTime()) ? null : excelVal;
    }

    const s = String(excelVal).trim();
    if (!s) return null;

    // dd-mm-yyyy or d-m-yyyy (or /)
    const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
    if (m) {
      const dd = Number(m[1]), mm = Number(m[2]), yyyy = Number(m[3]);
      const d = new Date(yyyy, mm - 1, dd);
      return isNaN(d.getTime()) ? null : d;
    }

    const d2 = new Date(s);
    return isNaN(d2.getTime()) ? null : d2;
  }

  // Local day range -> ISO (UTC) boundaries for filtering “date-only” columns safely
  function localDayRangeIso(excelVal) {
    const d = parseExcelDateLocal(excelVal);
    if (!d) return null;

    const startLocal = new Date(d.getFullYear(), d.getMonth(), d.getDate());
    const endLocal = new Date(startLocal.getTime() + 86400000);
    return { startIso: startLocal.toISOString(), endIso: endLocal.toISOString()
