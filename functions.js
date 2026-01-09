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

    redirectUri: "https://dndent.github.io/Database/functions.html"
  };

  const GRAPH_SCOPE = ["https://graph.microsoft.com/.default"];

  function shortErr(e) {
    const msg = (e && (e.message || e.errorMessage || String(e))) || "Unknown error";
    return msg.length > 230 ? msg.slice(0, 230) + "…" : msg;
  }

  function escapeODataString(s) {
    return String(s).replace(/'/g, "''");
  }

  function to2D(v) {
    return Array.isArray(v) ? v : [[v]];
  }
  function shape2D(a) {
    return { rows: a.length, cols: a[0]?.length ?? 0 };
  }

  // ---------- Date handling ----------
  function parseExcelDateLocal(excelVal) {
    if (excelVal === null || excelVal === undefined || excelVal === "") return null;

    // Excel serial number (days since 1899-12-30)
    if (typeof excelVal === "number" && isFinite(excelVal)) {
      const epoch = new Date(1899, 11, 30); // local
      return new Date(epoch.getTime() + excelVal * 86400000);
    }

    if (excelVal instanceof Date) {
      return isNaN(excelVal.getTime()) ? null : excelVal;
    }

    const s = String(excelVal).trim();
    if (!s) return null;

    // NL dd-mm-yyyy (or /)
    const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
    if (m) {
      const dd = Number(m[1]), mm = Number(m[2]), yyyy = Number(m[3]);
      const d = new Date(yyyy, mm - 1, dd);
      return isNaN(d.getTime()) ? null : d;
    }

    const d2 = new Date(s); // ISO etc.
    return isNaN(d2.getTime()) ? null : d2;
  }

  // Use local-day boundaries to avoid SharePoint date-only timezone shifts
  function localDayRangeIso(excelVal) {
    const d = parseExcelDateLocal(excelVal);
    if (!d) return null;

    const startLocal = new Date(d.getFullYear(), d.getMonth(), d.getDate());
    const endLocal = new Date(startLocal.getTime() + 86400000);
    return { startIso: startLocal.toISOString(), endIso: endLocal.toISOString() };
  }

  // ---------- MSAL (lazy) ----------
  let msalApp = null;
  let tokenCache = null;

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

  async function ensureMsal() {
    if (typeof msal !== "undefined" && msal?.PublicClientApplication) return;

    // Load msal-browser if not already loaded
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

    if (accounts.length) {
      try {
        const r = await app.acquireTokenSilent({ scopes: GRAPH_SCOPE, account: accounts[0] });
        tokenCache = r.accessToken;
        return tokenCache;
      } catch (_) {}
    }

    // Interactive login (popup must be allowed)
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

  // ---------- Data fetch ----------
  const memo = new Map(); // key -> value

  async function fetchValue(identifier, excelDate) {
    const id = (identifier ?? "").toString().trim();
    const range = localDayRangeIso(excelDate);
    if (!id || !range) return "";

    const key = `${id}__${range.startIso.slice(0, 10)}`;
    if (memo.has(key)) return memo.get(key);

    const filter =
      `fields/${CONFIG.fieldIdentifier} eq '${escapeODataString(id)}'` +
      ` and fields/${CONFIG.fieldDate} ge ${range.startIso}` +
      ` and fields/${CONFIG.fieldDate} lt ${range.endIso}`;

    // Select all relevant fields; helps debug internal names
    const select = [CONFIG.fieldIdentifier, CONFIG.fieldDate, CONFIG.fieldValue].join(",");

    const url =
      `https://graph.microsoft.com/v1.0/sites/${CONFIG.siteId}` +
      `/lists/${CONFIG.listId}/items` +
      `?$expand=fields($select=${encodeURIComponent(select)})` +
      `&$filter=${encodeURIComponent(filter)}` +
      `&$top=1`;

    const data = await graphGet(url);

    if (!data.value?.length) return ""; // no match

    const fields = data.value[0]?.fields || {};
    const val = fields[CONFIG.fieldValue];

    // If internal field name differs, show keys so you can fix CONFIG.fieldValue
    if (val === undefined || val === null || val === "") {
      return `ERR: No '${CONFIG.fieldValue}' field. Keys: ${Object.keys(fields).join(", ")}`;
    }

    memo.set(key, val);
    return val;
  }

  // ---------- Custom function (matrix in/matrix out) ----------
  CustomFunctions.associate("DATA", async (identifier, date) => {
    try {
      const ids = to2D(identifier);
      const dts = to2D(date);

      const s1 = shape2D(ids);
      const s2 = shape2D(dts);

      const rows = Math.max(s1.rows, s2.rows);
      const cols = Math.max(s1.cols, s2.cols);

      const out = Array.from({ length: rows }, () => Array(cols).fill(""));

      const tasks = [];
      for (let r = 0; r < rows; r++) {
        for (let c = 0; c < cols; c++) {
          const idVal = ids[Math.min(r, s1.rows - 1)][Math.min(c, s1.cols - 1)];
          const dtVal = dts[Math.min(r, s2.rows - 1)][Math.min(c, s2.cols - 1)];

          tasks.push(
            fetchValue(idVal, dtVal)
              .then(v => { out[r][c] = v; })
              .catch(e => { out[r][c] = `ERR: ${shortErr(e)}`; })
          );
        }
      }

      await Promise.all(tasks);
      return out;
    } catch (e) {
      return [[`ERR: ${shortErr(e)}`]];
    }
  });
})();
