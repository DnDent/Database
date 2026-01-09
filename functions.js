/* global CustomFunctions, msal */

const CONFIG = {
  tenantId: "37f72dd3-3cb6-4a3c-9f03-a0e9ad17c700",
  clientId: "a59c788a-3f0e-4fa0-a450-746734ef6fcd",

  siteId: "vinhuys.sharepoint.com,9ed18a60-56d2-4f3a-8d2f-eeea9517f7a1,039a6441-1114-4ac1-b5c2-93e74de479c2",
  listId: "882d9fb5-7971-45d9-a219-de43b6545661",

  fieldIdentifier: "Identifier",
  fieldDate: "Date",
  fieldValue: "Value"
};

// IMPORTANT: redirectUri must be stable & match what you configured in Entra/AAD.
// Using the current page without query/hash is safest.
const redirectUri = (() => {
  try {
    const u = new URL(window.location.href);
    u.search = "";
    u.hash = "";
    return u.toString();
  } catch {
    return window.location.href;
  }
})();

const msalApp = new msal.PublicClientApplication({
  auth: {
    clientId: CONFIG.clientId,
    authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
    redirectUri
  },
  cache: { cacheLocation: "sessionStorage" }
});

// Newer msal-browser versions require initialize() before calling login/acquireToken
const msalReady = typeof msalApp.initialize === "function"
  ? msalApp.initialize()
  : Promise.resolve();

// With admin consent granted, .default is simplest.
// If you want delegated scopes instead, use e.g. ["Sites.Read.All"] (and grant consent).
const GRAPH_SCOPE = ["https://graph.microsoft.com/.default"];

let tokenCache = null;
const memo = new Map(); // key -> value

function shortErr(e) {
  const msg = (e && (e.message || e.errorMessage || String(e))) || "Unknown error";
  return msg.length > 200 ? msg.slice(0, 200) + "…" : msg;
}

async function getToken() {
  await msalReady;

  if (tokenCache) return tokenCache;

  const accounts = msalApp.getAllAccounts();
  if (accounts.length) {
    try {
      const r = await msalApp.acquireTokenSilent({
        scopes: GRAPH_SCOPE,
        account: accounts[0]
      });
      tokenCache = r.accessToken;
      return tokenCache;
    } catch (e) {
      // fall through to interactive
    }
  }

  // Interactive login
  await msalApp.loginPopup({ scopes: GRAPH_SCOPE });

  const account = msalApp.getAllAccounts()[0];
  const r2 = await msalApp.acquireTokenSilent({
    scopes: GRAPH_SCOPE,
    account
  });

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

function to2D(v) {
  return Array.isArray(v) ? v : [[v]];
}

function shape2D(a) {
  return { rows: a.length, cols: a[0]?.length ?? 0 };
}

function escapeODataString(s) {
  return String(s).replace(/'/g, "''");
}

/**
 * Parse Excel date input into a JS Date (local time).
 * Accepts:
 * - Excel serial number (number)
 * - JS Date
 * - String (supports ISO, and also NL-ish "D-M-YYYY" / "DD-MM-YYYY")
 */
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

  // NL style: D-M-YYYY or DD-MM-YYYY (also allow /)
  const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m) {
    const dd = Number(m[1]);
    const mm = Number(m[2]);
    const yyyy = Number(m[3]);
    const d = new Date(yyyy, mm - 1, dd);
    return isNaN(d.getTime()) ? null : d;
  }

  // Try native parse (works for ISO "2026-01-02" etc.)
  const d2 = new Date(s);
  return isNaN(d2.getTime()) ? null : d2;
}

/**
 * Convert date to local-day boundaries (start/end) and return ISO strings (UTC) for OData filter.
 * This avoids SharePoint "date-only" timezone shifting issues.
 */
function localDayRangeIso(excelVal) {
  const d = parseExcelDateLocal(excelVal);
  if (!d) return null;

  const startLocal = new Date(d.getFullYear(), d.getMonth(), d.getDate()); // local midnight
  const endLocal = new Date(startLocal.getTime() + 86400000);

  return {
    startIso: startLocal.toISOString(), // converted to UTC
    endIso: endLocal.toISOString()
  };
}

async function fetchValue(identifier, excelDate) {
  const id = (identifier ?? "").toString().trim();
  const range = localDayRangeIso(excelDate);
  if (!id || !range) return "";

  const key = `${id}__${range.startIso.slice(0, 10)}`;
  if (memo.has(key)) return memo.get(key);

  // NOTE: OData DateTimeOffset literal in Graph filters is typically unquoted.
  const filter =
    `fields/${CONFIG.fieldIdentifier} eq '${escapeODataString(id)}'` +
    ` and fields/${CONFIG.fieldDate} ge ${range.startIso}` +
    ` and fields/${CONFIG.fieldDate} lt ${range.endIso}`;

  const select = [CONFIG.fieldIdentifier, CONFIG.fieldDate, CONFIG.fieldValue].join(",");

  const url =
    `https://graph.microsoft.com/v1.0/sites/${CONFIG.siteId}` +
    `/lists/${CONFIG.listId}/items` +
    `?$expand=fields($select=${encodeURIComponent(select)})` +
    `&$filter=${encodeURIComponent(filter)}` +
    `&$top=1`;

  const data = await graphGet(url);

  // If no match, return blank (or change to "NO MATCH" while debugging)
  if (!data.value?.length) return "";

  const fields = data.value[0]?.fields || {};
  const val = fields[CONFIG.fieldValue];

  // If value field missing, return keys to help you fix internal field name quickly
  if (val === undefined || val === null || val === "") {
    // Comment out next line if you don't want debug output:
    return `ERR: No '${CONFIG.fieldValue}' field. Keys: ${Object.keys(fields).join(", ")}`;
  }

  memo.set(key, val);
  return val;
}

/**
 * =TESLIN.DATA(identifier, date)
 * (Namespace TESLIN comes from manifest. Function name here must match functions.json name: "DATA".)
 *
 * Supports scalars or ranges (spills).
 * Broadcasting rules:
 * - If one input is 1x1, it broadcasts over the other input's shape.
 * - Otherwise, result shape is max(rows), max(cols) with edge broadcasting.
 */
CustomFunctions.associate("DATA", async (identifier, date) => {
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
          .catch(e => {
            // Prevent one failure from turning the whole function into #VALUE
            out[r][c] = `ERR: ${shortErr(e)}`;
          })
      );
    }
  }

  await Promise.all(tasks);
  return out;
});
