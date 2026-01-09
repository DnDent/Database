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

const msalApp = new msal.PublicClientApplication({
  auth: {
    clientId: CONFIG.clientId,
    authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
    // Must be same origin as where this file is hosted.
    redirectUri: window.location.href
  },
  cache: { cacheLocation: "sessionStorage" }
});

// With admin consent granted, .default is the simplest for MVP.
const GRAPH_SCOPE = ["https://graph.microsoft.com/.default"];

let tokenCache = null;
const memo = new Map(); // key -> value

async function getToken() {
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
    } catch {
      // fall through to popup
    }
  }

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
    const txt = await r.text();
    throw new Error(`Graph ${r.status}: ${txt}`);
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
 * Convert Excel date input to a UTC date-only start and end.
 * Accepts:
 * - Excel serial number (most common in custom functions)
 * - Date string
 * - JS Date
 */
function excelToUtcDayRange(excelVal) {
  if (excelVal === null || excelVal === undefined || excelVal === "") return null;

  let d;
  if (typeof excelVal === "number") {
    // Excel serial days since 1899-12-30
    const epoch = new Date(Date.UTC(1899, 11, 30));
    d = new Date(epoch.getTime() + excelVal * 86400000);
  } else if (excelVal instanceof Date) {
    d = excelVal;
  } else {
    d = new Date(excelVal);
  }

  if (isNaN(d.getTime())) return null;

  // Use UTC date-only boundaries to match "date-only" fields robustly
  const start = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()));
  const end = new Date(start.getTime() + 86400000);
  return { start, end };
}

async function fetchValue(identifier, excelDate) {
  const id = (identifier ?? "").toString().trim();
  const range = excelToUtcDayRange(excelDate);
  if (!id || !range) return "";

  const key = `${id}__${range.start.toISOString().slice(0, 10)}`;
  if (memo.has(key)) return memo.get(key);

  const filter =
    `fields/${CONFIG.fieldIdentifier} eq '${escapeODataString(id)}'` +
    ` and fields/${CONFIG.fieldDate} ge ${range.start.toISOString()}` +
    ` and fields/${CONFIG.fieldDate} lt ${range.end.toISOString()}`;

  const select = CONFIG.fieldValue;

  const url =
    `https://graph.microsoft.com/v1.0/sites/${CONFIG.siteId}` +
    `/lists/${CONFIG.listId}/items` +
    `?$expand=fields($select=${encodeURIComponent(select)})` +
    `&$filter=${encodeURIComponent(filter)}` +
    `&$top=1`;

  const data = await graphGet(url);
  const val = data.value?.[0]?.fields?.[CONFIG.fieldValue] ?? "";

  memo.set(key, val);
  return val;
}

/**
 * =TESLIN.DATA(identifier, date)
 * Supports scalars or ranges (spills).
 * Broadcasting rules:
 * - If one input is 1x1, it broadcasts over the other input's shape.
 * - Otherwise, result shape is max(rows), max(cols) with edge broadcasting.
 */
CustomFunctions.associate("TESLIN.DATA", async (identifier, date) => {
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
      tasks.push(fetchValue(idVal, dtVal).then(v => { out[r][c] = v; }));
    }
  }

  await Promise.all(tasks);
  return out;
});
