/* global CustomFunctions, msal */

var CONFIG = {
  tenantId: "37f72dd3-3cb6-4a3c-9f03-a0e9ad17c700",
  clientId: "a59c788a-3f0e-4fa0-a450-746734ef6fcd",

  siteId: "vinhuys.sharepoint.com,9ed18a60-56d2-4f3a-8d2f-eeea9517f7a1,039a6441-1114-4ac1-b5c2-93e74de479c2",
  listId: "882d9fb5-7971-45d9-a219-de43b6545661",

  fieldIdentifier: "Identifier",
  fieldDate: "Date",
  fieldValue: "Value",

  redirectUri: "https://dndent.github.io/Database/functions.html"
};

var GRAPH_SCOPE = ["https://graph.microsoft.com/.default"];

var msalApp = null;
var tokenCache = null;

// simple cache: key -> value
var memo = {};

function shortErr(e) {
  var msg = "Unknown error";
  try {
    msg = (e && (e.message || e.errorMessage || String(e))) || msg;
  } catch (x) {}
  if (msg.length > 220) msg = msg.slice(0, 220) + "...";
  return msg;
}

function escapeODataString(s) {
  return String(s).replace(/'/g, "''");
}

function to2D(v) {
  return Array.isArray(v) ? v : [[v]];
}

function shape2D(a) {
  return { rows: a.length, cols: (a[0] ? a[0].length : 0) };
}

/**
 * Parse Excel date:
 * - number: Excel serial
 * - Date object
 * - string: supports NL dd-mm-yyyy (or /) and ISO
 */
function parseExcelDateLocal(excelVal) {
  if (excelVal === null || excelVal === undefined || excelVal === "") return null;

  if (typeof excelVal === "number" && isFinite(excelVal)) {
    var epoch = new Date(1899, 11, 30); // local
    return new Date(epoch.getTime() + excelVal * 86400000);
  }

  if (excelVal instanceof Date) {
    return isNaN(excelVal.getTime()) ? null : excelVal;
  }

  var s = String(excelVal).trim();
  if (!s) return null;

  var m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m) {
    var dd = Number(m[1]);
    var mm = Number(m[2]);
    var yyyy = Number(m[3]);
    var d = new Date(yyyy, mm - 1, dd);
    return isNaN(d.getTime()) ? null : d;
  }

  var d2 = new Date(s);
  return isNaN(d2.getTime()) ? null : d2;
}

/**
 * Use local-midnight boundaries (prevents SharePoint date-only timezone misses)
 * Returns ISO strings.
 */
function localDayRangeIso(excelVal) {
  var d = parseExcelDateLocal(excelVal);
  if (!d) return null;

  var startLocal = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  var endLocal = new Date(startLocal.getTime() + 86400000);

  return { startIso: startLocal.toISOString(), endIso: endLocal.toISOString() };
}

function ensureMsalApp() {
  try {
    if (typeof msal === "undefined" || !msal.PublicClientApplication) {
      return Promise.reject(new Error("MSAL is missing. Check functions.html loads msal-browser."));
    }

    if (msalApp) return Promise.resolve(msalApp);

    msalApp = new msal.PublicClientApplication({
      auth: {
        clientId: CONFIG.clientId,
        authority: "https://login.microsoftonline.com/" + CONFIG.tenantId,
        redirectUri: CONFIG.redirectUri
      },
      cache: { cacheLocation: "sessionStorage" }
    });

    if (typeof msalApp.initialize === "function") {
      return msalApp.initialize().then(function () { return msalApp; });
    }

    return Promise.resolve(msalApp);
  } catch (e) {
    return Promise.reject(e);
  }
}

function getToken() {
  if (tokenCache) return Promise.resolve(tokenCache);

  return ensureMsalApp().then(function (app) {
    var accounts = app.getAllAccounts ? app.getAllAccounts() : [];
    if (accounts && accounts.length) {
      return app.acquireTokenSilent({ scopes: GRAPH_SCOPE, account: accounts[0] })
        .then(function (r) {
          tokenCache = r.accessToken;
          return tokenCache;
        })
        .catch(function () {
          // fallthrough to interactive
          return app.loginPopup({ scopes: GRAPH_SCOPE })
            .then(function () {
              var acct = app.getAllAccounts()[0];
              return app.acquireTokenSilent({ scopes: GRAPH_SCOPE, account: acct });
            })
            .then(function (r2) {
              tokenCache = r2.accessToken;
              return tokenCache;
            });
        });
    }

    // No account yet -> interactive
    return app.loginPopup({ scopes: GRAPH_SCOPE })
      .then(function () {
        var acct2 = app.getAllAccounts()[0];
        return app.acquireTokenSilent({ scopes: GRAPH_SCOPE, account: acct2 });
      })
      .then(function (r3) {
        tokenCache = r3.accessToken;
        return tokenCache;
      });
  });
}

function graphGet(url) {
  return getToken().then(function (token) {
    return fetch(url, { headers: { Authorization: "Bearer " + token } });
  }).then(function (r) {
    if (!r.ok) {
      return r.text().catch(function () { return ""; }).then(function (txt) {
        throw new Error("Graph " + r.status + ": " + (txt || r.statusText));
      });
    }
    return r.json();
  });
}

function fetchValue(identifier, excelDate) {
  var id = (identifier === null || identifier === undefined) ? "" : String(identifier).trim();
  var range = localDayRangeIso(excelDate);
  if (!id || !range) return Promise.resolve("");

  var key = id + "__" + range.startIso.slice(0, 10);
  if (memo.hasOwnProperty(key)) return Promise.resolve(memo[key]);

  var filter =
    "fields/" + CONFIG.fieldIdentifier + " eq '" + escapeODataString(id) + "'" +
    " and fields/" + CONFIG.fieldDate + " ge " + range.startIso +
    " and fields/" + CONFIG.fieldDate + " lt " + range.endIso;

  var select = CONFIG.fieldIdentifier + "," + CONFIG.fieldDate + "," + CONFIG.fieldValue;

  var url =
    "https://graph.microsoft.com/v1.0/sites/" + CONFIG.siteId +
    "/lists/" + CONFIG.listId + "/items" +
    "?$expand=fields($select=" + encodeURIComponent(select) + ")" +
    "&$filter=" + encodeURIComponent(filter) +
    "&$top=1";

  return graphGet(url).then(function (data) {
    if (!data || !data.value || !data.value.length) return "";

    var fields = data.value[0] && data.value[0].fields ? data.value[0].fields : {};
    var val = fields[CONFIG.fieldValue];

    // If internal field name differs, show keys (fastest debug)
    if (val === undefined || val === null || val === "") {
      var keys = [];
      for (var k in fields) if (fields.hasOwnProperty(k)) keys.push(k);
      return "ERR: No '" + CONFIG.fieldValue + "' field. Keys: " + keys.join(", ");
    }

    memo[key] = val;
    return val;
  });
}

// IMPORTANT: only associate "DATA" (no TESLIN.TESLIN.DATA, no TESLIN.DATA)
CustomFunctions.associate("DATA", function (identifier, date) {
  try {
    var ids = to2D(identifier);
    var dts = to2D(date);

    var s1 = shape2D(ids);
    var s2 = shape2D(dts);

    var rows = Math.max(s1.rows, s2.rows);
    var cols = Math.max(s1.cols, s2.cols);

    var out = [];
    for (var r = 0; r < rows; r++) {
      var row = [];
      for (var c = 0; c < cols; c++) row.push("");
      out.push(row);
    }

    var promises = [];
    for (var rr = 0; rr < rows; rr++) {
      for (var cc = 0; cc < cols; cc++) {
        (function (r2, c2) {
          var idVal = ids[Math.min(r2, s1.rows - 1)][Math.min(c2, s1.cols - 1)];
          var dtVal = dts[Math.min(r2, s2.rows - 1)][Math.min(c2, s2.cols - 1)];

          var p = fetchValue(idVal, dtVal)
            .then(function (v) { out[r2][c2] = v; })
            .catch(function (e) { out[r2][c2] = "ERR: " + shortErr(e); });

          promises.push(p);
        })(rr, cc);
      }
    }

    return Promise.all(promises).then(function () { return out; })
      .catch(function (e) { return [["ERR: " + shortErr(e)]]; });
  } catch (e2) {
    return [["ERR: " + shortErr(e2)]];
  }
});
