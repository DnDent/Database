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

    // IMPORTANT: Put your GitHub Pages functions.html here and register it as a Redirect URI in Entra.
    // Example: https://dndent.github.io/Database/functions.html
    redirectUri: "https://dndent.github.io/Database/functions.html"
  };

  const GRAPH_SCOPE = ["https://graph.microsoft.com/.default"];

  let msalApp = null;
  let tokenCache = null;
  const memo = new Map(); // cache results per (id, day)

  function shortErr(e) {
    const msg = (e && (e.message || e.errorMessage || e.name || String(e))) || "Unknown error";
    return msg.length > 240 ? msg.slice(0, 240) + "…" : msg;
  }

  // If runtime errors occur, Excel may still show #VALUE, but this helps during web devtools debugging.
  try {
    CustomFunctions.onRuntimeError = (err) => {
      // eslint-disable-next-line no-console
      console.error("CustomFunctions runtime error:", err);
    };
  } catch (_) {}

  function loadScript(url) {
    return new Promise((resolve, reject) => {
      try {
        const s = document.createElement("script");
        s.src = url;
        s.async = true;
        s.onload = resolve;
        s.onerror = () => reject(new Error(`Failed to load script: ${url}`));
        document.head.appendChild(s);
      } catch (e) {
        reject(e);
      }
    });
  }

  async function ensureMsal() {
    // If msal already exists, great.
    if (typeof msal !== "undefined" && msal?.PublicClientApplication) return;

    // Load MSAL from Microsoft's CDN (stable-ish pinned version).
    // You can bump the version later; pinning avoids sudden breaking changes.
    await loadScript("https://alcdn.msauth.net/browser/2.38.0/js/msal-browser.min.js");

    if (typeof msal === "undefined" || !msal?.PublicClientApplication) {
      throw new Error("MSAL did not load (msal is undefined after loading CDN script).");
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

    // Some MSAL versions require initialize()
    if (typeof msalApp.initialize === "function") {
      await msalApp.initialize();
    }

    return msalApp;
  }

  async function getToken() {
    if (tokenCache) return tokenCac
