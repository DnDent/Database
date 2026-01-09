/* global CustomFunctions, msal */

// MSAL configuration
var msalConfig = {
  auth: {
    clientId: "a59c788a-3f0e-4fa0-a450-746734ef6fcd",
    authority: "https://login.microsoftonline.com/common",
    redirectUri: "https://dndent.github.io/Database/functions.html"
  },
  cache: {
    cacheLocation: "sessionStorage"
  }
};

var msalInstance = null;
var tokenRequest = {
  scopes: ["https://graph.microsoft.com/Sites.ReadWrite.All"]
};

// Initialize MSAL
function initMsal() {
  if (!msalInstance && typeof msal !== "undefined") {
    msalInstance = new msal.PublicClientApplication(msalConfig);
  }
}

// Get access token
async function getToken() {
  initMsal();
  if (!msalInstance) {
    throw new Error("MSAL not initialized");
  }
  
  try {
    // Try silent token acquisition first
    var accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      tokenRequest.account = accounts[0];
      var response = await msalInstance.acquireTokenSilent(tokenRequest);
      return response.accessToken;
    }
    
    // If no account, do interactive login
    var response = await msalInstance.acquireTokenPopup(tokenRequest);
    return response.accessToken;
  } catch (error) {
    throw new Error("Auth failed: " + error.message);
  }
}

CustomFunctions.associate("DATA", function (identifier, date) {
  return new CustomFunctions.StreamingInvocation(async (setResult) => {
    try {
      setResult([["Authenticating..."]]);
      var token = await getToken();
      setResult([["Auth OK! Token: " + token.substring(0, 20) + "..."]]);
    } catch (error) {
      setResult([["ERROR: " + error.message]]);
    }
  });
});
