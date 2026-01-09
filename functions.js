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

CustomFunctions.associate("DATA", function (identifier, date, invocation) {
  // Check if MSAL is available
  if (typeof msal === "undefined") {
    return [["ERROR: MSAL library not loaded"]];
  }
  
  // Initialize MSAL instance
  if (!msalInstance) {
    try {
      msalInstance = new msal.PublicClientApplication(msalConfig);
    } catch (e) {
      return [["ERROR: MSAL init failed - " + e.message]];
    }
  }
  
  return [["MSAL initialized OK"]];
});
