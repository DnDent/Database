/* global CustomFunctions, msal */

CustomFunctions.associate("DATA", function (identifier, date) {
  try {
    // Check if MSAL is available
    if (typeof msal === "undefined") {
      return [["ERROR: MSAL library not loaded"]];
    }
    
    return [["MSAL library loaded OK"]];
  } catch (e) {
    return [["ERROR: " + (e.message || String(e))]];
  }
});
