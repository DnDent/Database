/* global CustomFunctions, msal */

CustomFunctions.associate("DATA", function (identifier, date) {
  try {
    var hasMsal = (typeof msal !== "undefined") && msal && msal.PublicClientApplication;
    var msg = hasMsal ? "MSAL_OK" : "MSAL_MISSING";
    // Must return a matrix because your JSON says dimensionality: "matrix"
    return [[msg]];
  } catch (e) {
    return [["ERR: " + (e && e.message ? e.message : String(e))]];
  }
});
