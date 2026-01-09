/* global CustomFunctions, msal */

CustomFunctions.associate("DATA", function (identifier, date) {
  try {
    // First, just test if we can return a simple value
    return [["Testing: " + identifier]];
  } catch (e) {
    return [["ERROR: " + (e.message || String(e))]];
  }
});
