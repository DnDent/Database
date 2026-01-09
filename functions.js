/* global CustomFunctions, OfficeRuntime */
CustomFunctions.associate("DATA", function (identifier, date) {
  // Use Office's built-in auth instead of MSAL
  // This is simpler and designed specifically for Office Add-ins
  return [["Using Office auth instead"]];
});
