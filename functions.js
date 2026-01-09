/* global CustomFunctions, msal */
CustomFunctions.associate("DATA", function (identifier, date) {
  var hasMsal = (typeof msal !== "undefined");
  return [[hasMsal ? "MSAL loaded" : "MSAL missing"]];
});
