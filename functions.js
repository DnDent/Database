/* global CustomFunctions */
CustomFunctions.associate("DATA", function (identifier, date) {
  var url = "https://vinhuys.sharepoint.com/sites/VinHuys/_api/web/lists/getbytitle('Vinhuys database')/items";
  return [["URL ready: " + url.substring(0, 50) + "..."]];
});
