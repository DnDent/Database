/* global CustomFunctions */

// SharePoint site and list details
var siteUrl = "https://vinhuys.sharepoint.com/sites/VinHuys";
var listName = "Vinhuys database";

CustomFunctions.associate("DATA", function (identifier, date) {
  // Build REST API URL to get list items
  var apiUrl = siteUrl + "/_api/web/lists/getbytitle('" + listName + "')/items";
  
  // For now, just return the URL we would call
  return [["Will call: " + apiUrl]];
});
