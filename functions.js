/* global CustomFunctions */
CustomFunctions.associate("DATA", function (identifier, date) {
  try {
    var xhr = new XMLHttpRequest();
    var url = "https://vinhuys.sharepoint.com/sites/VinHuys/_api/web/lists/getbytitle('Vinhuys database')/items";
    xhr.open("GET", url, false);
    xhr.setRequestHeader("Accept", "application/json;odata=verbose");
    xhr.send();
    
    return [["Status: " + xhr.status]];
  } catch (e) {
    return [["ERROR: " + e.message]];
  }
});
