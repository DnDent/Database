/* global CustomFunctions */
CustomFunctions.associate("DATA", function (identifier, date) {
  var url = "https://vinhuys.sharepoint.com/sites/VinHuys/_api/web/lists/getbytitle('Vinhuys database')/items";
  
  var xhttp = new XMLHttpRequest();
  return new Promise(function(resolve, reject) {
    xhttp.onreadystatechange = function() {
      if (xhttp.readyState !== 4) return;
      
      if (xhttp.status == 200) {
        resolve([["Success! Status: 200"]]);
      } else if (xhttp.status == 0) {
        resolve([["ERROR: Status 0 - CORS or network issue"]]);
      } else {
        resolve([["ERROR: Status " + xhttp.status]]);
      }
    };
    
    xhttp.onerror = function() {
      resolve([["ERROR: Network error occurred"]]);
    };
    
    xhttp.open("GET", url, true);
    xhttp.setRequestHeader("Accept", "application/json;odata=verbose");
    xhttp.send();
  });
});
