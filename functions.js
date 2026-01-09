/* global CustomFunctions */
CustomFunctions.associate("DATA", function (identifier, date) {
  var url = "https://vinhuys.sharepoint.com/sites/VinHuys/_api/web/lists/getbytitle('Vinhuys database')/items";
  
  var xhttp = new XMLHttpRequest();
  return new Promise(function(resolve, reject) {
    xhttp.onreadystatechange = function() {
      if (xhttp.readyState !== 4) return;
      
      if (xhttp.status == 200) {
        resolve([["Success! Status: " + xhttp.status]]);
      } else {
        reject({ status: xhttp.status, statusText: xhttp.statusText });
      }
    };
    
    xhttp.open("GET", url, true);
    xhttp.setRequestHeader("Accept", "application/json;odata=verbose");
    xhttp.send();
  });
});
