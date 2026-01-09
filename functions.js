var DATA_CACHE = null;
var BLOB_URL = "https://vinhuysstorage.blob.core.windows.net/data/vinhuys.json?sp=r&st=2026-01-09T14:30:06Z&se=2027-01-09T22:45:06Z&spr=https&sv=2024-11-04&sr=b&sig=2j6d753ENDCFlkHpxZQcifIri54p2TEWYGS2jZjHTk8%3D";

function getData(identifier, date) {
  return fetch(BLOB_URL)
    .then(function(response) { 
      if (!response.ok) {
        return "Fetch error: " + response.status;
      }
      return response.json(); 
    })
    .then(function(data) {
      if (typeof data === "string") {
        return data;
      }
      for (var i = 0; i < data.length; i++) {
        if (data[i].Identifier === identifier && data[i].Date === date) {
          return data[i].Value;
        }
      }
      return "Not found";
    })
    .catch(function(error) {
      return "Error: " + error.message;
    });
}

CustomFunctions.associate("DATA", getData);
