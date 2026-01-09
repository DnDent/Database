/* global CustomFunctions, OfficeRuntime */

var DATA_CACHE = null;
var FUNCTION_URL = "https://vinhuys-function-crh8gsfwajc2d4dr.westeurope-01.azurewebsites.net/api/getData";

async function loadData() {
    if (DATA_CACHE) {
        return DATA_CACHE;
    }

    var response = await fetch(FUNCTION_URL);

    if (!response.ok) {
        var errorText = await response.text();
        throw new Error("API error: " + response.status + " " + errorText);
    }

    var data = await response.json();
    DATA_CACHE = data;
    return data;
}

async function getData(identifier, date) {
    try {
        var data = await loadData();
        
        for (var i = 0; i < data.length; i++) {
            if (data[i].Identifier === identifier && data[i].Date === date) {
                return data[i].Value;
            }
        }
        return "Not found";
    } catch (error) {
        return "Error: " + error.message;
    }
}

CustomFunctions.associate("DATA", getData);
