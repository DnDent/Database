/* global CustomFunctions */

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

function lookupValue(data, identifier, date) {
    for (var i = 0; i < data.length; i++) {
        if (data[i].Identifier === identifier && data[i].Date === date) {
            return data[i].Value;
        }
    }
    return "Not found";
}

async function getData(identifier, date) {
    try {
        var data = await loadData();
        
        // Get dimensions - use MAX of both inputs
        var numRows = Math.max(identifier.length, date.length);
        var numCols = Math.max(identifier[0].length, date[0].length);
        
        // Build result matrix
        var result = [];
        for (var row = 0; row < numRows; row++) {
            var resultRow = [];
            for (var col = 0; col < numCols; col++) {
                // If one input is smaller, reuse its value (e.g., single cell applies to all)
                var idRow = row < identifier.length ? row : 0;
                var idCol = col < identifier[0].length ? col : 0;
                var dtRow = row < date.length ? row : 0;
                var dtCol = col < date[0].length ? col : 0;
                
                var id = identifier[idRow][idCol];
                var dt = date[dtRow][dtCol];
                resultRow.push(lookupValue(data, id, dt));
            }
            result.push(resultRow);
        }
        
        return result;
        
    } catch (error) {
        return [["Error: " + error.message]];
    }
}

CustomFunctions.associate("DATA", getData);
