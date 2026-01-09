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
        
        var idRows = identifier.length;
        var idCols = identifier[0].length;
        var dtRows = date.length;
        var dtCols = date[0].length;
        
        // Check for dimension mismatch (unless one is a single cell)
        var idIsSingle = (idRows === 1 && idCols === 1);
        var dtIsSingle = (dtRows === 1 && dtCols === 1);
        
        if (!idIsSingle && !dtIsSingle && (idRows !== dtRows || idCols !== dtCols)) {
            return [["Error: Dimension mismatch - ranges must be same size or single cell"]];
        }
        
        // Use dimensions from the larger input (or the non-single one)
        var numRows = Math.max(idRows, dtRows);
        var numCols = Math.max(idCols, dtCols);
        
        // Build result matrix
        var result = [];
        for (var row = 0; row < numRows; row++) {
            var resultRow = [];
            for (var col = 0; col < numCols; col++) {
                var id = idIsSingle ? identifier[0][0] : identifier[row][col];
                var dt = dtIsSingle ? date[0][0] : date[row][col];
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
