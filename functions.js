/* global CustomFunctions */

var DATA_CACHE = {};
var FUNCTION_URL = "https://vinhuys-function-crh8gsfwajc2d4dr.westeurope-01.azurewebsites.net/api/getData";

async function loadData(listName) {
    if (DATA_CACHE[listName]) {
        return DATA_CACHE[listName];
    }

    var url = FUNCTION_URL + "?list=" + encodeURIComponent(listName);
    var response = await fetch(url);

    if (!response.ok) {
        var errorText = await response.text();
        throw new Error("API error: " + response.status + " " + errorText);
    }

    var data = await response.json();
    DATA_CACHE[listName] = data;
    return data;
}

function formatDate(dateValue) {
    // Handle Excel date serial numbers
    if (typeof dateValue === "number") {
        var excelEpoch = new Date(1899, 11, 30);
        var date = new Date(excelEpoch.getTime() + dateValue * 86400000);
        return date.toISOString().split("T")[0]; // Returns YYYY-MM-DD
    }
    // Handle DD/MM/YYYY string from user input
    if (typeof dateValue === "string" && dateValue.includes("/")) {
        var parts = dateValue.split("/");
        if (parts.length === 3) {
            // DD/MM/YYYY -> YYYY-MM-DD
            return parts[2] + "-" + parts[1].padStart(2, "0") + "-" + parts[0].padStart(2, "0");
        }
    }
    // Already YYYY-MM-DD or other format
    return String(dateValue);
}

function formatSharePointDate(spDate) {
    // Convert SharePoint ISO date "2026-01-12T08:00:00Z" to "YYYY-MM-DD"
    if (!spDate) return "";
    return spDate.split("T")[0];
}

function lookupValue(data, entity, time, metric, listName) {
    var formattedTime = formatDate(time);
    
    for (var i = 0; i < data.length; i++) {
        var row = data[i];
        var match = false;
        
        if (listName === "Fund.data") {
            // Match on Identifier + Date
            var rowDate = formatSharePointDate(row.Date);
            match = row.Identifier === entity && rowDate === formattedTime;
        } else if (listName === "Companies") {
            // Match on Company + Period (and version if applicable)
            match = row.Company === entity && row.Period === time;
        }
        // Add more list types here as needed
        
        if (match) {
            // Return the requested metric
            if (row.hasOwnProperty(metric)) {
                return row[metric];
            } else {
                return "Metric not found: " + metric;
            }
        }
    }
    return "Not found";
}

function determineList(type) {
    var typeLower = type.toLowerCase();
    
    if (typeLower === "fund" || typeLower === "data") {
        return "Fund.data";
    } else if (typeLower === "positions") {
        return "Fund.positions";
    } else if (typeLower === "trades") {
        return "Fund.trades";
    } else if (typeLower === "benchmark" || typeLower === "benchmarks") {
        return "Benchmarks";
    } else if (typeLower === "company" || typeLower === "companies") {
        return "Companies";
    } else {
        // Not a known source - assume it's a company model version
        return "Companies";
    }
}

async function teslinGet(entity, type, time, metric, version) {
    try {
        // Get dimensions from time and metric (the spilling parameters)
        var timeRows = time.length;
        var timeCols = time[0].length;
        var metRows = metric.length;
        var metCols = metric[0].length;
        
        // Check if inputs are single cells
        var timeIsSingle = (timeRows === 1 && timeCols === 1);
        var metIsSingle = (metRows === 1 && metCols === 1);
        
        // Determine orientation
        var timeIsVertical = (timeRows > 1 && timeCols === 1);
        var timeIsHorizontal = (timeRows === 1 && timeCols > 1);
        var metIsVertical = (metRows > 1 && metCols === 1);
        var metIsHorizontal = (metRows === 1 && metCols > 1);
        
        // DIMENSION MISMATCH CHECK
        // Error only if: both are ranges, same orientation, different sizes
        if (!timeIsSingle && !metIsSingle) {
            if (timeIsVertical && metIsVertical && timeRows !== metRows) {
                return [["Error: Dimension mismatch"]];
            }
            if (timeIsHorizontal && metIsHorizontal && timeCols !== metCols) {
                return [["Error: Dimension mismatch"]];
            }
        }
        
        // DETERMINE OUTPUT DIMENSIONS
        var numRows, numCols;
        var matrixMode = false; // Flag to track if we're in matrix mode
        
        if (timeIsSingle && metIsSingle) {
            // Both single: 1x1 output
            numRows = 1;
            numCols = 1;
        } else if (timeIsSingle) {
            // Only metric varies: output matches metric shape
            numRows = metRows;
            numCols = metCols;
        } else if (metIsSingle) {
            // Only time varies: output matches time shape
            numRows = timeRows;
            numCols = timeCols;
        } else if (timeIsVertical && metIsHorizontal) {
            // MATRIX MODE: time vertical, metric horizontal
            // Output: rows = time count, cols = metric count
            numRows = timeRows;
            numCols = metCols;
            matrixMode = true;
        } else if (timeIsHorizontal && metIsVertical) {
            // MATRIX MODE (transposed): time horizontal, metric vertical
            // Output: rows = metric count, cols = time count
            numRows = metRows;
            numCols = timeCols;
            matrixMode = true;
        } else {
            // Same orientation: paired lookup (like old code)
            numRows = Math.max(timeRows, metRows);
            numCols = Math.max(timeCols, metCols);
        }
        
        // Get entity and type values (always use first cell)
        var ent = entity[0][0];
        var typeValue = type[0][0];
        var listName = determineList(typeValue);
        
        // Load data from cache or API
        var data = await loadData(listName);
        
        // BUILD RESULT MATRIX
        var result = [];
        for (var row = 0; row < numRows; row++) {
            var resultRow = [];
            for (var col = 0; col < numCols; col++) {
                var t, m;
                
                if (matrixMode) {
                    // Matrix mode: time determines row, metric determines col
                    if (timeIsVertical && metIsHorizontal) {
                        t = time[row][0];
                        m = metric[0][col];
                    } else {
                        // timeIsHorizontal && metIsVertical
                        t = time[0][col];
                        m = metric[row][0];
                    }
                } else {
                    // Non-matrix mode: same logic as old working code
                    t = timeIsSingle ? time[0][0] : time[row][col];
                    m = metIsSingle ? metric[0][0] : metric[row][col];
                }
                
                resultRow.push(lookupValue(data, ent, t, m, listName));
            }
            result.push(resultRow);
        }
        
        return result;
        
    } catch (error) {
        return [["Error: " + error.message]];
    }
}

CustomFunctions.associate("GET", teslinGet);
