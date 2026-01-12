/* global CustomFunctions, OfficeRuntime */

var DATA_CACHE = {};
var TOKEN_CACHE = null;
var FUNCTION_URL = "https://vinhuys-function-crh8gsfwajc2d4dr.westeurope-01.azurewebsites.net/api/getData";

async function getToken() {
    if (TOKEN_CACHE) {
        return TOKEN_CACHE;
    }
    
    var token = await OfficeRuntime.auth.getAccessToken({
        allowSignInPrompt: true,
        allowConsentPrompt: true
    });
    
    TOKEN_CACHE = token;
    return token;
}

async function loadData(listName) {
    if (DATA_CACHE[listName]) {
        return DATA_CACHE[listName];
    }

    // Get SSO token
    var token = await getToken();
    
    var url = FUNCTION_URL + "?list=" + encodeURIComponent(listName);
    var response = await fetch(url, {
        headers: {
            "Authorization": "Bearer " + token
        }
    });

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
        var days = dateValue - 25569; // Days since Unix epoch
        var ms = days * 86400000;
        var date = new Date(ms);
        var year = date.getUTCFullYear();
        var month = String(date.getUTCMonth() + 1).padStart(2, "0");
        var day = String(date.getUTCDate()).padStart(2, "0");
        return year + "-" + month + "-" + day;
    }
    // Handle DD/MM/YYYY string from user input
    if (typeof dateValue === "string" && dateValue.includes("/")) {
        var parts = dateValue.split("/");
        if (parts.length === 3) {
            return parts[2] + "-" + parts[1].padStart(2, "0") + "-" + parts[0].padStart(2, "0");
        }
    }
    // Already YYYY-MM-DD or other format
    return String(dateValue);
}

function formatSharePointDate(spDate) {
    if (!spDate) return "";

    // If it's already a date-only string, keep it
    var s = String(spDate);
    if (!s.includes("T")) return s;

    // Convert ISO Z time to LOCAL date (fixes 23:00Z -> next day in CET)
    var d = new Date(s);
    var year = d.getFullYear();
    var month = String(d.getMonth() + 1).padStart(2, "0");
    var day = String(d.getDate()).padStart(2, "0");
    return year + "-" + month + "-" + day;
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
        var matrixMode = false;
        
        if (timeIsSingle && metIsSingle) {
            numRows = 1;
            numCols = 1;
        } else if (timeIsSingle) {
            numRows = metRows;
            numCols = metCols;
        } else if (metIsSingle) {
            numRows = timeRows;
            numCols = timeCols;
        } else if (timeIsVertical && metIsHorizontal) {
            // MATRIX MODE: time vertical, metric horizontal
            numRows = timeRows;
            numCols = metCols;
            matrixMode = true;
        } else if (timeIsHorizontal && metIsVertical) {
            // MATRIX MODE (transposed): time horizontal, metric vertical
            numRows = metRows;
            numCols = timeCols;
            matrixMode = true;
        } else {
            // Same orientation: paired lookup
            numRows = Math.max(timeRows, metRows);
            numCols = Math.max(timeCols, metCols);
        }
        
        // Get entity and type values (always use first cell)
        var ent = entity[0][0];
        var typeValue = type[0][0];
        var listName = determineList(typeValue);
        
        // Load data from cache or API (now with authentication)
        var data = await loadData(listName);
        
        // BUILD RESULT MATRIX
        var result = [];
        for (var row = 0; row < numRows; row++) {
            var resultRow = [];
            for (var col = 0; col < numCols; col++) {
                var t, m;
                
                if (matrixMode) {
                    if (timeIsVertical && metIsHorizontal) {
                        t = time[row][0];
                        m = metric[0][col];
                    } else {
                        t = time[0][col];
                        m = metric[row][0];
                    }
                } else {
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
