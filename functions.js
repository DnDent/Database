/* global CustomFunctions */

var DATA_CACHE = {};
var FUNCTION_URL = "https://vinhuys-function-crh8gsfwajc2d4dr.westeurope-01.azurewebsites.net/api/getData";

// Known source types that route to specific lists
var KNOWN_SOURCES = ["fund", "positions", "trades", "benchmark", "benchmarks", "company", "companies"];

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
        var day = String(date.getDate()).padStart(2, "0");
        var month = String(date.getMonth() + 1).padStart(2, "0");
        var year = date.getFullYear();
        return day + "/" + month + "/" + year;
    }
    // Already a string
    return String(dateValue);
}

function lookupValue(data, entity, time, metric, listName) {
    var formattedTime = formatDate(time);
    
    for (var i = 0; i < data.length; i++) {
        var row = data[i];
        var match = false;
        
        if (listName === "Fund.data") {
            // Match on Fund + NAV date
            match = row.Fund === entity && formatDate(row.NAVdate) === formattedTime;
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
        // Get dimensions from all inputs
        var entRows = entity.length;
        var entCols = entity[0].length;
        var typeRows = type.length;
        var typeCols = type[0].length;
        var timeRows = time.length;
        var timeCols = time[0].length;
        var metRows = metric.length;
        var metCols = metric[0].length;
        
        // Check if inputs are single cells
        var entIsSingle = (entRows === 1 && entCols === 1);
        var typeIsSingle = (typeRows === 1 && typeCols === 1);
        var timeIsSingle = (timeRows === 1 && timeCols === 1);
        var metIsSingle = (metRows === 1 && metCols === 1);
        
        // For spilling: time and metric are the ones that typically vary
        // Entity and type are usually single values
        
        // Dimension mismatch check for time and metric (the spilling parameters)
        if (!timeIsSingle && !metIsSingle) {
            // Both are ranges - check if they're compatible
            var timeSameOrientation = (timeRows > 1 && timeCols === 1) || (timeRows === 1 && timeCols > 1);
            var metSameOrientation = (metRows > 1 && metCols === 1) || (metRows === 1 && metCols > 1);
            
            var timeIsVertical = timeRows > 1 && timeCols === 1;
            var timeIsHorizontal = timeRows === 1 && timeCols > 1;
            var metIsVertical = metRows > 1 && metCols === 1;
            var metIsHorizontal = metRows === 1 && metCols > 1;
            
            // If same orientation and different sizes = error
            if (timeIsVertical && metIsVertical && timeRows !== metRows) {
                return [["Error: Dimension mismatch"]];
            }
            if (timeIsHorizontal && metIsHorizontal && timeCols !== metCols) {
                return [["Error: Dimension mismatch"]];
            }
        }
        
        // Determine output dimensions based on time and metric orientation
        var numRows, numCols;
        
        if (timeIsSingle && metIsSingle) {
            numRows = 1;
            numCols = 1;
        } else if (timeIsSingle) {
            numRows = metRows;
            numCols = metCols;
        } else if (metIsSingle) {
            numRows = timeRows;
            numCols = timeCols;
        } else {
            // Both are ranges - create matrix based on orientation
            var timeIsVertical = timeRows > 1 && timeCols === 1;
            var metIsHorizontal = metRows === 1 && metCols > 1;
            
            if (timeIsVertical && metIsHorizontal) {
                // Time vertical, metric horizontal = matrix
                numRows = timeRows;
                numCols = metCols;
            } else if (!timeIsVertical && !metIsHorizontal) {
                // Time horizontal, metric vertical = matrix (transposed)
                numRows = metRows;
                numCols = timeCols;
            } else {
                // Same orientation = paired lookup
                numRows = Math.max(timeRows, metRows);
                numCols = Math.max(timeCols, metCols);
            }
        }
        
        // Get the list name from type
        var typeValue = type[0][0];
        var listName = determineList(typeValue);
        
        // Load data from cache or API
        var data = await loadData(listName);
        
        // Build result matrix
        var result = [];
        for (var row = 0; row < numRows; row++) {
            var resultRow = [];
            for (var col = 0; col < numCols; col++) {
                // Get entity value (usually single)
                var ent = entIsSingle ? entity[0][0] : entity[Math.min(row, entRows - 1)][Math.min(col, entCols - 1)];
                
                // Get time value
                var t;
                if (timeIsSingle) {
                    t = time[0][0];
                } else if (timeRows > 1 && timeCols === 1) {
                    // Vertical time range
                    t = time[Math.min(row, timeRows - 1)][0];
                } else if (timeRows === 1 && timeCols > 1) {
                    // Horizontal time range
                    t = time[0][Math.min(col, timeCols - 1)];
                } else {
                    t = time[Math.min(row, timeRows - 1)][Math.min(col, timeCols - 1)];
                }
                
                // Get metric value
                var m;
                if (metIsSingle) {
                    m = metric[0][0];
                } else if (metRows > 1 && metCols === 1) {
                    // Vertical metric range
                    m = metric[Math.min(row, metRows - 1)][0];
                } else if (metRows === 1 && metCols > 1) {
                    // Horizontal metric range
                    m = metric[0][Math.min(col, metCols - 1)];
                } else {
                    m = metric[Math.min(row, metRows - 1)][Math.min(col, metCols - 1)];
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
