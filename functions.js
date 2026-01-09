/* global CustomFunctions, OfficeRuntime */

var DATA_CACHE = null;
var FUNCTION_URL = "https://vinhuys-function.azurewebsites.net/api/getData";

async function getAccessToken() {
    try {
        var token = await OfficeRuntime.auth.getAccessToken({
            allowSignInPrompt: true,
            allowConsentPrompt: true
        });
        return token;
    } catch (error) {
        throw new Error("Auth failed: " + error.message);
    }
}

async function loadData() {
    if (DATA_CACHE) {
        return DATA_CACHE;
    }

    var token = await getAccessToken();
    
    var response = await fetch(FUNCTION_URL, {
        headers: {
            'Authorization': 'Bearer ' + token
        }
    });

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
