/* global CustomFunctions, OfficeRuntime */

async function getToken() {
    var steps = [];
    
    try {
        // Step 1: Check if OfficeRuntime exists
        steps.push("Step 1: Checking OfficeRuntime...");
        if (typeof OfficeRuntime === "undefined") {
            return "FAIL Step 1: OfficeRuntime is undefined";
        }
        steps.push("Step 1: OK - OfficeRuntime exists");
        
        // Step 2: Check if auth exists
        steps.push("Step 2: Checking OfficeRuntime.auth...");
        if (!OfficeRuntime.auth) {
            return "FAIL Step 2: OfficeRuntime.auth is undefined";
        }
        steps.push("Step 2: OK - OfficeRuntime.auth exists");
        
        // Step 3: Check if getAccessToken exists
        steps.push("Step 3: Checking getAccessToken...");
        if (!OfficeRuntime.auth.getAccessToken) {
            return "FAIL Step 3: getAccessToken is undefined";
        }
        steps.push("Step 3: OK - getAccessToken exists");
        
        // Step 4: Try to get token
        steps.push("Step 4: Calling getAccessToken...");
        var token = await OfficeRuntime.auth.getAccessToken({
            allowSignInPrompt: true,
            allowConsentPrompt: true
        });
        
        // Step 5: Check token
        steps.push("Step 5: Checking token...");
        if (!token) {
            return "FAIL Step 5: Token is empty/null";
        }
        
        // Success!
        return "SUCCESS! Token length: " + token.length + " | First 20 chars: " + token.substring(0, 20) + "...";
        
    } catch (error) {
        // Detailed error info
        var errorInfo = "FAIL at " + steps[steps.length - 1] + " | ";
        errorInfo += "Error code: " + (error.code || "none") + " | ";
        errorInfo += "Message: " + (error.message || String(error));
        return errorInfo;
    }
}

CustomFunctions.associate("TOKEN", getToken);