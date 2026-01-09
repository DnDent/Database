/* global CustomFunctions, OfficeRuntime */

CustomFunctions.associate("DATA", function (identifier, date) {
  var steps = [];
  
  try {
    steps.push("1: Creating promise");
    
    return new CustomFunctions.CancelablePromise(async (resolve, reject) => {
      try {
        steps.push("2: Inside promise");
        
        // Check OfficeRuntime
        try {
          if (typeof OfficeRuntime === "undefined") {
            resolve([["Steps: " + steps.join(" -> ") + " -> FAIL: OfficeRuntime undefined"]]);
            return;
          }
          steps.push("3: OfficeRuntime OK");
        } catch (e) {
          resolve([["Steps: " + steps.join(" -> ") + " -> ERROR checking OfficeRuntime: " + e.message]]);
          return;
        }
        
        // Check auth
        try {
          if (!OfficeRuntime.auth) {
            resolve([["Steps: " + steps.join(" -> ") + " -> FAIL: auth missing"]]);
            return;
          }
          steps.push("4: auth OK");
        } catch (e) {
          resolve([["Steps: " + steps.join(" -> ") + " -> ERROR checking auth: " + e.message]]);
          return;
        }
        
        // Check getAccessToken
        try {
          if (!OfficeRuntime.auth.getAccessToken) {
            resolve([["Steps: " + steps.join(" -> ") + " -> FAIL: getAccessToken missing"]]);
            return;
          }
          steps.push("5: getAccessToken OK");
        } catch (e) {
          resolve([["Steps: " + steps.join(" -> ") + " -> ERROR checking getAccessToken: " + e.message]]);
          return;
        }
        
        // Try to get token
        try {
          steps.push("6: Calling getAccessToken");
          var token = await OfficeRuntime.auth.getAccessToken({
            allowSignInPrompt: true,
            allowConsentPrompt: true,
            forMSGraphAccess: true
          });
          steps.push("7: Token received");
          resolve([["SUCCESS! Steps: " + steps.join(" -> ") + " | Token: " + token.substring(0, 10)]]);
        } catch (e) {
          resolve([["Steps: " + steps.join(" -> ") + " -> ERROR calling getAccessToken: " + (e.code || e.message || String(e))]]);
          return;
        }
        
      } catch (error) {
        resolve([["Steps: " + steps.join(" -> ") + " -> EXCEPTION in promise: " + (error.message || String(error))]]);
      }
    });
    
  } catch (error) {
    return [["Steps: " + steps.join(" -> ") + " -> EXCEPTION creating promise: " + (error.message || String(error))]];
  }
});
