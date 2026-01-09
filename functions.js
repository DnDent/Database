/* global CustomFunctions, OfficeRuntime */

CustomFunctions.associate("DATA", function (identifier, date) {
  return new CustomFunctions.CancelablePromise(async (resolve, reject) => {
    try {
      // Check if OfficeRuntime.auth exists
      if (typeof OfficeRuntime === "undefined" || !OfficeRuntime.auth) {
        resolve([["ERROR: OfficeRuntime.auth not available"]]);
        return;
      }
      
      // Try to get token
      var token = await OfficeRuntime.auth.getAccessToken({
        allowSignInPrompt: true,
        allowConsentPrompt: true,
        forMSGraphAccess: true
      });
      
      resolve([["Token: " + token.substring(0, 20) + "..."]]);
    } catch (error) {
      resolve([["ERROR: " + (error.code || error.message || String(error))]]);
    }
  });
});
