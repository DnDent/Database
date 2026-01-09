/* global CustomFunctions, OfficeRuntime */

async function getAccessToken() {
  try {
    // Use Office's built-in SSO to get a token
    var token = await OfficeRuntime.auth.getAccessToken({
      allowSignInPrompt: true,
      allowConsentPrompt: true,
      forMSGraphAccess: true
    });
    return token;
  } catch (error) {
    throw new Error("Auth failed: " + error.code);
  }
}

CustomFunctions.associate("DATA", function (identifier, date) {
  return new CustomFunctions.CancelablePromise(async (resolve, reject) => {
    try {
      var token = await getAccessToken();
      resolve([["Token: " + token.substring(0, 20) + "..."]]);
    } catch (error) {
      resolve([["ERROR: " + error.message]]);
    }
  });
});
