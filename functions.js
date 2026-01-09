/* global CustomFunctions */
CustomFunctions.associate("DATA", function (identifier, date) {
  try {
    var xhr = new XMLHttpRequest();
    // Try to call a simple public API (httpbin.org returns whatever you send)
    xhr.open("GET", "https://httpbin.org/get", false); // false = synchronous
    xhr.send();
    
    if (xhr.status === 200) {
      return [["HTTP request worked! Status: " + xhr.status]];
    } else {
      return [["HTTP failed with status: " + xhr.status]];
    }
  } catch (e) {
    return [["ERROR: " + e.message]];
  }
});
