/* global CustomFunctions */
CustomFunctions.associate("DATA", function (identifier, date) {
  // Check if XMLHttpRequest exists
  if (typeof XMLHttpRequest !== "undefined") {
    return [["XMLHttpRequest available"]];
  } else {
    return [["XMLHttpRequest NOT available"]];
  }
});
