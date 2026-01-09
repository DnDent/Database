/* global CustomFunctions */

function safeAssociate(name, fn) {
  try {
    CustomFunctions.associate(name, fn);
  } catch (e) {
    // Ignore "not in metadata" / duplicate registration errors
  }
}

safeAssociate("TESLIN.DATA", () => "WORKS!");
safeAssociate("DATA", () => "WORKS!");
