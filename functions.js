/* global CustomFunctions */

function safeAssociate(name, fn) {
  try {
    CustomFunctions.associate(name, fn);
  } catch (e) {
    // Ignore "not in metadata"/duplicate registration errors
  }
}

// Keep 2 params (because your JSON defines 2 parameters),
// and return a 2D array (because your JSON defines matrix output).
function works(identifier, date) {
  return [["WORKS!"]];
}

// Register common name variants so you don’t have to think about caching/name wiring.
safeAssociate("TESLIN.DATA", works);
safeAssociate("DATA", works);
safeAssociate("TESLIN.TESLIN.DATA", works);
