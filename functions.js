/* global CustomFunctions */

function safeAssociate(name, fn) {
  try {
    CustomFunctions.associate(name, fn);
  } catch (e) {
    // Ignore registration errors (e.g., name not in metadata)
  }
}

// Keep 2 params so it matches your JSON signature, but ignore them.
function works(identifier, date) {
  return "WORKS!";
}

// Register both just in case the metadata expects either variant
safeAssociate("TESLIN.DATA", works);
safeAssociate("DATA", works);
