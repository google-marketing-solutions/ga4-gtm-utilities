/**
 * Fixture for the harness self-test: top-level const/let in one file must be
 * visible to every other file (V8 Apps Script runtime semantics).
 */

const SHARED_CONFIG = {prefix: 'cfg:'};
let sharedCounter = 0;

/** @return {number} */
function bump() {
  sharedCounter++;
  return sharedCounter;
}
