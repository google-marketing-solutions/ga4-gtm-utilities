/**
 * Fixture for the harness self-test: loaded FIRST (features/ sorts before
 * shared/ and utils/). It only references later files' declarations at call
 * time, which is legal in Apps Script's single global scope.
 */

/** @return {string} Uses a const from shared/ and a function from utils/. */
function featureA() {
  return SHARED_CONFIG.prefix + helperC('a');
}

/** @return {!Array<number>} Exercises the shared mutable top-level state. */
function featureCounter() {
  return [bump(), bump()];
}
