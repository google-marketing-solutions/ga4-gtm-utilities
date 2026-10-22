/**
 * Fixture for the harness self-test: loaded LAST. Top-level code may use
 * declarations from files loaded earlier (shared/config.js).
 */

const helperArrow = (x) => `[${x}]`;

/** @return {string} */
function helperC(x) {
  return helperArrow(x).toUpperCase();
}

const DERIVED = SHARED_CONFIG.prefix + 'derived';
