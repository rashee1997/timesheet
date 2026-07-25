/**
 * SHARED.GS — single canonical home for helpers used across multiple files.
 * All .gs files share one global scope, so duplicate definitions silently
 * resolve to whichever file loads last; keeping one copy here prevents that.
 */

/**
 * Neutralizes leading =, +, -, @ so free-text saved into a cell can never be
 * interpreted as a formula (formula/DDE injection via job order text, which
 * later also flows into the Excel export).
 */
function sanitizeSheetText(v) {
  var s = String(v == null ? '' : v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

/** Case-insensitive lookup of an employee's email in the configured map. */
function findEmailForEmployee(name, emailMap) {
  if (emailMap[name]) return emailMap[name];
  var lower = String(name).toLowerCase().trim();
  var keys = Object.keys(emailMap);
  for (var i = 0; i < keys.length; i++) {
    if (keys[i].toLowerCase().trim() === lower) return emailMap[keys[i]];
  }
  return null;
}

/** Two-row-buffer Levenshtein edit distance (memory-lean). */
function levenshtein(a, b) {
  var m = a.length, n = b.length;
  if (m === 0) return n;
  if (n === 0) return m;
  var prev = [];
  for (var j = 0; j <= n; j++) prev[j] = j;
  for (var i = 1; i <= m; i++) {
    var curr = [i];
    for (var j2 = 1; j2 <= n; j2++) {
      var cost = a.charAt(i - 1) === b.charAt(j2 - 1) ? 0 : 1;
      curr[j2] = Math.min(prev[j2] + 1, curr[j2 - 1] + 1, prev[j2 - 1] + cost);
    }
    prev = curr;
  }
  return prev[n];
}

/**
 * Runs fn under the script lock so read-modify-write cycles on Script
 * Properties (templates, employee emails, schedule config) can't lose
 * one writer's changes when two run concurrently.
 */
function withScriptLock_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    throw new Error('Another change is being saved right now. Wait a few seconds and try again.');
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}
