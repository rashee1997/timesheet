/**
 * WEB APP JSON API — bridge for the standalone (Vercel) frontend.
 * ---------------------------------------------------------------------
 * Deploy this project as a Web App: Execute as "Me", Who has access "Anyone".
 *
 * Trust boundary: the Next.js proxy (app/api/gas/route.ts) is the one that
 * verifies the caller's Google ID token (via google-auth-library, since
 * Apps Script has no npm access to do that properly itself — see the design
 * doc). This file verifies an HMAC-SHA256 signature computed with a secret
 * only the Next.js server process knows, then trusts the actorEmail it was
 * handed for audit logging.
 *
 * Signature scheme (mirrored in the proxy):
 *   signature = base64(HMAC-SHA256(secret,
 *     timestamp + "." + action + "." + (actorEmail || "") + "." + JSON.stringify(payload || {})))
 * Requests older/newer than API_MAX_CLOCK_SKEW_MS are rejected (replay
 * protection). A legacy body.sharedSecret compare remains as a TEMPORARY
 * fallback until the signed frontend deploy is confirmed live — remove it then.
 *
 * Apps Script's doPost(e) cannot read custom HTTP headers (Google Issue
 * Tracker #67764685, official won't-fix), so auth travels in the JSON body.
 */

var API_SHARED_SECRET_PROPERTY = 'GAS_SHARED_SECRET';
var API_MAX_CLOCK_SKEW_MS = 5 * 60 * 1000;

/** Wraps raw-value handlers (arrays, plain objects, booleans) in the standard envelope. */
function apiData_(value) {
  return { success: true, data: value };
}

/** action name -> handler(payload, actorEmail). Whitelist — nothing outside this map is reachable. */
var API_ACTIONS = {
  getEmployeesForSheet: function (p) { return apiData_(getEmployeesForSheet(p.sheetName)); },
  getJobOrderSuggestions: function () { return apiData_(getJobOrderSuggestions()); },
  listTimesheetSheets: function () { return apiData_(listTimesheetSheets()); },
  submitTimeEntry: function (p, actorEmail) {
    return submitTimeEntry(Object.assign({}, p.formData, { actorEmail: actorEmail }));
  },
  submitBulkTimeEntries: function (p, actorEmail) {
    return submitBulkTimeEntries(Object.assign({}, p.formData, { actorEmail: actorEmail }));
  },
  undoLastSave: function (p, actorEmail) { return undoLastSave(actorEmail); },
  getEntryForEdit: function (p) { return getEntryForEdit(p); },
  updateTimeEntry: function (p, actorEmail) {
    return updateTimeEntry(Object.assign({}, p.formData, { actorEmail: actorEmail }));
  },
  getShiftTemplates: function () { return apiData_(getShiftTemplates()); },
  saveShiftTemplate: function (p) { return saveShiftTemplate(p.name, p.startTime, p.endTime, p.jobOrder); },
  deleteShiftTemplate: function (p) { return deleteShiftTemplate(p.id); },
  checkAiCredentials: function () { return apiData_(checkAiCredentials()); },
  parseNaturalLanguageEntries: function (p) { return parseNaturalLanguageEntries(p.text, p.sheetName); },
  parseNaturalLanguageEntriesWithImage: function (p) {
    return parseNaturalLanguageEntriesWithImage(p.text, p.imageBase64, p.mimeType, p.sheetName);
  },
  getRecentEntriesPreview: function (p) { return apiData_(getRecentEntriesPreview(p.sheetName)); },

  // ---- OT reports, scheduling, company settings ----
  getDefaultReportRange: function () { return apiData_(getDefaultReportRange()); },
  getAllEmployees: function () { return apiData_(getAllEmployees()); },
  searchContacts: function (p) { return apiData_(searchContacts(p.query)); },
  previewOtReport: function (p) { return previewOtReport(p.startDate, p.endDate, p.selectedEmployees); },
  sendOtReport: function (p, actorEmail) {
    return sendOtReport(Object.assign({}, p, { actorEmail: actorEmail }));
  },
  getScheduledReportConfig: function () { return apiData_(getScheduledReportConfig()); },
  setScheduledReportConfig: function (p) { return setScheduledReportConfig(p.config); },
  disableScheduledReport: function () { return disableScheduledReport(); },
  getCompanyInfo: function () { return apiData_(getCompanyInfo()); },
  saveCompanyInfo: function (p) { return saveCompanyInfo(p.formData); },
  removeCompanyLogo: function () { return removeCompanyLogo(); },

  // ---- Timesheets list (filters + fuzzy search) ----
  listTimesheetEntries: function (p) { return listTimesheetEntries(p); },

  // ---- Dashboard ----
  getRecentAuditEntries: function (p) { return apiData_(getRecentAuditEntries(p && p.limit)); },

  // ---- Notifications ----
  getNotifications: function (p) { return apiData_(getNotifications(p.userEmail, p.includeRead)); },
  markNotificationsAsRead: function (p) { return markNotificationsAsRead(p.notificationIds); },

  // ---- AI features ----
  smartCheckEntries: function (p) { return smartCheckEntries(p); },
  getDashboardNarrative: function (p) { return getDashboardNarrative(p); },
  askTimesheetQuery: function (p) { return askTimesheetQuery(p); },

  // ---- Settings & per-employee reports ----
  getRestDays: function () { return apiData_(getRestDays()); },
  saveRestDays: function (p) { return saveRestDays(p.days); },
  getEmployeeEmailSettings: function () { return apiData_(getEmployeeEmailSettings()); },
  saveEmployeeEmails: function (p) { return saveEmployeeEmails(p.map); },
  sendPerEmployeeReports: function (p, actorEmail) {
    return sendPerEmployeeReports(Object.assign({}, p, { actorEmail: actorEmail }));
  },
  previewPerEmployeeReports: function (p) { return previewPerEmployeeReports(p.startDate, p.endDate); }
};

/** Constant-time-ish equality: compares SHA-256 digests so string length/content can't leak via timing. */
function timingSafeEqual_(a, b) {
  var da = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(a), Utilities.Charset.UTF_8);
  var db = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(b), Utilities.Charset.UTF_8);
  var diff = 0;
  for (var i = 0; i < da.length; i++) diff |= da[i] ^ db[i];
  return diff === 0;
}

/**
 * Verifies the request body's auth. Returns true only if either:
 *  1. body.signature is a valid, fresh HMAC over (timestamp, action, actorEmail, payload), or
 *  2. LEGACY fallback: body.sharedSecret equals the stored secret (digest compare).
 */
function verifyApiAuth_(body, secret) {
  if (body.signature) {
    var ts = Number(body.timestamp);
    if (!isFinite(ts) || Math.abs(Date.now() - ts) > API_MAX_CLOCK_SKEW_MS) return false;

    var message = String(body.timestamp) + '.' + String(body.action || '') + '.' +
      String(body.actorEmail || '') + '.' + JSON.stringify(body.payload || {});
    var expected = Utilities.base64Encode(
      Utilities.computeHmacSha256Signature(message, secret, Utilities.Charset.UTF_8)
    );
    return timingSafeEqual_(body.signature, expected);
  }

  // TEMPORARY legacy path — delete once the signing frontend is confirmed live.
  return !!body.sharedSecret && timingSafeEqual_(body.sharedSecret, secret);
}

function doPost(e) {
  var result;
  try {
    // Parse request first
    if (!e || !e.postData || !e.postData.contents) {
      // Use generic error to avoid leaking info before auth
      return jsonOutput_({ success: false, error: 'Not authorized.' });
    }
    var body;
    try {
      body = JSON.parse(e.postData.contents);
    } catch (parseErr) {
      return jsonOutput_({ success: false, error: 'Not authorized.' });
    }

    // Auth check MUST happen before any other processing
    var expectedSecret = PropertiesService.getScriptProperties().getProperty(API_SHARED_SECRET_PROPERTY);
    if (!expectedSecret || !verifyApiAuth_(body, expectedSecret)) {
      return jsonOutput_({ success: false, error: 'Not authorized.' });
    }

    // Validate action exists
    var handler = API_ACTIONS[body.action];
    if (!handler) {
      return jsonOutput_({ success: false, error: 'Invalid action.' });
    }

    result = handler(body.payload || {}, body.actorEmail || null);
  } catch (err) {
    // Handlers return { success:false, error } for user-facing problems; anything
    // that escapes as an exception is unexpected — keep the detail in the log.
    Logger.log('doPost unexpected error for action "' + (typeof body === 'object' && body ? body.action : '?') + '": ' + (err && err.stack ? err.stack : err));
    result = { success: false, error: 'Something went wrong on the server. Try again, and check the Apps Script logs if it persists.' };
  }
  return jsonOutput_(result);
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Run this once from the Apps Script editor after deploying the Web App.
 * Generates a random shared secret and stores it in Script Properties.
 * The secret is RETURNED (view it in the editor's execution result — it is
 * deliberately not logged) so it can be copied into the Vercel project's
 * GAS_SHARED_SECRET env var. Re-running rotates the secret (update Vercel
 * too, or every proxied call will start failing with "Not authorized").
 */
function setupApiSharedSecret() {
  var secret = Utilities.getUuid() + Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperty(API_SHARED_SECRET_PROPERTY, secret);
  return secret;
}
