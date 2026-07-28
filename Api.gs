/**
 * WEB APP JSON API - bridge for the standalone (Vercel) frontend.
 * ---------------------------------------------------------------------
 * Deploy this project as a Web App: Execute as "Me", Who has access "Anyone".
 *
 * Trust boundary: the Next.js proxy (app/api/gas/route.ts) is the one that
 * verifies the caller's Google ID token (via google-auth-library, since
 * Apps Script has no npm access to do that properly itself - see the design
 * doc). This file verifies an HMAC-SHA256 signature computed with a secret
 * only the Next.js server process knows, then trusts the actorEmail it was
 * handed for audit logging.
 *
 * Signature scheme (mirrored in the proxy):
 *   signature = base64(HMAC-SHA256(secret,
 *     timestamp + "." + action + "." + (actorEmail || "") + "." + JSON.stringify(payload || {})))
 * Requests older/newer than API_MAX_CLOCK_SKEW_MS are rejected (replay
 * protection).
 *
 * Apps Script's doPost(e) cannot read custom HTTP headers (Google Issue
 * Tracker #67764685, official won't-fix), so auth travels in the JSON body.
 */

var API_SHARED_SECRET_PROPERTY = 'GAS_SHARED_SECRET';
var API_MAX_CLOCK_SKEW_MS = 5 * 60 * 1000;
var API_VERSION_LABEL = '2026-07-28-reliability';

/** Wraps raw-value handlers (arrays, plain objects, booleans) in the standard envelope. */
function apiData_(value) {
  return { success: true, data: value };
}

function apiAction_(kind, validate, handler, wrapData) {
  return { kind: kind, validate: validate || {}, handler: handler, wrapData: !!wrapData };
}

/** action name -> metadata + handler. Whitelist: nothing outside this map is reachable. */
var API_ACTIONS = {
  getEmployeesForSheet: apiAction_('read', { requiredStrings: ['sheetName'] }, function (p) { return getEmployeesForSheet(p.sheetName); }, true),
  getJobOrderSuggestions: apiAction_('read', {}, function () { return getJobOrderSuggestions(); }, true),
  listTimesheetSheets: apiAction_('read', {}, function () { return listTimesheetSheets(); }, true),
  submitTimeEntry: apiAction_('write', { requiredObjects: ['formData'] }, function (p, actorEmail) {
    return submitTimeEntry(Object.assign({}, p.formData, { actorEmail: actorEmail }));
  }),
  submitBulkTimeEntries: apiAction_('write', { requiredObjects: ['formData'] }, function (p, actorEmail) {
    return submitBulkTimeEntries(Object.assign({}, p.formData, { actorEmail: actorEmail }));
  }),
  undoLastSave: apiAction_('write', {}, function (p, actorEmail) { return undoLastSave(actorEmail); }),
  getEntryForEdit: apiAction_('read', { requiredStrings: ['sheetName', 'date'], requiredNumbers: ['employeeColumn'] }, function (p) { return getEntryForEdit(p); }),
  updateTimeEntry: apiAction_('write', { requiredObjects: ['formData'] }, function (p, actorEmail) {
    return updateTimeEntry(Object.assign({}, p.formData, { actorEmail: actorEmail }));
  }),
  getShiftTemplates: apiAction_('read', {}, function () { return getShiftTemplates(); }, true),
  saveShiftTemplate: apiAction_('write', { requiredStrings: ['name', 'startTime', 'endTime'] }, function (p) { return saveShiftTemplate(p.name, p.startTime, p.endTime, p.jobOrder); }),
  deleteShiftTemplate: apiAction_('write', { requiredStrings: ['id'] }, function (p) { return deleteShiftTemplate(p.id); }),
  checkAiCredentials: apiAction_('read', {}, function () { return checkAiCredentials(); }, true),
  parseNaturalLanguageEntries: apiAction_('read', { requiredStrings: ['text', 'sheetName'] }, function (p) { return parseNaturalLanguageEntries(p.text, p.sheetName); }),
  parseNaturalLanguageEntriesWithImage: apiAction_('read', { requiredStrings: ['text', 'imageBase64', 'mimeType', 'sheetName'] }, function (p) {
    return parseNaturalLanguageEntriesWithImage(p.text, p.imageBase64, p.mimeType, p.sheetName);
  }),
  getRecentEntriesPreview: apiAction_('read', { requiredStrings: ['sheetName'] }, function (p) { return getRecentEntriesPreview(p.sheetName); }, true),

  // ---- OT reports, scheduling, company settings ----
  getDefaultReportRange: apiAction_('read', {}, function () { return getDefaultReportRange(); }, true),
  getAllEmployees: apiAction_('read', {}, function () { return getAllEmployees(); }, true),
  searchContacts: apiAction_('read', { optionalStrings: ['query'] }, function (p) { return searchContacts(p.query); }, true),
  previewOtReport: apiAction_('read', { isoDates: ['startDate', 'endDate'], optionalArrays: ['selectedEmployees'] }, function (p) { return previewOtReport(p.startDate, p.endDate, p.selectedEmployees); }),
  sendOtReport: apiAction_('write', { isoDates: ['startDate', 'endDate'], requiredArrays: ['emails'] }, function (p, actorEmail) {
    return sendOtReport(Object.assign({}, p, { actorEmail: actorEmail }));
  }),
  getScheduledReportConfig: apiAction_('read', {}, function () { return getScheduledReportConfig(); }, true),
  setScheduledReportConfig: apiAction_('write', { requiredObjects: ['config'] }, function (p) { return setScheduledReportConfig(p.config); }),
  disableScheduledReport: apiAction_('write', {}, function () { return disableScheduledReport(); }),
  getCompanyInfo: apiAction_('read', {}, function () { return getCompanyInfo(); }, true),
  saveCompanyInfo: apiAction_('write', { requiredObjects: ['formData'] }, function (p) { return saveCompanyInfo(p.formData); }),
  removeCompanyLogo: apiAction_('write', {}, function () { return removeCompanyLogo(); }),

  // ---- Timesheets list (filters + fuzzy search) ----
  listTimesheetEntries: apiAction_('read', { isoDates: ['startDate', 'endDate'], optionalArrays: ['employees'], optionalNumbers: ['minHours', 'maxHours', 'page', 'pageSize'], optionalBooleans: ['otOnly'], optionalStrings: ['jobOrder', 'search'] }, function (p) { return listTimesheetEntries(p); }),

  // ---- Dashboard ----
  getRecentAuditEntries: apiAction_('read', { optionalNumbers: ['limit'] }, function (p) { return getRecentAuditEntries(p && p.limit); }, true),

  // ---- Notifications ----
  getNotifications: apiAction_('read', { requiredStrings: ['userEmail'], optionalBooleans: ['includeRead'] }, function (p) { return getNotifications(p.userEmail, p.includeRead); }, true),
  markNotificationsAsRead: apiAction_('write', { requiredArrays: ['notificationIds'] }, function (p) { return markNotificationsAsRead(p.notificationIds); }),

  // ---- AI features ----
  smartCheckEntries: apiAction_('read', { requiredStrings: ['sheetName'], requiredArrays: ['entries'] }, function (p) { return smartCheckEntries(p); }),
  getDashboardNarrative: apiAction_('read', { isoDates: ['startDate', 'endDate'] }, function (p) { return getDashboardNarrative(p); }),
  askTimesheetQuery: apiAction_('read', { requiredArrays: ['messages'], optionalObjects: ['context'] }, function (p) { return askTimesheetQuery(p); }),

  // ---- Settings & per-employee reports ----
  getRestDays: apiAction_('read', {}, function () { return getRestDays(); }, true),
  saveRestDays: apiAction_('write', { requiredArrays: ['days'] }, function (p) { return saveRestDays(p.days); }),
  getEmployeeEmailSettings: apiAction_('read', {}, function () { return getEmployeeEmailSettings(); }, true),
  saveEmployeeEmails: apiAction_('write', { requiredObjects: ['map'] }, function (p) { return saveEmployeeEmails(p.map); }),
  sendPerEmployeeReports: apiAction_('write', { isoDates: ['startDate', 'endDate'] }, function (p, actorEmail) {
    return sendPerEmployeeReports(Object.assign({}, p, { actorEmail: actorEmail }));
  }),
  previewPerEmployeeReports: apiAction_('read', { isoDates: ['startDate', 'endDate'] }, function (p) { return previewPerEmployeeReports(p.startDate, p.endDate); }),
  getBackendStatus: apiAction_('read', {}, function () { return getBackendStatus_(); }, true)
};

/** Constant-time-ish equality: compares SHA-256 digests so string length/content can't leak via timing. */
function timingSafeEqual_(a, b) {
  var da = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(a), Utilities.Charset.UTF_8);
  var db = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(b), Utilities.Charset.UTF_8);
  var diff = 0;
  for (var i = 0; i < da.length; i++) diff |= da[i] ^ db[i];
  return diff === 0;
}

/** Verifies a fresh HMAC over (timestamp, action, actorEmail, payload). */
function verifyApiAuth_(body, secret) {
  if (!body || !body.signature) return false;

  var ts = Number(body.timestamp);
  if (!isFinite(ts) || Math.abs(Date.now() - ts) > API_MAX_CLOCK_SKEW_MS) return false;

  var message = String(body.timestamp) + '.' + String(body.action || '') + '.' +
    String(body.actorEmail || '') + '.' + JSON.stringify(body.payload || {});
  var expected = Utilities.base64Encode(
    Utilities.computeHmacSha256Signature(message, secret, Utilities.Charset.UTF_8)
  );
  return timingSafeEqual_(body.signature, expected);
}

function doPost(e) {
  var result;
  var body = null;
  var requestId = createApiRequestId_();
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonOutput_({ success: false, error: 'Not authorized.' });
    }
    try {
      body = JSON.parse(e.postData.contents);
    } catch (parseErr) {
      return jsonOutput_({ success: false, error: 'Not authorized.' });
    }

    var expectedSecret = PropertiesService.getScriptProperties().getProperty(API_SHARED_SECRET_PROPERTY);
    if (!expectedSecret || !verifyApiAuth_(body, expectedSecret)) {
      return jsonOutput_({ success: false, error: 'Not authorized.' });
    }

    var action = API_ACTIONS[body.action];
    if (!action) {
      return jsonOutput_({ success: false, error: 'Invalid action.' });
    }

    var payload = normalizeApiPayload_(body.payload);
    var validationError = validateApiPayload_(payload, action.validate);
    if (validationError) {
      return jsonOutput_({ success: false, error: validationError });
    }

    Logger.log('api request ' + requestId + ' action=' + body.action + ' kind=' + action.kind + ' actor=' + (body.actorEmail || ''));
    result = normalizeApiResult_(action.handler(payload, body.actorEmail || null), action.wrapData);
  } catch (err) {
    Logger.log('api request ' + requestId + ' unexpected error for action "' + (body && body.action ? body.action : '?') + '": ' + (err && err.stack ? err.stack : err));
    result = { success: false, error: 'Something went wrong on the server. Reference: ' + requestId + '.' };
  }
  return jsonOutput_(result);
}

function normalizeApiPayload_(payload) {
  return payload && typeof payload === 'object' && !Array.isArray(payload) ? payload : {};
}

function normalizeApiResult_(value, wrapData) {
  if (wrapData) return apiData_(value);
  if (value && typeof value === 'object' && typeof value.success === 'boolean') return value;
  return { success: true, data: value };
}

function validateApiPayload_(payload, spec) {
  spec = spec || {};
  var err = validateRequiredStrings_(payload, spec.requiredStrings || []) ||
    validateOptionalStrings_(payload, spec.optionalStrings || []) ||
    validateRequiredArrays_(payload, spec.requiredArrays || []) ||
    validateOptionalArrays_(payload, spec.optionalArrays || []) ||
    validateRequiredObjects_(payload, spec.requiredObjects || []) ||
    validateOptionalObjects_(payload, spec.optionalObjects || []) ||
    validateOptionalNumbers_(payload, spec.optionalNumbers || []) ||
    validateRequiredNumbers_(payload, spec.requiredNumbers || []) ||
    validateOptionalBooleans_(payload, spec.optionalBooleans || []) ||
    validateIsoDates_(payload, spec.isoDates || []);
  return err || null;
}

function validateRequiredStrings_(payload, names) {
  for (var i = 0; i < names.length; i++) {
    var name = names[i];
    if (typeof payload[name] !== 'string' || payload[name].trim() === '') return 'Missing or invalid field: ' + name + '.';
  }
  return null;
}

function validateOptionalStrings_(payload, names) {
  for (var i = 0; i < names.length; i++) {
    var name = names[i];
    if (payload[name] !== undefined && payload[name] !== null && typeof payload[name] !== 'string') return 'Invalid field: ' + name + '.';
  }
  return null;
}

function validateRequiredArrays_(payload, names) {
  for (var i = 0; i < names.length; i++) {
    var name = names[i];
    if (!Array.isArray(payload[name])) return 'Missing or invalid field: ' + name + '.';
  }
  return null;
}

function validateOptionalArrays_(payload, names) {
  for (var i = 0; i < names.length; i++) {
    var name = names[i];
    if (payload[name] !== undefined && payload[name] !== null && !Array.isArray(payload[name])) return 'Invalid field: ' + name + '.';
  }
  return null;
}

function validateRequiredObjects_(payload, names) {
  for (var i = 0; i < names.length; i++) {
    var name = names[i];
    var value = payload[name];
    if (!value || typeof value !== 'object' || Array.isArray(value)) return 'Missing or invalid field: ' + name + '.';
  }
  return null;
}

function validateOptionalObjects_(payload, names) {
  for (var i = 0; i < names.length; i++) {
    var name = names[i];
    var value = payload[name];
    if (value !== undefined && value !== null && (typeof value !== 'object' || Array.isArray(value))) return 'Invalid field: ' + name + '.';
  }
  return null;
}

function validateRequiredNumbers_(payload, names) {
  for (var i = 0; i < names.length; i++) {
    var name = names[i];
    if (typeof payload[name] !== 'number' || !isFinite(payload[name])) return 'Missing or invalid field: ' + name + '.';
  }
  return null;
}

function validateOptionalNumbers_(payload, names) {
  for (var i = 0; i < names.length; i++) {
    var name = names[i];
    if (payload[name] !== undefined && payload[name] !== null && (typeof payload[name] !== 'number' || !isFinite(payload[name]))) return 'Invalid field: ' + name + '.';
  }
  return null;
}

function validateOptionalBooleans_(payload, names) {
  for (var i = 0; i < names.length; i++) {
    var name = names[i];
    if (payload[name] !== undefined && payload[name] !== null && typeof payload[name] !== 'boolean') return 'Invalid field: ' + name + '.';
  }
  return null;
}

function validateIsoDates_(payload, names) {
  for (var i = 0; i < names.length; i++) {
    var name = names[i];
    if (typeof payload[name] !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(payload[name])) return 'Missing or invalid field: ' + name + '.';
  }
  return null;
}

function getBackendStatus_() {
  var props = PropertiesService.getScriptProperties();
  var spreadsheetName = '';
  try {
    spreadsheetName = SpreadsheetApp.getActiveSpreadsheet().getName();
  } catch (e) {
    spreadsheetName = null;
  }
  return {
    version: API_VERSION_LABEL,
    timezone: Session.getScriptTimeZone(),
    authMode: 'hmac-sha256-timestamp',
    aiConfigured: !!(props.getProperty('CLOUDFLARE_ACCOUNT_ID') && props.getProperty('CLOUDFLARE_API_TOKEN')),
    spreadsheetName: spreadsheetName
  };
}

function createApiRequestId_() {
  return 'api_' + Utilities.getUuid().slice(0, 8);
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Run this once from the Apps Script editor after deploying the Web App.
 * Generates a random shared secret and stores it in Script Properties.
 * The secret is RETURNED (view it in the editor's execution result - it is
 * deliberately not logged) so it can be copied into the Vercel project's
 * GAS_SHARED_SECRET env var. Re-running rotates the secret (update Vercel
 * too, or every proxied call will start failing with "Not authorized").
 */
function setupApiSharedSecret() {
  var secret = Utilities.getUuid() + Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperty(API_SHARED_SECRET_PROPERTY, secret);
  return secret;
}

function runApiSelfTests() {
  var secret = 'test-secret';
  var payload = { sheetName: 'July 2026' };
  var body = {
    action: 'getEmployeesForSheet',
    payload: payload,
    actorEmail: 'tester@example.com',
    timestamp: Date.now()
  };
  body.signature = Utilities.base64Encode(Utilities.computeHmacSha256Signature(
    String(body.timestamp) + '.' + body.action + '.' + body.actorEmail + '.' + JSON.stringify(payload),
    secret,
    Utilities.Charset.UTF_8
  ));

  assertApiTest_(verifyApiAuth_(body, secret), 'valid HMAC should pass');
  assertApiTest_(!verifyApiAuth_(Object.assign({}, body, { signature: 'bad' }), secret), 'invalid HMAC should fail');
  assertApiTest_(!verifyApiAuth_(Object.assign({}, body, { timestamp: Date.now() - API_MAX_CLOCK_SKEW_MS - 1000 }), secret), 'stale HMAC should fail');
  assertApiTest_(!verifyApiAuth_({ action: body.action, payload: payload, actorEmail: body.actorEmail, sharedSecret: secret }, secret), 'legacy sharedSecret should fail');

  assertApiTest_(validateApiPayload_({}, { requiredStrings: ['sheetName'] }) === 'Missing or invalid field: sheetName.', 'required string validation');
  assertApiTest_(validateApiPayload_({ startDate: '2026-07-01', endDate: '2026-07-31' }, { isoDates: ['startDate', 'endDate'] }) === null, 'ISO date validation');
  assertApiTest_(normalizeApiResult_(['a'], true).data.length === 1, 'wrapped result normalization');
  assertApiTest_(normalizeApiResult_({ success: true, entries: [] }, false).entries.length === 0, 'existing success shape preserved');

  try {
    CacheService.getScriptCache().put(TIMESHEET_SCAN_CACHE_INDEX_KEY, '{bad json', 60);
    invalidateTimesheetScanCache_();
  } catch (e) {
    throw new Error('cache invalidation should tolerate corrupt index: ' + e.message);
  }

  return { success: true, message: 'API self-tests passed.' };
}

function assertApiTest_(condition, message) {
  if (!condition) throw new Error('API self-test failed: ' + message);
}
