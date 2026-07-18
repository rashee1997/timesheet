/**
 * WEB APP JSON API — bridge for the standalone (Vercel) frontend — Phase 1.
 * ---------------------------------------------------------------------
 * Deploy this project as a Web App: Execute as "Me", Who has access "Anyone".
 *
 * Trust boundary: the Next.js proxy (app/api/gas/route.ts) is the one that
 * verifies the caller's Google ID token (via google-auth-library, since
 * Apps Script has no npm access to do that properly itself — see the design
 * doc). This file only checks a shared secret that only the Next.js server
 * process knows, then trusts the actorEmail it was handed for audit logging.
 *
 * Apps Script's doPost(e) cannot read custom HTTP headers (Google Issue
 * Tracker #67764685, official won't-fix), so the secret travels in the
 * JSON body instead of a header.
 */

var API_SHARED_SECRET_PROPERTY = 'GAS_SHARED_SECRET';

/** action name -> handler(payload, actorEmail). Whitelist — nothing outside this map is reachable. */
var API_ACTIONS = {
  getEmployeesForSheet: function (p) { return getEmployeesForSheet(p.sheetName); },
  getJobOrderSuggestions: function () { return getJobOrderSuggestions(); },
  listTimesheetSheets: function () { return listTimesheetSheets(); },
  submitTimeEntry: function (p, actorEmail) {
    return submitTimeEntry(Object.assign({}, p.formData, { actorEmail: actorEmail }));
  },
  submitBulkTimeEntries: function (p, actorEmail) {
    return submitBulkTimeEntries(Object.assign({}, p.formData, { actorEmail: actorEmail }));
  },
  undoLastSave: function (p, actorEmail) { return undoLastSave(actorEmail); },
  getShiftTemplates: function () { return getShiftTemplates(); },
  saveShiftTemplate: function (p) { return saveShiftTemplate(p.name, p.startTime, p.endTime, p.jobOrder); },
  deleteShiftTemplate: function (p) { return deleteShiftTemplate(p.id); },
  checkAiCredentials: function () { return checkAiCredentials(); },
  parseNaturalLanguageEntries: function (p) { return parseNaturalLanguageEntries(p.text, p.sheetName); },
  parseNaturalLanguageEntriesWithImage: function (p) {
    return parseNaturalLanguageEntriesWithImage(p.text, p.imageBase64, p.mimeType, p.sheetName);
  },
  getRecentEntriesPreview: function (p) { return getRecentEntriesPreview(p.sheetName); },

  // ---- Phase 2: OT reports, scheduling, company settings ----
  getDefaultReportRange: function () { return getDefaultReportRange(); },
  getAllEmployees: function () { return getAllEmployees(); },
  searchContacts: function (p) { return searchContacts(p.query); },
  previewOtReport: function (p) { return previewOtReport(p.startDate, p.endDate, p.selectedEmployees); },
  sendOtReport: function (p) { return sendOtReport(p); },
  getScheduledReportConfig: function () { return getScheduledReportConfig(); },
  setScheduledReportConfig: function (p) { return setScheduledReportConfig(p.config); },
  disableScheduledReport: function () { return disableScheduledReport(); },
  getCompanyInfo: function () { return getCompanyInfo(); },
  saveCompanyInfo: function (p) { return saveCompanyInfo(p.formData); },
  removeCompanyLogo: function () { return removeCompanyLogo(); },

  // ---- Phase 3: timesheets list (filters + fuzzy search) ----
  listTimesheetEntries: function (p) { return listTimesheetEntries(p); }
};

function doPost(e) {
  var result;
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('Missing request body.');
    }
    var body = JSON.parse(e.postData.contents);

    var expectedSecret = PropertiesService.getScriptProperties().getProperty(API_SHARED_SECRET_PROPERTY);
    if (!expectedSecret || body.sharedSecret !== expectedSecret) {
      result = { success: false, error: 'Not authorized.' };
      return jsonOutput_(result);
    }

    var handler = API_ACTIONS[body.action];
    if (!handler) {
      result = { success: false, error: 'Unknown action: ' + body.action };
      return jsonOutput_(result);
    }

    result = handler(body.payload || {}, body.actorEmail || null);
  } catch (err) {
    Logger.log('doPost error: ' + err);
    result = { success: false, error: err.message || String(err) };
  }
  return jsonOutput_(result);
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Run this once from the Apps Script editor after deploying the Web App.
 * Generates a random shared secret, stores it in Script Properties, and
 * logs it so it can be copied into the Vercel project's GAS_SHARED_SECRET
 * env var. Re-running rotates the secret (update Vercel too, or every
 * proxied call will start failing with "Not authorized").
 */
function setupApiSharedSecret() {
  var secret = Utilities.getUuid() + Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperty(API_SHARED_SECRET_PROPERTY, secret);
  Logger.log('GAS_SHARED_SECRET set. Copy this into Vercel:\n' + secret);
  return secret;
}
