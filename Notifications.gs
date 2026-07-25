/**
 * Notifications.gs
 * Manages notification creation, retrieval, and status updates.
 * (sanitizeSheetText lives in Shared.gs — one canonical copy.)
 */

// Define the name of the sheet where notifications will be stored
const NOTIFICATIONS_SHEET_NAME = "Notifications";

// Only the tail of the sheet is ever scanned — the sheet grows forever, and
// reading the whole thing on every poll gets linearly slower over time.
const NOTIFICATIONS_MAX_SCAN_ROWS = 500;

const NOTIFICATION_COLS = { ID: 0, RECIPIENT: 1, MESSAGE: 2, TYPE: 3, LINK: 4, TIMESTAMP: 5, READ: 6 };

/**
 * Ensures the 'Notifications' sheet exists with the correct headers.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The Notification sheet.
 */
function getOrCreateNotificationsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(NOTIFICATIONS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(NOTIFICATIONS_SHEET_NAME);
    const headers = ["ID", "RecipientEmail", "Message", "Type", "Link", "Timestamp", "Read"];
    sheet.appendRow(headers);
    sheet.setFrozenRows(1); // Freeze the header row
  }
  return sheet;
}

/**
 * Creates a new notification for a specific user.
 * @param {string} recipientEmail The email of the user to notify.
 * @param {string} message The notification message.
 * @param {string} type The type of notification (e.g., "timesheet_approved", "reminder").
 * @param {string} [link=""] Optional URL link associated with the notification.
 * @returns {object} The created notification object.
 */
function createNotification(recipientEmail, message, type, link = "") {
  const sheet = getOrCreateNotificationsSheet();
  const id = Utilities.getUuid(); // Generate a unique ID for the notification
  const timestamp = new Date().toISOString();
  // Sanitize all user-provided text that goes into sheet cells
  const newRow = [id, recipientEmail, sanitizeSheetText(message), sanitizeSheetText(type), sanitizeSheetText(link), timestamp, false]; // 'false' for unread

  sheet.appendRow(newRow);

  return { id, recipientEmail, message, type, link, timestamp, read: false };
}

/** Reads only the last NOTIFICATIONS_MAX_SCAN_ROWS data rows. */
function readNotificationTail_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { rows: [], firstSheetRow: 2 };
  const firstSheetRow = Math.max(2, lastRow - NOTIFICATIONS_MAX_SCAN_ROWS + 1);
  const rows = sheet.getRange(firstSheetRow, 1, lastRow - firstSheetRow + 1, 7).getValues();
  return { rows: rows, firstSheetRow: firstSheetRow };
}

/**
 * Retrieves notifications for a given user (most recent NOTIFICATIONS_MAX_SCAN_ROWS only).
 * @param {string} userEmail The email of the user whose notifications to retrieve.
 * @param {boolean} [includeRead=false] Whether to include read notifications (defaults to unread only).
 * @returns {Array<object>} An array of notification objects.
 */
function getNotifications(userEmail, includeRead = false) {
  const sheet = getOrCreateNotificationsSheet();
  const tail = readNotificationTail_(sheet);
  const notifications = [];

  tail.rows.forEach(function (row, i) {
    if (row[NOTIFICATION_COLS.RECIPIENT] !== userEmail) return;
    const isRead = row[NOTIFICATION_COLS.READ];
    if (!includeRead && isRead) return;
    notifications.push({
      id: row[NOTIFICATION_COLS.ID],
      recipientEmail: row[NOTIFICATION_COLS.RECIPIENT],
      message: row[NOTIFICATION_COLS.MESSAGE],
      type: row[NOTIFICATION_COLS.TYPE],
      link: row[NOTIFICATION_COLS.LINK],
      timestamp: row[NOTIFICATION_COLS.TIMESTAMP],
      read: isRead,
      rowIndex: tail.firstSheetRow + i
    });
  });
  return notifications;
}

/**
 * Marks a list of notifications as read. Writes are batched into contiguous
 * runs instead of one setValue per row.
 * @param {Array<string>} notificationIds An array of notification IDs to mark as read.
 * @returns {object} An object indicating success and count of updated notifications.
 */
function markNotificationsAsRead(notificationIds) {
  if (!notificationIds || notificationIds.length === 0) {
    return { success: false, message: "No notification IDs provided." };
  }

  const sheet = getOrCreateNotificationsSheet();
  const tail = readNotificationTail_(sheet);
  if (tail.rows.length === 0) return { success: false, message: "No notifications found." };

  const wanted = {};
  notificationIds.forEach(function (id) { wanted[id] = true; });

  const targetRows = [];
  tail.rows.forEach(function (row, i) {
    if (wanted[row[NOTIFICATION_COLS.ID]] && row[NOTIFICATION_COLS.READ] === false) {
      targetRows.push(tail.firstSheetRow + i);
    }
  });

  // Group contiguous sheet rows into single setValues writes.
  let updatedCount = 0;
  let runStart = -1, runLen = 0;
  const flushRun = function () {
    if (runLen === 0) return;
    const values = [];
    for (let k = 0; k < runLen; k++) values.push([true]);
    sheet.getRange(runStart, NOTIFICATION_COLS.READ + 1, runLen, 1).setValues(values);
    updatedCount += runLen;
    runStart = -1;
    runLen = 0;
  };
  targetRows.forEach(function (rowNum) {
    if (runLen > 0 && rowNum === runStart + runLen) {
      runLen++;
    } else {
      flushRun();
      runStart = rowNum;
      runLen = 1;
    }
  });
  flushRun();

  return { success: true, updatedCount: updatedCount };
}
