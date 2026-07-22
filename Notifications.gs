/**
 * Notifications.gs
 * Manages notification creation, retrieval, and status updates.
 */

// Define the name of the sheet where notifications will be stored
const NOTIFICATIONS_SHEET_NAME = "Notifications";

/**
 * Neutralizes leading =, +, -, @ so free-text saved into a cell can never be
 * interpreted as a formula (formula/DDE injection protection).
 */
function sanitizeSheetText(v) {
  const s = String(v == null ? '' : v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

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

/**
 * Retrieves notifications for a given user.
 * @param {string} userEmail The email of the user whose notifications to retrieve.
 * @param {boolean} [includeRead=false] Whether to include read notifications (defaults to unread only).
 * @returns {Array<object>} An array of notification objects.
 */
function getNotifications(userEmail, includeRead = false) {
  const sheet = getOrCreateNotificationsSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return []; // No headers or data

  const headers = data[0];
  const notifications = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const recipientEmail = row[headers.indexOf("RecipientEmail")];

    if (recipientEmail === userEmail) {
      const isRead = row[headers.indexOf("Read")];
      if (includeRead || !isRead) {
        notifications.push({
          id: row[headers.indexOf("ID")],
          recipientEmail: recipientEmail,
          message: row[headers.indexOf("Message")],
          type: row[headers.indexOf("Type")],
          link: row[headers.indexOf("Link")],
          timestamp: row[headers.indexOf("Timestamp")],
          read: isRead,
          rowIndex: i + 1 // Store row index for easy updating
        });
      }
    }
  }
  return notifications;
}

/**
 * Marks a list of notifications as read.
 * @param {Array<string>} notificationIds An array of notification IDs to mark as read.
 * @returns {object} An object indicating success and count of updated notifications.
 */
function markNotificationsAsRead(notificationIds) {
  if (!notificationIds || notificationIds.length === 0) {
    return { success: false, message: "No notification IDs provided." };
  }

  const sheet = getOrCreateNotificationsSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { success: false, message: "No notifications found." };

  const headers = data[0];
  const idColIndex = headers.indexOf("ID");
  const readColIndex = headers.indexOf("Read");
  let updatedCount = 0;

  for (let i = 1; i < data.length; i++) {
    const notificationId = data[i][idColIndex];
    if (notificationIds.includes(notificationId)) {
      if (data[i][readColIndex] === false) { // Only update if not already read
        sheet.getRange(i + 1, readColIndex + 1).setValue(true);
        updatedCount++;
      }
    }
  }
  return { success: true, updatedCount: updatedCount };
}

// Example usage (for testing within Apps Script editor)
function testNotifications() {
  const user = Session.getActiveUser().getEmail(); // Get current user's email
  Logger.log(`Current user: ${user}`);

  // Create some notifications
  createNotification(user, "Reminder: Please submit your timesheet for the last week.", "reminder");
  createNotification("another@example.com", "This is for another user.", "info");

  // Get unread notifications
  const unread = getNotifications(user);
  Logger.log("Unread notifications:");
  unread.forEach(n => Logger.log(JSON.stringify(n)));

  // Mark one as read
  if (unread.length > 0) {
    markNotificationsAsRead([unread[0].id]);
    Logger.log(`Marked notification ID ${unread[0].id} as read.`);
  }

  // Get all notifications (including read)
  const all = getNotifications(user, true);
  Logger.log("All notifications (including read):");
  all.forEach(n => Logger.log(JSON.stringify(n)));
}
