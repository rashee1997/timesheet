var UNDO_PROPERTY_KEY = 'lastWriteUndoData';
var UNDO_TTL_MS = 60 * 1000; // 1 minute

/** Undo state is per-actor so one user can't undo another user's save. */
function undoKeyForActor_(actorEmail) {
  var email = actorEmail;
  if (!email) {
    try { email = Session.getActiveUser().getEmail(); } catch (e) { email = ''; }
  }
  return UNDO_PROPERTY_KEY + '_' + String(email || 'anonymous').toLowerCase();
}

/** Date cell values don't survive JSON round-trips; tag them explicitly. */
function encodeCellValues_(vals) {
  return vals.map(function (v) {
    return v instanceof Date ? { __date: v.getTime() } : v;
  });
}

function decodeCellValues_(vals) {
  return vals.map(function (v) {
    return v && typeof v === 'object' && typeof v.__date === 'number' ? new Date(v.__date) : v;
  });
}

function saveUndoData(sheetName, entries, actorEmail) {
  var data = {
    timestamp: Date.now(),
    sheetName: sheetName,
    entries: entries.map(function (e) {
      return {
        row: e.targetRow,
        col: e.employeeColumn,
        prevValues: e.prevValues ? encodeCellValues_(e.prevValues) : null,
        writtenValues: e.writtenValues ? encodeCellValues_(e.writtenValues) : null
      };
    })
  };
  PropertiesService.getScriptProperties().setProperty(undoKeyForActor_(actorEmail), JSON.stringify(data));
}

function getUndoData_(actorEmail) {
  var key = undoKeyForActor_(actorEmail);
  var raw = PropertiesService.getScriptProperties().getProperty(key);
  if (!raw) return null;
  var data;
  try {
    data = JSON.parse(raw);
  } catch (e) {
    PropertiesService.getScriptProperties().deleteProperty(key);
    return null;
  }
  if (!data || Date.now() - data.timestamp > UNDO_TTL_MS) {
    PropertiesService.getScriptProperties().deleteProperty(key);
    return null;
  }
  return data;
}

function undoLastSave(actorEmail) {
  var data = getUndoData_(actorEmail);
  if (!data) return { success: false, error: 'No recent save to undo (undo expires after 1 minute).' };

  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) {
    return { success: false, error: 'The sheet is busy right now. Wait a few seconds and try again — nothing was changed.' };
  }
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(data.sheetName);
    if (!sheet) return { success: false, error: 'Sheet "' + data.sheetName + '" no longer exists.' };

    // Verify nothing changed since the save this undo belongs to — restoring
    // over someone else's newer edit would silently destroy their work.
    for (var i = 0; i < data.entries.length; i++) {
      var entry = data.entries[i];
      if (!entry.writtenValues) continue;
      var current = sheet.getRange(entry.row, entry.col, 1, entry.writtenValues.length).getValues()[0];
      if (JSON.stringify(encodeCellValues_(current)) !== JSON.stringify(entry.writtenValues)) {
        return {
          success: false,
          error: 'These cells were changed again after your save (possibly by someone else) — undo cancelled to avoid overwriting the newer values.'
        };
      }
    }

    var restored = 0;
    data.entries.forEach(function (entry) {
      if (entry.prevValues && entry.prevValues.length > 0) {
        sheet.getRange(entry.row, entry.col, 1, entry.prevValues.length)
          .setValues([decodeCellValues_(entry.prevValues)]);
        restored++;
      }
    });

    SpreadsheetApp.flush();
    invalidateTimesheetScanCache_();
    PropertiesService.getScriptProperties().deleteProperty(undoKeyForActor_(actorEmail));
    logAudit('undo', data.sheetName, null, null, 'Undid ' + restored + ' entry blocks', actorEmail);
    return { success: true, message: 'Undid last save (' + restored + ' entries restored).' };
  } catch (e) {
    return { success: false, error: 'Undo failed: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}
