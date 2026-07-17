/**
 * Simple onEdit trigger: if someone manually types over a TOTAL/NORMAL/OT
 * formula cell (breaking the hour calculation for that row), put the
 * correct formula straight back.
 *
 * Reads each edited column's own LABEL_ROW text ("TOTAL HRS"/"NORMAL"/"OT")
 * rather than assuming a fixed column grid - a block is always internally
 * 6 columns wide (COLS_PER_EMPLOYEE), but where that block STARTS on the
 * sheet can drift, so grid arithmetic alone can point at the wrong block.
 */
function onEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();

  try {
    if (!isTimesheetSheet(sheet)) return;
  } catch (err) {
    return;
  }

  var lastCol = sheet.getLastColumn();
  var labelRow = sheet.getRange(CONFIG.LABEL_ROW, 1, 1, lastCol).getDisplayValues()[0];
  var startRow = e.range.getRow();
  var startCol = e.range.getColumn();
  var reverted = false;

  for (var r = 0; r < e.range.getNumRows(); r++) {
    var row = startRow + r;
    if (row < CONFIG.DATA_START_ROW) continue;

    for (var c = 0; c < e.range.getNumColumns(); c++) {
      var col = startCol + c;
      if (col > lastCol) continue;

      var label = String(labelRow[col - 1] || '').trim().toUpperCase();
      var offset = label === 'TOTAL HRS' ? 2 : (label === 'NORMAL' ? 3 : (label === 'OT' ? 4 : -1));
      if (offset === -1) continue;

      var blockCol = col - offset; // this block's own START column
      if (blockCol < 1) continue;

      var f = buildHourFormulas(row, blockCol);
      var formula = offset === 2 ? f.total : (offset === 3 ? f.normal : f.ot);

      var cell = sheet.getRange(row, col);
      cell.setFormula(formula);
      cell.setNumberFormat('[h]:mm');
      reverted = true;
    }
  }

  if (reverted) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'TOTAL/NORMAL/OT are calculated automatically and can\'t be edited directly — your change was reverted.',
      'Timesheet', 5
    );
  }
}
