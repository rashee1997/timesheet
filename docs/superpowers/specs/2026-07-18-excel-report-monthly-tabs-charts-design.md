# OT Report Excel: per-employee tabs, summary charts, and statistics

Status: approved for planning
Date: 2026-07-18
Backend repo: `F:\timesheet` (`rashee1997/timesheet`)
Related frontend repo: `F:\timesheet-web` (`rashee1997/timesheet-web`) — no changes needed there; this is purely inside the existing "Attach report as Excel" flow triggered from `/reports`.

## Context

`ExcelGen.gs`'s `buildReportExcel(report)` currently produces a single-tab
workbook: company header, report date range, and one row per employee with
aggregate totals (`Normal Hrs`, `OT Hrs`, `Total Hrs`, `Days`, `Entry Range`).
It's attached to the OT report email when the user checks "Attach report as
Excel" (`/reports` page and the Sheets-bound `ReportForm.html`, both calling
`sendOtReport` → `processSendReport` → `buildReportExcel`).

The user wants each employee's full daily breakdown for the report's date
range as its own tab, a first "Summary" page carrying charts and statistics
(not just the totals table it has today), and everything styled consistently
with the existing report look — including the auto-sized columns
`setColumnWidthsToFitContent` already gives the totals table.

## Goals

1. Workbook becomes: **Summary tab** (today's totals table + new stat banner
   + two charts) followed by **one tab per employee** with their day-by-day
   entries for the exact selected date range.
2. Every tab uses the same header/title styling and auto-fit column widths as
   the existing Summary table — no tab should look visually inconsistent with
   another.
3. No behavior change to email sending, the "no data" confirmation flow, or
   the frontend — this is entirely inside `buildReportExcel` and its direct
   caller.

## Non-goals

- Expanding a partial date-range selection to a full calendar month. Tabs
  cover exactly `report.startDate`–`report.endDate`, matching the summary
  table and the email body.
- Any new frontend UI. The existing "Attach report as Excel" checkbox is the
  only trigger.
- Per-employee charts. Charts live only on the Summary tab (bar chart:
  Normal vs OT per employee; pie chart: org-wide Normal vs OT split).

## Architecture

### 1. `EntriesList.gs` — extract the raw scan

`listTimesheetEntries` currently does sheet-scanning and filtering/pagination
in one function. Pull the scan (everything that builds the unfiltered
`entries` array — the `SpreadsheetApp.getActiveSpreadsheet().getSheets().forEach(...)`
block) into:

```js
/**
 * Raw per-shift rows for every employee across all timesheet sheets touching
 * [startDate, endDate], unfiltered and unpaginated. Shared by
 * listTimesheetEntries (which filters/sorts/paginates on top of this) and
 * the Excel report builder (which groups by employee for per-employee tabs).
 */
function collectTimesheetEntries(startDateStr, endDateStr) {
  // ... existing scan body, returns { entries, skippedSheets }
}
```

`listTimesheetEntries` becomes:

```js
function listTimesheetEntries(filters) {
  // ... existing validation ...
  var scan = collectTimesheetEntries(filters.startDate, filters.endDate);
  var filtered = scan.entries.filter(/* existing filter predicate, unchanged */);
  // ... existing sort/paginate/response, unchanged ...
}
```

This removes the only meaningful duplication risk (a third copy of the same
~70-line scan loop for the Excel builder) without touching
`listTimesheetEntries`'s existing behavior or response shape.

### 2. `ReportCode.js` — pass raw entries alongside the aggregate report

In `processSendReport`, where `attachPdf` is handled:

```js
if (payload.attachPdf) {
  try {
    var scan = collectTimesheetEntries(payload.startDate, payload.endDate);
    attachments = buildReportExcel(report, scan.entries);
  } catch (xlsxErr) {
    // unchanged error handling
  }
}
```

No change to `generateOtReport`, the "no data" confirmation gate, or the
email-sending block below it.

### 3. `ExcelGen.gs` — the workbook itself

`buildReportExcel(report, entries)` (new second parameter; `entries` is the
raw per-shift array from `collectTimesheetEntries`, each row already shaped
`{ date, sheetName, employeeName, employeeKey, startTime, endTime, jobOrder,
totalHours, normalHours, otHours }` — note `employeeKey` must be included in
what `collectTimesheetEntries` returns, since `listTimesheetEntries`'s
current paginated output strips it but the Excel builder needs it to group
rows by employee).

**Shared formatting rule:** every tab — Summary and each employee tab — ends
its build with the same three calls already used today, so nothing drifts
out of visual sync:

```js
sheet.setFrozenRows(headerRow);
setColumnWidthsToFitContent(sheet, numCols, headers, data);
```

`setColumnWidthsToFitContent` is reused as-is (it's already generic over
column count/headers/data — no changes needed there). The existing header
styling (`14pt`/`12pt`/`9pt` title stack, `#fbf0e6` header background, `#d1d5db`
borders, `9pt` data font) is factored into one small helper so both the
Summary table and each employee table call the identical styling code
instead of it being copy-pasted per tab:

```js
/** Applies the report's standard header-row + border + font styling. Used by
 * every tab so Summary and per-employee tables never visually drift apart. */
function styleReportTable(sheet, headerRow, numCols, dataRowCount) {
  sheet.getRange(headerRow, 1, 1, numCols)
    .setBackground('#fbf0e6').setFontWeight('bold').setFontSize(9)
    .setHorizontalAlignment('center');
  if (dataRowCount > 0) {
    sheet.getRange(headerRow, 1, dataRowCount + 1, numCols)
      .setBorder(true, true, true, true, true, true, '#d1d5db', SpreadsheetApp.BorderStyle.SOLID);
  }
}
```

**Summary tab** (same sheet/position as today — `ss.getSheets()[0]`, named
`'Summary'` instead of `'OT Report'` since it's no longer the only tab):

- Title stack (rows 1–3): unchanged — company name, "Timesheet OT Report",
  date range.
- **New stat banner** (row 4, merged across all 6 columns): one line,
  `Total Hours: 245.50   |   OT Hours: 32.00 (13.0%)   |   Active Employees: 18   |   Days Covered: 30`.
  Computed from `report.rows` (`reduce` for totals, `.length` for employee
  count, `report.rangeCalendarDays` for days covered) — no new inputs
  needed. Styled distinctly (`10pt`, bold, `#6b7280`) but same merge-and-center
  pattern as the title rows above it.
- Row 5: blank spacer (keeps the banner visually separated from the table,
  consistent with existing spacing rhythm).
- Existing totals table, shifted down two rows (header row 5 → 7) — same
  columns, same per-row styling, same `setColumnWidthsToFitContent` call.
- **Pie-chart source table**, placed at column H (clear of the 6-column
  table, so it doesn't intrude on the auto-fit width calculation for columns
  A–F): two rows, `Normal | <sum of report.rows normalHours>` and
  `OT | <sum of report.rows otHours>`, header `Split`/`Hours`. Small, auto-fit
  with its own `setColumnWidthsToFitContent` call scoped to just those two
  columns.
- **Bar chart**: `sheet.newChart().setChartType(Charts.ChartType.COLUMN)`,
  stacked, sourced directly from the existing table's Employee/Normal
  Hrs/OT Hrs columns (`addRange` three times — no data duplication).
  Positioned below the totals table.
- **Pie chart**: `Charts.ChartType.PIE`, sourced from the two-row helper
  table above. Positioned beside the bar chart.
- Both charts built with `Charts.ChartType` constants and inserted via
  `sheet.insertChart(chart)` before the `SpreadsheetApp.flush()` that already
  precedes the xlsx export fetch — standard column/pie chart types export as
  native, editable Excel charts through the existing
  `/export?format=xlsx` URL fetch (confirmed against Apps Script's chart
  export behavior; no new export mechanism needed).
- Skipped-sheets warning row: unchanged, just repositioned below the (now
  taller) table.

**Employee tabs** — one per `report.rows` entry, in the same order
`report.rows` is already sorted (alphabetical by name), appended after
Summary via `ss.insertSheet(name)`:

- Tab name: `sanitizeSheetName(row.name)` (new helper, see below).
- Header rows (1–2): employee name (14pt bold) + date range subtitle (9pt,
  `#6b7280`) — same visual weight as the Summary title stack, just narrower
  (this tab only needs 8 columns: Date, Day, Start, End, Job Order, Normal,
  OT, Total).
- Header row 4: column headers, styled via the same `styleReportTable`
  helper used on the Summary tab.
- Data rows: this employee's entries (`entries.filter(e => e.employeeKey ===
  normalizeNameKey(row.name))`), sorted by `date` ascending. `Day` column is
  derived (`Utilities.formatDate(date, tz, 'EEE')`) — not present on
  `entries` today, computed inline like `ReportCode.js` already does for OT
  rest-day logic.
- Totals row directly below the data, bold, same border treatment: sums of
  Normal/OT/Total plus a `days worked` count — mirrors the Summary table's
  per-employee row so the two are cross-checkable at a glance.
- Same `sheet.setFrozenRows(headerRow)` + `setColumnWidthsToFitContent` calls
  as every other tab.
- Safety net: if `entries.filter(...)` comes up empty for a name that's in
  `report.rows` (shouldn't happen — both come from the same underlying scan
  — but the scan and the totals aggregation are two separate passes), the tab
  still renders: header + a single "No entries found for this range." row +
  a zeroed totals row, instead of throwing and failing the whole export.

### 4. Sheet-name sanitizing (`Guard.gs` or `Code.js`, next to `sanitizeSheetText`)

```js
/**
 * Makes an employee name safe as both a Google Sheets and an Excel sheet
 * tab name (Excel's 31-char cap is the binding constraint), and de-dupes
 * against names already used in this workbook.
 */
function sanitizeSheetName(name, usedNames) {
  var s = String(name || '').replace(/[\[\]*?/\\:]/g, ' ').trim();
  if (s.length > 31) s = s.substring(0, 31).trim();
  if (!s) s = 'Employee';
  var base = s, n = 2;
  while (usedNames[s]) {
    var suffix = ' (' + n + ')';
    s = base.substring(0, 31 - suffix.length) + suffix;
    n++;
  }
  usedNames[s] = true;
  return s;
}
```

Called once per employee tab with a `usedNames` object shared across the
whole `buildReportExcel` call (so two employees who'd truncate to the same
31 characters still get distinct tabs).

## Data flow

```
processSendReport(payload)
  ├─ generateOtReport(...)              → report (aggregate totals, unchanged)
  ├─ collectTimesheetEntries(...)       → raw per-shift entries (new, shared with listTimesheetEntries)
  └─ buildReportExcel(report, entries)
        ├─ Summary tab   ← report.rows (table + stat banner + both charts)
        └─ per-employee tabs ← entries, grouped by employeeKey, filtered to report.rows names
```

## Error handling

Unchanged at the boundary: `processSendReport`'s existing try/catch around
the `buildReportExcel` call already surfaces any failure as "Excel
generation failed: ... uncheck 'Attach report as Excel' and try again,"
which now also covers the new per-employee/chart code paths. No new
user-facing error states are introduced. The one internal safety net (empty
entries for a known employee) is handled inline as described above rather
than thrown, since it's a data-consistency edge case, not a real failure.

## Testing

This repo has no Apps Script unit-test harness (the one test file,
`test-cloudflare-models.mjs`, is an unrelated standalone Node script for the
AI parser and doesn't run inside Apps Script). Verification is manual, same
as prior features touching `ExcelGen.gs` in this project: after
implementation, send a real report (with "Attach report as Excel" checked)
to a test address covering a multi-employee date range, and confirm in the
downloaded `.xlsx`:

- Summary tab: stat banner numbers match manual sums of the totals table;
  bar chart shows one bar-pair per employee; pie chart shows the correct
  Normal/OT split; column widths are auto-fit (no clipped text, no
  excessive whitespace).
- Each employee tab: row count matches that employee's `daysWorked` from the
  Summary table; totals row matches the Summary table's numbers for that
  employee; tab name is legible and, for a name >31 chars, correctly
  truncated.
- A date range with zero entries for the whole company still sends
  successfully with just the Summary tab (empty-report path, `confirmEmpty`)
  — no employee tabs, no chart-source-table errors on an empty totals array.
