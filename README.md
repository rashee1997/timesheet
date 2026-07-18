# Timesheet — Apps Script Backend

Google Apps Script project (bound to a Google Sheet) that handles timesheet entry,
validation, OT report generation, and AI-assisted ("QuickFill") parsing of pasted
shift text. Originally used only via Sheets-bound dialogs; now also exposed as a
JSON Web App so the standalone frontend below can call it.

**Frontend repo:** [rashee1997/timesheet-web](https://github.com/rashee1997/timesheet-web) — the Next.js app deployed on Vercel that talks to this backend.

## What's here

- `Code.js` — core entry validation/writing (`coreProcessEntries`), employee/job-order lookups
- `Api.gs` — `doPost` dispatcher exposing a whitelisted set of actions as a Web App JSON API
- `ai.js` — Cloudflare-Workers-AI-backed natural language ("QuickFill") shift parsing
- `EntriesList.gs` — filtered/searched/paginated timesheet entry listing
- `ReportCode.js`, `ScheduledReports.gs`, `ExcelGen.gs` — OT report preview/send/scheduling + Excel export
- `Undo.gs`, `Templates.gs`, `Audit.gs`, `CompanyInfo.gs`, `Guard.gs` — undo-last-save, shift templates, audit log, company branding, sheet-layout guards
- `Form.html`, `ReportForm.html`, `*.html` — the original Sheets-bound dialogs (menu-launched)

## Development

This is a [clasp](https://github.com/google/clasp) project.

```
clasp login
clasp push        # push local changes to the Apps Script editor (HEAD)
clasp deploy -i <deploymentId> -d "description"   # cut a new Web App version
```

`clasp push` only updates the editor's HEAD code — it does **not** update an
already-deployed Web App version. You must `clasp deploy -i <id>` to make code
changes live at the `/exec` URL.

## Web App deployment

`appsscript.json` must declare the webapp block explicitly:

```json
"webapp": { "executeAs": "USER_DEPLOYING", "access": "ANYONE_ANONYMOUS" }
```

`ANYONE_ANONYMOUS` is required (not `ANYONE`) because the frontend calls this
Web App server-to-server, with no Google session of its own — the real
authentication happens in the frontend's proxy layer via a verified Google ID
token, not via Apps Script's own access gate.
