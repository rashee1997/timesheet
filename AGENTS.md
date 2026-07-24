# Timesheet — Apps Script Backend

Local path: `F:\timesheet` (this repo). GitHub: `rashee1997/timesheet`.

Google Apps Script project (bound to a Google Sheet). Handles timesheet entry
validation/writing, OT report generation, AI-assisted QuickFill parsing, and
exposes a whitelisted JSON Web App API (`Api.gs`) consumed by the standalone
frontend.

## Sibling frontend project

Local path: `F:\timesheet-web`. GitHub: `rashee1997/timesheet-web`.

Next.js app deployed on Vercel (`https://timesheet-web-woad.vercel.app`) that
calls this backend's Web App through a server-side proxy
(`src/app/api/gas/route.ts`), which verifies the caller's Google ID token and
forwards requests with a shared secret. When a frontend request needs a new
backend capability, add the handler here and whitelist it in `Api.gs`'s
`API_ACTIONS` map before wiring the frontend call.

## Key files

- `Code.js` — core entry validation/writing (`coreProcessEntries`), employee/job-order lookups
- `Api.gs` — `doPost` dispatcher; `API_ACTIONS` is the whitelist of everything the Web App can reach
- `ai.js` — Cloudflare Workers AI-backed QuickFill natural-language shift parsing
- `EntriesList.gs` — filtered/searched/paginated timesheet entry listing
- `ReportCode.js`, `ScheduledReports.gs`, `ExcelGen.gs` — OT report preview/send/scheduling + Excel export
- `Undo.gs`, `Templates.gs`, `Audit.gs`, `CompanyInfo.gs`, `Guard.gs`, `fileformat.gs` — undo-last-save, shift templates, audit log, company branding, sheet-layout guards, file format handling
- `Form.html`, `ReportForm.html`, `*.html` — original Sheets-bound menu dialogs (still live, separate from the Web App API)

## Development / deployment

This is a [clasp](https://github.com/google/clasp) project.

```
clasp push                                          # push local changes to the Apps Script editor (HEAD)
clasp deploy -i <deploymentId> -d "description"      # cut a new Web App version — this is what goes live
```

`clasp push` alone does **not** update the live `/exec` Web App URL — always
follow it with `clasp deploy -i <productionDeploymentId>` (find the ID via
`clasp deployments`) when a change needs to reach the frontend.

`appsscript.json` must keep `"webapp": { "executeAs": "USER_DEPLOYING", "access": "ANYONE_ANONYMOUS" }` — `ANYONE_ANONYMOUS` is required because the frontend calls this Web App server-to-server with no Google session of its own; real auth happens in the frontend's proxy layer.
