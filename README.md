<img src="icon.svg" width="64" height="64" alt="Timesheet logo">

# Timesheet — Apps Script Backend

Google Apps Script project (bound to a Google Sheet) that handles timesheet entry,
validation, OT report generation, and AI-assisted ("QuickFill") parsing of pasted
shift text. Originally used only via Sheets-bound dialogs; now also exposed as a
JSON Web App so the standalone frontend below can call it.

**Frontend repo:** [rashee1997/timesheet-web](https://github.com/rashee1997/timesheet-web) — the Next.js app deployed on Vercel that talks to this backend.

## What's here

- `Code.js` — core entry validation/writing (`coreProcessEntries`), employee/job-order lookups
- `Api.gs` — `doPost` dispatcher exposing a whitelisted set of actions as a Web App JSON API
- `ai.js` — Cloudflare Workers AI-backed natural language ("QuickFill") shift parsing
- `EntriesList.gs` — filtered/searched/paginated timesheet entry listing
- `ReportCode.js`, `ScheduledReports.gs`, `ExcelGen.gs` — OT report preview/send/scheduling + Excel export
- `Undo.gs`, `Templates.gs`, `Audit.gs`, `CompanyInfo.gs`, `Guard.gs`, `fileformat.gs` — undo-last-save, shift templates, audit log, company branding, sheet-layout guards, file format handling
- `Form.html`, `ReportForm.html`, `*.html` — the original Sheets-bound dialogs (menu-launched)

## Architecture

```
 Browser (Google account holder)
        │
        ▼
 Next.js app (Vercel) ── verifies Google ID token
        │
        │  POST /exec  { action, payload, actorEmail, sharedSecret }
        ▼
 Apps Script Web App (this repo) ── Api.gs::doPost
        │  1. checks sharedSecret against Script Properties
        │  2. looks up `action` in the API_ACTIONS whitelist
        │  3. calls the matching handler, e.g. Code.gs::submitTimeEntry
        ▼
 Google Sheet (source of truth)
```

**Two independent checks stack on every request:**

1. **Identity** — the Next.js proxy (`app/api/gas/route.ts` in the frontend repo) verifies the caller is a real, logged-in company Google account via `google-auth-library`, since Apps Script has no npm access to do this itself.
2. **Origin** — this Web App only trusts requests carrying the shared secret stored in Script Properties. `doPost` rejects anything without a matching secret before touching the whitelist, so even a leaked `/exec` URL is useless without it.

`doPost` cannot read custom HTTP headers (Google Issue Tracker #67764685, official won't-fix), so the secret travels inside the JSON body instead of a header.

Only actions listed in `Api.gs`'s `API_ACTIONS` map are reachable from the outside — everything else in the script is unreachable from the Web App, by construction rather than by convention.

## Why Cloudflare Workers AI

QuickFill (natural-language shift parsing) and the log-chat feature both send timesheet data to an LLM. Cloudflare Workers AI was chosen over other free-tier AI providers for its stronger default data-handling posture — timesheet data includes real employee names and work patterns, so the inference provider matters. See `ai.gs` for the integration.

## Development

This is a [clasp](https://github.com/google/clasp) project.

```
clasp login
clasp push        # push local changes to the Apps Script editor (HEAD)
clasp deploy -i <deploymentId> -d "description"   # cut a new Web App version
```

`clasp push` only updates the editor's HEAD code — it does **not** update an already-deployed Web App version. You must `clasp deploy -i <id>` to make code changes live at the `/exec` URL.

## Web App deployment

`appsscript.json` must declare the webapp block explicitly:

```json
"webapp": { "executeAs": "USER_DEPLOYING", "access": "ANYONE_ANONYMOUS" }
```

`ANYONE_ANONYMOUS` is required (not `ANYONE`) because the frontend calls this Web App server-to-server, with no Google session of its own — the real authentication happens in the frontend's proxy layer via a verified Google ID token, not via Apps Script's own access gate.

### Rotating the shared secret

Run `setupApiSharedSecret()` once from the Apps Script editor to generate and store a new secret in Script Properties. Re-running it **rotates** the secret — the old one stops working immediately, so the new value must also be updated in the Vercel project's `GAS_SHARED_SECRET` environment variable in the same maintenance window, or every proxied request will start failing with "Not authorized."

## Sibling frontend project

Next.js app deployed on Vercel that calls this backend's Web App through a server-side proxy (`src/app/api/gas/route.ts`), which verifies the caller's Google ID token and forwards requests with the shared secret. When a frontend request needs a new backend capability: add the handler in the relevant `.gs` file here, then whitelist it in `Api.gs`'s `API_ACTIONS` map before wiring the frontend call. See [rashee1997/timesheet-web](https://github.com/rashee1997/timesheet-web).
