# Current Session Context

Last updated: 2026-05-04

This file is the current working handoff for the LibreChat customization effort. Use this as the primary reference for future turns instead of relying on prior chat history.

## Objective

This LibreChat fork is being turned into an internal enterprise AI platform for GovCloud/GCC High use with:

- Microsoft Entra ID / GCC High SSO
- Outlook inbox and calendar integration through delegated Microsoft Graph access
- enterprise usage tracking and admin reporting
- document and spreadsheet transformation workflows
- future MCP/Jira/Confluence/AWS Bedrock integration
- future Teams archive ingestion and retrieval

## Current Product State

### Core platform

- LibreChat is customized and deployed via Docker Compose.
- AWS Bedrock Claude Sonnet models are the primary LLM backend.
- Token balance / usage controls exist in the product and are surfaced to users.
- Admin reporting exists and has been moved toward a full-page workspace layout instead of a cramped settings-only view.

### Outlook workspace

The Outlook integration is the main active workstream.

Implemented:

- Outlook workspace in the sidebar
- internal workspace tabs for:
  - `Inbox`
  - `Calendar`
- delegated Microsoft Graph access via OBO flow
- inbox list with:
  - focused / other / all views
  - search
  - unread styling
  - bulk selection
  - bulk delete
- email detail view with thread display
- AI actions:
  - analyze email
  - create reply draft
  - daily brief
  - find meeting times
  - schedule meeting
- floating AI assistant panel:
  - resizable
  - sticky controls
  - persists size
- calendar view with:
  - day/week view
  - shared left-side time scale
  - current-time indicator
  - overlapping meetings rendered side by side
  - create/edit/delete event support

### Admin / enterprise features

Implemented or discussed previously in this repo:

- request-level usage tracking
- token accounting persistence
- admin usage endpoints
- user token progress / balance UI
- admin reporting UI
- issue reporting / user feedback flow

### Guided tutorials

Implemented:

- native tutorial overlay system in the client
- manual tutorial launcher in `Settings -> Account`
- no auto-start behavior for new users
- tutorial coverage for:
  - Cortex overview
  - Outlook workspace
  - admin reporting
- stable `data-tour` anchors on:
  - sidebar navigation
  - account menu button
  - Outlook workspace sections
  - admin reporting sections

Current behavior:

- users open Settings, go to Account, and click `Start tutorial`
- selecting a tutorial closes settings and launches the guided overlay
- the tutorial can open the Outlook or admin workspace as needed
- the Outlook tutorial can force the Inbox tab before highlighting inbox-only controls

### File workflows

Implemented or discussed previously:

- Word document transformation flow
- spreadsheet transformation flow
- emphasis on preserving formatting better than plaintext extraction
- special handling for document outputs returned by tools

Current spreadsheet architecture direction:

- keep the existing JavaScript spreadsheet transformer as the default/fallback path
- add an optional Python spreadsheet worker for higher-fidelity `.xlsx` processing
- route both tool-driven and direct file spreadsheet transforms through the shared spreadsheet service
- use environment gating so the Python worker can be enabled gradually
- keep the LLM in a planning role; do not allow arbitrary Python generation/execution

### Teams archive

New backend foundation implemented for Teams archive ingestion and search:

- persistence schemas for:
  - `TeamsArchiveConversation`
  - `TeamsArchiveMessage`
  - `TeamsArchiveSyncJob`
- backend service:
  - [api/server/services/TeamsArchiveService.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/services/TeamsArchiveService.js)
- API routes:
  - `GET /api/teams-archive/status`
  - `POST /api/teams-archive/sync`
  - `GET /api/teams-archive/conversations`
  - `GET /api/teams-archive/conversations/:chatId/messages`
  - `GET /api/teams-archive/search?q=...`
- built-in Cortex tool:
  - `teams_archive_search`
  - actions: `status`, `sync_archive`, `search_messages`, `list_conversations`, `get_messages`
- startup config now exposes `teamsArchiveEnabled`
- `.env.example` now includes `TEAMS_ARCHIVE_*` variables

Current v1 scope:

- user-scoped Teams chat archive sync using delegated Graph access
- stored message body HTML + cleaned text for search
- chat conversation listing and search over archived messages
- no Slack/Teams bot integration
- no channel export support yet
- no UI yet beyond API/config exposure

## Most Recent Session Changes

These are the most recent changes made in this session.

### Outlook calendar UI

- Replaced per-day repeated hour labels with one shared left-side timescale.
- Standardized all-day event spacing so timed grids line up across visible days.
- Changed overlapping meetings from stacked-on-top rendering to side-by-side lane rendering.
- Changed calendar visible hour window from `6 AM - 8 PM` to `9 AM - 7 PM`.

### Spreadsheet worker foundation

Files added:

- [services/spreadsheet-worker/app.py](/Users/praneetkotah/Desktop/Development/LibreChat/services/spreadsheet-worker/app.py)
- [services/spreadsheet-worker/Dockerfile](/Users/praneetkotah/Desktop/Development/LibreChat/services/spreadsheet-worker/Dockerfile)
- [services/spreadsheet-worker/requirements.txt](/Users/praneetkotah/Desktop/Development/LibreChat/services/spreadsheet-worker/requirements.txt)
- [api/server/services/Files/Spreadsheets/workerClient.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/services/Files/Spreadsheets/workerClient.js)

Implemented:

- FastAPI-based internal spreadsheet worker
- workbook inspection endpoint:
  - sheet names
  - header detection
  - preview rows
  - formula samples
  - merged cell discovery
  - detected tables
  - per-column profiles with inferred types and numeric summaries
- transform endpoint for the first high-fidelity operation set:
  - keep/remove/redact columns
  - `add_column`
  - `add_row`
  - `update_cells`
  - `sort_rows`
  - `add_totals_row`
- worker-side workbook validation before returning output
- Node-side worker client plus env-gated routing in the shared spreadsheet service
- JS fallback path retained for:
  - CSV
  - unsupported worker operations
  - worker unavailability
  - disabled worker environments

Current scope:

- Python worker is intentionally limited to `.xlsx` routing for now.
- `add_row` now inherits translated formulas from the template row by default when a value is not explicitly supplied.
- Existing JS transformer still handles broader legacy operations like:
  - `reorder_rows`
  - `merge_sheets`
  - `split_sheet`

### Outlook calendar timezone work

There was a timezone bug where events were rendering in the wrong timezone.

Symptoms:

- Example: an `11 AM` meeting appeared at `3 PM`.

Fixes applied:

#### Backend

File: [api/server/services/OutlookService.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/services/OutlookService.js)

- Calendar fetch now sends a Graph `Prefer: outlook.timezone="..."` header.
- Graph-facing timezone preference now uses a Graph-compatible identifier.
- Mailbox settings timezone values are normalized for app use.
- Calendar event `start` / `end` timezone values are normalized before reaching the client.
- Calendar mutation validation no longer relies on `new Date(dateTime)` for `DateTimeTimeZone` values.
- Calendar create/update payload handling now preserves wall time instead of silently shifting through browser/server timezone assumptions.

#### Frontend

File: [client/src/components/SidePanel/Outlook/Panel.tsx](/Users/praneetkotah/Desktop/Development/LibreChat/client/src/components/SidePanel/Outlook/Panel.tsx)

- Calendar bucket assignment uses timezone-aware parsing.
- Calendar event placement uses timezone-aware parsing.
- Calendar current-time line uses the mailbox/calendar timezone.
- Calendar event labels use timezone-aware formatting.
- Calendar edit form hydration uses timezone-aware parsing.
- Calendar create/edit submission now sends wall time plus the calendar timezone rather than converting everything to UTC on the client.
- Calendar display conversion no longer only handles `UTC` source events.
- Calendar rendering now resolves event wall times against the mailbox timezone first and browser timezone second.
- Calendar bucket keys and current-time matching now use the resolved calendar display timezone rather than raw local-date assumptions.

#### Shared type update

File: [packages/data-provider/src/types/outlook.ts](/Users/praneetkotah/Desktop/Development/LibreChat/packages/data-provider/src/types/outlook.ts)

- `OutlookCalendarResponse` now includes `timeZone`.

### Runtime fix after timezone changes

There was a follow-up runtime error:

- `"CalendarTimeZone" is not defined`

Root cause:

- `OutlookPanel` callbacks referenced `calendarTimeZone`, but the variable only existed inside `CalendarWorkspace`.

Fix:

- Added `const calendarTimeZone = calendarData?.timeZone || calendarData?.workingHours?.timeZone;`
  in `OutlookPanel` component scope.

## Files Most Relevant Right Now

### Outlook frontend

- [client/src/components/SidePanel/Outlook/Panel.tsx](/Users/praneetkotah/Desktop/Development/LibreChat/client/src/components/SidePanel/Outlook/Panel.tsx)

### Outlook backend

- [api/server/services/OutlookService.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/services/OutlookService.js)
- [api/server/routes/outlook.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/routes/outlook.js)

### Outlook client data layer

- [client/src/data-provider/Outlook/queries.ts](/Users/praneetkotah/Desktop/Development/LibreChat/client/src/data-provider/Outlook/queries.ts)
- [packages/data-provider/src/data-service.ts](/Users/praneetkotah/Desktop/Development/LibreChat/packages/data-provider/src/data-service.ts)
- [packages/data-provider/src/api-endpoints.ts](/Users/praneetkotah/Desktop/Development/LibreChat/packages/data-provider/src/api-endpoints.ts)
- [packages/data-provider/src/types/outlook.ts](/Users/praneetkotah/Desktop/Development/LibreChat/packages/data-provider/src/types/outlook.ts)

### Admin / audit / usage

- [api/server/routes/admin/outlookAudit.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/routes/admin/outlookAudit.js)
- [packages/data-schemas/src/schema/outlookAudit.ts](/Users/praneetkotah/Desktop/Development/LibreChat/packages/data-schemas/src/schema/outlookAudit.ts)
- admin reporting UI files under `client/src/components/Nav/SettingsTabs/Admin/` and workspace routing files

## Known Open Items

### Spreadsheet worker

- Default assistant spreadsheet/Word tool outputs now normalize LangChain `content_and_artifact` tuples correctly in `ToolService`, and generated `.xlsx` / `.docx` files are attached to the final assistant message so they render as downloadable files even when the Python worker path is not used.
- Spreadsheet processing is now intended to be Python-primary for supported spreadsheet files instead of silently splitting semantics between JS and Python:
  - worker support now covers `.xlsx`, `.xlsm`, and `.csv`
  - worker-side operations now include `add_column`, `add_row`, `update_cells`, `sort_rows`, `add_totals_row`, `reorder_rows`, `merge_sheets`, and `split_sheet`
  - JS fallback is now opt-in via `SPREADSHEET_WORKER_FALLBACK_TO_JS=true` instead of the default behavior
- Validate the Python worker end-to-end in a running Compose deployment.
- Decide whether `.xlsm` should remain JS-only or get a separate macro-preserving path later.
- Expand worker operations beyond the first set if finance workflows require:
  - richer formula/table semantics
- Add workbook diff / audit detail if users need more explicit change logs.

### Outlook calendar

- Verify timezone fix end-to-end on the deployed instance.
- Confirm:
  - existing events render at correct wall time
  - edit form shows correct wall time
  - newly created events save at correct wall time
- Consider adding an early/late events strip because the visible grid is now limited to `9 AM - 7 PM`.

### Outlook inbox / AI

- Potential future feature: natural-language mailbox querying instead of standard chat.
- Potential future feature: mailbox switching for multiple mailboxes.
- Potential future feature: better rendered email HTML / vendor-rich email display.

### Calendar UX

Suggested next refinements:

- density mode for calendar events
- `+N more` handling when overlap columns get too narrow
- hover/focus detail popover for compact events

### Admin / enterprise

- Keep validating the full-page admin workspace behavior after redeploys.
- Keep verifying usage metrics data sources if numbers look inconsistent.

## Azure / Entra Notes

Observed behavior during testing:

- Outlook Graph delegated access can work in a normal managed Edge profile but fail in incognito with `AADSTS53000`.
- This strongly suggests Conditional Access / compliant-device policy behavior tied to the browser profile/session context.
- Dev and prod differences should be diagnosed through Entra sign-in logs, especially:
  - `Application`
  - `Resource`
  - `Conditional Access`
  - whether `Microsoft Graph` or `All resources` is still in policy scope

## Deployment Notes

For timezone-related Outlook changes, both the API and client need redeploying.

Recommended validation after deploy:

1. Open Outlook workspace.
2. Confirm the page loads without runtime errors.
3. Check a known existing event:
   - displayed slot
   - event detail time
   - edit form time
4. Create a test event at a known wall time and verify Outlook stores it correctly.

## Validation Commands Recently Used

From repo root:

```bash
npm run build:client
node --check api/server/services/OutlookService.js
```

These passed after the latest changes. The existing build warnings about large chunks and PWA glob patterns remain, but no new build errors were introduced.

## Working Rule For Future Turns

Use this document as the primary reference for current project state. Treat chat history as secondary. If new work changes the Outlook/calendar/admin behavior materially, update this file again before ending the session.
