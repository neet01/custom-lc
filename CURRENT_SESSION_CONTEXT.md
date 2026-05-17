# Current Session Context

Last updated: 2026-05-17

This file is the current working handoff for the LibreChat customization effort. Use this as the primary reference for future turns instead of relying on prior chat history.

## Objective

This LibreChat fork is being turned into an internal enterprise AI platform for GovCloud/GCC High use with:

- Microsoft Entra ID / GCC High SSO
- Outlook inbox and calendar integration through delegated Microsoft Graph access
- enterprise usage tracking and admin reporting
- document and spreadsheet transformation workflows
- future MCP/Jira/Confluence/AWS Bedrock integration
- future Teams archive ingestion and retrieval

## Strategic North Star

Cortex is not intended to become just another enterprise chatbot or a thin RAG wrapper.

The long-term goal is to evolve Cortex into an enterprise memory and operational intelligence platform that continuously synthesizes fragmented organizational data into evidence-backed operational insight.

The core problem is that modern enterprises operate across disconnected systems:

- Jira
- Confluence
- GitLab
- Slack
- Teams
- AWS infrastructure
- CI/CD pipelines
- documents
- meetings
- operational telemetry
- and later ERP / EDW systems

Humans currently perform the cross-system reasoning manually. Important knowledge is tribal, fragmented, temporal, and difficult to retrieve.

Cortex exists to continuously build and maintain an evolving understanding of how the organization operates.

Architectural direction:

Raw Enterprise Systems
-> Normalization Layer
-> Canonical Entity / Event Model
-> Enterprise Knowledge Graph / Event Graph
-> Reasoning & Insight Engine
-> Chat / Agents / Dashboards / Automation

The intended canonical model should eventually represent:

- people
- teams
- systems
- repositories
- deployments
- incidents
- projects
- risks
- discussions
- decisions
- documents
- operational events
- ownership relationships
- and temporal change over time

Platform priorities:

- evidence-backed reasoning
- source attribution
- temporal awareness
- cross-system relationships
- delegated-user permissions
- organizational memory
- operational awareness
- explainability
- secure enterprise deployment

Important principle:

- the moat is not the chat UI or the LLM itself
- the moat is the continuously evolving operational memory and reasoning layer built from real enterprise workflows, relationships, and historical context

Engineering implications for future work:

- every integration should be treated as a feed into organizational memory, not as an isolated point feature
- preserve raw source records for auditability
- normalize only enough structure to support reasoning and relationships
- do not overfit the architecture to vector-only retrieval
- prioritize entity resolution, relationship mapping, event correlation, ownership tracking, and timeline analysis
- prioritize incremental operational value over platform sprawl

## Current Execution Priorities

Before expanding the broader enterprise memory layer, Cortex should prioritize the current execution order below.

### Priority 1: Stabilize agent/runtime context handling

- fix `input too long` failures before tool execution
- make summarization/compaction engage earlier and more reliably
- add overflow recovery with a compaction + retry path
- keep tool outputs aggressively bounded and observable

Reason:

- if turns fail before tools run, every additional memory source increases fragility instead of value

### Priority 2: Finish retrieval correctness for current sources

- support full-body retrieval for truncated source records
- guarantee archive/source fallback when memory-layer retrieval returns sparse or empty results
- expose completeness and truncation metadata clearly
- prefer exact source resolution over broad preview search

Reason:

- Cortex cannot become an enterprise memory layer if retrieval is fragmentary or ambiguous

### Priority 3: Build a repeatable testing/eval loop

- service/API integration tests for retrieval and sync flows
- tool-call telemetry to confirm actual action selection
- a fixed eval set of known facts across Teams and uploaded files
- regression checks for fallback, long-body recovery, and participant filters

Reason:

- ad hoc prompting is not a sufficient validation strategy for a memory platform

### Priority 4: Productize reusable uploaded-file workflows

- first-class “use existing file” flow in the chat composer
- multi-select from the user file library
- clear distinction between attach-for-this-run and search-across-my-library

Reason:

- the repo already partially supports reuse of persisted uploads, but the UX is not yet a deliberate product feature

### Priority 5: Complete the document pipeline

- durable `Document` / `DocumentVersion` / `DocumentJob` processing
- async extraction/chunking/indexing workers
- provenance-preserving chunk storage
- processing status / retry visibility

Reason:

- retrieval should not be expanded on top of unstable or incomplete ingestion

### Priority 6: Add user-scoped uploaded-document retrieval

- search/retrieve across the user’s uploaded document corpus
- return file/version/chunk provenance
- expose a dedicated uploaded-documents retrieval tool for agents

Reason:

- this is the first real document memory feed and should be durable before broader cross-source expansion

### Priority 7: Expand the shared enterprise memory layer

- only after the above is stable should additional sources be pulled deeper into the canonical entity/event model

Reason:

- cross-source reasoning is only valuable if the underlying retrieval and ingestion layers are trustworthy

## Current Product State

### Core platform

- LibreChat is customized and deployed via Docker Compose.
- AWS Bedrock Claude Sonnet models are the primary LLM backend.
- Token balance / usage controls exist in the product and are surfaced to users.
- Admin reporting exists and has been moved toward a full-page workspace layout instead of a cramped settings-only view.
- Main chat UI has been simplified from stock LibreChat behavior:
  - visible message forking UI removed
  - visible multi-conversation UI removed
  - assistant sender label displays `Cortex` instead of `bedrock`
  - prompt editing is allowed, assistant-response editing is not
  - user chat bubbles use a translucent Hermeus yellow background
- Chat sidebar and export tweaks implemented:
  - bookmark actions available from the conversation three-dot menu
  - chat export modal now closes after successful export
- Upload menu terminology updated:
  - `Upload as Text`
  - `Upload Custom File`

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
- attachment awareness in the message list and message detail view
- mailbox pagination controls for older/newer messages
- folder picker for Outlook mail folders
- AI actions:
  - analyze email
  - create reply draft
  - daily brief
  - find meeting times
  - schedule meeting
- floating AI assistant panel:
  - opens by default in Outlook workspace
  - draggable
  - resizable
  - sticky controls
  - persists size
  - animated open/close transitions
- calendar view with:
  - day/week view
  - shared left-side time scale
  - current-time indicator
  - overlapping meetings rendered side by side
  - create/edit/delete event support
  - click-open empty slots to create an event
  - Teams meeting join link surfaced on event detail when available

Known Outlook status:

- timezone rendering bug was addressed in both backend and frontend paths
- folder loading/search regressions were corrected by moving back toward safer Graph request shapes
- attachment metadata and download code paths have been revised several times and should be treated as implemented but worth regression-testing after major Outlook changes

### Admin / enterprise features

Implemented or discussed previously in this repo:

- request-level usage tracking
- token accounting persistence
- admin usage endpoints
- user token progress / balance UI
- admin reporting UI
- issue reporting / user feedback flow
- finance CSV export from admin usage reporting
- feature request option added to issue reporting

### Guided tutorials

Implemented:

- native tutorial overlay system in the client
- manual tutorial launcher in `Settings -> Account`
- no auto-start behavior for new users
- tutorial coverage now narrowed to:
  - chat interface + enterprise agents
  - Outlook email analysis flow
- stable `data-tour` anchors on:
  - sidebar navigation
  - chat model selector / composer / send flow
  - Outlook workspace sections and AI assistant actions
  - account menu button

Current behavior:

- users open Settings, go to Account, and click `Start tutorial`
- selecting a tutorial closes settings and launches the guided overlay
- the tutorial can switch between the main chat surface and the Outlook workspace as needed
- the Outlook tutorial can force the Inbox tab and open the floating AI assistant before highlighting the analysis controls
- the overlay/highlight system was refactored away from the large box-shadow spotlight to a four-panel cutout overlay for cleaner formatting

### File workflows

Implemented or discussed previously:

- Word document transformation flow
- spreadsheet transformation flow
- emphasis on preserving formatting better than plaintext extraction
- special handling for document outputs returned by tools
- default spreadsheet transform file-return path fixed so generated files are attached back into chat correctly

Current spreadsheet architecture direction:

- keep the existing JavaScript spreadsheet transformer as the default/fallback path
- add an optional Python spreadsheet worker for higher-fidelity `.xlsx` processing
- route both tool-driven and direct file spreadsheet transforms through the shared spreadsheet service
- use environment gating so the Python worker can be enabled gradually
- keep the LLM in a planning role; do not allow arbitrary Python generation/execution

Current worker direction:

- Python is the intended primary engine for supported spreadsheet processing
- JS remains as fallback/legacy coverage for operations not yet moved over cleanly

### Bedrock file upload workaround

Implemented:

- Bedrock message attachments that exceed Bedrock's document upload limit are now intercepted during `/api/files` upload processing.
- Instead of storing the file as a provider-bound document attachment that later fails at Bedrock Converse time, Cortex now:
  - detects the oversize Bedrock document upload
  - extracts text locally using the existing document parser / text parser path
  - stores the attachment as a `FileSources.text` message attachment

Effect:

- users can continue using `Upload Custom File` with oversized engineering spreadsheets/docs
- the file is still available to the chat as extracted text context
- the provider-side Bedrock document rejection is avoided for this class of files

Current limitation:

- this is a text-extraction fallback, not a true provider-side native document upload
- fidelity for very large/complex spreadsheets is limited to the extracted text representation
- old conversations that already contain provider-bound Bedrock document attachments can still replay those stale attachments on later turns; a fresh conversation/upload path is required to validate the Phase 0 fix cleanly

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
  - actions: `status`, `sync_archive`, `search_messages`, `advanced_search_messages`, `recent_messages`, `list_conversations`, `get_messages`, `get_messages_window`, `summarize_conversation`
- startup config now exposes `teamsArchiveEnabled`
- `.env.example` now includes `TEAMS_ARCHIVE_*` variables

Current v1 scope:

- user-scoped Teams chat archive sync using delegated Graph access
- stored message body HTML + cleaned text for search
- chat conversation listing and search over archived messages
- no Slack/Teams bot integration
- no channel export support yet
- Teams sync/status panel now lives in the MCP builder side-panel surface
- access today is via API routes or the built-in `teams_archive_search` tool

Enterprise memory projection status:

- Teams archive now also projects into a canonical enterprise memory layer after successful sync
- new canonical persistence introduced for:
  - `EnterpriseMemoryEntity`
  - `EnterpriseMemoryRelationship`
  - `EnterpriseMemoryChunk`
  - `EnterpriseMemoryJob`
- current projection scope is intentionally conservative:
  - Teams conversations become `conversation` entities
  - Teams participants/senders/mentions become `person` entities
  - conversation-to-participant edges are stored as relationships
  - Teams messages become retrieval chunks with provenance metadata
- projection failure is recorded separately and does not discard the underlying Teams archive sync result

Enterprise retrieval status:

- Phase 2 has started for Teams retrieval
- new retrieval service:
  - [api/server/services/EnterpriseMemory/retrieval.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/services/EnterpriseMemory/retrieval.js)
- current Phase 2 behavior:
  - `advanced_search_messages` and `recent_messages` attempt `EnterpriseMemoryChunk` retrieval first
  - if enterprise-memory retrieval is unavailable or errors, the service falls back to source-archive retrieval
  - this retrieval path is still lexical/structured, not semantic/vector-based
- Phase 3 has started for Teams retrieval
- current Phase 3 behavior:
  - `summarize_conversation` returns bounded conversation summaries with participant/date-range/highlight metadata
  - `get_messages_window` returns a small local message window around an anchor message or latest topic hit
  - tool execution now logs selected Teams actions so backend traces show when the model chose `advanced_search_messages`, `summarize_conversation`, or `get_messages_window`
  - the intent is to stop pulling full raw threads into prompt context for routine questions
- Teams sync hardening for Monday rollout is now in place:
  - per-user sync execution is protected by an atomic Mongo-backed lease
  - global active syncs are protected by a slot-based concurrency cap
  - new envs:
    - `TEAMS_ARCHIVE_SYNC_STALE_MINUTES`
    - `TEAMS_ARCHIVE_MAX_CONCURRENT_SYNCS`
  - status now exposes `activeSyncs` and `maxConcurrentSyncs` for operator visibility

## Most Recent Session Changes

These are the most recent changes made in this session.

### Document intelligence planning and Phase 1 kickoff

New architecture/design document added:

- [DOCUMENT_INTELLIGENCE_SYSTEM.md](/Users/praneetkotah/Desktop/Development/LibreChat/DOCUMENT_INTELLIGENCE_SYSTEM.md)

Reasoning captured there:

- provider-native uploads are not a stable foundation for enterprise-scale document reasoning
- Cortex needs its own document ingestion, extraction, chunking, retrieval, and lineage model
- S3 + Mongo + later OpenSearch is the preferred storage split
- LibreChat services should remain the processing/orchestration layer, not the durable corpus store

Phase 1 scaffolding implemented in repo:

- new persistence schemas/models/methods for:
  - `Document`
  - `DocumentVersion`
  - `DocumentJob`
- new upload registration service:
  - [api/server/services/Documents/register.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/services/Documents/register.js)
- existing upload flow now auto-registers indexable document-like files into the document pipeline

Current Phase 1 behavior:

- image/audio/video uploads are ignored by the document pipeline
- document-like uploads create:
  - a canonical `Document`
  - initial `DocumentVersion`
  - initial pending `DocumentJob`
- text-backed uploads created by the Bedrock oversize fallback enter the pipeline with:
  - `extractionKind = text`
  - initial job type `chunk`
- binary/provider-backed documents enter the pipeline with:
  - `extractionKind = none`
  - initial job type `extract`

Intent:

- establish durable Cortex-owned document lineage without changing the current chat/file upload contract
- create the substrate for later extraction workers and retrieval flows

Deployment status:

- Phase 1 is implemented in the repo
- production/runtime validation of the `Document`, `DocumentVersion`, and `DocumentJob` side effects is still pending unless explicitly deployed and tested after this handoff

### Bedrock oversized upload Phase 0 validation

What was verified:

- the upload route now resolves `endpointType=bedrock` correctly for agent chats
- oversized Bedrock PDFs are intercepted during `/api/files` upload
- the fallback warning appears in logs:
  - `Falling back to text extraction for oversized Bedrock attachment ...`

Important nuance discovered during testing:

- if a conversation already contains an old provider-bound PDF attachment from before the fallback was deployed, later turns in that same conversation can still replay that stale raw PDF into Bedrock
- this makes the prompt fail even though the new upload path is fixed

Operational guidance:

- validate the fallback in a fresh conversation with a fresh upload
- do not reuse older conversations when checking whether Phase 0 works

Files most relevant to the Phase 0 fix:

- [api/server/services/Files/process.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/services/Files/process.js)
- [client/src/components/Chat/Input/Files/AttachFileMenu.tsx](/Users/praneetkotah/Desktop/Development/LibreChat/client/src/components/Chat/Input/Files/AttachFileMenu.tsx)

### Enterprise memory layer Phase 1 kickoff

New canonical memory scaffolding implemented:

- new data-schemas layer for:
  - `EnterpriseMemoryEntity`
  - `EnterpriseMemoryRelationship`
  - `EnterpriseMemoryChunk`
  - `EnterpriseMemoryJob`
- new DB methods for:
  - entity upsert
  - relationship bulk upsert
  - chunk bulk upsert
  - projection job create/update
  - basic entity/chunk lookup
- new Teams projection service:
  - [api/server/services/EnterpriseMemory/teamsProjection.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/services/EnterpriseMemory/teamsProjection.js)
- Teams sync now triggers that projection after a successful archive sync and returns a `memoryProjection` result block

Current Phase 1 memory behavior:

- source records remain the source of truth:
  - `TeamsArchiveConversation`
  - `TeamsArchiveMessage`
  - `TeamsArchiveSyncJob`
- canonical enterprise memory is a projection layer on top of source records, not a replacement
- visibility is currently user-scoped
- chunks are currently Mongo-backed retrieval units; no OpenSearch indexing yet
- entity resolution is intentionally conservative:
  - `conversation` entities keyed by Teams chat id
  - `person` entities keyed by AAD user id, then email, then display name fallback
- no cross-source linking yet between Teams and Outlook/docs/Jira/Confluence/etc.

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

Latest observability update:

- Added API-side spreadsheet routing logs in:
  - [api/server/services/Files/Spreadsheets/service.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/services/Files/Spreadsheets/service.js)
- Added Python worker request/success/failure logs in:
  - [services/spreadsheet-worker/app.py](/Users/praneetkotah/Desktop/Development/LibreChat/services/spreadsheet-worker/app.py)
- Enabled explicit Uvicorn access logs in:
  - [services/spreadsheet-worker/Dockerfile](/Users/praneetkotah/Desktop/Development/LibreChat/services/spreadsheet-worker/Dockerfile)

Purpose:

- Make it obvious whether a spreadsheet job was:
  - selected for Python worker routing
  - sent to the worker successfully
  - failed and fell back to JS
  - never selected because env or file extension gating prevented it

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

### Document intelligence

- Phase 0 is validated for fresh uploads, but stale historical provider-bound attachments can still contaminate older conversations
- Phase 1 persistence scaffolding exists, but downstream worker consumption of `DocumentJob` records is not implemented yet
- no `DocumentChunk` persistence yet
- no retrieval/indexing over the new document pipeline yet

### Teams archive

- no Teams channel post ingestion yet
- no vector/hybrid retrieval yet; search is stored-text based
- enterprise memory projection exists for Teams, but only as a first source adapter
- no tenant-wide/shared visibility model yet; current projection remains user-scoped
- no OpenSearch or hybrid retrieval over enterprise memory chunks yet
- current Phase 2 chunk retrieval exists for Teams, but semantic ranking, neighbor-window expansion, and cross-source retrieval are not implemented yet

### Outlook

- attachment handling and downloads should continue to be regression-tested after any further Outlook changes
- no native inline attachment preview for PDFs/images yet
- mailbox switching across multiple mailboxes is still future work

### Spreadsheet processing

- worker observability is in place, but end-to-end production confidence still depends on continued validation with real finance workbooks
- the Python-primary vs JS-fallback boundary should be kept explicit to avoid split user experiences until migration is complete

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

For the new Outlook attachment work, both the API and client also need redeploying. The selected email/thread view now expects enriched attachment metadata from the backend and uses a same-origin download route for file retrieval.

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
node --check api/server/routes/outlook.js
node --check api/server/services/Files/Spreadsheets/service.js
python3 -m py_compile services/spreadsheet-worker/app.py
```

Additional passing validation:

```bash
cd api && npx jest --config jest.config.js server/services/OutlookService.spec.js --runInBand
```

These passed after the latest changes. The existing build warnings about large chunks and PWA glob patterns remain, but no new build errors were introduced.

Note:

- `cd api && npx jest --config jest.config.js server/services/Files/Spreadsheets/service.spec.js --runInBand`
  is currently blocked by an unrelated repo test bootstrap issue involving `packages/data-schemas` capability exports, not by the spreadsheet routing changes themselves.

## Latest Outlook Attachment Update

- Outlook inbox list now shows a paperclip indicator for messages with attachments.
- Selected messages and thread messages now render non-inline attachments as downloadable cards in the Cortex UI.
- Backend now fetches attachment metadata for the selected thread only, keeping the mailbox list lightweight.
- New route added:
  - `GET /api/outlook/messages/:messageId/attachments/:attachmentId/download`
- Outlook audit logging now includes:
  - `attachment_downloaded`

## Latest Admin Finance Export Update

- Admin reporting now exposes a finance CSV export from the `Usage by user` tab.
- New backend route:
  - `GET /api/admin/usage/finance-report.csv?days=30`
- Export contents:
  - per-user request and token totals
  - cache token totals
  - estimated input / output / cache write / cache read cost in USD
  - estimated total cost in USD
  - priced vs unpriced request counts
  - model list seen for each user
  - final `TOTAL` row
- Pricing basis:
  - repo model pricing table (`getMultiplier` / `getCacheMultiplier`)
  - intended as an estimate for finance review, not invoice truth
- UI note now explicitly tells admins to reconcile exported estimates against AWS billing / CUR data.

Validation added for this work:

```bash
cd LibreChat && npx jest --config packages/api/jest.config.mjs packages/api/src/admin/usage.spec.ts --runInBand
cd LibreChat && npm run build:client
node --check LibreChat/api/server/routes/admin/usage.js
```

## Working Rule For Future Turns

Use this document as the primary reference for current project state. Treat chat history as secondary. If new work changes the Outlook/calendar/admin behavior materially, update this file again before ending the session.
