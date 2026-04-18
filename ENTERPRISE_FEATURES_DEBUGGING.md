# Enterprise Features Debugging Guide

This file documents the custom LibreChat changes added in this fork so another LLM or engineer can quickly understand:

- what was added
- where the code lives
- how the major flows work
- how to debug failures

This guide only covers the custom features added in this fork. It does not try to restate all stock LibreChat behavior.

## Features Added

### 1. Request-Level Usage Tracking

Added request-level usage persistence for LLM calls, including fields such as:

- user
- session/conversation
- provider/model
- input/output/total tokens
- latency
- timestamps

Primary files:

- [packages/data-schemas/src/schema/usage.ts](/Users/praneetkotah/Desktop/Development/LibreChat/packages/data-schemas/src/schema/usage.ts:1)
- [packages/data-schemas/src/methods/usage.ts](/Users/praneetkotah/Desktop/Development/LibreChat/packages/data-schemas/src/methods/usage.ts:1)
- [packages/api/src/usage/service.ts](/Users/praneetkotah/Desktop/Development/LibreChat/packages/api/src/usage/service.ts:1)
- [packages/api/src/admin/usage.ts](/Users/praneetkotah/Desktop/Development/LibreChat/packages/api/src/admin/usage.ts:1)
- [api/server/routes/admin/usage.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/routes/admin/usage.js:1)
- [config/migrate-usage-records.js](/Users/praneetkotah/Desktop/Development/LibreChat/config/migrate-usage-records.js:1)

Key endpoints:

- `GET /api/admin/usage`
- `GET /api/admin/usage/summary`

Feature flag:

- `USAGE_TRACKING_ENABLED=true`

### 2. Admin Usage Dashboard

Added an admin-only usage view in Settings that shows:

- overview cards
- user rollups
- recent requests

Primary files:

- [client/src/components/Nav/SettingsTabs/Admin/Admin.tsx](/Users/praneetkotah/Desktop/Development/LibreChat/client/src/components/Nav/SettingsTabs/Admin/Admin.tsx:1)
- [client/src/data-provider/Admin/queries.ts](/Users/praneetkotah/Desktop/Development/LibreChat/client/src/data-provider/Admin/queries.ts:1)

### 3. Cost/Budget Progress Bar

Added a user-facing green-to-red progress bar in the balance UI.

Important note:

- LibreChat native balance is cost-credit based, not literal raw tokens
- current UI may still be interpreted by users as tokens unless relabeled later

Primary files:

- [client/src/components/Nav/SettingsTabs/Balance/TokenUsageProgress.tsx](/Users/praneetkotah/Desktop/Development/LibreChat/client/src/components/Nav/SettingsTabs/Balance/TokenUsageProgress.tsx:1)
- [client/src/components/Nav/SettingsTabs/Balance/Balance.tsx](/Users/praneetkotah/Desktop/Development/LibreChat/client/src/components/Nav/SettingsTabs/Balance/Balance.tsx:1)
- [client/src/components/Nav/AccountSettings.tsx](/Users/praneetkotah/Desktop/Development/LibreChat/client/src/components/Nav/AccountSettings.tsx:1)
- [api/server/controllers/Balance.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/controllers/Balance.js:1)

### 4. Issue Reporting System

Added a dedicated issue-report flow on assistant messages so users can report:

- bad response
- faulty MCP tool
- bad file transformation
- timeout/error
- auth/permission issue

Reports are stored in Mongo and shown in the admin panel.

Primary files:

- [packages/data-schemas/src/schema/issueReport.ts](/Users/praneetkotah/Desktop/Development/LibreChat/packages/data-schemas/src/schema/issueReport.ts:1)
- [packages/data-schemas/src/methods/issueReport.ts](/Users/praneetkotah/Desktop/Development/LibreChat/packages/data-schemas/src/methods/issueReport.ts:1)
- [packages/api/src/issues.ts](/Users/praneetkotah/Desktop/Development/LibreChat/packages/api/src/issues.ts:1)
- [packages/api/src/admin/issues.ts](/Users/praneetkotah/Desktop/Development/LibreChat/packages/api/src/admin/issues.ts:1)
- [api/server/routes/issues.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/routes/issues.js:1)
- [api/server/routes/admin/issues.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/routes/admin/issues.js:1)
- [client/src/components/Chat/Messages/ReportIssueButton.tsx](/Users/praneetkotah/Desktop/Development/LibreChat/client/src/components/Chat/Messages/ReportIssueButton.tsx:1)
- [client/src/components/Chat/Messages/HoverButtons.tsx](/Users/praneetkotah/Desktop/Development/LibreChat/client/src/components/Chat/Messages/HoverButtons.tsx:1)
- [config/migrate-issue-reports.js](/Users/praneetkotah/Desktop/Development/LibreChat/config/migrate-issue-reports.js:1)

Key endpoints:

- `POST /api/issues`
- `GET /api/admin/issues`

### 5. Native Spreadsheet Transform Workflow

Added native spreadsheet handling that returns a new file back in chat.

Current capabilities:

- inspect workbook structure
- keep columns
- remove columns
- redact columns
- output `xlsx` or `csv`

Primary files:

- [api/server/services/Files/Spreadsheets/transform.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/services/Files/Spreadsheets/transform.js:1)
- [api/server/services/Files/Spreadsheets/service.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/services/Files/Spreadsheets/service.js:1)
- [api/app/clients/tools/util/spreadsheet.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/app/clients/tools/util/spreadsheet.js:1)
- [api/server/routes/files/files.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/routes/files/files.js:1)
- [api/app/clients/tools/manifest.json](/Users/praneetkotah/Desktop/Development/LibreChat/api/app/clients/tools/manifest.json:1)
- [packages/api/src/tools/registry/definitions.ts](/Users/praneetkotah/Desktop/Development/LibreChat/packages/api/src/tools/registry/definitions.ts:1)

Native tool name:

- `spreadsheet_transform`

Key endpoint:

- `POST /api/files/:file_id/transform/spreadsheet`

### 6. Native Word Document Workflow

Added native `.docx` handling that returns a new `.docx` file back in chat.

Current capabilities:

- inspect attached `.docx`
- replace exact text
- redact phrases
- prepend/append text
- fully rewrite the document body

Important limitation:

- current implementation regenerates a clean `.docx` from extracted text
- it does not preserve rich source formatting, tables, comments, or tracked changes

Primary files:

- [api/server/services/Files/WordDocuments/transform.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/services/Files/WordDocuments/transform.js:1)
- [api/server/services/Files/WordDocuments/service.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/services/Files/WordDocuments/service.js:1)
- [api/app/clients/tools/util/wordDocument.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/app/clients/tools/util/wordDocument.js:1)
- [api/server/routes/files/files.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/routes/files/files.js:1)
- [api/app/clients/tools/manifest.json](/Users/praneetkotah/Desktop/Development/LibreChat/api/app/clients/tools/manifest.json:1)
- [packages/api/src/tools/registry/definitions.ts](/Users/praneetkotah/Desktop/Development/LibreChat/packages/api/src/tools/registry/definitions.ts:1)

Native tool name:

- `word_document_transform`

Key endpoint:

- `POST /api/files/:file_id/transform/word-document`

## Runtime Expectations

### Docker

This fork was run locally via Docker Compose with a local build override. Important local files:

- [docker-compose.yml](/Users/praneetkotah/Desktop/Development/LibreChat/docker-compose.yml:1)
- [docker-compose.override.yml](/Users/praneetkotah/Desktop/Development/LibreChat/docker-compose.override.yml:1)
- [librechat.yaml](/Users/praneetkotah/Desktop/Development/LibreChat/librechat.yaml:1)
- [.env](/Users/praneetkotah/Desktop/Development/LibreChat/.env:1)

Important note:

- the stock compose file uses the upstream image by default
- the override file is what forces Docker to build the local fork and pick up custom UI/backend changes

### Mongo

The custom features depend on Mongo collections for:

- `Usage`
- `IssueReport`

If the app boots but admin/issue features behave strangely, verify migrations ran and the container is pointed at the expected Mongo DB.

## Debugging By Feature

### Usage Tracking Not Showing Data

Symptoms:

- admin dashboard says usage tracking is disabled
- `/api/admin/usage` returns `503`
- admin dashboard shows error state instead of zero state

Check:

1. Ensure `.env` includes `USAGE_TRACKING_ENABLED=true`
2. Restart the app after changing `.env`
3. Run:
   - `npm run migrate:usage-records`
4. Verify endpoint behavior:
   - unauthenticated should usually return `401`
   - disabled returns `503`
5. Inspect:
   - [packages/api/src/admin/usage.ts](/Users/praneetkotah/Desktop/Development/LibreChat/packages/api/src/admin/usage.ts:1)
   - [packages/api/src/usage/service.ts](/Users/praneetkotah/Desktop/Development/LibreChat/packages/api/src/usage/service.ts:1)

If data is still empty:

- confirm real model requests were made after tracking was enabled
- confirm the request path being exercised is one of the instrumented LLM flows

### Admin Dashboard Missing

Symptoms:

- Admin tab missing
- usage cards missing
- reported issues queue missing

Check:

1. Confirm Docker is building the local fork, not using the stock registry image
2. Confirm the signed-in user is actually `ADMIN`
3. Hard refresh the browser to avoid stale frontend bundles
4. Check:
   - [client/src/components/Nav/SettingsTabs/Admin/Admin.tsx](/Users/praneetkotah/Desktop/Development/LibreChat/client/src/components/Nav/SettingsTabs/Admin/Admin.tsx:1)

### Report Issue Button Missing

Symptoms:

- assistant messages do not show the issue-report action

Check:

1. Hover an assistant message, not a user message
2. Confirm frontend includes:
   - [client/src/components/Chat/Messages/ReportIssueButton.tsx](/Users/praneetkotah/Desktop/Development/LibreChat/client/src/components/Chat/Messages/ReportIssueButton.tsx:1)
   - [client/src/components/Chat/Messages/HoverButtons.tsx](/Users/praneetkotah/Desktop/Development/LibreChat/client/src/components/Chat/Messages/HoverButtons.tsx:1)
3. Rebuild the Docker image and refresh the browser

If submission fails:

1. Verify `POST /api/issues` is mounted
2. Run the migration:
   - `npm run migrate:issue-reports`
3. Inspect:
   - [api/server/routes/issues.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/routes/issues.js:1)
   - [packages/api/src/issues.ts](/Users/praneetkotah/Desktop/Development/LibreChat/packages/api/src/issues.ts:1)

### Spreadsheet Tool Fails

Symptoms:

- tool not available to the model
- route returns `400`
- generated file not attached back into chat

Check tool registration:

- [api/app/clients/tools/util/spreadsheet.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/app/clients/tools/util/spreadsheet.js:1)
- [api/app/clients/tools/util/handleTools.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/app/clients/tools/util/handleTools.js:1)
- [api/app/clients/tools/manifest.json](/Users/praneetkotah/Desktop/Development/LibreChat/api/app/clients/tools/manifest.json:1)
- [packages/api/src/tools/registry/definitions.ts](/Users/praneetkotah/Desktop/Development/LibreChat/packages/api/src/tools/registry/definitions.ts:1)

Check route/service:

- [api/server/routes/files/files.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/routes/files/files.js:1)
- [api/server/services/Files/Spreadsheets/service.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/services/Files/Spreadsheets/service.js:1)
- [api/server/services/Files/Spreadsheets/transform.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/services/Files/Spreadsheets/transform.js:1)

Common causes:

- source file is not a supported spreadsheet MIME type
- requested columns do not exist
- storage strategy does not support `saveBuffer`
- transformed file record is created but frontend is running an old bundle

### Word Document Tool Fails

Symptoms:

- tool not available
- route returns `400`
- file generated but content is missing or simplified

Check tool registration:

- [api/app/clients/tools/util/wordDocument.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/app/clients/tools/util/wordDocument.js:1)
- [api/app/clients/tools/util/handleTools.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/app/clients/tools/util/handleTools.js:1)
- [api/app/clients/tools/manifest.json](/Users/praneetkotah/Desktop/Development/LibreChat/api/app/clients/tools/manifest.json:1)
- [packages/api/src/tools/registry/definitions.ts](/Users/praneetkotah/Desktop/Development/LibreChat/packages/api/src/tools/registry/definitions.ts:1)

Check route/service:

- [api/server/routes/files/files.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/routes/files/files.js:1)
- [api/server/services/Files/WordDocuments/service.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/services/Files/WordDocuments/service.js:1)
- [api/server/services/Files/WordDocuments/transform.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/services/Files/WordDocuments/transform.js:1)

Common causes:

- file is not `.docx`
- no requested transformation matched any text
- another actor expects rich-format preservation, but current implementation intentionally rewrites a clean text-based docx

If the complaint is “formatting disappeared,” that is expected with the current implementation.

### Docker/User Management Confusion

Inside the Docker `api` container, `docker compose exec api ...` starts in `/app/api`, not `/app`.

That means:

- `create-user` can exist in both places
- `list-users` only exists at the monorepo root

Use:

```bash
docker compose exec api sh -lc 'cd /app && npm run list-users'
docker compose exec api sh -lc 'cd /app && npm run create-user'
```

Primary scripts:

- [package.json](/Users/praneetkotah/Desktop/Development/LibreChat/package.json:1)
- [config/create-user.js](/Users/praneetkotah/Desktop/Development/LibreChat/config/create-user.js:1)
- [config/list-users.js](/Users/praneetkotah/Desktop/Development/LibreChat/config/list-users.js:1)

## Useful Test Commands

Issue reporting:

```bash
cd packages/api && npx jest src/issues.spec.ts src/admin/issues.spec.ts --runInBand --coverage=false
cd api && npx jest server/routes/__tests__/issues.spec.js server/routes/__tests__/admin-issues.spec.js --runInBand
```

Word documents:

```bash
cd api && npx jest server/services/Files/WordDocuments/transform.spec.js --runInBand
cd api && npx jest server/routes/files/files.word-document.test.js --runInBand
cd api && npx jest test/app/clients/tools/util/wordDocument.test.js --runInBand
```

Spreadsheets:

```bash
cd api && npx jest server/services/Files/Spreadsheets/transform.spec.js --runInBand
cd api && npx jest server/routes/files/files.transform.test.js --runInBand
cd api && npx jest test/app/clients/tools/util/spreadsheet.test.js --runInBand
```

Shared package build:

```bash
cd packages/api && npm run build
git diff --check
```

## Current Known Limitations

- budget progress currently maps to LibreChat cost credits, not literal token counts
- Word document workflow does not preserve rich source formatting
- admin issues queue is read-only for now
- no commit has been made yet for these local changes unless someone commits them after reading this file

## Suggested First Debug Sequence

When a feature appears broken, use this order:

1. Confirm Docker is running the local fork
2. Confirm `.env` and `librechat.yaml` are mounted into the container
3. Confirm required migrations ran
4. Confirm the user role is correct
5. Hit the backend endpoint directly
6. Run the targeted Jest test for that feature
7. Inspect the specific service/tool/route files linked above

That order catches most failures quickly.
