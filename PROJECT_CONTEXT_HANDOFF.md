# LibreChat Enterprise Project Context Handoff

Last updated: 2026-06-16

This document is a high-signal handoff for another LLM or engineer.

It is meant to answer:

- what this project is trying to become
- what constraints it operates under
- what has already been implemented
- what related repositories exist
- what design decisions have already been made
- what is still pending

This is broader than [ENTERPRISE_FEATURES_DEBUGGING.md](/Users/praneetkotah/Desktop/Development/LibreChat/ENTERPRISE_FEATURES_DEBUGGING.md:1), which is focused more on how to debug the custom features in this LibreChat fork.

## Cross-Machine Note

There may be temporary divergence between this local workspace and the user's work laptop.

One specifically reported work-laptop fix was:

- file outputs from native document tools were not being streamed back through `ToolService.js`
- the fix added `FILE` handling in `packages/data-provider/src/types/runs.ts`
- the fix also updated `api/server/services/ToolService.js` so tools like `word_document_transform` and `spreadsheet_transform` can return downloadable files to the UI instead of only pasting text in chat

If document transforms work on one machine but not another, compare those files first before assuming the Word/spreadsheet services are broken.

## Project Goal

This LibreChat fork is being turned into an internal enterprise AI platform, effectively an `EnterpriseGPT`, for a regulated environment.

Long-term goals include:

- enterprise chat UI built on LibreChat
- usage observability and admin oversight
- budget awareness and anti-runaway-spend controls
- native file transformation workflows
- integrations for Jira, Confluence, SharePoint, Slack, and other Office 365 / enterprise tools
- use of AWS Bedrock models and Bedrock-based agent patterns
- future ability to execute enterprise tasks on behalf of users through safe integrations

The current implementation effort has focused on the LibreChat core platform first, not on every integration at once.

## Environment And Security Constraints

Assumptions discussed for this project:

- hosted entirely in AWS GovCloud (US)
- no reliance on public SaaS unless explicitly approved
- all data remains within controlled infrastructure
- private subnets preferred
- compatible with CUI / high-security enterprise environments
- strict network boundaries expected
- prefer AWS-native or self-hosted components

Identity assumptions:

- authentication handled via Microsoft Entra ID / Azure AD in GCC High
- OIDC / OAuth flows must work with GCC High endpoints
- RBAC should eventually derive from Entra group claims
- no local username/password auth should be introduced as the target production auth model

Important note:

- for local development, temporary local/test users were used where needed to keep the app testable
- that is a dev convenience, not the intended final enterprise auth design

## Repositories In Play

There are two repos relevant to this effort.

### 1. This Repo: LibreChat Fork

Path:

- [LibreChat](/Users/praneetkotah/Desktop/Development/LibreChat)

Purpose:

- core chat platform
- admin dashboard and enterprise UI enhancements
- usage tracking and budget display
- native file workflows
- issue reporting
- future place where integrations are surfaced to the user

### 2. Separate MCP Repo

Path:

- `/Users/praneetkotah/Desktop/Development/enterprise-mcp-services`

Purpose:

- separate MCP services for enterprise integrations
- currently intended for things like Jira and Confluence
- meant to be deployable separately from LibreChat

That repo was created specifically because the integration layer should evolve independently from LibreChat itself.

## What Has Been Implemented In This LibreChat Fork

The following items are implemented locally in this fork.

### 1. Request-Level Usage Tracking

Implemented:

- request-level usage capture across LLM calls
- persistence in Mongo via a custom `Usage` collection
- admin APIs for recent usage and summary rollups

Tracked fields include:

- user
- session ID
- conversation ID
- provider
- model
- input tokens
- output tokens
- total tokens
- latency
- timestamps

Main code areas:

- [packages/data-schemas/src/schema/usage.ts](/Users/praneetkotah/Desktop/Development/LibreChat/packages/data-schemas/src/schema/usage.ts:1)
- [packages/data-schemas/src/methods/usage.ts](/Users/praneetkotah/Desktop/Development/LibreChat/packages/data-schemas/src/methods/usage.ts:1)
- [packages/api/src/usage/service.ts](/Users/praneetkotah/Desktop/Development/LibreChat/packages/api/src/usage/service.ts:1)
- [packages/api/src/admin/usage.ts](/Users/praneetkotah/Desktop/Development/LibreChat/packages/api/src/admin/usage.ts:1)
- [api/server/routes/admin/usage.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/routes/admin/usage.js:1)

Feature flag:

- `USAGE_TRACKING_ENABLED=true`

### 2. Admin Usage Dashboard

Implemented:

- admin-only Settings dashboard for enterprise usage visibility
- usage overview cards
- user-level rollups
- recent request tracking

Main code areas:

- [client/src/components/Nav/SettingsTabs/Admin/Admin.tsx](/Users/praneetkotah/Desktop/Development/LibreChat/client/src/components/Nav/SettingsTabs/Admin/Admin.tsx:1)
- [client/src/data-provider/Admin/queries.ts](/Users/praneetkotah/Desktop/Development/LibreChat/client/src/data-provider/Admin/queries.ts:1)

### 3. Budget / Balance Progress Bar

Implemented:

- user-facing green-to-red progress bar
- relabeled away from fake "tokens remaining" semantics
- now treated as estimated budget / spend instead of pretending to be raw tokens

Important design conclusion:

- LibreChat native balance is cost-credit based
- it should be used as a budget proxy, not as the source of truth for raw token counts
- raw token counts come from `Usage` records

Current pricing modeling is based on AWS GovCloud usage of:

- Claude Sonnet 3.7
- Claude Sonnet 4.5

Modeled rates:

- input: `$3 / 1M`
- output: `$15 / 1M`
- cache write: `$3.75 / 1M`
- cache read: `$0.30 / 1M`

Main code areas:

- [client/src/components/Nav/SettingsTabs/Balance/TokenUsageProgress.tsx](/Users/praneetkotah/Desktop/Development/LibreChat/client/src/components/Nav/SettingsTabs/Balance/TokenUsageProgress.tsx:1)
- [packages/data-schemas/src/methods/tx.ts](/Users/praneetkotah/Desktop/Development/LibreChat/packages/data-schemas/src/methods/tx.ts:1)

Important numeric rule:

- `1,000,000` balance credits = `$1`

Example sizing:

- `$10` budget = `10,000,000`
- `$25` budget = `25,000,000`
- `$50` budget = `50,000,000`

### 4. Issue Reporting System

Implemented:

- dedicated `Report issue` action on assistant messages
- categories for:
  - bad response
  - faulty MCP tool
  - bad file transformation
  - timeout/error
  - auth/permission issue
- persistent `IssueReport` collection
- admin issue queue in the admin dashboard

Main code areas:

- [packages/data-schemas/src/schema/issueReport.ts](/Users/praneetkotah/Desktop/Development/LibreChat/packages/data-schemas/src/schema/issueReport.ts:1)
- [packages/api/src/issues.ts](/Users/praneetkotah/Desktop/Development/LibreChat/packages/api/src/issues.ts:1)
- [packages/api/src/admin/issues.ts](/Users/praneetkotah/Desktop/Development/LibreChat/packages/api/src/admin/issues.ts:1)
- [api/server/routes/issues.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/routes/issues.js:1)
- [api/server/routes/admin/issues.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/routes/admin/issues.js:1)
- [client/src/components/Chat/Messages/ReportIssueButton.tsx](/Users/praneetkotah/Desktop/Development/LibreChat/client/src/components/Chat/Messages/ReportIssueButton.tsx:1)

Current scope:

- read-only issue queue for admins
- no triage workflow or status changes yet

### 5. Outlook Workspace And AI Inbox

Implemented:

- Outlook workspace in the LibreChat sidebar with `Inbox` and `Calendar` tabs
- delegated Microsoft Graph access through the shared OBO token path
- folder-aware mailbox browsing with search, focused/other/all views, thread display, unread state updates, and delete actions
- attachment awareness in both the mailbox list and selected-message view, plus same-origin attachment downloads
- AI actions for:
  - message analysis
  - draft generation
  - daily brief
  - meeting-slot proposal
  - meeting creation
- calendar visibility and event CRUD inside LibreChat, including timezone-aware rendering and Teams meeting join links when present

Main code areas:

- [api/server/routes/outlook.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/routes/outlook.js:1)
- [api/server/services/OutlookService.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/services/OutlookService.js:1)
- [client/src/components/SidePanel/Outlook/Panel.tsx](/Users/praneetkotah/Desktop/Development/LibreChat/client/src/components/SidePanel/Outlook/Panel.tsx:1)
- [client/src/data-provider/Outlook/queries.ts](/Users/praneetkotah/Desktop/Development/LibreChat/client/src/data-provider/Outlook/queries.ts:1)

Feature flags/config:

- `OUTLOOK_AI_ENABLED=true`
- `.env.example` includes `OUTLOOK_GRAPH_*` and `OUTLOOK_AI_*` settings

### 6. Outlook Audit Trail And Admin Visibility

Implemented:

- persistent `OutlookAudit` records for mailbox, calendar, attachment, and AI actions
- admin endpoint for paginated Outlook audit review
- admin dashboard surface for the audit stream
- usage/balance persistence for Outlook AI actions through the same usage accounting path used elsewhere in the fork

Main code areas:

- [packages/data-schemas/src/schema/outlookAudit.ts](/Users/praneetkotah/Desktop/Development/LibreChat/packages/data-schemas/src/schema/outlookAudit.ts:1)
- [packages/data-schemas/src/methods/outlookAudit.ts](/Users/praneetkotah/Desktop/Development/LibreChat/packages/data-schemas/src/methods/outlookAudit.ts:1)
- [api/server/routes/outlook.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/routes/outlook.js:1)
- [api/server/routes/admin/outlookAudit.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/routes/admin/outlookAudit.js:1)
- [client/src/components/Nav/SettingsTabs/Admin/Admin.tsx](/Users/praneetkotah/Desktop/Development/LibreChat/client/src/components/Nav/SettingsTabs/Admin/Admin.tsx:1)
- [config/migrate-outlook-audit.js](/Users/praneetkotah/Desktop/Development/LibreChat/config/migrate-outlook-audit.js:1)

Key endpoints:

- `GET /api/outlook/*`
- `POST /api/outlook/*`
- `PATCH /api/outlook/*`
- `DELETE /api/outlook/*`
- `GET /api/admin/outlook-audit`

### 7. Teams Archive And Retrieval

Implemented:

- user-scoped Teams chat archive sync with delegated Graph access
- status, sync, cancel, and reset flows
- archived conversation listing and message retrieval
- lexical message search plus richer tool-driven retrieval paths
- account-level Teams sync status card in the side-panel surface

Main code areas:

- [api/server/routes/teamsArchive.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/routes/teamsArchive.js:1)
- [api/server/services/TeamsArchiveService.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/services/TeamsArchiveService.js:1)
- [api/app/clients/tools/util/teamsArchive.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/app/clients/tools/util/teamsArchive.js:1)
- [client/src/components/Nav/SettingsTabs/Account/TeamsArchiveStatus.tsx](/Users/praneetkotah/Desktop/Development/LibreChat/client/src/components/Nav/SettingsTabs/Account/TeamsArchiveStatus.tsx:1)
- [packages/data-schemas/src/schema/teamsArchiveConversation.ts](/Users/praneetkotah/Desktop/Development/LibreChat/packages/data-schemas/src/schema/teamsArchiveConversation.ts:1)
- [packages/data-schemas/src/schema/teamsArchiveMessage.ts](/Users/praneetkotah/Desktop/Development/LibreChat/packages/data-schemas/src/schema/teamsArchiveMessage.ts:1)
- [packages/data-schemas/src/schema/teamsArchiveSyncJob.ts](/Users/praneetkotah/Desktop/Development/LibreChat/packages/data-schemas/src/schema/teamsArchiveSyncJob.ts:1)
- [packages/data-schemas/src/schema/teamsArchiveBackfillState.ts](/Users/praneetkotah/Desktop/Development/LibreChat/packages/data-schemas/src/schema/teamsArchiveBackfillState.ts:1)
- [packages/data-schemas/src/schema/teamsArchiveSyncLease.ts](/Users/praneetkotah/Desktop/Development/LibreChat/packages/data-schemas/src/schema/teamsArchiveSyncLease.ts:1)

Current API/tool surface:

- routes:
  - `GET /api/teams-archive/status`
  - `POST /api/teams-archive/sync`
  - `POST /api/teams-archive/cancel`
  - `POST /api/teams-archive/reset`
  - `GET /api/teams-archive/conversations`
  - `GET /api/teams-archive/conversations/:chatId/messages`
  - `GET /api/teams-archive/search`
- built-in tool:
  - `teams_archive_search`
  - actions include `recent_meeting_chats`, `conversation_dossier`, `conversation_recent_messages`, `conversation_sender_messages`, `conversation_activity_diagnostics`, `sender_identity_report`, `get_message_body`, `get_messages_window`, and `summarize_conversation`

Current scope/limits:

- chat archive only, not Teams channel export
- access today is through API routes and the built-in tool, not a dedicated Teams workspace UI
- retrieval is source-backed and bounded; it is not yet a fully semantic/vector stack

### 8. Enterprise Memory Projection For Teams

Implemented:

- canonical memory persistence for:
  - `EnterpriseMemoryEntity`
  - `EnterpriseMemoryRelationship`
  - `EnterpriseMemoryChunk`
  - `EnterpriseMemoryJob`
- Teams archive sync projects conversations, people, relationships, and message chunks into that canonical layer
- retrieval paths such as `advanced_search_messages` and `recent_messages` attempt enterprise-memory lookup first and fall back to the source archive when needed

Main code areas:

- [api/server/services/EnterpriseMemory/teamsProjection.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/services/EnterpriseMemory/teamsProjection.js:1)
- [api/server/services/EnterpriseMemory/retrieval.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/services/EnterpriseMemory/retrieval.js:1)
- [packages/data-schemas/src/schema/enterpriseMemoryEntity.ts](/Users/praneetkotah/Desktop/Development/LibreChat/packages/data-schemas/src/schema/enterpriseMemoryEntity.ts:1)
- [packages/data-schemas/src/schema/enterpriseMemoryRelationship.ts](/Users/praneetkotah/Desktop/Development/LibreChat/packages/data-schemas/src/schema/enterpriseMemoryRelationship.ts:1)
- [packages/data-schemas/src/schema/enterpriseMemoryChunk.ts](/Users/praneetkotah/Desktop/Development/LibreChat/packages/data-schemas/src/schema/enterpriseMemoryChunk.ts:1)
- [packages/data-schemas/src/schema/enterpriseMemoryJob.ts](/Users/praneetkotah/Desktop/Development/LibreChat/packages/data-schemas/src/schema/enterpriseMemoryJob.ts:1)

Current limits:

- retrieval is still lexical/structured, not OpenSearch/vector-backed
- cross-source linking between Teams, Outlook, documents, Jira, and Confluence is not implemented yet
- projection failure is tracked separately and does not invalidate the underlying Teams archive sync

### 8a. Slack Archive / GovSlack Scaffold

Implemented:

- GovSlack archive persistence scaffolding for:
  - `SlackArchiveConversation`
  - `SlackArchiveMessage`
  - `SlackArchiveSyncJob`
  - `SlackArchiveSyncLease`
- backend routes for:
  - `GET /api/slack-archive/status`
  - `GET /api/slack-archive/oauth/start`
  - `GET /api/slack-archive/oauth/callback`
  - `POST /api/slack-archive/sync`
  - `POST /api/slack-archive/cancel`
  - `POST /api/slack-archive/reset`
  - `GET /api/slack-archive/conversations`
  - `GET /api/slack-archive/conversations/:conversationId/messages`
  - `GET /api/slack-archive/search`
- built-in Cortex tool:
  - `slack_archive_search`
- side-panel Slack archive status card now lives in the MCP builder surface directly under the Teams archive card

Main code areas:

- [api/server/services/SlackArchiveService.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/services/SlackArchiveService.js:1)
- [api/server/services/SlackArchiveOAuthService.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/services/SlackArchiveOAuthService.js:1)
- [api/server/routes/slackArchive.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/routes/slackArchive.js:1)
- [api/app/clients/tools/util/slackArchive.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/app/clients/tools/util/slackArchive.js:1)
- [client/src/components/Nav/SettingsTabs/Account/SlackArchiveStatus.tsx](/Users/praneetkotah/Desktop/Development/LibreChat/client/src/components/Nav/SettingsTabs/Account/SlackArchiveStatus.tsx:1)
- [client/src/components/SidePanel/MCPBuilder/MCPBuilderPanel.tsx](/Users/praneetkotah/Desktop/Development/LibreChat/client/src/components/SidePanel/MCPBuilder/MCPBuilderPanel.tsx:1)

Current scope/limits:

- current implementation is scaffold-level only for GovSlack archive support
- GovSlack OAuth install/callback is wired, but Slack Web API conversation/message ingestion is not implemented yet
- sync attempts currently fail intentionally with `501` until ingestion is added
- current token persistence is stored per Cortex user for development convenience; production should move the bot/workspace install into a dedicated workspace-scoped install model
- there is not yet a `SlackIdentityLink` or equivalent mapping from `team_id + slack_user_id` to the Cortex/Entra user, so future `/Cortex` DM or channel flows are not ready

GovSlack-specific design decisions:

- use GovSlack domains and app configuration, not commercial Slack defaults:
  - OAuth authorize host: `https://slack-gov.com`
  - API base: `https://slack-gov.com/api`
- keep the archive and the future bot as thin adapters into the existing Cortex runtime, retrieval, and enterprise-memory pipeline
- for the future interactive bot, prefer a dedicated `slack_worker` or equivalent event consumer rather than bolting Slack event handling into the main web process
- shared-channel and group-chat responses should default to the least-privileged safe surface, typically DM or ephemeral response patterns where possible

### 9. Document Pipeline Scaffolding

Implemented:

- canonical document persistence for:
  - `Document`
  - `DocumentVersion`
  - `DocumentJob`
- upload registration now creates durable document-side records for document-like uploads
- oversized Bedrock uploads can be converted into text-backed document pipeline inputs instead of remaining provider-bound attachments

Main code areas:

- [api/server/services/Documents/register.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/services/Documents/register.js:1)
- [api/server/services/Files/process.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/services/Files/process.js:1)
- [packages/data-schemas/src/schema/document.ts](/Users/praneetkotah/Desktop/Development/LibreChat/packages/data-schemas/src/schema/document.ts:1)
- [packages/data-schemas/src/schema/documentVersion.ts](/Users/praneetkotah/Desktop/Development/LibreChat/packages/data-schemas/src/schema/documentVersion.ts:1)
- [packages/data-schemas/src/schema/documentJob.ts](/Users/praneetkotah/Desktop/Development/LibreChat/packages/data-schemas/src/schema/documentJob.ts:1)

Current limit:

- worker consumption of `DocumentJob` records is not implemented yet, so this is durable scaffolding rather than a complete document-intelligence runtime

### 10. Native Spreadsheet File Workflow

Implemented natively in LibreChat, not as an MCP integration.

Capabilities added:

- inspect spreadsheets
- keep columns
- remove columns
- redact columns
- target specific sheets
- output transformed file as `xlsx` or `csv`
- return the generated file back into chat

Main code areas:

- [api/server/services/Files/Spreadsheets/transform.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/services/Files/Spreadsheets/transform.js:1)
- [api/server/services/Files/Spreadsheets/service.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/services/Files/Spreadsheets/service.js:1)
- [api/app/clients/tools/util/spreadsheet.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/app/clients/tools/util/spreadsheet.js:1)
- [api/server/routes/files/files.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/routes/files/files.js:1)

Native tool name:

- `spreadsheet_transform`

Also added:

- spreadsheet promotion into native code execution flow so richer analytical tasks can use the code environment

Important product direction:

- deterministic cleanup/export goes through the native spreadsheet tool
- deeper finance-style analysis can use code execution and then return a real file

Known limitation:

- `.xlsm` / macro-enabled spreadsheets are not a supported first-class workflow right now
- preserving VBA/macros is not currently implemented
- the agreed product direction was to ask finance users to upload non-macro spreadsheets for now

### 11. Native Word Document Workflow

Implemented natively in LibreChat, not as an MCP integration.

Capabilities added:

- inspect `.docx`
- replace exact text
- redact phrases
- prepend text
- append text
- rewrite the body and return a new `.docx`

Main code areas:

- [api/server/services/Files/WordDocuments/transform.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/services/Files/WordDocuments/transform.js:1)
- [api/server/services/Files/WordDocuments/service.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/services/Files/WordDocuments/service.js:1)
- [api/app/clients/tools/util/wordDocument.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/app/clients/tools/util/wordDocument.js:1)

Native tool name:

- `word_document_transform`

Known limitation:

- current implementation recreates a clean `.docx` from extracted text
- it does not preserve complex source formatting, tables, comments, or tracked changes
- if the backend document transform succeeds but the UI only shows pasted text, also inspect `ToolService.js` / run-content handling for file artifact streaming

### 12. Enterprise Debugging Guide

Added:

- [ENTERPRISE_FEATURES_DEBUGGING.md](/Users/praneetkotah/Desktop/Development/LibreChat/ENTERPRISE_FEATURES_DEBUGGING.md:1)

Purpose:

- gives another model or engineer concrete debug entry points for the enterprise features

## Local Development / Testing Workflow

The stated goal for development was to create a local loop where the app can be started, changed, tested, and validated autonomously.

What was established:

- local Mongo-backed LibreChat development
- local Docker-based startup
- migrations for custom collections
- targeted test runs for new features
- health checks and live endpoint validation

Important local runtime detail:

- stock LibreChat Docker Compose uses the upstream image
- a local override is needed so Docker builds this fork instead of running the stock LibreChat image

Important local files:

- [docker-compose.yml](/Users/praneetkotah/Desktop/Development/LibreChat/docker-compose.yml:1)
- [docker-compose.override.yml](/Users/praneetkotah/Desktop/Development/LibreChat/docker-compose.override.yml:1)
- [librechat.yaml](/Users/praneetkotah/Desktop/Development/LibreChat/librechat.yaml:1)
- [.env](/Users/praneetkotah/Desktop/Development/LibreChat/.env:1)

Important migration scripts:

- `npm run migrate:usage-records`
- `npm run migrate:issue-reports`

## Notes About Docker And User Management

An implementation/detail quirk discovered during local development:

- inside the Docker `api` container, some npm scripts are available from `/app`
- but `docker compose exec api` often lands in `/app/api`

Example:

- `npm run create-user` can appear to work from one path
- `npm run list-users` may fail unless run from `/app`

Working pattern used:

```bash
docker compose exec api sh -lc 'cd /app && npm run list-users'
docker compose exec api sh -lc 'cd /app && npm run create-user'
```

This mattered during local auth recovery and test-user setup.

## Current MCP / Integration Strategy

The design direction that was agreed to:

- external systems like Jira, Confluence, and probably Outlook belong in separate integration services
- those services should not be tightly coupled into LibreChat core logic
- MCP is appropriate for enterprise integrations
- native file handling should stay inside LibreChat, not be outsourced to MCP

The preferred pattern:

- `LibreChat` is the user-facing orchestration/UI layer
- `MCP services` are the integration contract
- `Bedrock agents` are used only where higher-order planning/orchestration adds value
- deterministic actions and retrieval should not always go through an agent

## Jira And Confluence Architecture Decisions

These discussions already happened and should not need to be rediscovered.

### Hosting Direction

Recommended direction:

- shared internal MCP services
- not one MCP server per user
- deploy as shared stateless services
- prefer ECS for hosted MCP services rather than EKS or EC2 unless there is a strong reason otherwise

Rationale:

- lower operational complexity than EKS
- better scaling story than per-user servers
- less reconnect churn
- better fit for internal containerized services

### Jira Direction

Recommended direction:

- common deterministic Jira operations should be explicit MCP tools
- complex analysis/workflow can use a Bedrock agent when needed

Examples of deterministic Jira tools:

- list issues for the user
- get issue
- search issues
- create issue
- transition issue

Pattern:

- simple reads/writes should not always go through a Bedrock agent first
- agent-assisted workflows are better reserved for complex reasoning tasks

### Confluence Direction

Recommended direction:

- retrieval-first, not agent-first
- use a retrieval service / vector search / normalized document pipeline
- allow the model to synthesize answers from retrieved content

Important background:

- the current user environment already has a Confluence setup that does vector search on XML documents stored in S3
- those are condensed versions of actual pages

Recommended architecture:

- Confluence ingestion/indexing service
- retrieval API
- thin MCP facade

### Atlassian Auth Decision

Important conclusion:

- for self-hosted Jira/Confluence Data Center, passing the Entra token directly did not work
- per-user delegated access should use Atlassian OAuth 2.0 incoming links
- PATs are not the right interactive delegated model

Chosen direction:

- LibreChat should own the OAuth flow and token storage where possible
- MCP servers should accept the bearer token from LibreChat and forward it to Jira/Confluence
- do not build an entirely separate custom auth UI in the MCP service if LibreChat’s MCP OAuth support can be used

The separate MCP repo was updated accordingly to support delegated bearer pass-through.

## Separate MCP Repo Status

In the separate `enterprise-mcp-services` repo, the following direction was taken:

- Dockerized MCP runtime
- Jira MCP server
- Confluence MCP scaffold
- direct tools for basic operations
- optional Bedrock-backed analysis paths for more complex tasks
- request-context support for delegated auth headers

Key design choice:

- basic queries should go directly to the API
- more complex reasoning tasks can route to a Bedrock agent

This repo is intentionally separate from LibreChat.

## Outlook Implementation Context

Outlook is implemented in this repo today through LibreChat itself.

Current scope:

- mailbox and folder access
- thread/message viewing
- attachment visibility and downloads
- read-state updates and delete actions
- AI analysis and drafting actions
- daily brief / meeting-slot proposal / meeting creation flows
- calendar browsing plus event create/update/delete
- audit logging and usage tracking around Outlook actions

The older GCC High Outlook add-in discussion is still relevant only as future context:

- Office add-ins can require broker-style redirects such as `brk-multihub://...`
- GCC High still requires government-cloud endpoints rather than public-cloud defaults
- if an Outlook add-in is built later, the better pattern is a thin shell over this existing LibreChat Outlook backend rather than a separate duplicate implementation

## Things Explicitly Not Done Yet

The following items were discussed but not implemented yet:

- full per-user token limit enforcement using the new `Usage` records
- per-workspace budget enforcement
- full Entra/GCC High RBAC and group claim integration
- richer admin analytics pages beyond current usage/issues view
- issue triage workflow with status changes
- high-fidelity `.docx` round-trip with full formatting preservation
- true `.xlsm` macro-preserving spreadsheet support
- thin Outlook add-in shell, if that product surface is still wanted later
- Teams channel export / channel archive support
- cross-source enterprise-memory linking across Teams, Outlook, documents, Jira, Confluence, and Slack
- semantic/vector retrieval on top of enterprise memory
- worker-driven execution of `DocumentJob` records
- final production deployment architecture in AWS GovCloud

## Current Product Position

At the current point in the project:

- LibreChat has enterprise-specific observability additions
- Outlook workspace is implemented inside the product
- Outlook actions write audit and usage records
- Teams archive, status tracking, and archive-backed retrieval are implemented
- Teams now project into a canonical enterprise-memory layer
- users can report failures inside the platform
- admins can see usage and issue reports
- users have a budget-style progress indicator
- native file workflows now exist for spreadsheets and Word docs
- document-pipeline persistence scaffolding is implemented
- external Jira/Confluence-style integrations are still intended to live in the separate MCP services repo

This is a strong platform foundation, but not the finished EnterpriseGPT end state.

## Recommended Next Areas Of Work

If another model picks this up, the likely next high-value threads are:

1. Finish budget and usage governance
- per-user limits
- per-workspace limits
- clearer admin cost views

2. Harden Teams retrieval and enterprise memory
- validate throttling/backoff behavior under real sync load
- improve retrieval completeness and diagnostics
- continue moving high-value queries off raw archive scans where practical

3. Mature native file workflows and the document pipeline
- richer spreadsheet analytical outputs
- better Word fidelity
- worker execution for `DocumentJob`

4. Wire remaining enterprise integrations into LibreChat
- start with Jira/Confluence end-to-end
- then SharePoint/Slack-style sources

5. Add enterprise auth/RBAC
- Entra group claims
- admin/user role enforcement tied to enterprise identity

## Important Operational Reminder

Many of these changes are local development work in the fork and may not all be committed or deployed everywhere yet.

Any new model or engineer should verify:

- current git status
- current branch
- whether Docker is using the local build override
- whether migrations have run
- whether the target machine has the correct `.env` and `librechat.yaml`

## Quick Pointers

For implementation debugging:

- see [ENTERPRISE_FEATURES_DEBUGGING.md](/Users/praneetkotah/Desktop/Development/LibreChat/ENTERPRISE_FEATURES_DEBUGGING.md:1)

For the main admin dashboard code:

- see [Admin.tsx](/Users/praneetkotah/Desktop/Development/LibreChat/client/src/components/Nav/SettingsTabs/Admin/Admin.tsx:1)

For native spreadsheet support:

- see [api/server/services/Files/Spreadsheets/service.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/services/Files/Spreadsheets/service.js:1)

For native Word support:

- see [api/server/services/Files/WordDocuments/service.js](/Users/praneetkotah/Desktop/Development/LibreChat/api/server/services/Files/WordDocuments/service.js:1)

For request-level usage tracking:

- see [packages/api/src/usage/service.ts](/Users/praneetkotah/Desktop/Development/LibreChat/packages/api/src/usage/service.ts:1)
