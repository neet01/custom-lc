# LibreChat Enterprise Project Context Handoff

This document is a high-signal handoff for another LLM or engineer.

It is meant to answer:

- what this project is trying to become
- what constraints it operates under
- what has already been implemented
- what related repositories exist
- what design decisions have already been made
- what is still pending

This is broader than [ENTERPRISE_FEATURES_DEBUGGING.md](/Users/praneetkotah/Desktop/Development/LibreChat/ENTERPRISE_FEATURES_DEBUGGING.md:1), which is focused more on how to debug the custom features in this LibreChat fork.

## Project Goal

This LibreChat fork is being turned into an internal enterprise AI platform, effectively an `EnterpriseGPT`, for a regulated environment.

Long-term goals include:

- enterprise chat UI built on LibreChat
- usage observability and admin oversight
- budget awareness and anti-runaway-spend controls
- native file transformation workflows
- integrations for Jira, Confluence, Outlook, and eventually other Office 365 / enterprise tools
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

### 5. Native Spreadsheet File Workflow

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

### 6. Native Word Document Workflow

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

### 7. Enterprise Debugging Guide

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

## Outlook Planning Context

This was discussed as planning only and has not been implemented yet.

The user had tried building a custom Outlook add-in in GCC High and ran into `brk://` / broker redirect issues.

Planning conclusion:

- integrate Outlook with LibreChat first
- treat the custom Outlook add-in as a later phase if needed

Reasoning:

- LibreChat already provides admin controls, issue reporting, usage tracking, and future tool orchestration
- building Outlook features in LibreChat first is lower-risk than immediately debugging Office host behavior plus GCC High auth plus deployment issues

Planning notes already established:

- Outlook add-ins using modern Office auth patterns can require broker-style redirects such as `brk-multihub://...`
- GCC High requires government cloud endpoints, not public cloud defaults
- likely future Outlook capabilities should include:
  - search mail
  - summarize threads
  - draft replies
  - create calendar events
  - possibly action flows gated by approval or policy

Recommended future direction:

- make Outlook another governed enterprise integration surfaced through LibreChat
- later, if needed, build an Outlook add-in as a thin shell that launches or complements LibreChat

## Things Explicitly Not Done Yet

The following items were discussed but not implemented yet:

- full per-user token limit enforcement using the new `Usage` records
- per-workspace budget enforcement
- full Entra/GCC High RBAC and group claim integration
- richer admin analytics pages beyond current usage/issues view
- issue triage workflow with status changes
- high-fidelity `.docx` round-trip with full formatting preservation
- true `.xlsm` macro-preserving spreadsheet support
- full Outlook integration
- final production deployment architecture in AWS GovCloud

## Current Product Position

At the current point in the project:

- LibreChat has enterprise-specific observability additions
- users can report failures inside the platform
- admins can see usage and issue reports
- users have a budget-style progress indicator
- native file workflows now exist for spreadsheets and Word docs
- external integrations are planned with a separate MCP services repo

This is a strong platform foundation, but not the finished EnterpriseGPT end state.

## Recommended Next Areas Of Work

If another model picks this up, the likely next high-value threads are:

1. Finish budget and usage governance
- per-user limits
- per-workspace limits
- clearer admin cost views

2. Mature native file workflows
- richer spreadsheet analytical outputs
- better Word fidelity

3. Wire enterprise integrations into LibreChat
- start with Jira/Confluence end-to-end
- then Outlook

4. Add enterprise auth/RBAC
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

