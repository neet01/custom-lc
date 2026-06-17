F3# Low-Level App Call Diagram

Last updated: 2026-05-16

## Purpose

This document is a review aid for the current LibreChat fork. It focuses on low-level call paths rather than product features so changes can be traced from UI action to backend route, service layer, external dependency, and persistence layer.

It reflects the current repo state, including custom Outlook, Teams archive, enterprise memory, MCP, and document-processing work.

## Runtime Topology

```mermaid
flowchart LR
    Browser["Browser / React Client"]

    subgraph Docker["Docker Compose Runtime"]
      API["LibreChat API Container\nExpress + Controllers + Services"]
      Mongo["MongoDB\nchat + archive + memory + config data"]
      Meili["Meilisearch\nchat search / app search"]
      Rag["rag_api\nretrieval API"]
      Vector["pgvector / Postgres\nvector storage"]
    end

    subgraph External["External Systems"]
      Bedrock["AWS Bedrock\nmodels / agents"]
      Graph["Microsoft Graph GCC High"]
      MCP["Remote MCP Servers"]
      Storage["Local uploads or S3-compatible storage"]
    end

    Browser -->|HTTP / SSE| API
    API --> Mongo
    API --> Meili
    API --> Rag
    Rag --> Vector
    API --> Bedrock
    API --> Graph
    API --> MCP
    API --> Storage
```

## Server Boot Path

```mermaid
sequenceDiagram
    participant DC as Docker Compose
    participant API as api/server/index.js
    participant Mongo as MongoDB
    participant Config as App Config
    participant MCP as MCP Init
    participant Jobs as Stream / Generation Jobs

    DC->>API: start container
    API->>Mongo: connectDb()
    API->>Mongo: indexSync() in background
    API->>Mongo: seedDatabase()
    API->>Config: getAppConfig(baseOnly)
    API->>API: initializeFileStorage()
    API->>API: performStartupChecks()
    API->>API: updateInterfacePermissions()
    API->>API: mount middleware + routes
    API->>MCP: initializeMCPs()
    API->>API: initializeOAuthReconnectManager()
    API->>API: checkMigrations()
    API->>Jobs: createStreamServices() + GenerationJobManager.initialize()
```

## Main API Surface

```mermaid
flowchart TD
    App["Express App"]

    App --> Auth["/api/auth\nOIDC / JWT / token exchange"]
    App --> Agents["/api/agents\nchat, SSE stream, tools"]
    App --> Files["/api/files\nupload, transform, delete"]
    App --> Outlook["/api/outlook\nmail, calendar, AI actions"]
    App --> Teams["/api/teams-archive\nsync, search, retrieval"]
    App --> MCP["/api/mcp\nMCP server registry / auth"]
    App --> Convos["/api/convos + /api/messages"]
    App --> Admin["/api/admin/*"]
    App --> Search["/api/search"]
    App --> Memories["/api/memories"]
```

## Chat / Agent Generation Path

This is the main call path when a user sends a normal chat or agent message.

```mermaid
sequenceDiagram
    participant UI as React Client
    participant API as /api/agents/*
    participant MW as Middleware
    participant Ctrl as AgentController
    participant Init as initializeClient()
    participant Tools as ToolService
    participant MCP as MCP Service
    participant Model as Bedrock / model endpoint
    participant Jobs as GenerationJobManager
    participant DB as MongoDB

    UI->>API: POST /api/agents or /api/agents/:endpoint
    API->>MW: moderateText + auth + access + convo validation + endpoint build
    API->>Ctrl: AgentController(req,res,next,...)
    Ctrl->>Init: build model client / agent runtime config
    Ctrl->>DB: load conversation + messages + agent config
    Ctrl->>Tools: resolve enabled tools
    Tools->>MCP: hydrate MCP tools if selected
    Ctrl->>Jobs: create generation job / stream state
    Ctrl->>Model: send prompt + context + tool definitions
    alt model calls tool
        Model->>Tools: tool call
        alt MCP tool
            Tools->>MCP: execute tool over MCP transport
        else built-in tool
            Tools->>DB: read/write app data as needed
            Tools->>API: call internal services
        end
        Tools-->>Model: tool result
    end
    Model-->>Jobs: streamed tokens / run steps
    Jobs-->>UI: SSE on /api/agents/chat/stream/:streamId
    Ctrl->>DB: persist messages / artifacts / attachments
```

### Notes

- Streaming is stateful through `GenerationJobManager`, not just raw one-shot HTTP responses.
- Tool resolution is dynamic and includes built-in tools, file tools, web/image tools, action tools, and MCP tools.
- MCP auth and reconnection logic are centralized in the server, not the browser.

## Outlook Call Path

Outlook uses delegated Entra auth with OBO token exchange before every Graph access path that needs a downstream token.

```mermaid
sequenceDiagram
    participant UI as Outlook Workspace UI
    participant Route as /api/outlook/*
    participant Service as OutlookService
    participant OBO as GraphTokenService
    participant Entra as Entra Token Endpoint
    participant Graph as Microsoft Graph
    participant AI as OutlookAIService / Bedrock
    participant DB as MongoDB

    UI->>Route: GET/POST/PATCH /api/outlook/*
    Route->>Service: OutlookService.<operation>()
    Service->>OBO: getGraphApiToken(user, federated access token, scopes)
    OBO->>Entra: OBO token grant
    Entra-->>OBO: Graph access token
    OBO-->>Service: delegated token
    Service->>Graph: mail/calendar/folder/message request
    Graph-->>Service: Graph payload
    alt AI action requested
        Service->>AI: analyze / draft / summarize / daily brief
        AI->>AI: synthesize prompt and context
        AI->>AI: call Bedrock model
        AI-->>Service: structured result + usage
    end
    Service-->>Route: normalized result
    Route->>DB: write Outlook audit + usage/balance records
    Route-->>UI: JSON response
```

### Current implementation details

- OBO is centralized in `GraphTokenService`.
- Outlook route handlers are thin wrappers around `OutlookService`.
- AI actions record usage and audit separately from raw Graph access.
- The read/unread update path now writes to Graph directly rather than waiting for a passive refresh.

## Teams Archive Sync Path

This is the ingestion path that builds the archive and then projects it into enterprise memory.

```mermaid
sequenceDiagram
    participant UI as Teams UI / Tool / API caller
    participant Route as /api/teams-archive/sync
    participant Service as TeamsArchiveService
    participant OBO as GraphTokenService
    participant Entra as Entra Token Endpoint
    participant Graph as Microsoft Graph
    participant DB as MongoDB
    participant Memory as Enterprise Memory Projection

    UI->>Route: POST /api/teams-archive/sync
    Route->>Service: syncUserArchive(user, options)
    Service->>Service: acquire sync slot / lease / job state
    Service->>OBO: getGraphApiToken(...teams scopes...)
    OBO->>Entra: OBO token grant
    Entra-->>OBO: Graph access token
    loop chats + messages
        Service->>Graph: /me/chats
        Graph-->>Service: chat pages
        Service->>Graph: /chats/:id/messages
        Graph-->>Service: message pages
        Service->>DB: upsert TeamsArchiveConversation
        Service->>DB: upsert TeamsArchiveMessage
        Service->>DB: update TeamsArchiveSyncJob heartbeat/progress
    end
    alt enterprise memory projection available
        Service->>Memory: projectTeamsArchiveSyncToMemory()
        Memory->>DB: upsert entities
        Memory->>DB: upsert relationships
        Memory->>DB: bulk upsert enterprise memory chunks
        Memory->>DB: update EnterpriseMemoryJob
    end
    Service-->>Route: sync result
    Route-->>UI: JSON result
```

### What the slow Mongo logs mean

- Slow `update` operations on `enterprisememorychunks` are expected during projection because the current projection path bulk-upserts chunk records after sync.
- Those are write-heavy projection operations, not user search operations.

## Teams Search Split: Archive Regex vs Enterprise Memory

This is the most important current review detail because the product behavior can look similar while the backend path is very different.

```mermaid
flowchart TD
    Start["Teams search request"]

    Start --> Kind{"Which retrieval path?"}

    Kind -->|search_messages| Basic["TeamsArchiveService.searchMessages()"]
    Kind -->|recent_messages| Recent["TeamsArchiveService.recentMessages()"]
    Kind -->|advanced_search_messages| Advanced["TeamsArchiveService.advancedSearchMessages()"]
    Kind -->|summarize / window| Context["Conversation/window retrieval"]

    Basic --> BasicDB["Mongo regex search on teamsarchivemessages\nbodyText/bodyPreview/bodyContent/etc."]
    BasicDB --> BasicOut["Compact message previews"]

    Recent --> RecentMem{"enterprise memory available?"}
    Advanced --> AdvMem{"enterprise memory available?"}

    RecentMem -->|yes| ChunkSearch["searchTeamsMemoryChunks()"]
    RecentMem -->|no| RecentDB["Mongo query on teamsarchivemessages"]

    AdvMem -->|yes| ChunkSearch
    AdvMem -->|no| AdvDB["Mongo query on teamsarchivemessages + convo filters"]

    ChunkSearch --> ChunkDB["Mongo query on enterprisememorychunks\nplus conversation lookups"]
    ChunkDB --> ChunkOut["Enterprise-memory previews"]

    RecentDB --> RecentOut["Recent message previews"]
    AdvDB --> AdvOut["Advanced message previews"]
    Context --> WindowDB["Conversation/message fetch by chatId"]
    WindowDB --> ContextOut["Summary or message window"]
```

### Current status from logs

- The logs showing `COLLSCAN` on `teamsarchivemessages` with regex filters are evidence of the archive-message fallback/search path still being exercised.
- If the user experience improved without a clearly visible explicit advanced-search tool call, the likely reason is one of:
  - the agent is choosing a better Teams action mix even without naming it clearly
  - some requests are using enterprise-memory retrieval while others are still using archive regex search
  - summarized/conversation-bounded retrieval is improving answer quality after the initial search

## File Upload and Transform Path

This is the path behind document/spreadsheet processing and generated file attachments.

```mermaid
sequenceDiagram
    participant UI as Chat UI / File UI
    participant Route as /api/files
    participant Proc as Files/process.js
    participant Store as File strategy\nlocal / S3 / provider storage
    participant Reg as Document registration / file DB
    participant Xform as Spreadsheet/Word service
    participant Worker as Optional Python spreadsheet worker
    participant JS as JS transform fallback
    participant DB as MongoDB

    UI->>Route: upload or transform request
    alt upload
        Route->>Proc: processFileUpload()
        Proc->>Store: save original file
        Proc->>DB: create file record
        Proc->>Reg: register document pipeline metadata when applicable
        Proc-->>UI: attachment metadata
    else transform
        Route->>Xform: transformSpreadsheetFile() / transformWordDocumentFile()
        Xform->>Store: download source file buffer
        alt Python worker enabled and supported
            Xform->>Worker: inspect/transform workbook
            Worker-->>Xform: generated buffer
        else JS path
            Xform->>JS: inspectSpreadsheetBuffer() / transformSpreadsheetBuffer()
            JS-->>Xform: generated buffer
        end
        Xform->>Store: save generated output
        Xform->>DB: create generated file record
        Xform-->>UI: output file attachment metadata
    end
```

### Current implementation details

- The spreadsheet service chooses between Python worker and JS fallback per file.
- Generated files are now re-registered into the live tool/file context so chained transforms can target the newly created file in the same run.
- Agent callback handling now defers intermediate file artifacts so only the final transform output is surfaced back to the user during a chained run.

## MCP Tool Path

```mermaid
sequenceDiagram
    participant UI as Chat UI
    participant Agent as Agent runtime
    participant ToolSvc as ToolService
    participant MCP as MCP.js
    participant OAuth as MCP OAuth/token store
    participant Remote as Remote MCP Server
    participant DB as MongoDB

    UI->>Agent: ask question that can use MCP
    Agent->>ToolSvc: request tool definitions
    ToolSvc->>MCP: resolve server configs + available tools
    MCP->>DB: read stored MCP OAuth tokens / config
    alt token missing or stale
        MCP->>OAuth: start or refresh OAuth flow
        OAuth-->>UI: auth URL / reconnect event
    end
    Agent->>ToolSvc: tool call
    ToolSvc->>MCP: execute MCP tool
    MCP->>Remote: invoke remote tool over configured transport
    Remote-->>MCP: tool result
    MCP-->>ToolSvc: normalized tool output
    ToolSvc-->>Agent: tool result
```

## Persistence Map

```mermaid
flowchart LR
    Chat["Chat messages / convos"] --> Mongo["MongoDB"]
    OutlookAudit["Outlook audit + usage/balance"] --> Mongo
    TeamsArchive["TeamsArchiveConversation / TeamsArchiveMessage / SyncJob"] --> Mongo
    EnterpriseMemory["EnterpriseMemoryEntity / Relationship / Chunk / Job"] --> Mongo
    Files["File records"] --> Mongo
    SearchIdx["Chat search index"] --> Meili["Meilisearch"]
    RagReq["RAG API"] --> Vector["pgvector / Postgres"]
    UploadBytes["Original / generated files"] --> Blob["Local bind mount or S3"]
```

## Review Hotspots

If you need to audit changes quickly, these are the best places to start:

1. `api/server/index.js`
   This is the real server boot graph and route mount map.

2. `api/server/routes/agents/*` and `api/server/services/ToolService.js`
   This is the core chat/tool orchestration path.

3. `api/server/services/GraphTokenService.js`
   This is the shared OBO choke point for Outlook, Teams, and SharePoint-style delegated Graph access.

4. `api/server/services/OutlookService.js` and `api/server/routes/outlook.js`
   This is the main enterprise workspace integration path.

5. `api/server/services/TeamsArchiveService.js`
   This is the archive ingestion, basic search, advanced search fallback, and summary path.

6. `api/server/services/EnterpriseMemory/*`
   This is the new cross-source retrieval layer and the source of `enterprisememorychunks` write load.

7. `api/server/services/Files/Spreadsheets/*`, `api/server/services/Files/WordDocuments/*`, and `api/server/routes/files/files.js`
   This is the document transformation platform and generated-file return path.

## Current Caveats

- Teams basic search still uses regex scans against `teamsarchivemessages` and can show `COLLSCAN` in Mongo logs.
- Enterprise memory improves retrieval quality, but it also adds significant write load during sync because chunks are bulk-upserted.
- Outlook, Teams, and SharePoint-style picker flows all rely on the same OBO token-exchange pattern; auth regressions in that area can affect multiple product surfaces.
- The document platform now has multiple execution paths: upload pipeline, JS transform path, optional Python spreadsheet worker path, and agent callback attachment handling. That is the correct direction, but it means file-workflow regressions need end-to-end testing rather than route-only testing.

