# Cortex Document Intelligence System

Last updated: 2026-05-17

## Purpose

This document captures the plan for turning Cortex from a chat application with file uploads into a central enterprise intelligence layer.

The immediate trigger was Bedrock Converse file upload limits. That issue exposed a broader architectural constraint: provider-native file handling is not a stable foundation for enterprise-scale document reasoning.

The correct long-term design is a Cortex-owned document pipeline.

## Strategic Context

Document intelligence is not an isolated feature area.

It exists inside the broader Cortex objective: building an enterprise memory and operational intelligence platform that can reason across fragmented organizational systems over time.

That means uploaded files should be treated as one source feed into a larger canonical memory layer alongside:

- Teams
- Outlook
- SharePoint
- Jira
- Confluence
- GitLab
- operational telemetry
- and future enterprise systems

The long-term value is not just semantic file retrieval. It is the ability to connect documents to:

- people
- teams
- systems
- projects
- incidents
- decisions
- operational events
- ownership relationships
- and temporal change

This has two design consequences:

1. document ingestion must preserve provenance and source truth
2. document retrieval must evolve toward canonical entity/event alignment, not remain a standalone vector silo

In other words, file upload and RAG work should be implemented as feeds into organizational memory, not as a one-off document chatbot subsystem.

## Near-Term Execution Order

Document intelligence work should follow the execution order below so the platform does not outrun its operational reliability.

### 1. Stabilize runtime compaction and prompt-budget behavior

Before more document context is introduced, Cortex needs reliable context compaction and overflow recovery.

Why:

- broader retrieval is counterproductive if turns fail before tools run

### 2. Finish retrieval correctness on existing sources

Current retrieval paths must support:

- full-body recovery for truncated source content
- deterministic fallback when memory-layer retrieval is empty
- explicit completeness/truncation signaling

Why:

- users need trustworthy evidence retrieval before broader memory abstractions are added

### 3. Build a real test/eval loop

Why:

- document and memory features cannot be validated by ad hoc prompting alone

Required:

- service/API tests
- tool-call telemetry
- fixed retrieval eval sets

### 4. Productize reusable uploaded-file access

Why:

- persisted uploads already exist, but users need a first-class way to reference them again in later chats

### 5. Complete the document pipeline

Why:

- user-scoped RAG should sit on top of durable ingestion, versioning, jobs, and chunk provenance

### 6. Add user-scoped uploaded-document retrieval

Why:

- this is the correct first durable document-memory feature, and it should be implemented before broader cross-source graph ambitions

## Why This Needs To Exist

Provider-native uploads are useful, but they are not a reliable system of record.

They fail against enterprise requirements:

- provider size limits block legitimate engineering and finance documents
- rich files lose too much fidelity when reduced to one flat text blob
- retrieval quality degrades as file sizes scale
- provider APIs do not give Cortex enough control over chunking, lineage, and citations
- different providers impose different document constraints, which leads to inconsistent user behavior

If Cortex is going to be the enterprise intelligence layer, it needs to own:

- document ingestion
- document metadata
- extraction strategy
- chunking
- retrieval
- evidence lineage

The model should reason over Cortex-managed evidence, not raw enterprise files whenever scale matters.

## Design Principles

### 1. Keep original files as the source of truth

Every uploaded file should have a durable original artifact.

### 2. Separate storage from reasoning

The LLM should not be the document system. Cortex should manage storage, structure, and retrieval, then supply the relevant slices to the model.

### 3. Prefer structure over flat text

For PDFs, spreadsheets, Word docs, and engineering documentation, preserving structure matters:

- page boundaries
- headings
- tables
- ranges
- sheet names
- captions
- adjacency

### 4. Retrieval before generation

Large files should be answered through retrieval and synthesis, not by pushing the whole artifact into a single model call.

### 5. Evidence must be auditable

Answers should be traceable to:

- document version
- chunk(s)
- pages or sheets
- extraction strategy

## Target Architecture

### Durable storage

- **S3** for original files and derived artifacts
- **Mongo** for document metadata, jobs, versions, chunks, and lineage
- **OpenSearch** later for full-text and hybrid retrieval

### Processing layer

LibreChat’s containerized services should remain the orchestration and extraction layer:

- upload intake
- extraction workers
- chunking workers
- retrieval API
- spreadsheet and document-specific transforms

The containers are the processing layer, not the long-term corpus store.

## Cross-Source Memory Layer

Document intelligence is only one part of the broader enterprise system.

Cortex needs a canonical layer that can align:

- documents
- Teams chats
- Outlook mail and calendar data
- SharePoint content
- Jira and Confluence records
- future GitLab and Slack sources

The first canonical memory primitives are:

- `EnterpriseMemoryEntity`
- `EnterpriseMemoryRelationship`
- `EnterpriseMemoryChunk`
- `EnterpriseMemoryJob`

Purpose of each:

- `EnterpriseMemoryEntity`
  - canonical people, conversations, projects, documents, issues, topics
- `EnterpriseMemoryRelationship`
  - graph edges between canonical entities
- `EnterpriseMemoryChunk`
  - source-backed retrieval units with provenance and visibility metadata
- `EnterpriseMemoryJob`
  - projection/enrichment/indexing job state

This layer is not a replacement for source-native archives.

The correct pattern is:

1. preserve source-native records
2. project them into canonical memory records
3. retrieve across memory chunks
4. later enrich and link across sources

This is the bridge between today's upload/RAG work and the longer-term operational intelligence layer.

The target state is not:

- files in one silo
- chats in another silo
- tickets in another silo

The target state is a shared canonical memory model where those records can be correlated for:

- evidence-backed reasoning
- timeline analysis
- ownership inference
- cross-system relationship discovery
- and operational insight generation

## Core Entities

### Document

Represents the canonical Cortex record for an uploaded artifact.

Fields:

- user
- source file id
- filename
- mime type
- bytes
- source
- context
- pipeline status
- latest version id
- current job id

### Document Version

Represents one extracted/reasoned-over version of a document.

Fields:

- document id
- source file id
- version number
- source filepath
- extraction kind
- text length
- chunk count
- status

### Document Job

Represents queued or running work against a document version.

Fields:

- document id
- document version id
- user
- job type
- status
- attempts
- started/completed timestamps
- error state

### Future: Document Chunk

This is intentionally deferred until the next phase. The long-term chunk model should preserve order and adjacency.

Expected fields:

- document version id
- order index
- prev chunk id
- next chunk id
- page range
- sheet name
- section path
- token estimate
- content

## Phased Rollout

## Phase 0: Reliability Stopgap

Status: implemented and validated

Goal:

- prevent Bedrock file-size failures from breaking chats

What was done:

- oversized Bedrock-compatible uploads now fall back to text extraction during upload rather than failing during Converse request assembly

What was validated:

- fresh uploads in Bedrock agent chats now hit the fallback path
- the upload logs show `resolvedEndpoint=bedrock`
- the upload logs show the oversize fallback warning for large PDFs

Known caveat:

- conversations that already contain older provider-bound document attachments can still replay those stale raw attachments on later turns and trigger the old 4.5 MB Bedrock error
- this is a conversation-history problem, not a failure of the new upload fallback itself

Limitation:

- this is a safety net, not the target architecture

## Phase 1: Canonical Document Registration

Status: implemented in repo, deployment/runtime validation pending

Goal:

- create durable Cortex-owned document records at upload time

Scope:

- register uploaded document-like files as `Document` records
- create first `DocumentVersion`
- create initial `DocumentJob`
- link back to the existing file upload record by `sourceFileId`

Reasoning:

- this introduces the document pipeline without changing the chat upload contract
- uploads remain stable
- Cortex starts building document lineage immediately
- future extraction and retrieval workers get a durable queueable substrate

Non-goals:

- no retrieval yet
- no chunk persistence yet
- no OpenSearch dependency yet
- no UI changes yet

Implemented artifacts:

- `Document` schema/model/methods
- `DocumentVersion` schema/model/methods
- `DocumentJob` schema/model/methods
- upload-time registration service
- upload pipeline hook that registers indexable files into the document pipeline

Current behavior:

- image/audio/video uploads are ignored by the document pipeline
- provider-bound binary documents start with:
  - `extractionKind = none`
  - initial job type `extract`
- text-backed files created by the Bedrock oversize fallback start with:
  - `extractionKind = text`
  - initial job type `chunk`

## Phase 2: Enterprise Retrieval Over Memory Chunks

Status: started for Teams

Goal:

- move retrieval away from raw source transcript dumps and toward Cortex-owned chunk retrieval

Current implementation:

- Teams archive sync already projects messages into `EnterpriseMemoryChunk`
- new retrieval service:
  - `api/server/services/EnterpriseMemory/retrieval.js`
- Teams tool actions `advanced_search_messages` and `recent_messages` now attempt enterprise-memory chunk retrieval first
- if Phase 2 retrieval is unavailable or errors, the system falls back to source-archive retrieval

What this solves:

- starts separating archival storage from retrieval behavior
- reduces dependence on replaying full raw Teams threads into model context
- creates a stable path for later chunk summarization, neighbor-window retrieval, and source fusion

What it does not solve yet:

- no semantic/vector search yet
- no OpenSearch dependency yet
- no neighbor-window expansion yet
- no cross-source ranking yet

### Teams retrieval Phase 3: bounded context and summaries

Status: started

Goal:

- stop loading full raw Teams threads into model context for normal research questions

Current implementation:

- `teams_archive_search` now exposes:
  - `summarize_conversation`
  - `get_messages_window`
- `summarize_conversation` returns a bounded summary for one archived chat:
  - participant list
  - date range
  - total/matched message counts
  - top senders
  - recent highlights
- `get_messages_window` returns a small local slice around:
  - a specific message id, or
  - the most recent query match inside a specific chat

Why this matters:

- the current failure mode for raw tool retrieval is prompt bloat
- Monday rollout also needs operational guardrails, not just better retrieval

### Teams sync operational hardening

Status: implemented

Current implementation:

- per-user Teams sync acquisition now uses an atomic Mongo-backed lease
- global active Teams syncs now use a slot-based concurrency cap
- stale-job reconciliation still exists and now works alongside the lease model
- sync status exposes:
  - `activeSyncs`
  - `maxConcurrentSyncs`

Operator envs:

- `TEAMS_ARCHIVE_SYNC_STALE_MINUTES`
- `TEAMS_ARCHIVE_MAX_CONCURRENT_SYNCS`

Why this matters:

- prevents duplicate sync races for the same user across API replicas
- bounds how many concurrent backfills can compete with normal app traffic
- provides a safer bridge until the sync engine is moved to a dedicated worker/ECS architecture
- bounded windows and summaries let Cortex answer high-level questions without replaying entire threads
- this is the bridge between source archive search and later hybrid/semantic enterprise retrieval

Infrastructure impact:

- no new infrastructure is required for this first Phase 2 slice
- Mongo-backed enterprise memory collections are enough for the initial rollout
- OpenSearch becomes relevant in later phases when hybrid retrieval and scale requirements justify it

## Phase 3: Structured Extraction

Goal:

- extract structured representations for supported document types

Target support:

- PDF
- DOCX
- XLSX
- CSV
- HTML / Markdown / plain text

Output should preserve structure where possible:

- pages
- sections
- headings
- sheets
- tables
- formulas

Reasoning:

- flat text is too lossy for enterprise engineering and finance docs

## Phase 4: Chunk Graph

Goal:

- persist retrieval-ready document chunks with order and adjacency

Requirements:

- structural chunking, not naive character splitting
- overlap between adjacent chunks
- page/sheet/section metadata
- chunk lineage to version/source document

Reasoning:

- this is what allows large documents to be queried without handing the whole file to a model

## Phase 5: Retrieval Layer

Goal:

- answer document questions through retrieval instead of raw file injection

Modes:

- exact keyword
- metadata-filtered lookup
- semantic retrieval later
- adjacency expansion

Routing logic should decide between:

- native provider upload
- Cortex-managed chunk retrieval
- spreadsheet-specific analysis flow

## Phase 6: Hierarchical Synthesis

Goal:

- support summaries and answers over large corpora

Patterns:

- chunk-level summarization
- section-level summarization
- final synthesis over retrieved evidence

Reasoning:

- enterprise documents will exceed single-context assumptions regularly

## Phase 6: Search Infrastructure

Goal:

- add a proper search backend once chunking/retrieval justify it

Recommended direction:

- OpenSearch for full-text + later hybrid retrieval

Reasoning:

- Mongo is fine for metadata and lineage
- Mongo is not the right long-term primary retrieval engine for enterprise-scale corpora

## Phase 7: Operations And Governance

Goal:

- make the system observable and auditable

Required admin surfaces:

- ingestion success/failure counts
- extraction job durations
- chunk counts per document
- retrieval hit rates
- oversized-provider fallback frequency
- top file types
- top document-heavy users or workflows

## Enterprise Memory Phase 1

Status: implemented in repo for Teams as the first source adapter

Goal:

- establish a canonical cross-source memory substrate without waiting for full RAG/indexing infrastructure

What is implemented:

- canonical Mongo persistence for:
  - `EnterpriseMemoryEntity`
  - `EnterpriseMemoryRelationship`
  - `EnterpriseMemoryChunk`
  - `EnterpriseMemoryJob`
- Teams archive projection service that runs after successful sync

Current Teams projection behavior:

- each Teams chat becomes a `conversation` entity
- participants/senders/mentions become `person` entities
- conversation membership is stored as `has_participant` relationships
- each Teams message becomes a `message` chunk with:
  - source provenance
  - parent conversation reference
  - linked entity ids
  - normalized text body

Why this order is correct:

- it keeps source-native Teams data intact
- it creates a reusable retrieval substrate now
- it avoids coupling early enterprise memory work to OpenSearch before the canonical data model exists

Current limitations:

- user-scoped visibility only
- no cross-source entity linking yet
- no tenant-wide memory governance model yet
- no OpenSearch/hybrid retrieval over enterprise memory chunks yet
- no SharePoint/Jira/Confluence/GitLab/Slack projection adapters yet

## Infrastructure Position

### What should stay in LibreChat services

- upload handling
- extraction orchestration
- spreadsheet/doc processing
- document registration
- retrieval API

### What should not live inside containers as durable storage

- the long-term RAG corpus
- original enterprise documents
- durable retrieval indexes

### Preferred storage split

- S3: originals and derived artifacts
- Mongo: metadata, jobs, versions, chunks, lineage
- OpenSearch: retrieval indexes when Phase 6 is reached

## Phase 1 Implementation Decision

Phase 1 is deliberately conservative.

It is implemented as an additive layer on top of the existing `File` model:

- uploads still create `File` records exactly as before
- document-like uploads are additionally registered as `Document` records
- the initial version and initial pending job are created automatically

This is the right first step because it:

- avoids destabilizing the current upload and chat workflow
- gives Cortex durable document lineage immediately
- creates a clean handoff point for later extraction and chunking workers

## Immediate Next Steps After Phase 1

1. Add a worker/service that consumes pending `DocumentJob` records.
2. Persist structured extraction output for PDFs, Word docs, and spreadsheets.
3. Introduce `DocumentChunk` persistence.
4. Route oversized Bedrock documents to chunk retrieval instead of plain-text fallback.
5. Add citations and admin visibility.
