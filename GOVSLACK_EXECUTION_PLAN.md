# Slack <> Cortex Execution Plan

Last updated: 2026-05-21

## Purpose

This document converts the current Slack planning notes into an execution plan for Cortex.

The plan assumes three phases:

1. Slack archive and context
2. Shared Cortex sessions
3. Cortex in GovSlack

## Important Note

Polishing the current Teams functionality first is strongly recommended.

Reason:

- the Slack archive and retrieval model should reuse the same source-native archive, enterprise memory, and retrieval patterns already being built for Teams
- if Teams ingestion and retrieval remain incomplete, Slack implementation time will increase because the same archive/indexing issues will reappear in a second source
- poor ingestion quality will lead to sparse or inaccurate retrieval, which will weaken Slack answer quality later

This is not a separate phase in the plan below, but it is a practical prerequisite.

## Overall Timeline

- Phase 1: `5-8 weeks`
- Phase 2: `6-10 weeks`
- Phase 3: `4-7 weeks`

Sequential total:

- `15-25 weeks`

Expected timeline slip is normal because:

- Slack/GovSlack access and app setup can create external dependency delays
- shared sessions require core runtime refactoring
- archive correctness and access control will require iteration

## Phase 1: Slack Archive and Cortex Context

### Objective

Make Slack a read-only context source for Cortex.

This phase should allow Cortex to:

- ingest historical Slack data
- index channels, private channels, and DMs
- retrieve relevant Slack context on demand
- preserve that context in enterprise memory

### Estimated Timeline

- `5-8 weeks`

### Implementation Steps

1. Stabilize the Teams indexing and retrieval pattern.
   Fix current gaps in message chunking, retrieval completeness, and source fidelity so the same archive-to-memory model can be reused for Slack.

2. Improve entity and relationship modeling.
   Expand the current enterprise memory projection so message history, people, and conversation relationships are captured more reliably.

3. Validate retrieval quality.
   Ensure RAG and enterprise-memory retrieval return correct and sufficiently complete results before Slack is layered on top.

4. Prepare infrastructure for indexing at scale.
   Validate storage, indexing, sync throughput, and observability so the platform can support more users and another collaboration source.

5. Set up GovSlack connectivity.
   Configure a GovSlack-compatible application and the required credentials/endpoints for archive ingestion and future runtime support.

6. Build the Slack data pipeline.
   Add source-native Slack archive ingestion, normalization, and sync handling using the same broad pattern as the Teams archive service.

7. Add Slack observability and sync controls.
   Track indexing health, sync progress, failures, and resumable sync behavior.

8. Test with a pilot subset of users.
   Validate archive completeness, load behavior, and enterprise-memory projection on a limited Slack scope before broader rollout.

### Longest Parts

The work most likely to take the longest in Phase 1 is:

- fixing ingestion completeness and retrieval correctness in the existing Teams-derived pattern
- building a permission-bounded Slack archive model
- validating that sync and retrieval quality hold up at pilot scale

### Critical Path

`teams retrieval stability -> enterprise memory quality -> Slack connectivity -> Slack archive pipeline -> retrieval validation -> pilot rollout`

## Phase 2: Shared Cortex Sessions

### Objective

Allow multiple internal users to collaborate in the same Cortex session.

This phase should make Cortex collaborative rather than single-user only.

### Estimated Timeline

- `6-10 weeks`

### Implementation Steps

1. Define the shared-session model.
   Decide how sessions relate to existing conversations, files, citations, tools, prompts, and retained history.

2. Build the shared-session persistence layer.
   Add new shared-session data models for ownership, membership, invite state, permissions, and audit metadata.

3. Build session management services and APIs.
   Add backend support for creating sessions, inviting users, joining, revoking access, and managing session-level state.

4. Refactor the Cortex chat runtime.
   Change the runtime from single-user conversation ownership to session-scoped multi-user access.

5. Add collaboration UI.
   Let users create shared sessions, invite participants, and view membership/session state in the UI.

6. Validate multi-user behavior.
   Test session history, file visibility, citations, tool calls, and user attribution across participants.

### Longest Parts

The work most likely to take the longest in Phase 2 is:

- refactoring the core chat runtime from owner-centric to session-centric behavior
- access-control enforcement across shared sessions
- preserving correct attribution and visibility for files, citations, and tool outputs

### Critical Path

`shared-session model -> persistence -> backend access control -> runtime refactor -> collaboration UI -> multi-user validation`

## Phase 3: Cortex in GovSlack

### Objective

Allow users to invoke Cortex directly from GovSlack.

This phase should allow:

- `@Cortex` mentions
- thread invocation
- direct-message usage
- private-channel usage within approved boundaries
- preservation of the Slack-originated interaction inside Cortex history

### Estimated Timeline

- `4-7 weeks`

### Implementation Steps

1. Confirm the GovSlack app/runtime model.
   Finalize support for OAuth/app auth, mentions, threads, DMs, private channels, and channel boundary behavior.

2. Build the GovSlack event runtime.
   Decide between bot/event mode and Socket Mode, then implement the service that receives `@Cortex` invocations and hands them off to Cortex.

3. Map Slack interactions to Cortex sessions.
   Store the Slack channel/thread context and connect invocations to the corresponding Cortex session and history model.

4. Route GovSlack invocations through the Cortex runtime.
   Send the Slack request into the normal Cortex agent/LLM path and return the response back to GovSlack.

5. Connect group-chat invocations to shared Cortex sessions.
   Use the shared-session model from Phase 2 so Slack group interactions can map to persistent collaborative Cortex work.

6. Validate end-to-end behavior.
   Pilot with a limited group and test channels, threads, DMs, retries, response correctness, and failure handling.

### Longest Parts

The work most likely to take the longest in Phase 3 is:

- GovSlack runtime setup and event-handling decisions
- mapping Slack threads and users cleanly into Cortex session/history behavior
- operational hardening around retries, failed requests, and channel-safe responses

### Critical Path

`GovSlack app setup -> event runtime -> Slack-to-Cortex mapping -> response loop -> shared-session linkage -> pilot validation`

## Summary

The expected delivery sequence is:

1. Stabilize the archive/retrieval pattern through Teams and Phase 1 Slack archive work
2. Build collaborative shared sessions in Cortex
3. Expose Cortex directly inside GovSlack once the underlying memory and session model is stable

The likely longest phase by engineering effort is Phase 2.

The likely highest uncertainty from external dependency is Phase 3.

The largest quality risk across the whole effort is Phase 1, because Slack answer quality will depend on archive completeness, indexing quality, and retrieval correctness.

