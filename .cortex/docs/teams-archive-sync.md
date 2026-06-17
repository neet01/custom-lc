# Teams Archive Sync — Architecture & Known Issues

## Overview

`api/server/services/TeamsArchiveService.js` — The main sync service (201KB). Syncs Microsoft Teams
chat history into MongoDB via Microsoft Graph API (GCC High endpoint). All Graph calls go through
the delegated OBO token flow in `GraphTokenService.js`.

## Sync Flow

```
syncUserArchive(userId, options)
  1. Acquire per-user lease (user:{userId})
  2. Acquire global slot lease (slot:0 ... slot:N) — default max 1 concurrent sync
  3. Phase: discovering_chats
     - Paginate /me/chats (50 chats/page)
     - For each page: mapWithConcurrency(chats, discoveryConcurrency, listChatMembers)
     - Heartbeat + refreshBackfillStateSnapshot after every page
  4. Phase: syncing_messages
     - For each discovered conversation: fetch messages page-by-page (50/page)
     - sleep(messagePageDelayMs) between pages to avoid Graph bursts
     - Heartbeat every 5 conversations or 30 seconds
  5. Completion: update SyncJob → 'success', release leases, queue memory projection
  6. Finally: release leases, clear module-level Maps
```

## Key Config Constants

| Constant | Default | Env Override |
|----------|---------|--------------|
| `DEFAULT_MAX_CONCURRENT_SYNCS` | 1 | `TEAMS_ARCHIVE_MAX_CONCURRENT_SYNCS` |
| `DEFAULT_DISCOVERY_CONCURRENCY` | **2** | `TEAMS_ARCHIVE_DISCOVERY_CONCURRENCY` |
| `DEFAULT_MESSAGES_PER_CHAT` | 250 | `TEAMS_ARCHIVE_MAX_MESSAGES_PER_CHAT` |
| `DEFAULT_MESSAGE_PAGE_DELAY_MS` | **300** | `TEAMS_ARCHIVE_MESSAGE_PAGE_DELAY_MS` |
| `DEFAULT_GRAPH_RETRY_ATTEMPTS` | 5 | `TEAMS_ARCHIVE_GRAPH_RETRY_ATTEMPTS` |
| `DEFAULT_GRAPH_RETRY_BASE_MS` | 1000 | `TEAMS_ARCHIVE_GRAPH_RETRY_BASE_MS` |
| `DEFAULT_GRAPH_RETRY_MAX_MS` | 60000 | `TEAMS_ARCHIVE_GRAPH_RETRY_MAX_MS` |
| `ENSURE_ACTIVE_CHECK_INTERVAL_MS` | **10000** | — |
| `BACKFILL_SNAPSHOT_THROTTLE_MS` | **60000** | — |
| `HEARTBEAT_MIN_INTERVAL_MS` | 30000 | — |
| `HEARTBEAT_CHAT_INTERVAL` | 5 | — |

**Bold** = changed from original; these are the key tuning levers.

## Graph Retry Logic

`graphRequest()` at line ~281. Retries on status 429, 503, 504 only. Reads `Retry-After` header;
falls back to exponential backoff with jitter. Does NOT retry 500 or 502 (these are skipped at
the conversation level via `isRecoverableChatMessageError`).

## Recoverable vs Fatal Errors

`isRecoverableChatMessageError()` determines whether a Graph error skips the current conversation
or aborts the entire sync. Currently recoverable: **403, 404, 500, 502**.

- 403/404: access denied / chat not found — expected, skip
- 500/502: transient Graph server errors — skip conversation, retry on next sync
- 503/504: fully-exhausted retry → these escape `graphRequest` already; also recoverable now

## Rate Limiting Applied (as of 2026-06-16)

1. **Discovery concurrency 4 → 2**: halves concurrent member-lookup chains per chat page.
   This is the primary Graph burst reduction.
2. **Message page delay 300ms**: adds a 300ms pause between consecutive message-page fetches
   within a conversation. The message loop is already fully serial (one conversation at a time,
   one page at a time), so this is a rate governor, not a parallelism fix. Trades latency for
   smoother request shaping during rollout.
3. **500/502 retried at `isRetryableGraphStatus`**: transient Graph server errors now get the
   same exponential backoff as 429/503/504, across all Graph call sites. If all retries exhaust,
   the error still propagates normally.
4. **`ensureSyncJobActive` time-gated**: DB check at most once per 10s per sync job.
5. **`refreshBackfillStateSnapshot` throttled**: within a 60s window, skips the 6 parallel
   `countDocuments` queries and persists only the `updates` fields (e.g. `nextChatPageLink`,
   `discoveryComplete`). Existing count values in MongoDB are preserved by the `$set` semantics
   of the upsert. `force: true` at sync start/end bypasses throttle for accurate counts there.

## Module-Level State

Two Maps are used for throttling — both cleared in the `finally` block of `syncUserArchive`:

```js
const ensureActiveLastChecked = new Map();       // syncJobId → lastCheckedTimestamp
const backfillSnapshotLastRefreshed = new Map(); // userId → lastRefreshedTimestamp
```

These are process-scoped. In a multi-process/cluster setup they don't coordinate across workers,
but since `maxConcurrentSyncs` is MongoDB-lease-backed (not in-memory), this is safe.

## MongoDB Collections

| Collection | Purpose |
|------------|---------|
| `TeamsArchiveConversation` | One doc per Teams chat, stores `syncCursor`, `syncStatus` |
| `TeamsArchiveMessage` | Individual messages |
| `TeamsArchiveSyncJob` | One doc per sync run, tracks phase/checkpoint/stats |
| `TeamsArchiveSyncLease` | Atomic lease records (user and slot leases) |
| `TeamsArchiveBackfillState` | Per-user backfill progress snapshot (UI status) |

## Known Remaining Gaps

- No per-conversation message cursor checkpointing in `SyncJob`: if process dies mid-conversation,
  that conversation restarts from its last `syncCursor` page, but there's no job-level record of
  which conversation was in-progress.
- Enterprise memory projection (`EnterpriseMemory/teamsProjection.js`) runs async after sync
  completion; projection failures are logged but don't affect sync job status.
- Teams archive covers **chats only**, not channel exports.
- Enterprise memory retrieval is lexical (not semantic/vector).
