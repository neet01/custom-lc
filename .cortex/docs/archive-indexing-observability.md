# Archive Indexing Observability

## Purpose

The archive indexing diagnostics view answers the operational question: "Why isn't this Slack channel or Teams chat appearing after a sync?"

It separates the archive pipeline into three visible stages:

1. Discovery: a Slack conversation or Teams chat exists in the archive conversation collection.
2. History capture: messages were stored under that conversation.
3. Memory projection: chunkable messages were projected into `EnterpriseMemoryChunk` records that retrieval can search.

The dashboard is metadata-only. It does not display raw Slack or Teams message text.

## Admin UI

Location: Admin reporting -> Archive indexing.

The tab supports:

- Source filter: Slack or Teams.
- Channel/chat search: matches display metadata such as channel name, topic, purpose, or source ID.
- User filter: optional LibreChat user ID for per-user archive debugging.
- Type filter: Slack conversation types or Teams chat types.
- Sync status filter: `complete`, `failed`, `pending`, `running`, or `deferred_failed`.

The table shows per-conversation health:

- `Healthy`: archived messages are projected into memory chunks.
- `Sync Failed`: the archive sync recorded a failed or deferred-failed state.
- `Pending` / `Running`: the conversation is still being processed.
- `No Messages`: the conversation was discovered, but no archive messages were stored.
- `No Chunkable Messages`: messages exist, but normalization marked all of them unchunkable.
- `Not Projected`: chunkable messages exist, but no enterprise-memory chunks exist.
- `Stale Projection`: newer meaningful archive messages exist than the latest indexed chunk.

## API

Endpoint:

```text
GET /api/admin/archive-diagnostics
```

Query params:

- `source`: `slack` or `teams`; defaults to `slack`.
- `userId`: optional LibreChat user ID.
- `q`: optional channel/chat metadata search.
- `type`: optional source-specific conversation type.
- `status`: optional archive sync status.
- `limit`: page size, capped server-side.
- `offset`: page offset.

The route requires the same admin reporting access as the rest of the admin usage APIs.

## Debug Workflow

When a user says a channel is missing after sync:

1. Open Admin reporting -> Archive indexing.
2. Select the source and search for the channel name or source conversation ID.
3. If no row appears, the conversation was not discovered or stored for that source/user. Check latest sync status, OAuth scopes, app installation, and whether the bot/user has access to that Slack channel or Teams chat.
4. If the row is `Sync Failed`, inspect the sync error in the row and latest sync job card.
5. If the row is `No Messages`, discovery worked but history capture did not store messages.
6. If the row is `No Chunkable Messages`, history capture worked but message normalization skipped everything. Check the skipped-message reason breakdown.
7. If the row is `Not Projected`, the archive has chunkable messages but the memory projection stage did not produce chunks.
8. If the row is `Stale Projection`, rerun sync/projection or inspect projection job timing.
9. If the row is `Healthy` but agent retrieval fails, debug the retrieval tool/query path rather than the sync/indexing pipeline.

## Implementation Notes

The diagnostics service reads these model families:

- Slack: `SlackArchiveConversation`, `SlackArchiveMessage`, `SlackArchiveSyncJob`.
- Teams: `TeamsArchiveConversation`, `TeamsArchiveMessage`, `TeamsArchiveSyncJob`.
- Projection: `EnterpriseMemoryChunk`, `EnterpriseMemoryJob`.

The service compares archive conversation/message metadata against chunk counts by source parent record ID. This is intentionally a read-only diagnostic surface and should not mutate archive, sync, or memory records.
