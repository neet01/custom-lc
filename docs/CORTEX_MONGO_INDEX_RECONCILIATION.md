# Cortex Mongo Index Reconciliation

This runbook covers manual index reconciliation for Cortex Teams archive and enterprise memory collections.

The script does not rely on `MONGO_AUTO_INDEX`. It uses the configured `MONGO_URI`, prints current indexes, checks duplicates for every unique index, creates non-unique performance indexes first, and only creates unique indexes when duplicate checks pass.

## Script

```bash
config/reconcile-cortex-indexes.js
```

Target collections:

- `teamsarchivemessages`
- `teamsarchiveconversations`
- `enterprisememorychunks`
- `enterprisememoryentities`
- `enterprisememoryrelationships`

## Modes

Dry run is the default.

```bash
node config/reconcile-cortex-indexes.js --dry-run
```

Apply missing indexes:

```bash
node config/reconcile-cortex-indexes.js --apply
```

Verify required indexes exist:

```bash
node config/reconcile-cortex-indexes.js --verify
```

Run query plan checks:

```bash
node config/reconcile-cortex-indexes.js --explain
```

Apply, verify, and explain in one run:

```bash
node config/reconcile-cortex-indexes.js --apply --verify --explain
```

## Docker Examples

Staging dry run:

```bash
UID=$(id -u) GID=$(id -g) docker compose exec -T api node /app/config/reconcile-cortex-indexes.js --dry-run
```

Staging apply:

```bash
UID=$(id -u) GID=$(id -g) docker compose exec -T api node /app/config/reconcile-cortex-indexes.js --apply
```

Production verify:

```bash
UID=$(id -u) GID=$(id -g) docker compose exec -T api node /app/config/reconcile-cortex-indexes.js --verify
```

Production explain checks:

```bash
UID=$(id -u) GID=$(id -g) docker compose exec -T api node /app/config/reconcile-cortex-indexes.js --explain
```

## Recommended Production Sequence

1. Wait for the active Teams sync/projection to finish.
2. Take and verify a MongoDB backup.
3. Run dry run in production.
4. Review duplicate output for all unique indexes.
5. If duplicates are present, do not apply unique indexes until duplicates are cleaned up.
6. Run apply.
7. Run verify.
8. Run explain checks and confirm important paths use indexes instead of `COLLSCAN`.

## Duplicate Handling

The script checks duplicates before creating each unique index. If duplicates exist, that unique index is skipped and the script prints duplicate key samples and counts.

Manual cleanup is required before rerunning `--apply`.

The high-risk unique indexes are:

- `teamsarchivemessages`: `{ user, graphMessageId }`
- `teamsarchiveconversations`: `{ user, graphChatId }`
- `enterprisememorychunks`: projection key across scope, user, source, source record, chunk type, and order
- `enterprisememoryentities`: canonical entity key
- `enterprisememoryrelationships`: canonical relationship key

## Explain Checks

The script runs `explain("executionStats")` for:

- Teams message lookup by `user + graphChatId`, sorted by `sentDateTime`
- Teams recent message lookup by `user + graphChatId`, sorted by `sentDateTime` descending
- Teams sender fallback lookup by `user + fromEmail`, sorted by `sentDateTime` descending
- Teams sender fallback lookup by `user + fromDisplayName`, sorted by `sentDateTime` descending
- Teams conversation lookup by `user + graphChatId`
- Enterprise memory chunk lookup by `user + source + sourceRecordType`
- Enterprise memory chunk lookup by `user + source + metadata.chatType + chunkType`
- Enterprise memory chunk unique upsert key

Healthy output should show `IXSCAN` stages and low `totalDocsExamined`. If output still shows `COLLSCAN`, inspect missing or mismatched indexes in `--verify`.

## Notes

- Index creation uses `background: true` where supported.
- The script is idempotent and skips indexes that already exist with the expected key and uniqueness.
- If the script is updated with additional indexes, rerun `--dry-run`, then `--apply`; existing indexes will be skipped and only missing indexes will be created.
- If an index exists with the same key but different uniqueness, the script prints a manual-action warning and does not modify that index.
- MongoDB reports text indexes internally as `_fts/_ftsx`; the script verifies the configured text index by name to avoid false missing-index reports.
