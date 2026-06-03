require('dotenv').config();

const mongoose = require('mongoose');
const { connectDb } = require('../api/db/connect');

const TARGET_COLLECTIONS = [
  'teamsarchivemessages',
  'teamsarchiveconversations',
  'enterprisememorychunks',
  'enterprisememoryentities',
  'enterprisememoryrelationships',
];

const INDEXES = {
  teamsarchivemessages: [
    {
      name: 'teams_message_by_chat_sent_asc',
      key: { user: 1, graphChatId: 1, sentDateTime: 1, createdAt: 1 },
      reason: 'Projection and dossier reads by chat in chronological order',
    },
    {
      name: 'teams_message_by_chat_sent_desc',
      key: { user: 1, graphChatId: 1, sentDateTime: -1 },
      reason: 'Recent chat message reads',
    },
    {
      name: 'teams_message_by_sender_sent_desc',
      key: { user: 1, fromUserId: 1, sentDateTime: -1 },
      reason: 'Sender-scoped archive search',
    },
    {
      name: 'teams_message_by_user_sent_desc',
      key: { user: 1, sentDateTime: -1 },
      reason: 'Recent archive search',
    },
    {
      name: 'teams_message_by_chat_created_desc',
      key: { user: 1, graphChatId: 1, createdAt: -1 },
      reason: 'Fallback chat message ordering',
    },
    {
      name: 'teams_message_by_tenant_chat_sent_desc',
      key: { tenantId: 1, graphChatId: 1, sentDateTime: -1 },
      reason: 'Tenant-scoped archive reads',
    },
    {
      name: 'teams_message_unique_user_graph_message',
      key: { user: 1, graphMessageId: 1 },
      unique: true,
      duplicateFields: ['user', 'graphMessageId'],
      reason: 'Idempotent Teams message upsert',
    },
  ],
  teamsarchiveconversations: [
    {
      name: 'teams_conversation_by_last_message',
      key: { user: 1, lastMessageAt: -1 },
      reason: 'Conversation list ordering',
    },
    {
      name: 'teams_conversation_sync_queue',
      key: { user: 1, syncStatus: 1, sourceUpdatedAt: -1, sourceLastMessageAt: -1, updatedAt: -1 },
      reason: 'Message sync queue selection',
    },
    {
      name: 'teams_conversation_by_source_last_message',
      key: { user: 1, sourceLastMessageAt: -1, updatedAt: -1 },
      reason: 'Backfill/discovery ordering',
    },
    {
      name: 'teams_conversation_by_participant_degraded',
      key: { user: 1, participantDegraded: 1, updatedAt: -1 },
      reason: 'Searchability diagnostics',
    },
    {
      name: 'teams_conversation_by_chat_type',
      key: { user: 1, chatType: 1, lastMessageAt: -1, updatedAt: -1 },
      reason: 'Chat-type-scoped conversation lookup',
    },
    {
      name: 'teams_conversation_by_tenant_last_message',
      key: { tenantId: 1, lastMessageAt: -1 },
      reason: 'Tenant-scoped conversation ordering',
    },
    {
      name: 'teams_conversation_unique_user_graph_chat',
      key: { user: 1, graphChatId: 1 },
      unique: true,
      duplicateFields: ['user', 'graphChatId'],
      reason: 'Idempotent Teams conversation upsert',
    },
  ],
  enterprisememorychunks: [
    {
      name: 'em_chunk_by_parent_timestamp',
      key: { parentEntityId: 1, sourceTimestamp: -1 },
      reason: 'Parent entity context retrieval',
    },
    {
      name: 'em_chunk_by_source_type_timestamp',
      key: { source: 1, sourceRecordType: 1, sourceTimestamp: -1 },
      reason: 'Source/type scoped chunk scans',
    },
    {
      name: 'em_chunk_by_user_source_type_timestamp',
      key: { user: 1, source: 1, sourceRecordType: 1, sourceTimestamp: -1 },
      reason: 'Enterprise memory retrieval and counts',
    },
    {
      name: 'em_chunk_by_user_parent_timestamp',
      key: { user: 1, sourceParentRecordId: 1, sourceTimestamp: -1 },
      reason: 'Conversation-scoped chunk retrieval',
    },
    {
      name: 'em_chunk_by_user_source_type_parent',
      key: { user: 1, source: 1, sourceRecordType: 1, sourceParentRecordId: 1 },
      reason: 'Projected searchable conversation diagnostics',
    },
    {
      name: 'em_chunk_by_user_source_chat_type_chunk_parent',
      key: { user: 1, source: 1, 'metadata.chatType': 1, chunkType: 1, sourceParentRecordId: 1 },
      reason: 'Searchability diagnostics by Teams chat type',
    },
    {
      name: 'em_chunk_text_search',
      key: { title: 'text', text: 'text', summary: 'text' },
      reason: 'Text search fallback support',
    },
    {
      name: 'em_chunk_unique_projection_key',
      key: {
        visibilityScope: 1,
        user: 1,
        tenantId: 1,
        source: 1,
        sourceRecordType: 1,
        sourceRecordId: 1,
        chunkType: 1,
        orderIndex: 1,
      },
      unique: true,
      duplicateFields: [
        'visibilityScope',
        'user',
        'tenantId',
        'source',
        'sourceRecordType',
        'sourceRecordId',
        'chunkType',
        'orderIndex',
      ],
      reason: 'Idempotent enterprise memory chunk projection upsert',
    },
  ],
  enterprisememoryentities: [
    {
      name: 'em_entity_by_source_type_display',
      key: { source: 1, entityType: 1, displayName: 1 },
      reason: 'Entity lookup by source/type/display name',
    },
    {
      name: 'em_entity_by_tenant_source_type_updated',
      key: { tenantId: 1, source: 1, entityType: 1, updatedAt: -1 },
      reason: 'Tenant-scoped entity listing',
    },
    {
      name: 'em_entity_by_user_source_type_record_type',
      key: { user: 1, source: 1, entityType: 1, sourceRecordType: 1 },
      reason: 'Projection coverage diagnostics',
    },
    {
      name: 'em_entity_unique_canonical_key',
      key: { visibilityScope: 1, user: 1, tenantId: 1, source: 1, entityType: 1, canonicalKey: 1 },
      unique: true,
      duplicateFields: ['visibilityScope', 'user', 'tenantId', 'source', 'entityType', 'canonicalKey'],
      reason: 'Idempotent canonical entity projection',
    },
  ],
  enterprisememoryrelationships: [
    {
      name: 'em_relationship_by_from_type',
      key: { fromEntityId: 1, relationshipType: 1 },
      reason: 'Forward graph traversal',
    },
    {
      name: 'em_relationship_by_to_type',
      key: { toEntityId: 1, relationshipType: 1 },
      reason: 'Reverse graph traversal',
    },
    {
      name: 'em_relationship_unique_projection_key',
      key: {
        visibilityScope: 1,
        user: 1,
        tenantId: 1,
        source: 1,
        relationshipType: 1,
        fromEntityId: 1,
        toEntityId: 1,
        sourceRecordType: 1,
        sourceRecordId: 1,
      },
      unique: true,
      duplicateFields: [
        'visibilityScope',
        'user',
        'tenantId',
        'source',
        'relationshipType',
        'fromEntityId',
        'toEntityId',
        'sourceRecordType',
        'sourceRecordId',
      ],
      reason: 'Idempotent enterprise relationship projection',
    },
  ],
};

function parseArgs(argv) {
  const args = new Set(argv.slice(2));
  return {
    apply: args.has('--apply'),
    dryRun: args.has('--dry-run') || !args.has('--apply'),
    verify: args.has('--verify'),
    explain: args.has('--explain'),
    help: args.has('--help') || args.has('-h'),
  };
}

function printHelp() {
  console.log(`
Usage:
  node config/reconcile-cortex-indexes.js [--dry-run] [--apply] [--verify] [--explain]

Modes:
  --dry-run   Print current indexes, duplicate checks, and planned work. Default.
  --apply     Create missing indexes after duplicate checks pass.
  --verify    Confirm required indexes exist. Does not create indexes by itself.
  --explain   Run explain() checks for key Teams/archive/memory query patterns.

Examples:
  node config/reconcile-cortex-indexes.js --dry-run
  node config/reconcile-cortex-indexes.js --apply
  node config/reconcile-cortex-indexes.js --verify
  node config/reconcile-cortex-indexes.js --explain
`);
}

function stableStringify(value) {
  if (value === null || typeof value !== 'object' || value instanceof Date) {
    return JSON.stringify(value);
  }

  if (Array.isArray(value)) {
    return `[${value.map((entry) => stableStringify(entry)).join(',')}]`;
  }

  return `{${Object.keys(value)
    .sort()
    .map((key) => `${JSON.stringify(key)}:${stableStringify(value[key])}`)
    .join(',')}}`;
}

function keysEqual(left, right) {
  return stableStringify(left) === stableStringify(right);
}

function normalizeIndexForOutput(index) {
  return {
    name: index.name,
    key: index.key,
    unique: Boolean(index.unique),
  };
}

function findExistingIndex(currentIndexes, target) {
  return currentIndexes.find((index) => keysEqual(index.key, target.key));
}

function groupIdFromFields(fields) {
  return fields.reduce((acc, field) => {
    acc[field.replace(/\./g, '_')] = `$${field}`;
    return acc;
  }, {});
}

async function printCurrentIndexes(db) {
  console.log('\nCURRENT INDEXES');
  for (const collectionName of TARGET_COLLECTIONS) {
    const indexes = await db.collection(collectionName).indexes().catch((error) => {
      if (error?.codeName === 'NamespaceNotFound') {
        return [];
      }
      throw error;
    });
    console.log(`\n${collectionName}`);
    console.log(JSON.stringify(indexes.map(normalizeIndexForOutput), null, 2));
  }
}

async function findDuplicates(collection, indexSpec, limit = 20) {
  const pipeline = [
    {
      $group: {
        _id: groupIdFromFields(indexSpec.duplicateFields),
        count: { $sum: 1 },
        ids: { $push: '$_id' },
      },
    },
    { $match: { count: { $gt: 1 } } },
    { $limit: limit },
    {
      $project: {
        _id: 0,
        key: '$_id',
        count: 1,
        ids: { $slice: ['$ids', 10] },
      },
    },
  ];

  return collection.aggregate(pipeline, { allowDiskUse: true }).toArray();
}

async function runDuplicateChecks(db) {
  console.log('\nDUPLICATE CHECKS FOR UNIQUE INDEXES');
  const results = {};

  for (const [collectionName, indexSpecs] of Object.entries(INDEXES)) {
    const collection = db.collection(collectionName);
    results[collectionName] = {};

    for (const indexSpec of indexSpecs.filter((index) => index.unique)) {
      const duplicates = await findDuplicates(collection, indexSpec);
      results[collectionName][indexSpec.name] = duplicates;

      if (duplicates.length === 0) {
        console.log(`[OK] ${collectionName}.${indexSpec.name}: no duplicates found`);
      } else {
        console.log(`[DUPLICATES] ${collectionName}.${indexSpec.name}: ${duplicates.length} duplicate key sample(s)`);
        for (const duplicate of duplicates) {
          console.log(JSON.stringify({ key: duplicate.key, count: duplicate.count, ids: duplicate.ids }, null, 2));
        }
      }
    }
  }

  return results;
}

async function createMissingIndexes(db, duplicateResults, { dryRun }) {
  console.log('\nINDEX RECONCILIATION');

  for (const [collectionName, indexSpecs] of Object.entries(INDEXES)) {
    const collection = db.collection(collectionName);
    const currentIndexes = await collection.indexes().catch((error) => {
      if (error?.codeName === 'NamespaceNotFound') {
        return [];
      }
      throw error;
    });

    const nonUniqueIndexes = indexSpecs.filter((index) => !index.unique);
    const uniqueIndexes = indexSpecs.filter((index) => index.unique);

    for (const indexSpec of [...nonUniqueIndexes, ...uniqueIndexes]) {
      const existing = findExistingIndex(currentIndexes, indexSpec);
      const existingUnique = Boolean(existing?.unique);
      const targetUnique = Boolean(indexSpec.unique);

      if (existing && existingUnique === targetUnique) {
        console.log(`[EXISTS] ${collectionName}.${indexSpec.name}`);
        continue;
      }

      if (existing && existingUnique !== targetUnique) {
        console.log(
          `[MANUAL] ${collectionName}.${indexSpec.name}: key exists as ${existingUnique ? 'unique' : 'non-unique'} index "${existing.name}", expected ${targetUnique ? 'unique' : 'non-unique'}`,
        );
        continue;
      }

      if (indexSpec.unique && duplicateResults?.[collectionName]?.[indexSpec.name]?.length > 0) {
        console.log(`[SKIP] ${collectionName}.${indexSpec.name}: duplicate records exist`);
        continue;
      }

      const options = {
        name: indexSpec.name,
        background: true,
        ...(indexSpec.unique ? { unique: true } : {}),
      };

      if (dryRun) {
        console.log(`[DRY-RUN] would create ${collectionName}.${indexSpec.name}: ${JSON.stringify(indexSpec.key)}`);
        continue;
      }

      console.log(`[CREATE] ${collectionName}.${indexSpec.name}: ${indexSpec.reason}`);
      await collection.createIndex(indexSpec.key, options);
    }
  }
}

async function verifyIndexes(db) {
  console.log('\nVERIFICATION');
  let missingCount = 0;

  for (const [collectionName, indexSpecs] of Object.entries(INDEXES)) {
    const currentIndexes = await db.collection(collectionName).indexes().catch((error) => {
      if (error?.codeName === 'NamespaceNotFound') {
        return [];
      }
      throw error;
    });

    for (const indexSpec of indexSpecs) {
      const existing = findExistingIndex(currentIndexes, indexSpec);
      const expectedUnique = Boolean(indexSpec.unique);
      const actualUnique = Boolean(existing?.unique);

      if (!existing) {
        missingCount += 1;
        console.log(`[MISSING] ${collectionName}.${indexSpec.name}`);
      } else if (actualUnique !== expectedUnique) {
        missingCount += 1;
        console.log(
          `[MISMATCH] ${collectionName}.${indexSpec.name}: unique=${actualUnique}, expected unique=${expectedUnique}`,
        );
      } else {
        console.log(`[OK] ${collectionName}.${indexSpec.name}`);
      }
    }
  }

  if (missingCount > 0) {
    console.log(`\nVerification failed: ${missingCount} missing/mismatched index(es).`);
    return false;
  }

  console.log('\nVerification passed: all required Cortex indexes exist.');
  return true;
}

function planStageSummary(plan) {
  if (!plan || typeof plan !== 'object') {
    return [];
  }

  const stages = [];
  const visit = (node) => {
    if (!node || typeof node !== 'object') {
      return;
    }
    if (node.stage) {
      stages.push({
        stage: node.stage,
        indexName: node.indexName,
        direction: node.direction,
      });
    }
    for (const value of Object.values(node)) {
      if (value && typeof value === 'object') {
        if (Array.isArray(value)) {
          value.forEach(visit);
        } else {
          visit(value);
        }
      }
    }
  };

  visit(plan);
  return stages.filter((stage, index) => index === 0 || stage.stage !== stages[index - 1].stage || stage.indexName);
}

async function explainFind(collection, label, filter, options = {}) {
  const cursor = collection.find(filter).limit(1);
  if (options.sort) {
    cursor.sort(options.sort);
  }

  const explain = await cursor.explain('executionStats');
  const winningPlan = explain?.queryPlanner?.winningPlan;
  console.log(`\n[EXPLAIN] ${label}`);
  console.log(JSON.stringify({
    collection: collection.collectionName,
    filter,
    sort: options.sort,
    stages: planStageSummary(winningPlan),
    totalKeysExamined: explain?.executionStats?.totalKeysExamined,
    totalDocsExamined: explain?.executionStats?.totalDocsExamined,
    executionTimeMillis: explain?.executionStats?.executionTimeMillis,
  }, null, 2));
}

async function runExplainChecks(db) {
  console.log('\nEXPLAIN CHECKS');

  const messages = db.collection('teamsarchivemessages');
  const conversations = db.collection('teamsarchiveconversations');
  const chunks = db.collection('enterprisememorychunks');

  const sampleMessage = await messages.findOne({ graphChatId: { $exists: true, $ne: '' } });
  if (sampleMessage) {
    await explainFind(
      messages,
      'Teams message lookup by user + graphChatId sorted by sentDateTime',
      { user: sampleMessage.user, graphChatId: sampleMessage.graphChatId },
      { sort: { sentDateTime: 1, createdAt: 1 } },
    );
  } else {
    console.log('[SKIP] No sample Teams archive message found.');
  }

  const sampleConversation = await conversations.findOne({ graphChatId: { $exists: true, $ne: '' } });
  if (sampleConversation) {
    await explainFind(
      conversations,
      'Teams conversation lookup by user + graphChatId',
      { user: sampleConversation.user, graphChatId: sampleConversation.graphChatId },
    );
  } else {
    console.log('[SKIP] No sample Teams archive conversation found.');
  }

  const sampleChunk = await chunks.findOne({ source: 'teams', sourceRecordType: { $exists: true, $ne: '' } });
  if (sampleChunk) {
    await explainFind(
      chunks,
      'Enterprise memory chunk lookup by user + source + sourceRecordType',
      { user: sampleChunk.user, source: sampleChunk.source, sourceRecordType: sampleChunk.sourceRecordType },
      { sort: { sourceTimestamp: -1, updatedAt: -1 } },
    );

    await explainFind(
      chunks,
      'Enterprise memory chunk unique upsert key',
      {
        visibilityScope: sampleChunk.visibilityScope || 'user',
        user: sampleChunk.user || null,
        tenantId: sampleChunk.tenantId || null,
        source: sampleChunk.source,
        sourceRecordType: sampleChunk.sourceRecordType,
        sourceRecordId: sampleChunk.sourceRecordId,
        chunkType: sampleChunk.chunkType,
        orderIndex: sampleChunk.orderIndex || 0,
      },
    );
  } else {
    console.log('[SKIP] No sample enterprise memory chunk found.');
  }

  const sampleChatTypeChunk = await chunks.findOne({
    source: 'teams',
    'metadata.chatType': { $exists: true, $ne: '' },
    chunkType: { $exists: true, $ne: '' },
  });
  if (sampleChatTypeChunk) {
    await explainFind(
      chunks,
      'Enterprise memory chunk lookup by user + source + metadata.chatType + chunkType',
      {
        user: sampleChatTypeChunk.user,
        source: sampleChatTypeChunk.source,
        'metadata.chatType': sampleChatTypeChunk.metadata.chatType,
        chunkType: sampleChatTypeChunk.chunkType,
      },
    );
  } else {
    console.log('[SKIP] No sample enterprise memory chunk with metadata.chatType found.');
  }
}

async function reconcileCortexIndexes(options = parseArgs(process.argv)) {
  if (options.help) {
    printHelp();
    return { ok: true };
  }

  await connectDb();
  const db = mongoose.connection.db;

  console.log(`Cortex index reconciliation mode: ${options.dryRun ? 'dry-run' : 'apply'}`);
  await printCurrentIndexes(db);
  const duplicateResults = await runDuplicateChecks(db);

  if (!options.verify && !options.explain) {
    await createMissingIndexes(db, duplicateResults, options);
  } else if (options.apply) {
    await createMissingIndexes(db, duplicateResults, options);
  }

  const shouldVerify = options.verify || options.apply;
  const verified = shouldVerify ? await verifyIndexes(db) : true;

  if (options.explain) {
    await runExplainChecks(db);
  }

  return { ok: verified };
}

if (require.main === module) {
  reconcileCortexIndexes()
    .then(async (result) => {
      await mongoose.disconnect();
      process.exit(result.ok ? 0 : 2);
    })
    .catch(async (error) => {
      console.error('[reconcile-cortex-indexes] Failed:', error);
      await mongoose.disconnect().catch(() => undefined);
      process.exit(1);
    });
}

module.exports = {
  INDEXES,
  TARGET_COLLECTIONS,
  reconcileCortexIndexes,
};
