#!/usr/bin/env node
require('dotenv').config();
require('module-alias/register');

const path = require('path');
const moduleAlias = require('module-alias');
const mongoose = require('mongoose');
const { createModels } = require('@librechat/data-schemas');
const { connectDb } = require('../api/db/connect');

moduleAlias.addAlias('~', path.resolve(__dirname, '..', 'api'));

const TeamsArchiveService = require('../api/server/services/TeamsArchiveService');

function parseArgs(argv) {
  const args = {};
  for (let index = 2; index < argv.length; index += 1) {
    const token = argv[index];
    if (!token.startsWith('--')) {
      continue;
    }
    const key = token.slice(2);
    const next = argv[index + 1];
    if (!next || next.startsWith('--')) {
      args[key] = true;
    } else {
      args[key] = next;
      index += 1;
    }
  }

  return {
    user: args.user || '',
    chatId: args.chatId || '',
    apply: args.apply === true,
    dryRun: args['dry-run'] === true || args.apply !== true,
    limit: args.limit ? Number(args.limit) : undefined,
    help: args.help === true || args.h === true,
  };
}

function printHelp() {
  console.log(`
Usage:
  node config/backfill-teams-conversation-recency.js --user <email> --dry-run
  node config/backfill-teams-conversation-recency.js --user <email> --apply
  node config/backfill-teams-conversation-recency.js --chatId <graphChatId-or-objectId> --dry-run
  node config/backfill-teams-conversation-recency.js --chatId <graphChatId-or-objectId> --apply
  node config/backfill-teams-conversation-recency.js --user <email> --chatId <id> --apply

Options:
  --user <email>     Limit to one Cortex user by email.
  --chatId <id>      Limit to one Teams archive conversation by graphChatId or Mongo ObjectId.
  --limit <n>        Max conversations to process for user-wide backfills.
  --dry-run          Print old/new recency fields without updating. Default.
  --apply            Apply recency field updates.
`);
}

async function findUserByEmail(email) {
  const User = mongoose.models.User;
  if (!User || !email) {
    return null;
  }

  return User.findOne({
    $or: [{ email }, { email: String(email).toLowerCase() }],
  }).lean();
}

async function findUsersForChat(chatId) {
  const TeamsArchiveConversation = mongoose.models.TeamsArchiveConversation;
  if (!TeamsArchiveConversation || !chatId) {
    return [];
  }

  const objectIdMatch = /^[a-f0-9]{24}$/i.test(String(chatId));
  const conversations = await TeamsArchiveConversation.find({
    $or: [
      { graphChatId: chatId },
      ...(objectIdMatch ? [{ _id: chatId }] : []),
    ],
  })
    .select({ user: 1, graphChatId: 1, topic: 1 })
    .lean();

  return conversations.map((conversation) => ({
    id: conversation.user?.toString?.() || String(conversation.user || ''),
    chatId: conversation.graphChatId,
    topic: conversation.topic || '',
  }));
}

function printConversationResult(result) {
  console.log(
    JSON.stringify(
      {
        topic: result.topic,
        graphChatId: result.graphChatId,
        archiveConversationId: result.archiveConversationId,
        totalMessageCount: result.totalMessageCount,
        loadedMessageCount: result.loadedMessageCount,
        truncated: result.truncated,
        wouldChange: result.wouldChange,
        didChange: result.didChange,
        oldRecency: result.oldRecency,
        newRecency: result.newRecency,
      },
      null,
      2,
    ),
  );
}

async function run() {
  const args = parseArgs(process.argv);
  if (args.help) {
    printHelp();
    return;
  }

  if (!args.user && !args.chatId) {
    throw new Error('Provide --user <email>, --chatId <id>, or both.');
  }

  await connectDb();
  createModels(mongoose);

  const targets = [];
  if (args.user) {
    const user = await findUserByEmail(args.user);
    if (!user) {
      throw new Error(`No user found for email: ${args.user}`);
    }
    targets.push({
      id: user._id?.toString?.() || user.id,
      email: user.email,
      chatId: args.chatId,
    });
  } else {
    const chatTargets = await findUsersForChat(args.chatId);
    if (chatTargets.length === 0) {
      throw new Error(`No Teams archive conversations found for chatId: ${args.chatId}`);
    }
    targets.push(...chatTargets);
  }

  const summaries = [];
  for (const target of targets) {
    console.log(
      `[backfill-teams-conversation-recency] ${args.apply ? 'Applying' : 'Dry run'} for user=${target.email || target.id} chatId=${target.chatId || args.chatId || 'ALL'}`,
    );
    const result = await TeamsArchiveService.backfillConversationRecency(
      { id: target.id, email: target.email },
      {
        chatId: target.chatId || args.chatId,
        apply: args.apply,
        limit: args.limit,
      },
    );
    result.conversations.forEach(printConversationResult);
    summaries.push({
      user: target.email || target.id,
      processedConversationCount: result.processedConversationCount,
      changedConversationCount: result.changedConversationCount,
      updatedConversationCount: result.updatedConversationCount,
    });
  }

  console.log('[backfill-teams-conversation-recency] Summary');
  console.log(JSON.stringify(summaries, null, 2));
}

if (require.main === module) {
  run()
    .then(async () => {
      await mongoose.disconnect();
      process.exit(0);
    })
    .catch(async (error) => {
      console.error('[backfill-teams-conversation-recency] Failed:', error?.message || error);
      await mongoose.disconnect().catch(() => undefined);
      process.exit(1);
    });
}

module.exports = {
  parseArgs,
  run,
};
