const { tool } = require('@langchain/core/tools');
const SlackArchiveService = require('~/server/services/SlackArchiveService');

const SLACK_ARCHIVE_TOOL_NAME = 'slack_archive_search';

const slackArchiveJsonSchema = {
  type: 'object',
  properties: {
    action: {
      type: 'string',
      enum: [
        'status',
        'sync_archive',
        'advanced_search_messages',
        'search_messages',
        'list_conversations',
        'get_messages',
      ],
      description:
        'Prefer advanced_search_messages for semantic/hybrid indexed retrieval over archived Slack memory. Use search_messages only for exact lexical fallback, list_conversations for channel/DM selection, get_messages to inspect one selected conversation, status to check readiness, and sync_archive to ingest GovSlack history.',
    },
    query: {
      type: 'string',
      description:
        'For search_messages: a short distilled keyword phrase for exact lexical fallback. Do not pass the full user question here.',
    },
    topic: {
      type: 'string',
      description:
        "For advanced_search_messages: pass the user's full natural-language question verbatim (e.g., \"who said they are responsible for landing gear\" or \"what did I tell people I would do today\"). The server infers the sender (including the signed-in user for first-person questions), the timeframe (today/this week/etc.), and the intent (commitments, ownership/responsibility), then re-ranks results — so do not pre-distill it.",
    },
    conversationId: {
      type: 'string',
      description:
        'Stable Slack conversation id, such as a channel id or DM id. Use this for follow-up retrieval in one conversation.',
    },
    senderUserId: {
      type: 'string',
      description: 'Optional Slack user id to restrict advanced_search_messages or search_messages to one sender.',
    },
    senderScope: {
      type: 'string',
      enum: ['any', 'me', 'others'],
      description:
        'For advanced_search_messages: restrict sender to any, the signed-in user, or others.',
    },
    conversationType: {
      type: 'string',
      enum: ['any', 'public_channel', 'private_channel', 'im', 'mpim'],
      description:
        'For advanced_search_messages: optionally restrict to public channels, private channels, DMs, or group DMs.',
    },
    participants: {
      type: 'array',
      items: { type: 'string' },
      maxItems: 10,
      description:
        'For advanced_search_messages: optional participant names, emails, Slack user ids, or channel names to narrow search.',
    },
    daysBack: {
      type: 'integer',
      minimum: 1,
      maximum: 3650,
      description: 'For advanced_search_messages: how many days back to search.',
    },
    sortBy: {
      type: 'string',
      enum: ['relevance', 'recent', 'oldest'],
      description:
        'For advanced_search_messages: relevance uses the Mongo text index score, recent/oldest use time ordering.',
    },
    limit: {
      type: 'integer',
      minimum: 1,
      maximum: 100,
      description: 'Optional result limit. Keep this small unless exhaustive retrieval is requested.',
    },
    offset: {
      type: 'integer',
      minimum: 0,
      maximum: 100000,
      description: 'Optional pagination offset.',
    },
    conversationLimit: {
      type: 'integer',
      minimum: 1,
      maximum: 10000,
      description: 'For sync_archive: maximum number of Slack conversations to ingest.',
    },
    messagesPerConversation: {
      type: 'integer',
      minimum: 1,
      maximum: 5000,
      description: 'For sync_archive: maximum number of messages to ingest per conversation.',
    },
  },
  required: ['action'],
};

function formatJsonResult(result) {
  return JSON.stringify(result);
}

function createSlackArchiveTool({ req }) {
  return tool(
    async ({
      action,
      query,
      topic,
      conversationId,
      senderUserId,
      senderScope,
      conversationType,
      participants,
      daysBack,
      sortBy,
      limit,
      offset,
      conversationLimit,
      messagesPerConversation,
    }) => {
      const user = req?.user;
      if (!user) {
        throw new Error('Authenticated user context is required for Slack archive access');
      }

      if (action === 'status') {
        return formatJsonResult(await SlackArchiveService.getStatus(user));
      }

      if (action === 'sync_archive') {
        return formatJsonResult(
          await SlackArchiveService.syncUserArchive(user, {
            conversationLimit,
            messagesPerConversation,
          }),
        );
      }

      if (action === 'search_messages') {
        return formatJsonResult(
          await SlackArchiveService.searchMessages(user, {
            query,
            conversationId,
            senderUserId,
            limit,
            offset,
          }),
        );
      }

      if (action === 'advanced_search_messages') {
        return formatJsonResult(
          await SlackArchiveService.advancedSearchMessages(user, {
            query,
            topic,
            conversationId,
            senderUserId,
            senderScope,
            conversationType,
            participants,
            daysBack,
            sortBy,
            limit,
            offset,
          }),
        );
      }

      if (action === 'list_conversations') {
        return formatJsonResult(
          await SlackArchiveService.listConversations(user, {
            limit,
            offset,
          }),
        );
      }

      if (action === 'get_messages') {
        return formatJsonResult(
          await SlackArchiveService.listConversationMessages(user, conversationId, {
            limit,
            offset,
          }),
        );
      }

      throw new Error(`Unsupported Slack archive action: ${action}`);
    },
    {
      name: SLACK_ARCHIVE_TOOL_NAME,
      description:
        'Searches and retrieves archived GovSlack messages. Prefer advanced_search_messages for indexed semantic/hybrid discovery using EnterpriseMemory text-index chunks; use get_messages for exact channel/thread context and search_messages only as lexical fallback.',
      schema: slackArchiveJsonSchema,
    },
  );
}

module.exports = {
  SLACK_ARCHIVE_TOOL_NAME,
  slackArchiveJsonSchema,
  createSlackArchiveTool,
};
