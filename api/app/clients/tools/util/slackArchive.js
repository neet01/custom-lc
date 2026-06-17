const { tool } = require('@langchain/core/tools');
const SlackArchiveService = require('~/server/services/SlackArchiveService');

const SLACK_ARCHIVE_TOOL_NAME = 'slack_archive_search';

const slackArchiveJsonSchema = {
  type: 'object',
  properties: {
    action: {
      type: 'string',
      enum: ['status', 'sync_archive', 'search_messages', 'list_conversations', 'get_messages'],
      description:
        'Use status to check readiness, sync_archive to begin Slack archive ingestion scaffolding, search_messages for lexical discovery, list_conversations for channel/DM selection, and get_messages to inspect one selected conversation.',
    },
    query: {
      type: 'string',
      description:
        'For search_messages: a short search phrase. Do not pass the full user question.',
    },
    conversationId: {
      type: 'string',
      description:
        'Stable Slack conversation id, such as a channel id or DM id. Use this for follow-up retrieval in one conversation.',
    },
    senderUserId: {
      type: 'string',
      description: 'Optional Slack user id to restrict search_messages to one sender.',
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
      conversationId,
      senderUserId,
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
        'Searches and retrieves archived Slack messages. Current scope supports status, GovSlack archive sync, lexical search, conversation listing, and per-conversation message retrieval.',
      schema: slackArchiveJsonSchema,
    },
  );
}

module.exports = {
  SLACK_ARCHIVE_TOOL_NAME,
  slackArchiveJsonSchema,
  createSlackArchiveTool,
};
