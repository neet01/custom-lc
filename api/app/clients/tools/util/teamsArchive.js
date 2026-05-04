const { tool } = require('@langchain/core/tools');
const TeamsArchiveService = require('~/server/services/TeamsArchiveService');

const TEAMS_ARCHIVE_TOOL_NAME = 'teams_archive_search';

const teamsArchiveJsonSchema = {
  type: 'object',
  properties: {
    action: {
      type: 'string',
      enum: ['status', 'sync_archive', 'search_messages', 'list_conversations', 'get_messages'],
      description:
        'Use status to check archive readiness, sync_archive to ingest Teams chat history, search_messages to query archived content, list_conversations to inspect available archived chats, and get_messages to read a specific archived chat thread.',
    },
    query: {
      type: 'string',
      description:
        'For search_messages: the search query to run against archived Teams chat content.',
    },
    chatId: {
      type: 'string',
      description:
        'For get_messages or search_messages: the archived Teams chat id to scope the request to.',
    },
    limit: {
      type: 'integer',
      minimum: 1,
      maximum: 100,
      description: 'Optional result limit for searches, conversation lists, and message retrieval.',
    },
    offset: {
      type: 'integer',
      minimum: 0,
      maximum: 100000,
      description: 'Optional pagination offset for list and search operations.',
    },
    chatLimit: {
      type: 'integer',
      minimum: 1,
      maximum: 250,
      description:
        'For sync_archive: maximum number of Teams chats to ingest during this sync request.',
    },
    messagesPerChat: {
      type: 'integer',
      minimum: 1,
      maximum: 1000,
      description:
        'For sync_archive: maximum number of messages to ingest per Teams chat during this sync request.',
    },
  },
  required: ['action'],
};

function formatJsonResult(result) {
  return JSON.stringify(result, null, 2);
}

function createTeamsArchiveTool({ req }) {
  return tool(
    async ({ action, query, chatId, limit, offset, chatLimit, messagesPerChat }) => {
      const user = req?.user;

      if (!user) {
        throw new Error('Authenticated user context is required for Teams archive access');
      }

      if (action === 'status') {
        return formatJsonResult(await TeamsArchiveService.getStatus(user));
      }

      if (action === 'sync_archive') {
        return formatJsonResult(
          await TeamsArchiveService.syncUserArchive(user, {
            mode: 'chats',
            chatLimit,
            messagesPerChat,
          }),
        );
      }

      if (action === 'search_messages') {
        return formatJsonResult(
          await TeamsArchiveService.searchMessages(user, {
            query,
            chatId,
            limit,
            offset,
          }),
        );
      }

      if (action === 'list_conversations') {
        return formatJsonResult(
          await TeamsArchiveService.listConversations(user, {
            limit,
            offset,
          }),
        );
      }

      if (action === 'get_messages') {
        return formatJsonResult(
          await TeamsArchiveService.listConversationMessages(user, chatId, {
            limit,
            offset,
          }),
        );
      }

      throw new Error(`Unsupported Teams archive action: ${action}`);
    },
    {
      name: TEAMS_ARCHIVE_TOOL_NAME,
      description:
        'Searches and retrieves archived Microsoft Teams chats that were previously ingested into Cortex. Use this to query old Teams discussions after Teams becomes read-only.',
      schema: teamsArchiveJsonSchema,
    },
  );
}

module.exports = {
  TEAMS_ARCHIVE_TOOL_NAME,
  createTeamsArchiveTool,
  teamsArchiveJsonSchema,
};
