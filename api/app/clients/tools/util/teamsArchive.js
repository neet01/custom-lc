const { tool } = require('@langchain/core/tools');
const TeamsArchiveService = require('~/server/services/TeamsArchiveService');

const TEAMS_ARCHIVE_TOOL_NAME = 'teams_archive_search';

const teamsArchiveJsonSchema = {
  type: 'object',
  properties: {
    action: {
      type: 'string',
      enum: [
        'status',
        'sync_archive',
        'search_messages',
        'advanced_search_messages',
        'recent_messages',
        'list_conversations',
        'get_messages',
      ],
      description:
        'Use status to check archive readiness, sync_archive to ingest Teams chat history, search_messages to query archived content, advanced_search_messages for structured discovery across topic, sender scope, chat type, participants, and recency, recent_messages to find messages the signed-in user sent recently, list_conversations to inspect available archived chats, and get_messages to read a specific archived chat thread.',
    },
    query: {
      type: 'string',
      description:
        'For search_messages or recent_messages: the search query to run against archived Teams chat content.',
    },
    topic: {
      type: 'string',
      description:
        'For advanced_search_messages: the core topic or subject to search for. Pass the distilled subject, not the full user question.',
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
    daysBack: {
      type: 'integer',
      minimum: 1,
      maximum: 3650,
      description:
        'For recent_messages or advanced_search_messages: how many days back to search.',
    },
    senderScope: {
      type: 'string',
      enum: ['any', 'me', 'others'],
      description:
        'For advanced_search_messages: whether to search messages from anyone, only the signed-in user, or everyone except the signed-in user.',
    },
    chatType: {
      type: 'string',
      enum: ['any', 'oneOnOne', 'group', 'meeting'],
      description:
        'For advanced_search_messages: optionally restrict to one-on-one chats, group chats, or meeting chats.',
    },
    participants: {
      type: 'array',
      items: { type: 'string' },
      maxItems: 10,
      description:
        'For advanced_search_messages: optional participant names or emails to narrow the search to specific chats.',
    },
    sortBy: {
      type: 'string',
      enum: ['recent', 'oldest'],
      description:
        'For advanced_search_messages: whether to return the newest matches first or the oldest matches first.',
    },
  },
  required: ['action'],
};

function formatJsonResult(result) {
  return JSON.stringify(result, null, 2);
}

function createTeamsArchiveTool({ req }) {
  return tool(
    async ({
      action,
      query,
      topic,
      chatId,
      limit,
      offset,
      chatLimit,
      messagesPerChat,
      daysBack,
      senderScope,
      chatType,
      participants,
      sortBy,
    }) => {
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

      if (action === 'advanced_search_messages') {
        return formatJsonResult(
          await TeamsArchiveService.advancedSearchMessages(user, {
            query,
            topic,
            limit,
            offset,
            daysBack,
            senderScope,
            chatType,
            participants,
            sortBy,
          }),
        );
      }

      if (action === 'recent_messages') {
        return formatJsonResult(
          await TeamsArchiveService.recentMessages(user, {
            query,
            limit,
            daysBack,
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
        'Searches and retrieves archived Microsoft Teams chats that were previously ingested into Cortex. Always provide an "action" parameter. Use action="advanced_search_messages" for questions like what has been discussed about a topic, optionally constrained by timeframe, participants, chat type, or sender scope. Use action="recent_messages" when the user asks what they sent recently or asks for their own recent Teams messages.',
      schema: teamsArchiveJsonSchema,
    },
  );
}

module.exports = {
  TEAMS_ARCHIVE_TOOL_NAME,
  createTeamsArchiveTool,
  teamsArchiveJsonSchema,
};
