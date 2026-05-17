const { tool } = require('@langchain/core/tools');
const { logger } = require('@librechat/data-schemas');
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
        'conversation_dossier',
        'get_messages',
        'get_messages_window',
        'summarize_conversation',
      ],
      description:
        'Use status to check archive readiness, sync_archive to ingest Teams chat history, search_messages for quick preview retrieval, advanced_search_messages for structured topic discovery across sender scope, chat type, participants, and recency, recent_messages to find messages the signed-in user sent recently, list_conversations to inspect available archived chats, conversation_dossier for exhaustive archive-backed retrieval of one resolved chat, get_messages for compact thread previews, get_messages_window to pull a bounded context window around a message or topic hit, and summarize_conversation to answer high-level questions without loading the whole thread.',
    },
    query: {
      type: 'string',
      description:
        'For search_messages or recent_messages: a short search phrase for preview retrieval. For list_conversations or conversation_dossier this may be used as a conversation topic hint. Do not pass the full user question.',
    },
    topic: {
      type: 'string',
      description:
        'For advanced_search_messages: the core topic or subject to search for. For list_conversations, conversation_dossier, or summarize_conversation this can also be used as a conversation topic hint. Pass the distilled subject, not the full user question.',
    },
    chatId: {
      type: 'string',
      description:
        'For conversation_dossier, get_messages, get_messages_window, summarize_conversation, or search_messages: the archived Teams chat id to scope the request to.',
    },
    limit: {
      type: 'integer',
      minimum: 1,
      maximum: 100,
      description:
        'Optional result limit for searches, conversation lists, and message retrieval. Keep this small unless the user explicitly asks for exhaustive output.',
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
      maximum: 10000,
      description:
        'For sync_archive: maximum number of Teams chats to ingest during this sync request.',
    },
    messagesPerChat: {
      type: 'integer',
      minimum: 1,
      maximum: 5000,
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
        'For advanced_search_messages, list_conversations, or conversation_dossier: optionally restrict to one-on-one chats, group chats, or meeting chats.',
    },
    participants: {
      type: 'array',
      items: { type: 'string' },
      maxItems: 10,
      description:
        'For advanced_search_messages, list_conversations, or conversation_dossier: optional participant names or emails to narrow the search to specific chats.',
    },
    sortBy: {
      type: 'string',
      enum: ['recent', 'oldest'],
      description:
        'For advanced_search_messages: whether to return the newest matches first or the oldest matches first.',
    },
    aroundMessageId: {
      type: 'string',
      description:
        'For get_messages_window: optional archived message id or Teams graph message id to center the window around.',
    },
    before: {
      type: 'integer',
      minimum: 0,
      maximum: 50,
      description:
        'For get_messages_window: how many messages to return before the anchor message. Prefer small windows.',
    },
    after: {
      type: 'integer',
      minimum: 0,
      maximum: 50,
      description:
        'For get_messages_window: how many messages to return after the anchor message. Prefer small windows.',
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
      aroundMessageId,
      before,
      after,
    }) => {
      const user = req?.user;

      if (!user) {
        throw new Error('Authenticated user context is required for Teams archive access');
      }

      logger.debug('[teams_archive_search] Executing action', {
        userId: user?.id || user?._id?.toString?.(),
        action,
        chatId,
        topic,
        query,
      });

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
            query,
            topic,
            limit,
            offset,
            daysBack,
            chatType,
            participants,
          }),
        );
      }

      if (action === 'conversation_dossier') {
        return formatJsonResult(
          await TeamsArchiveService.getConversationDossier(user, {
            chatId,
            query,
            topic,
            limit,
            daysBack,
            chatType,
            participants,
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

      if (action === 'get_messages_window') {
        return formatJsonResult(
          await TeamsArchiveService.getMessagesWindow(user, {
            chatId,
            aroundMessageId,
            query,
            before,
            after,
            limit,
          }),
        );
      }

      if (action === 'summarize_conversation') {
        return formatJsonResult(
          await TeamsArchiveService.summarizeConversation(user, {
            chatId,
            query,
            topic,
            daysBack,
            limit,
          }),
        );
      }

      throw new Error(`Unsupported Teams archive action: ${action}`);
    },
      {
        name: TEAMS_ARCHIVE_TOOL_NAME,
        description:
          'Searches and retrieves archived Microsoft Teams chats that were previously ingested into Cortex. Always provide an "action" parameter. For broad questions like what has been discussed about a topic, prefer action="advanced_search_messages". For questions about the signed-in user\'s recent messages, prefer action="recent_messages". For exact or completeness-sensitive requests like all messages from my one-on-one with a person, first use action="conversation_dossier" with participants and chatType="oneOnOne". Only fall back to list_conversations if the exact chat is ambiguous and you need a candidate list. Use action="search_messages" only for quick previews. Avoid running multiple broad searches in one turn, and avoid loading whole threads unless the user explicitly asks for that level of detail.',
        schema: teamsArchiveJsonSchema,
      },
  );
}

module.exports = {
  TEAMS_ARCHIVE_TOOL_NAME,
  createTeamsArchiveTool,
  teamsArchiveJsonSchema,
};
