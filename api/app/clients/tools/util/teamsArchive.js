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
        'recent_meeting_chats',
        'list_conversations',
        'conversation_dossier',
        'get_messages',
        'conversation_recent_messages',
        'conversation_sender_messages',
        'conversation_activity_diagnostics',
        'sender_identity_report',
        'get_message_body',
        'get_messages_window',
        'summarize_conversation',
      ],
      description:
        'Use status to check archive readiness, sync_archive to ingest Teams chat history, recent_meeting_chats for recent meeting chat requests, conversation_recent_messages for what is new/latest in one selected chat, conversation_sender_messages for messages from me/a person in one chat, conversation_activity_diagnostics to explain recency/searchability, sender_identity_report when sender matching is uncertain, search_messages for broad lexical discovery, advanced_search_messages for recall-safe union retrieval, recent_messages for messages sent by the signed-in user, list_conversations for disambiguation, conversation_dossier for exhaustive archive-backed retrieval, get_message_body for exact full text, get_messages_window for local context, and summarize_conversation for evidence-labeled summaries.',
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
        'Stable Teams conversation identity. For single-conversation actions pass selectedConversation.graphChatId from the previous result whenever available. Do not rediscover recurring meetings by title if a selectedConversation exists.',
    },
    priorGraphChatId: {
      type: 'string',
      description:
        'Optional prior selectedConversation.graphChatId from earlier Teams tool output. Always pass this for follow-up questions so same-title recurring meetings do not switch silently.',
    },
    priorTopic: {
      type: 'string',
      description:
        'Optional prior selectedConversation.topic from earlier Teams tool output. Use with priorGraphChatId to detect same-title recurring meeting switches.',
    },
    messageId: {
      type: 'string',
      description:
        'For get_message_body: the archived message id or Teams graph message id whose full body text should be returned.',
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
      enum: ['any', 'me', 'others', 'person', 'all'],
      description:
        'For advanced_search_messages use any/me/others. For conversation_sender_messages and sender_identity_report use me/person/all as supported by the action.',
    },
    senderName: {
      type: 'string',
      description:
        'For conversation_sender_messages or sender_identity_report with senderScope=person: sender display name to match.',
    },
    senderEmail: {
      type: 'string',
      description:
        'For conversation_sender_messages or sender_identity_report with senderScope=person: sender email to match.',
    },
    senderUserId: {
      type: 'string',
      description:
        'For conversation_sender_messages or sender_identity_report with senderScope=person: Teams/Entra sender user id to match.',
    },
    personName: {
      type: 'string',
      description:
        'For sender_identity_report with senderScope=person: person display name to inspect.',
    },
    personEmail: {
      type: 'string',
      description:
        'For sender_identity_report with senderScope=person: person email to inspect.',
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
    includeSystem: {
      type: 'boolean',
      description:
        'For conversation_recent_messages: include Teams system/empty activity. Defaults to false so latest/new means human-readable messages.',
    },
    includeRecentMessages: {
      type: 'boolean',
      description:
        'For conversation_activity_diagnostics: include newest human-readable message previews in the diagnostic output.',
    },
    sort: {
      type: 'string',
      enum: ['newest', 'oldest'],
      description:
        'For conversation_sender_messages: newest or oldest message ordering.',
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
  return JSON.stringify(result);
}

function stableStringify(value) {
  if (Array.isArray(value)) {
    return `[${value.map((entry) => stableStringify(entry)).join(',')}]`;
  }

  if (value && typeof value === 'object') {
    const entries = Object.entries(value)
      .filter(([, entryValue]) => entryValue !== undefined)
      .sort(([a], [b]) => a.localeCompare(b));
    return `{${entries
      .map(([key, entryValue]) => `${JSON.stringify(key)}:${stableStringify(entryValue)}`)
      .join(',')}}`;
  }

  return JSON.stringify(value);
}

function buildActionSignature(action, payload) {
  return `${action}:${stableStringify(payload)}`;
}

function clampActionLimit(action, limit) {
  const parsed = Number(limit);
  if (!Number.isFinite(parsed)) {
    switch (action) {
      case 'list_conversations':
      case 'recent_meeting_chats':
        return 3;
      case 'conversation_dossier':
        return 4;
      case 'get_messages':
      case 'conversation_recent_messages':
      case 'conversation_sender_messages':
      case 'conversation_activity_diagnostics':
      case 'sender_identity_report':
      case 'get_message_body':
      case 'get_messages_window':
      case 'summarize_conversation':
        return 6;
      case 'search_messages':
      case 'advanced_search_messages':
      case 'recent_messages':
        return 4;
      default:
        return undefined;
    }
  }

  const normalized = Math.max(1, Math.trunc(parsed));
  switch (action) {
    case 'list_conversations':
    case 'recent_meeting_chats':
      return Math.min(normalized, 5);
    case 'conversation_dossier':
      return Math.min(normalized, 6);
    case 'get_messages':
    case 'conversation_recent_messages':
    case 'conversation_sender_messages':
    case 'conversation_activity_diagnostics':
    case 'sender_identity_report':
    case 'get_message_body':
    case 'get_messages_window':
    case 'summarize_conversation':
      return Math.min(normalized, 8);
    case 'search_messages':
    case 'advanced_search_messages':
    case 'recent_messages':
      return Math.min(normalized, 6);
    default:
      return normalized;
  }
}

function createTeamsArchiveTool({ req }) {
  const actionCounts = new Map();
  const duplicateActionCounts = new Map();

  return tool(
    async ({
      action,
      query,
      topic,
      chatId,
      messageId,
      limit,
      offset,
      chatLimit,
      messagesPerChat,
      daysBack,
      senderScope,
      senderName,
      senderEmail,
      senderUserId,
      personName,
      personEmail,
      chatType,
      participants,
      sortBy,
      aroundMessageId,
      before,
      after,
      includeSystem,
      includeRecentMessages,
      sort,
      priorGraphChatId,
      priorTopic,
    }) => {
      const user = req?.user;

      if (!user) {
        throw new Error('Authenticated user context is required for Teams archive access');
      }

      const resolvedLimit = clampActionLimit(action, limit);
      const actionPayload = {
        action,
        query,
        topic,
        chatId,
        messageId,
        limit: resolvedLimit,
        offset,
        chatLimit,
        messagesPerChat,
        daysBack,
        senderScope,
        senderName,
        senderEmail,
        senderUserId,
        personName,
        personEmail,
        chatType,
        participants,
        sortBy,
        aroundMessageId,
        before,
        after,
        includeSystem,
        includeRecentMessages,
        sort,
        priorGraphChatId,
        priorTopic,
      };
      const actionSignature = buildActionSignature(action, actionPayload);
      actionCounts.set(action, (actionCounts.get(action) || 0) + 1);
      const duplicateCount = duplicateActionCounts.get(actionSignature) || 0;

      if (duplicateCount > 0) {
        logger.warn('[teams_archive_search] Suppressing duplicate tool call in same run', {
          userId: user?.id || user?._id?.toString?.(),
          action,
          duplicateCount: duplicateCount + 1,
        });
        return formatJsonResult({
          retrievalMode: 'duplicate_call_suppressed',
          action,
          guidance:
            'This exact Teams archive tool request was already executed in the current run. Reuse the earlier result instead of repeating the same call.',
        });
      }

      duplicateActionCounts.set(actionSignature, duplicateCount + 1);

      logger.debug('[teams_archive_search] Executing action', {
        userId: user?.id || user?._id?.toString?.(),
        action,
        chatId,
        topic,
        query,
        limit: resolvedLimit,
        invocationCount: actionCounts.get(action),
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
            limit: resolvedLimit,
            offset,
          }),
        );
      }

      if (action === 'advanced_search_messages') {
        return formatJsonResult(
          await TeamsArchiveService.advancedSearchMessages(user, {
            query,
            topic,
            chatId,
            limit: resolvedLimit,
            offset,
            daysBack,
            senderScope,
            chatType,
            participants,
            sortBy,
            priorGraphChatId,
            priorTopic,
          }),
        );
      }

      if (action === 'recent_messages') {
        return formatJsonResult(
          await TeamsArchiveService.recentMessages(user, {
            query,
            limit: resolvedLimit,
            daysBack,
          }),
        );
      }

      if (action === 'recent_meeting_chats') {
        return formatJsonResult(
          await TeamsArchiveService.recentMeetingChats(user, {
            query,
            topic,
            limit: resolvedLimit,
            offset,
            daysBack,
          }),
        );
      }

      if (action === 'list_conversations') {
        return formatJsonResult(
          await TeamsArchiveService.listConversations(user, {
            query,
            topic,
            limit: resolvedLimit,
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
            limit: resolvedLimit,
            daysBack,
            chatType,
            participants,
            priorGraphChatId,
            priorTopic,
          }),
        );
      }

      if (action === 'get_messages') {
        return formatJsonResult(
          await TeamsArchiveService.listConversationMessages(user, chatId, {
            limit: resolvedLimit,
            offset,
          }),
        );
      }

      if (action === 'conversation_recent_messages') {
        return formatJsonResult(
          await TeamsArchiveService.conversationRecentMessages(user, {
            chatId,
            limit: resolvedLimit,
            offset,
            daysBack,
            includeSystem,
          }),
        );
      }

      if (action === 'conversation_sender_messages') {
        return formatJsonResult(
          await TeamsArchiveService.conversationSenderMessages(user, {
            chatId,
            senderScope,
            senderName,
            senderEmail,
            senderUserId,
            limit: resolvedLimit,
            offset,
            includeSystem,
            sort,
            priorGraphChatId,
            priorTopic,
          }),
        );
      }

      if (action === 'conversation_activity_diagnostics') {
        return formatJsonResult(
          await TeamsArchiveService.conversationActivityDiagnostics(user, {
            chatId,
            includeRecentMessages,
            includeSystem,
            limit: resolvedLimit,
            priorGraphChatId,
            priorTopic,
          }),
        );
      }

      if (action === 'sender_identity_report') {
        return formatJsonResult(
          await TeamsArchiveService.senderIdentityReport(user, {
            chatId,
            senderScope,
            senderName,
            senderEmail,
            senderUserId,
            personName,
            personEmail,
            priorGraphChatId,
            priorTopic,
          }),
        );
      }

      if (action === 'get_message_body') {
        return formatJsonResult(
          await TeamsArchiveService.getMessageBody(user, {
            chatId,
            messageId,
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
            limit: resolvedLimit,
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
            limit: resolvedLimit,
            priorGraphChatId,
            priorTopic,
          }),
        );
      }

      throw new Error(`Unsupported Teams archive action: ${action}`);
    },
      {
        name: TEAMS_ARCHIVE_TOOL_NAME,
        description:
          'Searches and retrieves archived Microsoft Teams chats. Always pass priorGraphChatId from selectedConversation for follow-up questions. Use action="recent_meeting_chats" for recent meeting/standup requests because it ranks by lastMeaningfulMessageAt instead of Teams system activity. Use action="conversation_recent_messages" with selectedConversation.graphChatId for follow-ups like what is new/latest in that chat. Use action="conversation_sender_messages" for messages from me/a person in a selected chat, and inspect zeroResultDiagnostics before saying no messages exist. Use action="conversation_activity_diagnostics" to explain why a chat appears recent or risky. Use action="sender_identity_report" when sender matching is uncertain. Respect evidenceBudget and do not answer definitively when evidenceSufficient=false. Use graphChatId as the stable identity for recurring meetings; never treat title/topic alone as unique. If identityChanged=true or identityWarning is returned, ask for clarification.',
        schema: teamsArchiveJsonSchema,
      },
  );
}

module.exports = {
  TEAMS_ARCHIVE_TOOL_NAME,
  createTeamsArchiveTool,
  teamsArchiveJsonSchema,
};
