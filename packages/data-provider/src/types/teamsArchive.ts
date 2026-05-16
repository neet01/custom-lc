export type TeamsArchiveSyncStatus = 'running' | 'success' | 'failure';

export type TeamsArchiveLatestSync = {
  id: string;
  status: TeamsArchiveSyncStatus;
  mode?: string;
  conversationCount: number;
  messageCount: number;
  startedAt?: string;
  completedAt?: string;
  errorMessage?: string;
};

export type TeamsArchiveStatusResponse = {
  enabled: boolean;
  graphBaseUrl: string;
  graphScopes: string;
  syncModes: string[];
  channelSyncSupported: boolean;
  conversationCount: number;
  messageCount: number;
  latestSync: TeamsArchiveLatestSync | null;
};

export type TeamsArchiveSyncRequest = {
  mode?: 'chats';
  chatLimit?: number;
  messagesPerChat?: number;
};

export type TeamsArchiveSyncConversation = {
  id: string;
  graphChatId: string;
  topic: string;
  chatType: string;
  messageCount: number;
  lastMessageAt?: string;
};

export type TeamsArchiveSyncResponse = {
  syncJob: {
    _id?: string;
    id?: string;
    status?: TeamsArchiveSyncStatus;
    mode?: string;
    conversationCount?: number;
    messageCount?: number;
    startedAt?: string;
    completedAt?: string;
    errorMessage?: string;
  };
  mode: string;
  conversationCount: number;
  messageCount: number;
  conversations: TeamsArchiveSyncConversation[];
  memoryProjection?:
    | {
        status: 'success';
        jobId?: string;
        projectedConversationCount?: number;
        entityCount?: number;
        relationshipCount?: number;
        chunkCount?: number;
      }
    | {
        status: 'failure';
        errorMessage?: string;
      }
    | {
        status: 'skipped';
        reason?: string;
      }
    | null;
};
