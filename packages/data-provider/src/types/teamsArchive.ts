export type TeamsArchiveSyncStatus = 'running' | 'success' | 'failure' | 'cancelled';

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

export type TeamsArchiveLatestProjection = {
  id: string;
  status: 'pending' | 'running' | 'success' | 'failure';
  startedAt?: string;
  completedAt?: string;
  errorMessage?: string;
  stats?: Record<string, unknown>;
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
  latestProjection: TeamsArchiveLatestProjection | null;
};

export type TeamsArchiveSyncRequest = {
  mode?: 'chats';
  chatLimit?: number;
  messagesPerChat?: number;
  async?: boolean;
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

export type TeamsArchiveSyncAcceptedResponse = {
  accepted: true;
  status: 'running';
  mode: string;
  message: string;
};

export type TeamsArchiveCancelResponse = {
  cancelled: boolean;
  status: 'cancelled' | 'idle';
  syncJob?: {
    _id?: string;
    id?: string;
    status?: TeamsArchiveSyncStatus;
    mode?: string;
    conversationCount?: number;
    messageCount?: number;
    startedAt?: string;
    completedAt?: string;
    errorMessage?: string;
  } | null;
  message: string;
};
