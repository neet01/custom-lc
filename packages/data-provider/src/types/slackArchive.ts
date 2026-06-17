export type SlackArchiveSyncStatus = 'running' | 'success' | 'partial' | 'failure' | 'cancelled';

export type SlackArchiveLatestSync = {
  id: string;
  status: SlackArchiveSyncStatus;
  mode?: string;
  phase?: string | null;
  checkpoint?: Record<string, unknown>;
  stats?: Record<string, unknown>;
  requestedConversationLimit?: number;
  requestedMessagesPerConversation?: number;
  discoveredConversationCount?: number;
  processedConversationCount?: number;
  skippedConversationCount?: number;
  conversationCount: number;
  messageCount: number;
  startedAt?: string | null;
  completedAt?: string | null;
  errorMessage?: string;
};

export type SlackArchiveOAuthStatus = {
  installConfigured: boolean;
  redirectUri: string;
  connected: boolean;
  teamId: string | null;
  enterpriseId: string | null;
};

export type SlackArchiveStatusResponse = {
  enabled: boolean;
  apiBaseUrl: string;
  oauth: SlackArchiveOAuthStatus;
  userScopes: string;
  botScopes: string;
  syncModes: string[];
  conversationTypes: string[];
  threadSupport: boolean;
  conversationCount: number;
  messageCount: number;
  latestSync: SlackArchiveLatestSync | null;
};

export type SlackArchiveSyncRequest = {
  conversationLimit?: number;
  messagesPerConversation?: number;
  async?: boolean;
};

export type SlackArchiveSyncResponse = {
  syncJob: SlackArchiveLatestSync;
  nextRequirements?: string[];
};

export type SlackArchiveCancelResponse = {
  cancelled: boolean;
  status: 'cancelled' | 'idle';
  syncJob?: SlackArchiveLatestSync | null;
  message: string;
};

export type SlackArchiveResetResponse = {
  deleted: {
    conversations: number;
    messages: number;
    syncJobs: number;
    syncLeases: number;
  };
  message: string;
};
