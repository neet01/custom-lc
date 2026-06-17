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
  identityLinked?: boolean;
  teamId: string | null;
  enterpriseId: string | null;
};

export type SlackArchiveStatusResponse = {
  enabled: boolean;
  apiBaseUrl: string;
  oauth: SlackArchiveOAuthStatus;
  userScopes: string;
  botScopes: string;
  maxConcurrentSyncs?: number;
  activeSyncs?: number;
  syncModes: string[];
  conversationTypes: string[];
  threadSupport: boolean;
  conversationCount: number;
  messageCount: number;
  latestSync: SlackArchiveLatestSync | null;
  latestProjection?: {
    id: string | null;
    status?: string;
    jobType?: string | null;
    sourceRecordType?: string | null;
    sourceRecordId?: string | null;
    stats?: Record<string, unknown>;
    errorMessage?: string;
    startedAt?: string | null;
    completedAt?: string | null;
  } | null;
  projectionChunkCount?: number;
  projectionConversationCount?: number;
  projectionEntityCount?: number;
};

export type SlackArchiveSyncRequest = {
  conversationLimit?: number;
  messagesPerConversation?: number;
  async?: boolean;
};

export type SlackArchiveSyncResponse = {
  syncJob: SlackArchiveLatestSync;
  mode?: string;
  conversationCount?: number;
  messageCount?: number;
  skippedConversationCount?: number;
  processedConversationCount?: number;
  memoryProjection?: {
    status: string;
    queuedAt?: string;
    requestedConversationCount?: number;
    reason?: string;
  };
  conversations?: Array<Record<string, unknown>>;
  nextRequirements?: string[];
};

export type SlackArchiveSyncAcceptedResponse = {
  accepted: true;
  status: 'running';
  mode: string;
  message: string;
  syncJob: SlackArchiveLatestSync;
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
    projectionJobs?: number;
    chunks?: number;
    entities?: number;
    relationships?: number;
  };
  message: string;
};
