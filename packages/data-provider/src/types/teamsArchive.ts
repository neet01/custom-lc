export type TeamsArchiveSyncStatus = 'running' | 'success' | 'failure' | 'cancelled';

export type TeamsArchiveLatestSync = {
  id: string;
  status: TeamsArchiveSyncStatus;
  mode?: string;
  phase?: string | null;
  checkpoint?: Record<string, unknown>;
  stats?: Record<string, unknown>;
  requestedChatLimit?: number;
  requestedMessagesPerChat?: number;
  discoveredChatCount?: number;
  processedChatCount?: number;
  skippedChatCount?: number;
  projectionJobId?: string | null;
  conversationCount: number;
  messageCount: number;
  startedAt?: string;
  completedAt?: string;
  errorMessage?: string;
};

export type TeamsArchiveBackfillState = {
  status: 'idle' | 'discovering' | 'syncing' | 'paused' | 'complete' | 'failed';
  discoveryComplete: boolean;
  nextChatPageLinkPresent: boolean;
  discoveredChatCount: number;
  completedChatCount: number;
  pendingChatCount: number;
  runningChatCount: number;
  failedChatCount: number;
  totalMessageCount: number;
  lastSyncJobId?: string | null;
  lastProjectionJobId?: string | null;
  lastDiscoveredAt?: string | null;
  lastCompletedAt?: string | null;
  lastHeartbeatAt?: string | null;
  errorMessage?: string | null;
};

export type TeamsArchiveLatestProjection = {
  id: string;
  status: 'pending' | 'running' | 'success' | 'failure';
  startedAt?: string;
  completedAt?: string;
  errorMessage?: string;
  stats?: Record<string, unknown>;
};

export type TeamsArchiveProjectionCoverage = {
  indexedConversationCount: number;
  totalConversationCount: number;
  indexedChunkCount: number;
  searchableConversationCount: number;
  pendingConversationCount: number;
  fullyIndexed: boolean;
  coveragePercent: number;
};

export type TeamsArchiveStatusResponse = {
  enabled: boolean;
  graphBaseUrl: string;
  graphScopes: string;
  maxConcurrentSyncs: number;
  activeSyncs: number;
  syncModes: string[];
  channelSyncSupported: boolean;
  conversationCount: number;
  messageCount: number;
  backfillState: TeamsArchiveBackfillState | null;
  latestSync: TeamsArchiveLatestSync | null;
  latestProjection: TeamsArchiveLatestProjection | null;
  projectionCoverage: TeamsArchiveProjectionCoverage | null;
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
  syncStatus?: 'pending' | 'running' | 'complete' | 'failed';
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
  discovery?: {
    discoveredThisRun: number;
    discoveryComplete: boolean;
    nextChatPageLinkPresent: boolean;
  };
  skippedMessageChats?: number;
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

export type TeamsArchiveResetResponse = {
  deleted: {
    conversations: number;
    messages: number;
    syncJobs: number;
    syncLeases: number;
    backfillStates: number;
    projectionJobs: number;
    chunks: number;
    entities: number;
    relationships: number;
  };
  message: string;
};
