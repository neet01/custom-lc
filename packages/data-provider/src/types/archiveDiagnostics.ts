export type ArchiveDiagnosticsSource = 'slack' | 'teams';
export type ArchiveDiagnosticsSeverity = 'ok' | 'warning' | 'error';

export type ArchiveDiagnosticsParams = {
  source?: ArchiveDiagnosticsSource;
  userId?: string;
  q?: string;
  type?: string;
  status?: string;
  limit?: number;
  offset?: number;
};

export type ArchiveDiagnosticsCountItem = {
  key: string;
  count: number;
};

export type ArchiveDiagnosticsJob = {
  id: string | null;
  status: string | null;
  phase: string | null;
  jobType: string | null;
  sourceRecordType: string | null;
  errorMessage: string | null;
  stats: Record<string, unknown> | null;
  createdAt: string | null;
  startedAt: string | null;
  completedAt: string | null;
  updatedAt: string | null;
};

export type ArchiveDiagnosticsConversation = {
  id: string | null;
  userId: string;
  source: ArchiveDiagnosticsSource;
  sourceConversationId: string;
  displayName: string;
  type: string;
  syncStatus: string;
  syncError: string;
  messageCount: number;
  declaredMessageCount: number;
  meaningfulMessageCount: number;
  chunkableMessageCount: number;
  skippedMessageCount: number;
  chunkCount: number;
  messageChunkCount: number;
  windowChunkCount: number;
  lastMessageAt: string | null;
  lastMeaningfulMessageAt: string | null;
  lastMessageSyncAt: string | null;
  latestChunkAt: string | null;
  updatedAt: string | null;
  health: {
    state: string;
    severity: ArchiveDiagnosticsSeverity;
    reason: string;
  };
};

export type ArchiveDiagnosticsResponse = {
  source: ArchiveDiagnosticsSource;
  generatedAt: string;
  filters: {
    userId: string | null;
    q: string | null;
    type: string | null;
    status: string | null;
    limit: number;
    offset: number;
  };
  summary: {
    totalConversations: number;
    filteredConversations: number;
    totalMessages: number;
    totalChunks: number;
    healthyConversationCount: number;
    warningConversationCount: number;
    errorConversationCount: number;
  };
  breakdowns: {
    conversationsByType: ArchiveDiagnosticsCountItem[];
    conversationsByStatus: ArchiveDiagnosticsCountItem[];
    chunksByRecordType: ArchiveDiagnosticsCountItem[];
    chunksByChunkType: ArchiveDiagnosticsCountItem[];
    skippedMessageReasons: ArchiveDiagnosticsCountItem[];
  };
  latestSync: ArchiveDiagnosticsJob | null;
  latestProjection: ArchiveDiagnosticsJob | null;
  conversations: ArchiveDiagnosticsConversation[];
};
