import type { TeamsArchiveSyncStatus } from '~/schema/teamsArchiveSyncJob';

export type { TeamsArchiveSyncStatus };

export interface TeamsArchiveParticipantData {
  displayName?: string;
  email?: string;
  userId?: string;
  source?: 'graph' | 'inferred_from_messages' | 'inferred_from_mentions' | 'mixed' | 'unknown';
  confidence?: 'high' | 'medium' | 'low';
}

export interface TeamsArchiveConversationData {
  user: string;
  graphChatId: string;
  chatType?: string;
  topic?: string;
  webUrl?: string;
  participants?: TeamsArchiveParticipantData[];
  syncStatus?: 'pending' | 'running' | 'complete' | 'failed' | 'deferred_failed';
  syncCursor?: string;
  syncError?: string;
  syncAttemptCount?: number;
  syncDeferredAt?: Date;
  syncDeferredReason?: string;
  syncDeferredStatus?: number;
  syncNeedsIntervention?: boolean;
  nextRetryAt?: Date;
  lastErrorStatus?: number;
  sourceDiscoveredAt?: Date;
  sourceLastMessageAt?: Date;
  syncStartedAt?: Date;
  syncCompletedAt?: Date;
  lastMessageSyncAt?: Date;
  lastMessageAt?: Date;
  lastSyncedAt?: Date;
  sourceUpdatedAt?: Date;
  messageCount?: number;
  participantMetadataSource?: 'graph' | 'inferred_from_messages' | 'inferred_from_mentions' | 'mixed' | 'unknown';
  participantConfidence?: 'high' | 'medium' | 'low';
  participantDegraded?: boolean;
  participantStats?: Record<string, unknown>;
}

export interface TeamsArchiveAttachmentData {
  id?: string;
  name?: string;
  contentType?: string;
  contentUrl?: string;
}

export interface TeamsArchiveMentionData {
  id?: string;
  displayName?: string;
  mentionedUserId?: string;
}

export interface TeamsArchiveMessageData {
  user: string;
  graphChatId: string;
  graphMessageId: string;
  replyToId?: string;
  fromDisplayName?: string;
  fromEmail?: string;
  fromUserId?: string;
  subject?: string;
  summary?: string;
  importance?: string;
  messageType?: string;
  bodyContentType?: string;
  bodyPreview?: string;
  bodyContent?: string;
  bodyText?: string;
  attachments?: TeamsArchiveAttachmentData[];
  mentions?: TeamsArchiveMentionData[];
  webUrl?: string;
  sentDateTime?: Date;
  lastModifiedDateTime?: Date;
  deletedDateTime?: Date;
  etag?: string;
  normalizedTextLength?: number;
  isSystemLikeMessage?: boolean;
  isChunkable?: boolean;
  skipChunkReason?: string;
}

export interface TeamsArchiveSyncJobData {
  user: string;
  status: TeamsArchiveSyncStatus;
  mode?: string;
  phase?: string;
  checkpoint?: Record<string, unknown>;
  stats?: Record<string, unknown>;
  requestedChatLimit?: number;
  requestedMessagesPerChat?: number;
  discoveredChatCount?: number;
  processedChatCount?: number;
  skippedChatCount?: number;
  projectionJobId?: string;
  conversationCount?: number;
  messageCount?: number;
  errorMessage?: string;
  startedAt?: Date;
  completedAt?: Date;
}

export interface TeamsArchiveBackfillStateData {
  user: string;
  status: 'idle' | 'discovering' | 'syncing' | 'paused' | 'complete' | 'failed';
  nextChatPageLink?: string;
  discoveryComplete?: boolean;
  discoveredChatCount?: number;
  completedChatCount?: number;
  pendingChatCount?: number;
  runningChatCount?: number;
  failedChatCount?: number;
  totalMessageCount?: number;
  lastSyncJobId?: string;
  lastProjectionJobId?: string;
  lastDiscoveredAt?: Date;
  lastCompletedAt?: Date;
  lastHeartbeatAt?: Date;
  errorMessage?: string;
  tenantId?: string;
}
