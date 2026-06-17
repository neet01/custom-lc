import type { SlackArchiveSyncStatus } from '~/schema/slackArchiveSyncJob';

export type { SlackArchiveSyncStatus };

export interface SlackArchiveParticipantData {
  slackUserId?: string;
  displayName?: string;
  realName?: string;
  username?: string;
  email?: string;
  isBot?: boolean;
  isAppUser?: boolean;
}

export interface SlackArchiveConversationData {
  user: string;
  slackConversationId: string;
  teamId?: string;
  enterpriseId?: string;
  conversationType?: 'public_channel' | 'private_channel' | 'im' | 'mpim';
  name?: string;
  topic?: string;
  purpose?: string;
  isArchived?: boolean;
  isShared?: boolean;
  isExtShared?: boolean;
  isOrgShared?: boolean;
  isSlackConnect?: boolean;
  participants?: SlackArchiveParticipantData[];
  syncStatus?: 'pending' | 'running' | 'complete' | 'failed';
  syncCursor?: string;
  syncError?: string;
  syncAttemptCount?: number;
  sourceDiscoveredAt?: Date;
  sourceLastMessageAt?: Date;
  syncStartedAt?: Date;
  syncCompletedAt?: Date;
  lastMessageSyncAt?: Date;
  lastMessageAt?: Date;
  lastHumanMessageAt?: Date;
  lastMeaningfulMessageAt?: Date;
  lastSystemMessageAt?: Date;
  lastSyncedAt?: Date;
  sourceUpdatedAt?: Date;
  messageCount?: number;
  humanMessageCount?: number;
  systemMessageCount?: number;
  meaningfulMessageCount?: number;
}

export interface SlackArchiveReactionData {
  name?: string;
  count?: number;
  users?: string[];
}

export interface SlackArchiveMentionData {
  slackUserId?: string;
  displayName?: string;
}

export interface SlackArchiveAttachmentData {
  id?: string;
  title?: string;
  text?: string;
  fallback?: string;
  serviceName?: string;
  titleLink?: string;
}

export interface SlackArchiveFileData {
  id?: string;
  name?: string;
  title?: string;
  mimetype?: string;
  filetype?: string;
  prettyType?: string;
  urlPrivate?: string;
}

export interface SlackArchiveMessageData {
  user: string;
  slackConversationId: string;
  slackMessageTs: string;
  teamId?: string;
  enterpriseId?: string;
  slackUserId?: string;
  botId?: string;
  username?: string;
  displayName?: string;
  subtype?: string;
  text?: string;
  normalizedText?: string;
  threadTs?: string;
  parentUserId?: string;
  replyCount?: number;
  replyUsers?: string[];
  latestReplyTs?: string;
  reactions?: SlackArchiveReactionData[];
  mentions?: SlackArchiveMentionData[];
  attachments?: SlackArchiveAttachmentData[];
  files?: SlackArchiveFileData[];
  raw?: Record<string, unknown>;
  sentAt?: Date;
  editedAt?: Date;
  deletedAt?: Date;
  normalizedTextLength?: number;
  isSystemLikeMessage?: boolean;
  isChunkable?: boolean;
  skipChunkReason?: string;
}

export interface SlackArchiveSyncJobData {
  user: string;
  status: SlackArchiveSyncStatus;
  mode?: string;
  phase?: string;
  checkpoint?: Record<string, unknown>;
  stats?: Record<string, unknown>;
  requestedConversationLimit?: number;
  requestedMessagesPerConversation?: number;
  discoveredConversationCount?: number;
  processedConversationCount?: number;
  skippedConversationCount?: number;
  conversationCount?: number;
  messageCount?: number;
  errorMessage?: string;
  startedAt?: Date;
  completedAt?: Date;
}
