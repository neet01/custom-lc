import type { TeamsArchiveSyncStatus } from '~/schema/teamsArchiveSyncJob';

export type { TeamsArchiveSyncStatus };

export interface TeamsArchiveParticipantData {
  displayName?: string;
  email?: string;
  userId?: string;
}

export interface TeamsArchiveConversationData {
  user: string;
  graphChatId: string;
  chatType?: string;
  topic?: string;
  webUrl?: string;
  participants?: TeamsArchiveParticipantData[];
  lastMessageAt?: Date;
  lastSyncedAt?: Date;
  sourceUpdatedAt?: Date;
  messageCount?: number;
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
}

export interface TeamsArchiveSyncJobData {
  user: string;
  status: TeamsArchiveSyncStatus;
  mode?: string;
  conversationCount?: number;
  messageCount?: number;
  errorMessage?: string;
  startedAt?: Date;
  completedAt?: Date;
}
