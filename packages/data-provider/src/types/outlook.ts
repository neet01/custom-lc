export type OutlookEmailAddress = {
  name: string;
  address: string;
};

export type OutlookMessage = {
  id: string;
  conversationId?: string;
  subject: string;
  from: OutlookEmailAddress;
  toRecipients?: OutlookEmailAddress[];
  ccRecipients?: OutlookEmailAddress[];
  receivedDateTime?: string;
  sentDateTime?: string;
  bodyPreview: string;
  body?: string;
  importance: 'low' | 'normal' | 'high' | string;
  inferenceClassification?: 'focused' | 'other' | string;
  isRead: boolean;
  hasAttachments: boolean;
  webLink?: string;
  thread?: OutlookMessage[];
  threadMessageCount?: number;
};

export type OutlookStatusResponse = {
  enabled: boolean;
  connected: boolean;
  graphBaseUrl: string;
  scopes: string;
  requires: {
    openid: boolean;
    openidReuseTokens: boolean;
    delegatedGraphScopes: string;
  };
  calendarContextEnabled?: boolean;
  userContextEnabled?: boolean;
  directoryContextEnabled?: boolean;
  mailboxSettingsContextEnabled?: boolean;
  meetingSchedulingEnabled?: boolean;
};

export type OutlookMessagesParams = {
  folder?: 'inbox' | 'drafts' | 'sent' | 'sentitems' | 'all';
  inboxView?: 'focused' | 'other' | 'all';
  limit?: number;
};

export type OutlookMessagesResponse = {
  messages: OutlookMessage[];
};

export type OutlookInsights = {
  mode: 'local-extractive' | string;
  summary: string;
  suggestedActions: string[];
  riskSignals: string[];
  calendarSignals?: string[];
  identitySignals?: string[];
  generatedAt: string;
};

export type OutlookAnalyzeResponse = {
  messageId: string;
  insights: OutlookInsights;
};

export type OutlookDraftRequest = {
  instructions?: string;
  tone?: 'professional' | 'concise' | string;
};

export type OutlookDraftResponse = {
  sourceMessageId: string;
  draftId?: string;
  subject?: string;
  bodyPreview?: string;
  webLink?: string;
  message: string;
};

export type OutlookDeleteResponse = {
  messageId: string;
  message: string;
};

export type OutlookDateTimeTimeZone = {
  dateTime: string;
  timeZone: string;
};

export type OutlookMeetingAttendee = {
  name: string;
  address: string;
};

export type OutlookMeetingSlot = {
  id: string;
  confidence?: number;
  organizerAvailability?: string;
  suggestionReason?: string;
  attendeeAvailability?: unknown[];
  start: OutlookDateTimeTimeZone;
  end: OutlookDateTimeTimeZone;
};

export type OutlookMeetingSlotsRequest = {
  durationMinutes?: number;
  days?: number;
  maxCandidates?: number;
  subject?: string;
  attendees?: OutlookMeetingAttendee[];
};

export type OutlookMeetingSlotsResponse = {
  messageId: string;
  subject: string;
  attendees: OutlookMeetingAttendee[];
  durationMinutes: number;
  emptySuggestionsReason?: string;
  suggestions: OutlookMeetingSlot[];
};

export type OutlookCreateMeetingRequest = {
  slot: {
    start: OutlookDateTimeTimeZone;
    end: OutlookDateTimeTimeZone;
  };
  subject?: string;
  attendees?: OutlookMeetingAttendee[];
  instructions?: string;
  createReplyDraft?: boolean;
  sendInvites?: boolean;
};

export type OutlookCreateMeetingResponse = {
  sourceMessageId: string;
  event: {
    id?: string;
    subject: string;
    start?: OutlookDateTimeTimeZone;
    end?: OutlookDateTimeTimeZone;
    webLink?: string;
    onlineMeeting?: {
      joinUrl?: string;
      conferenceId?: string;
    };
  };
  attendees: OutlookMeetingAttendee[];
  draft?: {
    id?: string;
    subject?: string;
    webLink?: string;
  };
  message: string;
};
