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
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  bodyPreview: string;
  body?: string;
  bodyHtml?: string;
  bodyContentType?: 'text' | 'html' | string;
  importance: 'low' | 'normal' | 'high' | string;
  inferenceClassification?: 'focused' | 'other' | string;
  isRead: boolean;
  isDraft?: boolean;
  hasAttachments: boolean;
  webLink?: string;
  thread?: OutlookMessage[];
  threadMessageCount?: number;
  draftReplies?: OutlookMessage[];
  draftReplyCount?: number;
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
  search?: string;
};

export type OutlookMessagesResponse = {
  messages: OutlookMessage[];
  search?: string;
};

export type OutlookCalendarAttendee = {
  name: string;
  address: string;
  response?: 'none' | 'organizer' | 'tentativelyAccepted' | 'accepted' | 'declined' | string;
};

export type OutlookCalendarEvent = {
  id: string;
  subject: string;
  start?: OutlookDateTimeTimeZone;
  end?: OutlookDateTimeTimeZone;
  location?: string;
  organizer?: OutlookEmailAddress;
  showAs?: string;
  isAllDay?: boolean;
  isOnlineMeeting?: boolean;
  webLink?: string;
  bodyPreview?: string;
  type?: string;
  attendees?: OutlookCalendarAttendee[];
};

export type OutlookCalendarParams = {
  startDateTime?: string;
  endDateTime?: string;
  view?: 'day' | 'week' | 'agenda';
  limit?: number;
};

export type OutlookCalendarResponse = {
  startDateTime: string;
  endDateTime: string;
  view: 'day' | 'week' | 'agenda';
  events: OutlookCalendarEvent[];
  workingHours?: {
    daysOfWeek: string[];
    startTime: string;
    endTime: string;
    timeZone: string;
  };
};

export type OutlookCalendarEventMutationRequest = {
  subject: string;
  start: OutlookDateTimeTimeZone;
  end: OutlookDateTimeTimeZone;
  location?: string;
  attendees?: OutlookMeetingAttendee[];
  isOnlineMeeting?: boolean;
  body?: string;
};

export type OutlookCalendarEventMutationResponse = {
  message: string;
  event: OutlookCalendarEvent & {
    onlineMeeting?: {
      joinUrl?: string;
      conferenceId?: string;
    };
  };
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

export type OutlookBrief = {
  mode: 'local-extractive' | string;
  headline: string;
  summary: string;
  priorities: string[];
  followUps: string[];
  meetingHighlights: string[];
  notableEmails: string[];
  risks: string[];
  generatedAt: string;
};

export type OutlookAnalyzeSelectionRequest = {
  messageIds: string[];
};

export type OutlookAnalyzeSelectionResponse = {
  messageIds: string[];
  conversationIds: string[];
  messageCount: number;
  brief: OutlookBrief;
};

export type OutlookDailyBriefResponse = {
  windowStart: string;
  windowEnd: string;
  emailCount: number;
  meetingCount: number;
  messageIds: string[];
  brief: OutlookBrief;
};

export type OutlookDraftRequest = {
  instructions?: string;
  tone?: 'professional' | 'concise' | string;
  replyMode?: 'smart' | 'reply' | 'reply_all';
  replyAll?: boolean;
};

export type OutlookDraftResponse = {
  sourceMessageId: string;
  conversationId?: string;
  draftId?: string;
  subject?: string;
  bodyPreview?: string;
  webLink?: string;
  replyMode?: 'reply' | 'reply_all' | string;
  toRecipients?: OutlookEmailAddress[];
  ccRecipients?: OutlookEmailAddress[];
  message: string;
};

export type OutlookDeleteResponse = {
  messageId: string;
  message: string;
};

export type OutlookUpdateReadStateRequest = {
  isRead: boolean;
};

export type OutlookUpdateReadStateResponse = {
  messageId: string;
  isRead: boolean;
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
  confidenceReason?: string;
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
  schedulingAttendees?: OutlookMeetingAttendee[];
  externalAttendeesExcluded?: OutlookMeetingAttendee[];
  externalAttendeesWithThreadAvailability?: OutlookMeetingAttendee[];
  availabilityNotes?: string[];
  durationMinutes: number;
  workingHours?: {
    daysOfWeek: string[];
    startTime: string;
    endTime: string;
    timeZone: string;
  };
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
  meetingNotePreview?: string;
  meetingDraft?: {
    id?: string;
    subject?: string;
    webLink?: string;
  };
  draft?: {
    id?: string;
    subject?: string;
    webLink?: string;
  };
  message: string;
};
