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
