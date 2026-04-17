export type UsageSource = 'agent' | 'assistant' | 'tool' | 'system';

export interface UsageRecordData {
  user: string;
  conversationId: string;
  messageId?: string;
  requestId?: string;
  sessionId?: string;
  model?: string;
  provider?: string;
  endpoint?: string;
  context?: string;
  source?: UsageSource;
  inputTokens: number;
  outputTokens: number;
  totalTokens: number;
  cacheCreationTokens?: number;
  cacheReadTokens?: number;
  latencyMs?: number;
}
