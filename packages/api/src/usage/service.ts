import { logger } from '@librechat/data-schemas';
import type { UsageRecordData, UsageSource } from '@librechat/data-schemas';
import { isEnabled } from '~/utils/common';

export const USAGE_TRACKING_ENABLED = 'USAGE_TRACKING_ENABLED';

export interface UsageRecordInput {
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
  inputTokens?: number;
  outputTokens?: number;
  cacheCreationTokens?: number;
  cacheReadTokens?: number;
  latencyMs?: number;
}

export interface UsagePersistenceDeps {
  createUsageRecords: (records: UsageRecordData[]) => Promise<unknown>;
}

export function createUsageRecord(input: UsageRecordInput): UsageRecordData {
  const cacheCreationTokens = Math.max(input.cacheCreationTokens ?? 0, 0);
  const cacheReadTokens = Math.max(input.cacheReadTokens ?? 0, 0);
  const inputTokens = Math.max(input.inputTokens ?? 0, 0) + cacheCreationTokens + cacheReadTokens;
  const outputTokens = Math.max(input.outputTokens ?? 0, 0);

  return {
    user: input.user,
    conversationId: input.conversationId,
    messageId: input.messageId,
    requestId: input.requestId ?? input.messageId,
    sessionId: input.sessionId,
    model: input.model,
    provider: input.provider,
    endpoint: input.endpoint,
    context: input.context ?? 'message',
    source: input.source ?? 'system',
    inputTokens,
    outputTokens,
    totalTokens: inputTokens + outputTokens,
    cacheCreationTokens: cacheCreationTokens || undefined,
    cacheReadTokens: cacheReadTokens || undefined,
    latencyMs: input.latencyMs != null ? Math.max(input.latencyMs, 0) : undefined,
  };
}

export async function persistUsageRecords(
  deps: UsagePersistenceDeps,
  records: UsageRecordInput[],
  featureFlag: string | boolean | undefined = process.env[USAGE_TRACKING_ENABLED],
): Promise<void> {
  if (!isEnabled(featureFlag) || records.length === 0) {
    return;
  }

  try {
    await deps.createUsageRecords(records.map(createUsageRecord));
  } catch (error) {
    logger.error('[persistUsageRecords] Error persisting usage records', error);
  }
}
