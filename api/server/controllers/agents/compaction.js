const { countFormattedMessageTokens } = require('@librechat/api');
const { logger } = require('@librechat/data-schemas');
const { ContentTypes } = require('librechat-data-provider');

const PREVIEW_CHAR_LIMIT = 220;
const PREVIEW_ENTRY_LIMIT = 20;

const COMPACTION_MODES = {
  preflight: {
    triggerTokenRatio: 0.35,
    absoluteTriggerTokens: 14000,
    triggerMessageCount: 80,
    targetTokenRatio: 0.12,
    targetTailTokenCap: 4500,
    maxFormattedTokens: 9000,
    minRecentMessages: 4,
    maxRecentMessages: 8,
    maxSummaryChars: 2200,
    retainedMessageCharLimit: 260,
    preserveRecentMessagesVerbatim: 2,
  },
  overflow: {
    triggerTokenRatio: 0,
    absoluteTriggerTokens: 0,
    triggerMessageCount: 0,
    targetTokenRatio: 0.07,
    targetTailTokenCap: 2200,
    maxFormattedTokens: 4500,
    minRecentMessages: 3,
    maxRecentMessages: 5,
    maxSummaryChars: 1400,
    retainedMessageCharLimit: 120,
    preserveRecentMessagesVerbatim: 1,
  },
  emergency: {
    triggerTokenRatio: 0,
    absoluteTriggerTokens: 0,
    triggerMessageCount: 0,
    targetTokenRatio: 0.04,
    targetTailTokenCap: 1200,
    maxFormattedTokens: 2500,
    minRecentMessages: 2,
    maxRecentMessages: 3,
    maxSummaryChars: 900,
    retainedMessageCharLimit: 80,
    preserveRecentMessagesVerbatim: 1,
  },
};

function truncateText(text, limit) {
  if (typeof text !== 'string') {
    return '';
  }
  const normalized = text.replace(/\s+/g, ' ').trim();
  if (normalized.length <= limit) {
    return normalized;
  }
  return `${normalized.slice(0, Math.max(0, limit - 1)).trimEnd()}...`;
}

function getPartText(part) {
  if (part == null || typeof part !== 'object') {
    return '';
  }

  if (typeof part.text === 'string' && part.text.trim().length > 0) {
    return part.text;
  }

  if (typeof part.type === 'string') {
    const value = part[part.type];
    if (typeof value === 'string' && value.trim().length > 0) {
      return value;
    }
  }

  return '';
}

function flattenContent(content) {
  if (typeof content === 'string') {
    return content;
  }

  if (!Array.isArray(content)) {
    return '';
  }

  const fragments = [];
  for (const part of content) {
    if (part?.type === ContentTypes.THINK || part?.type === ContentTypes.ERROR) {
      continue;
    }

    if (part?.type === ContentTypes.TOOL_CALL && part.tool_call) {
      const toolName = part.tool_call.name ? `Tool ${part.tool_call.name}` : 'Tool call';
      const args =
        typeof part.tool_call.args === 'string' && part.tool_call.args.trim().length > 0
          ? ` args=${truncateText(part.tool_call.args, 160)}`
          : '';
      const output =
        typeof part.tool_call.output === 'string' && part.tool_call.output.trim().length > 0
          ? ` output=${truncateText(part.tool_call.output, PREVIEW_CHAR_LIMIT)}`
          : '';
      fragments.push(`${toolName}${args}${output}`);
      continue;
    }

    const text = getPartText(part);
    if (text) {
      fragments.push(text);
    }
  }

  return fragments.join('\n');
}

function getRoleLabel(message = {}) {
  if (message.role === 'user') {
    return 'User';
  }
  if (message.role === 'assistant') {
    return 'Assistant';
  }
  if (message.role === 'system') {
    return 'System';
  }
  if (message.role === 'tool') {
    return 'Tool';
  }
  return message.sender || message.role || 'Message';
}

function summarizeMessageLine(message = {}) {
  const label = getRoleLabel(message);
  const content = flattenContent(message.content ?? message.text ?? '');
  if (!content) {
    return null;
  }
  return `- ${label}: ${truncateText(content, PREVIEW_CHAR_LIMIT)}`;
}

function trimPartText(part, remaining) {
  if (remaining <= 0) {
    return [null, 0];
  }

  if (typeof part?.text === 'string') {
    const text = truncateText(part.text, remaining);
    return [{ ...part, text }, Math.max(0, remaining - text.length)];
  }

  if (typeof part?.type === 'string' && typeof part[part.type] === 'string') {
    const text = truncateText(part[part.type], remaining);
    return [{ ...part, [part.type]: text }, Math.max(0, remaining - text.length)];
  }

  return [part, remaining];
}

function trimRetainedMessage(message, charLimit) {
  if (charLimit <= 0 || message == null || typeof message !== 'object') {
    return message;
  }

  if (typeof message.content === 'string') {
    return { ...message, content: truncateText(message.content, charLimit) };
  }

  if (!Array.isArray(message.content)) {
    return message;
  }

  let remaining = charLimit;
  const trimmedContent = [];

  for (const part of message.content) {
    if (remaining <= 0) {
      break;
    }

    if (part?.type === ContentTypes.THINK || part?.type === ContentTypes.ERROR) {
      continue;
    }

    if (part?.type === ContentTypes.TOOL_CALL && part.tool_call) {
      trimmedContent.push({
        ...part,
        tool_call: {
          ...part.tool_call,
          args: typeof part.tool_call.args === 'string' ? '[trimmed]' : part.tool_call.args,
          output:
            typeof part.tool_call.output === 'string'
              ? truncateText(part.tool_call.output, Math.min(remaining, 60))
              : part.tool_call.output,
        },
      });
      remaining = Math.max(0, remaining - 60);
      continue;
    }

    const [trimmedPart, nextRemaining] = trimPartText(part, remaining);
    if (trimmedPart) {
      trimmedContent.push(trimmedPart);
    }
    remaining = nextRemaining;
  }

  return { ...message, content: trimmedContent };
}

function sumPayloadTokens(payload, encoding) {
  return payload.reduce((total, message) => {
    const messageTokenCount =
      typeof message.tokenCount === 'number' && Number.isFinite(message.tokenCount)
        ? message.tokenCount
        : countFormattedMessageTokens(message, encoding);
    return total + messageTokenCount;
  }, 0);
}

function buildCompactionSummary({ olderMessages, existingSummary, maxChars }) {
  const sections = [
    'Conversation summary of earlier turns retained to stay within model limits.',
  ];

  if (existingSummary?.text) {
    sections.push(`Previous summary:\n${truncateText(existingSummary.text, 1200)}`);
  }

  const lines = [];
  for (const message of olderMessages) {
    const line = summarizeMessageLine(message);
    if (line) {
      lines.push(line);
    }
    if (lines.length >= PREVIEW_ENTRY_LIMIT) {
      break;
    }
  }

  if (lines.length > 0) {
    sections.push(`Earlier exchanges:\n${lines.join('\n')}`);
  }

  sections.push('Treat the recent messages below as the source of truth if anything conflicts.');

  let summary = sections.join('\n\n').trim();
  if (summary.length > maxChars) {
    summary = `${summary.slice(0, Math.max(0, maxChars - 1)).trimEnd()}...`;
  }

  return summary;
}

function chooseRetainedSlice(payload, encoding, modeConfig) {
  const totalBudget = Math.max(2048, Math.floor(modeConfig.targetTokenRatio * 1000000));
  void totalBudget;
  const retained = [];
  let retainedTokens = 0;
  let retainedCount = 0;

  const targetTailTokens = Math.max(2048, modeConfig.targetTailTokens || 0);

  for (let i = payload.length - 1; i >= 0; i--) {
    const message = payload[i];
    const tokenCount =
      typeof message.tokenCount === 'number' && Number.isFinite(message.tokenCount)
        ? message.tokenCount
        : countFormattedMessageTokens(message, encoding);

    const mustKeep = retainedCount < modeConfig.minRecentMessages;
    const withinTokenBudget =
      retainedTokens + tokenCount <= targetTailTokens || retainedCount < modeConfig.maxRecentMessages;

    if (!mustKeep && !withinTokenBudget) {
      return i + 1;
    }

    retained.unshift(message);
    retainedTokens += tokenCount;
    retainedCount++;

    if (retainedCount >= modeConfig.maxRecentMessages && retainedTokens >= targetTailTokens) {
      return i;
    }
  }

  return 0;
}

function compactPayloadToTarget({
  payload,
  maxContextTokens,
  encoding,
  initialSummary,
  mode = 'preflight',
}) {
  const modeConfig = COMPACTION_MODES[mode] ?? COMPACTION_MODES.preflight;
  const totalTokens = sumPayloadTokens(payload, encoding);
  const triggerTokens =
    maxContextTokens != null && maxContextTokens > 0
      ? Math.floor(maxContextTokens * modeConfig.triggerTokenRatio)
      : null;
  const exceededAbsoluteTrigger =
    typeof modeConfig.absoluteTriggerTokens === 'number' &&
    modeConfig.absoluteTriggerTokens > 0 &&
    totalTokens > modeConfig.absoluteTriggerTokens;

  if (
    mode === 'preflight' &&
    !exceededAbsoluteTrigger &&
    triggerTokens != null &&
    totalTokens <= triggerTokens &&
    payload.length <= modeConfig.triggerMessageCount
  ) {
    return {
      compacted: false,
      payload,
      initialSummary,
      estimatedTokens: totalTokens,
      signature: `none:${payload.length}:${totalTokens}`,
    };
  }

  const targetTailTokens =
    maxContextTokens != null && maxContextTokens > 0
      ? Math.max(2048, Math.floor(maxContextTokens * modeConfig.targetTokenRatio))
      : 8192;
  const configuredMode = {
    ...modeConfig,
    targetTailTokens: Math.min(targetTailTokens, modeConfig.targetTailTokenCap || targetTailTokens),
  };
  const retainFrom = chooseRetainedSlice(payload, encoding, configuredMode);
  const olderMessages = payload.slice(0, retainFrom);
  const retainedMessages = payload.slice(retainFrom);

  if (olderMessages.length === 0) {
    return {
      compacted: false,
      payload,
      initialSummary,
      estimatedTokens: totalTokens,
      signature: `none:${payload.length}:${totalTokens}`,
    };
  }

  const summaryText = buildCompactionSummary({
    olderMessages,
    existingSummary: initialSummary,
    maxChars: modeConfig.maxSummaryChars,
  });

  const summaryMessage = {
    role: 'system',
    content: summaryText,
  };
  summaryMessage.tokenCount = countFormattedMessageTokens(summaryMessage, encoding);

  const trimmedRetainedMessages = retainedMessages.map((message, index) => {
    const preserveVerbatim =
      index >= retainedMessages.length - (modeConfig.preserveRecentMessagesVerbatim || 0);
    if (preserveVerbatim) {
      return message;
    }
    return trimRetainedMessage(message, modeConfig.retainedMessageCharLimit);
  });

  const compactedPayload = [summaryMessage, ...trimmedRetainedMessages];
  const summaryState = { text: summaryText, tokenCount: summaryMessage.tokenCount };
  const compactedTokens = sumPayloadTokens(compactedPayload, encoding);

  logger.debug('[AgentCompaction] Compacted payload', {
    mode,
    originalMessages: payload.length,
    retainedMessages: trimmedRetainedMessages.length,
    summarizedMessages: olderMessages.length,
    originalTokens: totalTokens,
    compactedTokens,
  });

  return {
    compacted: true,
    payload: compactedPayload,
    initialSummary: summaryState,
    estimatedTokens: compactedTokens,
    signature: `${mode}:${retainFrom}:${compactedPayload.length}:${compactedTokens}`,
  };
}

function isPromptOverflowError(error) {
  const message = String(error?.message || '').toLowerCase();
  return (
    message.includes('input is too long') ||
    message.includes('prompt is too long') ||
    message.includes('maximum') && message.includes('token') ||
    message.includes('too many tokens')
  );
}

module.exports = {
  COMPACTION_MODES,
  compactPayloadToTarget,
  isPromptOverflowError,
  sumPayloadTokens,
};
