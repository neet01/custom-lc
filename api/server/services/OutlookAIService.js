const { BedrockRuntimeClient, ConverseCommand } = require('@aws-sdk/client-bedrock-runtime');
const { NodeHttpHandler } = require('@smithy/node-http-handler');
const { HttpsProxyAgent } = require('https-proxy-agent');
const { logger } = require('@librechat/data-schemas');

const DEFAULT_REGION = 'us-gov-west-1';
const DEFAULT_MAX_TOKENS = 1200;
const DEFAULT_TEMPERATURE = 0.2;
const DEFAULT_DRAFT_STYLE =
  'direct, concise, assertive, and professionally neutral. Avoid excessive warmth, flattery, apologies, hedging, and attention-seeking phrasing.';

let bedrockClient;

function getAIConfig() {
  return {
    provider: String(process.env.OUTLOOK_AI_PROVIDER || 'local').toLowerCase(),
    modelId: process.env.OUTLOOK_AI_MODEL_ID,
    region:
      process.env.OUTLOOK_AI_BEDROCK_REGION ||
      process.env.BEDROCK_AWS_DEFAULT_REGION ||
      process.env.AWS_REGION ||
      DEFAULT_REGION,
    maxTokens: Number(process.env.OUTLOOK_AI_MAX_TOKENS) || DEFAULT_MAX_TOKENS,
    temperature:
      process.env.OUTLOOK_AI_TEMPERATURE !== undefined
        ? Number(process.env.OUTLOOK_AI_TEMPERATURE)
        : DEFAULT_TEMPERATURE,
    draftStyle: process.env.OUTLOOK_AI_DRAFT_STYLE || DEFAULT_DRAFT_STYLE,
  };
}

function isModelBackedAIEnabled() {
  const config = getAIConfig();
  return config.provider === 'bedrock' && Boolean(config.modelId);
}

function getBedrockClient() {
  const config = getAIConfig();
  if (bedrockClient?.config?.region === config.region) {
    return bedrockClient.client;
  }

  const clientConfig = { region: config.region };
  if (process.env.PROXY) {
    const proxyAgent = new HttpsProxyAgent(process.env.PROXY);
    clientConfig.requestHandler = new NodeHttpHandler({
      httpAgent: proxyAgent,
      httpsAgent: proxyAgent,
    });
  }

  bedrockClient = {
    config: { region: config.region },
    client: new BedrockRuntimeClient(clientConfig),
  };
  return bedrockClient.client;
}

function compactEmail(message) {
  return {
    subject: message.subject,
    from: message.from,
    toRecipients: message.toRecipients,
    ccRecipients: message.ccRecipients,
    receivedDateTime: message.receivedDateTime,
    importance: message.importance,
    hasAttachments: message.hasAttachments,
    body: message.body || message.bodyPreview || '',
    thread: Array.isArray(message.thread)
      ? message.thread.slice(0, 12).map((threadMessage) => ({
          subject: threadMessage.subject,
          from: threadMessage.from,
          receivedDateTime: threadMessage.receivedDateTime,
          body: threadMessage.body || threadMessage.bodyPreview || '',
        }))
      : undefined,
  };
}

function compactCalendarEvents(calendarEvents = []) {
  return calendarEvents.slice(0, 8).map((event) => ({
    subject: event.subject,
    start: event.start,
    end: event.end,
    location: event.location,
    organizer: event.organizer,
    showAs: event.showAs,
    isOnlineMeeting: event.isOnlineMeeting,
  }));
}

function compactMessagesForBrief(messages = []) {
  return messages.slice(0, 12).map((message) => compactEmail(message));
}

function compactOutlookContext(outlookContext = {}) {
  const compactParticipant = (participant) => ({
    name: participant.name,
    address: participant.address,
    roles: participant.roles,
    relationshipToSignedInUser: participant.relationshipToSignedInUser,
    profile: participant.profile
      ? {
          displayName: participant.profile.displayName,
          email: participant.profile.email,
          jobTitle: participant.profile.jobTitle,
          department: participant.profile.department,
          officeLocation: participant.profile.officeLocation,
        }
      : undefined,
  });

  return {
    signedInUser: outlookContext.signedInUser
      ? {
          displayName: outlookContext.signedInUser.displayName,
          email: outlookContext.signedInUser.email,
          userPrincipalName: outlookContext.signedInUser.userPrincipalName,
          jobTitle: outlookContext.signedInUser.jobTitle,
          department: outlookContext.signedInUser.department,
          officeLocation: outlookContext.signedInUser.officeLocation,
        }
      : undefined,
    manager: outlookContext.manager
      ? {
          displayName: outlookContext.manager.displayName,
          email: outlookContext.manager.email,
          jobTitle: outlookContext.manager.jobTitle,
          department: outlookContext.manager.department,
        }
      : undefined,
    mailboxSettings: outlookContext.mailboxSettings,
    participants: Array.isArray(outlookContext.participants)
      ? outlookContext.participants.slice(0, 20).map(compactParticipant)
      : [],
    rules: outlookContext.rules || [],
  };
}

function extractText(response) {
  const content = response?.output?.message?.content;
  if (!Array.isArray(content)) {
    return '';
  }
  return content
    .map((part) => part?.text || '')
    .filter(Boolean)
    .join('\n')
    .trim();
}

function extractUsage(response, fallback = {}) {
  const usage = response?.usage ?? {};
  const inputTokens = Number(usage.inputTokens ?? usage.input_tokens ?? 0) || 0;
  const outputTokens = Number(usage.outputTokens ?? usage.output_tokens ?? 0) || 0;
  const totalTokens =
    Number(usage.totalTokens ?? usage.total_tokens ?? inputTokens + outputTokens) ||
    inputTokens + outputTokens;

  if (inputTokens <= 0 && outputTokens <= 0 && totalTokens <= 0) {
    return null;
  }

  return {
    input_tokens: inputTokens,
    output_tokens: outputTokens,
    total_tokens: totalTokens,
    model: fallback.model,
    provider: fallback.provider,
  };
}

function parseJsonObject(value) {
  const text = String(value || '').trim();
  try {
    return JSON.parse(text);
  } catch {
    const match = text.match(/\{[\s\S]*\}/);
    if (!match) {
      throw new Error('Model response did not contain JSON');
    }
    return JSON.parse(match[0]);
  }
}

function normalizeStringArray(value, fallback) {
  if (!Array.isArray(value)) {
    return fallback;
  }
  const strings = value.map((item) => String(item || '').trim()).filter(Boolean);
  return strings.length > 0 ? strings : fallback;
}

async function callBedrock({ system, prompt }) {
  const config = getAIConfig();
  if (!config.modelId) {
    throw new Error('OUTLOOK_AI_MODEL_ID is required for Bedrock-backed Outlook AI');
  }

  const response = await getBedrockClient().send(
    new ConverseCommand({
      modelId: config.modelId,
      system: [{ text: system }],
      messages: [
        {
          role: 'user',
          content: [{ text: prompt }],
        },
      ],
      inferenceConfig: {
        maxTokens: config.maxTokens,
        temperature: Number.isFinite(config.temperature) ? config.temperature : DEFAULT_TEMPERATURE,
      },
    }),
  );

  return {
    text: extractText(response),
    usage: extractUsage(response, {
      model: config.modelId,
      provider: config.provider,
    }),
  };
}

async function generateAnalysis({ message, calendarEvents = [], outlookContext = {} }) {
  if (!isModelBackedAIEnabled()) {
    return null;
  }

  const system = [
    'You are an enterprise email assistant embedded in LibreChat.',
    'Analyze emails for busy internal users in a regulated environment.',
    'Do not invent facts. Do not include sensitive raw email text unless necessary.',
    'Return only valid JSON matching the requested schema.',
  ].join(' ');

  const prompt = JSON.stringify({
    task: 'Analyze this email conversation and produce concise action-oriented inbox insights.',
    schema: {
      summary: 'string, 2-4 sentences',
      suggestedActions: ['3-6 concrete action items'],
      riskSignals: ['0-5 urgency, compliance, dependency, or scheduling signals'],
      calendarSignals: ['0-4 signals from calendar context, if provided'],
      identitySignals: [
        '0-4 signals about who should respond, whether signedInUser is directly addressed, and hierarchy if known',
      ],
    },
    emailConversation: compactEmail(message),
    outlookContext: compactOutlookContext(outlookContext),
    calendarEvents: compactCalendarEvents(calendarEvents),
  });

  const response = await callBedrock({ system, prompt });
  const parsed = parseJsonObject(response.text);

  return {
    insights: {
      mode: 'bedrock',
      summary: String(parsed.summary || '').trim() || 'No summary was generated.',
      suggestedActions: normalizeStringArray(parsed.suggestedActions, [
        'Review the email and decide whether a reply is needed.',
      ]),
      riskSignals: normalizeStringArray(parsed.riskSignals, ['No obvious risk signals detected.']),
      calendarSignals: normalizeStringArray(parsed.calendarSignals, []),
      identitySignals: normalizeStringArray(parsed.identitySignals, []),
      generatedAt: new Date().toISOString(),
    },
    usage: response.usage,
  };
}

async function generateReplyDraft({
  message,
  instructions = '',
  tone = 'professional',
  calendarEvents = [],
  outlookContext = {},
  draftRecipients = {},
  replyMode = 'reply',
}) {
  if (!isModelBackedAIEnabled()) {
    return null;
  }

  const config = getAIConfig();
  const system = [
    'You are drafting an Outlook reply for the signed-in user.',
    'The signedInUser in outlookContext is the only person you are allowed to write as.',
    'Never sign as the sender, another recipient, a manager, an executive, or any person other than signedInUser.',
    'If signedInUser is one of multiple recipients, draft only the response signedInUser should send. Do not imply another recipient approved anything.',
    'If the correct author is ambiguous, draft a short clarification instead of pretending to be someone else.',
    'Use signedInUser.displayName for the sign-off only when a sign-off is appropriate.',
    'Write only the reply body, with no markdown fences and no analysis preamble.',
    `Default writing style: ${config.draftStyle}`,
    'Be clear about ownership, asks, next steps, dates, and blockers.',
    'Prefer short paragraphs and plain language. Cut filler.',
    'Do not use phrases like "I hope this email finds you well", "just checking in", "I would be happy to", or "please let me know if you need anything else" unless the user explicitly asks for a softer tone.',
    'Do not apologize, praise, or express enthusiasm unless the thread context makes it necessary.',
    'Do not beg for attention or over-explain. Make the ask directly.',
    'Do not promise actions the user did not approve.',
    'If the email asks for scheduling, use calendar context cautiously and suggest availability windows only when clear.',
    'Recipient alignment is mandatory: draft text must match the actual draft recipients list.',
    'If there is exactly one draft recipient, address only that person in the salutation.',
    'If there are multiple recipients, use a group salutation that matches the recipient set (for example "Hi all,").',
    'Never mention names that are not present in draft recipients.',
  ].join(' ');

  const prompt = JSON.stringify({
    task: 'Draft a reply to this email conversation.',
    tone,
    defaultStyle: config.draftStyle,
    userInstructions:
      instructions ||
      'Draft a direct, useful response that answers the thread, states the next step, and avoids unnecessary pleasantries.',
    identityRules: [
      'Author is signedInUser only.',
      'Sign-off must match signedInUser, not any other participant.',
      'Use participant job titles and relationship context when available, but do not invent missing hierarchy.',
      'If signedInUser should not be the responder, say so briefly and explain what needs clarification.',
    ],
    replyMode,
    draftRecipients: {
      toRecipients: Array.isArray(draftRecipients.toRecipients) ? draftRecipients.toRecipients : [],
      ccRecipients: Array.isArray(draftRecipients.ccRecipients) ? draftRecipients.ccRecipients : [],
    },
    emailConversation: compactEmail(message),
    outlookContext: compactOutlookContext(outlookContext),
    calendarEvents: compactCalendarEvents(calendarEvents),
  });

  const response = await callBedrock({ system, prompt });
  return {
    draft: response.text,
    usage: response.usage,
  };
}

async function generateMeetingInviteNote({
  message,
  subject,
  slot,
  instructions = '',
  calendarEvents = [],
  outlookContext = {},
}) {
  if (!isModelBackedAIEnabled()) {
    return null;
  }

  const system = [
    'You are drafting the body text for an Outlook calendar invite.',
    'Write a practical, concise meeting brief with direct language.',
    'Do not invent facts, owners, or decisions that are not grounded in the thread context.',
    'Return plain text only, no markdown, no JSON, no code fences.',
    'Structure should include objective, agenda bullets, and prep/decisions when relevant.',
  ].join(' ');

  const prompt = JSON.stringify({
    task: 'Generate a meeting invite note from this email context.',
    outputFormat: 'plain_text',
    constraints: {
      maxLength: 1200,
      style: 'direct, concise, neutral, and businesslike',
      include: ['Objective', 'Agenda (2-5 bullets)', 'Prep/decisions needed (when relevant)'],
    },
    subject,
    scheduledSlot: slot,
    organizerInstructions: instructions || undefined,
    emailConversation: compactEmail(message),
    outlookContext: compactOutlookContext(outlookContext),
    calendarEvents: compactCalendarEvents(calendarEvents),
  });

  const response = await callBedrock({ system, prompt });
  const text = String(response.text || '')
    .replace(/\r/g, '')
    .trim();
  return {
    note: text || null,
    usage: response.usage,
  };
}

async function generateSelectionBrief({ messages = [], outlookContext = {} }) {
  if (!isModelBackedAIEnabled()) {
    return null;
  }

  const system = [
    'You are an executive email triage assistant embedded in LibreChat.',
    'Synthesize multiple emails into an efficient, decision-oriented brief.',
    'Do not invent facts. Be concrete, concise, and operational.',
    'Return only valid JSON matching the requested schema.',
  ].join(' ');

  const prompt = JSON.stringify({
    task: 'Summarize the selected emails into a short actionable brief.',
    schema: {
      headline: 'string, one sentence',
      summary: 'string, 2-5 sentences',
      priorities: ['2-6 highest-priority items'],
      followUps: ['0-6 recommended follow-ups'],
      meetingHighlights: ['0-4 meeting-related notes if present'],
      notableEmails: ['2-6 notable emails with sender/topic emphasis'],
      risks: ['0-5 urgency, dependency, or compliance risks'],
    },
    selectedEmailCount: messages.length,
    emails: compactMessagesForBrief(messages),
    outlookContext: compactOutlookContext(outlookContext),
  });

  const response = await callBedrock({ system, prompt });
  const parsed = parseJsonObject(response.text);

  return {
    brief: {
      mode: 'bedrock',
      headline: String(parsed.headline || '').trim() || 'Selected email summary ready.',
      summary: String(parsed.summary || '').trim() || 'No summary was generated.',
      priorities: normalizeStringArray(parsed.priorities, [
        'Review the selected emails and identify the highest-priority reply.',
      ]),
      followUps: normalizeStringArray(parsed.followUps, []),
      meetingHighlights: normalizeStringArray(parsed.meetingHighlights, []),
      notableEmails: normalizeStringArray(parsed.notableEmails, []),
      risks: normalizeStringArray(parsed.risks, ['No obvious risks were detected.']),
      generatedAt: new Date().toISOString(),
    },
    usage: response.usage,
  };
}

async function generateDailyBrief({
  messages = [],
  meetings = [],
  outlookContext = {},
  windowHours = 24,
}) {
  if (!isModelBackedAIEnabled()) {
    return null;
  }

  const system = [
    'You are an executive daily-brief assistant embedded in LibreChat.',
    'Summarize the last 24 hours of email and meeting activity into a compact, high-signal brief.',
    'Focus on priorities, follow-ups, decisions, and risks.',
    'Do not invent facts. Return only valid JSON matching the requested schema.',
  ].join(' ');

  const prompt = JSON.stringify({
    task: 'Generate a daily brief from the last 24 hours of Outlook activity.',
    schema: {
      headline: 'string, one sentence',
      summary: 'string, 3-6 sentences',
      priorities: ['2-6 top priorities from emails and meetings'],
      followUps: ['0-6 recommended follow-ups'],
      meetingHighlights: ['0-6 meeting outcomes, topics, or follow-up needs'],
      notableEmails: ['2-6 notable emails with sender/topic emphasis'],
      risks: ['0-5 urgency, dependency, or compliance risks'],
    },
    windowHours,
    emailCount: messages.length,
    meetingCount: meetings.length,
    emails: compactMessagesForBrief(messages),
    meetings: compactCalendarEvents(meetings),
    outlookContext: compactOutlookContext(outlookContext),
  });

  const response = await callBedrock({ system, prompt });
  const parsed = parseJsonObject(response.text);

  return {
    brief: {
      mode: 'bedrock',
      headline: String(parsed.headline || '').trim() || 'Daily brief ready.',
      summary: String(parsed.summary || '').trim() || 'No summary was generated.',
      priorities: normalizeStringArray(parsed.priorities, [
        'Review the latest email and meeting activity for follow-up items.',
      ]),
      followUps: normalizeStringArray(parsed.followUps, []),
      meetingHighlights: normalizeStringArray(parsed.meetingHighlights, []),
      notableEmails: normalizeStringArray(parsed.notableEmails, []),
      risks: normalizeStringArray(parsed.risks, ['No obvious risks were detected.']),
      generatedAt: new Date().toISOString(),
    },
    usage: response.usage,
  };
}

function logModelFailure(operation, error) {
  logger.warn('[OutlookAIService] Model-backed Outlook AI failed; falling back safely', {
    operation,
    provider: getAIConfig().provider,
    modelConfigured: Boolean(getAIConfig().modelId),
    error: error?.message,
  });
}

module.exports = {
  getAIConfig,
  isModelBackedAIEnabled,
  generateAnalysis,
  generateReplyDraft,
  generateMeetingInviteNote,
  generateSelectionBrief,
  generateDailyBrief,
  logModelFailure,
  parseJsonObject,
  DEFAULT_DRAFT_STYLE,
  compactOutlookContext,
};
