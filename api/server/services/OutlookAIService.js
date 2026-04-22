const { BedrockRuntimeClient, ConverseCommand } = require('@aws-sdk/client-bedrock-runtime');
const { NodeHttpHandler } = require('@smithy/node-http-handler');
const { HttpsProxyAgent } = require('https-proxy-agent');
const { isEnabled } = require('@librechat/api');
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

  return extractText(response);
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

  const raw = await callBedrock({ system, prompt });
  const parsed = parseJsonObject(raw);

  return {
    mode: 'bedrock',
    summary: String(parsed.summary || '').trim() || 'No summary was generated.',
    suggestedActions: normalizeStringArray(parsed.suggestedActions, [
      'Review the email and decide whether a reply is needed.',
    ]),
    riskSignals: normalizeStringArray(parsed.riskSignals, ['No obvious risk signals detected.']),
    calendarSignals: normalizeStringArray(parsed.calendarSignals, []),
    identitySignals: normalizeStringArray(parsed.identitySignals, []),
    generatedAt: new Date().toISOString(),
  };
}

async function generateReplyDraft({
  message,
  instructions = '',
  tone = 'professional',
  calendarEvents = [],
  outlookContext = {},
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
    emailConversation: compactEmail(message),
    outlookContext: compactOutlookContext(outlookContext),
    calendarEvents: compactCalendarEvents(calendarEvents),
  });

  return callBedrock({ system, prompt });
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
  logModelFailure,
  parseJsonObject,
  DEFAULT_DRAFT_STYLE,
  compactOutlookContext,
};
