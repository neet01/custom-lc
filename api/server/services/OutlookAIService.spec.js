jest.mock('@aws-sdk/client-bedrock-runtime', () => {
  const send = jest.fn();
  return {
    __mockSend: send,
    BedrockRuntimeClient: jest.fn(() => ({ send })),
    ConverseCommand: jest.fn((input) => ({ input })),
  };
});

jest.mock('@librechat/api', () => ({
  isEnabled: (value) => value === true || value === 'true' || value === '1',
}));

jest.mock('@librechat/data-schemas', () => ({
  logger: {
    warn: jest.fn(),
    error: jest.fn(),
  },
}));

const bedrockRuntime = require('@aws-sdk/client-bedrock-runtime');
const OutlookAIService = require('./OutlookAIService');

describe('OutlookAIService', () => {
  const originalEnv = process.env;

  beforeEach(() => {
    jest.clearAllMocks();
    process.env = {
      ...originalEnv,
      OUTLOOK_AI_PROVIDER: 'bedrock',
      OUTLOOK_AI_MODEL_ID: 'amazon.nova-micro-v1:0',
      OUTLOOK_AI_BEDROCK_REGION: 'us-gov-west-1',
    };
    bedrockRuntime.__mockSend.mockResolvedValue({
      output: {
        message: {
          content: [{ text: 'I need the owner and due date before I can commit to the timeline.' }],
        },
      },
    });
  });

  afterEach(() => {
    process.env = originalEnv;
  });

  it('uses a direct assertive default style for reply drafting', async () => {
    await OutlookAIService.generateReplyDraft({
      message: {
        subject: 'Timeline request',
        from: { name: 'Ops', address: 'ops@example.mil' },
        toRecipients: [{ name: 'User Two', address: 'user2@example.mil' }],
        body: 'Can you take this on?',
      },
      outlookContext: {
        signedInUser: {
          displayName: 'User Two',
          email: 'user2@example.mil',
          jobTitle: 'Program Manager',
        },
        participants: [
          {
            name: 'Ops',
            address: 'ops@example.mil',
            relationshipToSignedInUser: 'internal_user',
            profile: { jobTitle: 'Director of Ops', department: 'Operations' },
          },
          {
            name: 'User Two',
            address: 'user2@example.mil',
            relationshipToSignedInUser: 'signed_in_user',
            profile: { jobTitle: 'Program Manager', department: 'Engineering' },
          },
        ],
      },
    });

    const commandInput = bedrockRuntime.ConverseCommand.mock.calls[0][0];
    const systemPrompt = commandInput.system.map((part) => part.text).join('\n');
    const userPrompt = JSON.parse(commandInput.messages[0].content[0].text);

    expect(systemPrompt).toContain('Default writing style');
    expect(systemPrompt).toContain('direct, concise, assertive');
    expect(systemPrompt).toContain('Do not beg for attention');
    expect(systemPrompt).toContain('only person you are allowed to write as');
    expect(userPrompt.outlookContext.signedInUser.displayName).toBe('User Two');
    expect(userPrompt.identityRules).toContain('Author is signedInUser only.');
    expect(userPrompt.userInstructions).toContain('avoids unnecessary pleasantries');
  });

  it('allows the default draft style to be overridden by environment', async () => {
    process.env.OUTLOOK_AI_DRAFT_STYLE = 'brief, executive, and firm';

    await OutlookAIService.generateReplyDraft({
      message: {
        subject: 'Decision needed',
        from: { name: 'Finance', address: 'finance@example.mil' },
        body: 'Should we proceed?',
      },
    });

    const commandInput = bedrockRuntime.ConverseCommand.mock.calls[0][0];
    const systemPrompt = commandInput.system.map((part) => part.text).join('\n');
    const userPrompt = JSON.parse(commandInput.messages[0].content[0].text);

    expect(systemPrompt).toContain('brief, executive, and firm');
    expect(userPrompt.defaultStyle).toBe('brief, executive, and firm');
  });
});
