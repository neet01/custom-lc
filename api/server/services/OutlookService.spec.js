jest.mock('@librechat/api', () => ({
  isEnabled: (value) => value === true || value === 'true' || value === '1',
}));

jest.mock('~/server/services/GraphTokenService', () => ({
  getGraphApiToken: jest.fn(),
}));

jest.mock('@librechat/data-schemas', () => ({
  logger: {
    warn: jest.fn(),
    error: jest.fn(),
  },
}));

const { getGraphApiToken } = require('~/server/services/GraphTokenService');
const OutlookService = require('./OutlookService');

const user = {
  id: 'user-1',
  provider: 'openid',
  openidId: 'entra-user',
  federatedTokens: {
    access_token: 'openid-access-token',
  },
};

describe('OutlookService', () => {
  const originalEnv = process.env;
  const originalFetch = global.fetch;

  beforeEach(() => {
    jest.clearAllMocks();
    process.env = {
      ...originalEnv,
      OUTLOOK_AI_ENABLED: 'true',
      OPENID_REUSE_TOKENS: 'true',
      OUTLOOK_GRAPH_BASE_URL: 'https://graph.microsoft.us',
      OUTLOOK_GRAPH_SCOPES: 'https://graph.microsoft.us/.default',
    };
    getGraphApiToken.mockResolvedValue({ access_token: 'graph-token' });
    global.fetch = jest.fn();
  });

  afterEach(() => {
    process.env = originalEnv;
    global.fetch = originalFetch;
  });

  it('reports delegated connection status without exchanging a token', () => {
    const status = OutlookService.getStatus(user);

    expect(status).toMatchObject({
      enabled: true,
      connected: true,
      graphBaseUrl: 'https://graph.microsoft.us/v1.0',
      scopes: 'https://graph.microsoft.us/.default',
    });
    expect(getGraphApiToken).not.toHaveBeenCalled();
  });

  it('rejects mailbox access when the feature is disabled', async () => {
    process.env.OUTLOOK_AI_ENABLED = 'false';

    await expect(OutlookService.listMessages(user)).rejects.toMatchObject({
      name: 'OutlookServiceError',
      status: 403,
    });
  });

  it('lists recent inbox messages from the GCC High Graph endpoint', async () => {
    global.fetch.mockResolvedValue({
      ok: true,
      status: 200,
      json: async () => ({
        value: [
          {
            id: 'message-1',
            conversationId: 'thread-1',
            subject: 'Budget follow-up',
            from: { emailAddress: { name: 'Finance', address: 'finance@example.mil' } },
            receivedDateTime: '2026-04-21T12:00:00Z',
            bodyPreview: 'Please review the runway numbers.',
            importance: 'high',
            inferenceClassification: 'focused',
            isRead: false,
            hasAttachments: true,
            webLink: 'https://outlook.example/message-1',
          },
        ],
      }),
    });

    const result = await OutlookService.listMessages(user, { limit: 10 });

    expect(result.messages).toHaveLength(1);
    expect(result.messages[0]).toMatchObject({
      id: 'message-1',
      subject: 'Budget follow-up',
      from: { name: 'Finance', address: 'finance@example.mil' },
      hasAttachments: true,
      inferenceClassification: 'focused',
    });
    expect(global.fetch).toHaveBeenCalledWith(
      expect.objectContaining({
        origin: 'https://graph.microsoft.us',
        pathname: '/v1.0/me/mailFolders/inbox/messages',
      }),
      expect.objectContaining({
        headers: expect.objectContaining({
          Authorization: 'Bearer graph-token',
        }),
      }),
    );
    const requestedUrl = global.fetch.mock.calls[0][0];
    expect(requestedUrl.searchParams.get('$filter')).toBe(
      "inferenceClassification eq 'focused'",
    );
  });

  it('deletes a message through Microsoft Graph', async () => {
    global.fetch.mockResolvedValue({
      ok: true,
      status: 204,
    });

    const result = await OutlookService.deleteMessage(user, 'message-to-delete');

    expect(result).toMatchObject({
      messageId: 'message-to-delete',
      message: 'Email moved to Deleted Items.',
    });
    expect(global.fetch).toHaveBeenCalledWith(
      expect.objectContaining({
        pathname: '/v1.0/me/messages/message-to-delete',
      }),
      expect.objectContaining({
        method: 'DELETE',
      }),
    );
  });

  it('creates a reply draft without sending mail', async () => {
    global.fetch
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          id: 'source-message',
          subject: 'Need input',
          from: { emailAddress: { name: 'Ops', address: 'ops@example.mil' } },
          body: { content: 'Can you review this?' },
          bodyPreview: 'Can you review this?',
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 201,
        json: async () => ({
          id: 'draft-message',
          subject: 'RE: Need input',
          bodyPreview: 'Thanks for reaching out.',
          webLink: 'https://outlook.example/draft-message',
        }),
      });

    const result = await OutlookService.createReplyDraft(user, 'source-message', {
      instructions: 'Ask for the due date.',
    });

    expect(result).toMatchObject({
      sourceMessageId: 'source-message',
      draftId: 'draft-message',
      message: 'Draft reply created. Review it in Outlook before sending.',
    });
    expect(global.fetch).toHaveBeenLastCalledWith(
      expect.objectContaining({
        pathname: '/v1.0/me/messages/source-message/createReply',
      }),
      expect.objectContaining({
        method: 'POST',
        body: expect.stringContaining('Ask for the due date.'),
      }),
    );
  });
});
