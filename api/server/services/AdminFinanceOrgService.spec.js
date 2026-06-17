jest.mock(
  '@librechat/data-schemas',
  () => ({
    logger: {
      debug: jest.fn(),
      warn: jest.fn(),
    },
  }),
  { virtual: true },
);

jest.mock('~/server/services/GraphTokenService', () => ({
  getGraphApiToken: jest.fn(),
}));

const { getGraphApiToken } = require('~/server/services/GraphTokenService');
const { resolveFinanceUserOrgMetadata } = require('./AdminFinanceOrgService');

describe('AdminFinanceOrgService', () => {
  const originalEnv = process.env;
  const originalFetch = global.fetch;
  const requester = {
    id: 'admin-1',
    openidId: 'admin-openid-1',
    federatedTokens: {
      access_token: 'admin-assertion',
    },
  };

  beforeEach(() => {
    jest.clearAllMocks();
    process.env = { ...originalEnv };
    delete process.env.ADMIN_USAGE_GRAPH_ORG_ENRICHMENT_ENABLED;
    delete process.env.ADMIN_USAGE_GRAPH_BASE_URL;
    delete process.env.ADMIN_USAGE_GRAPH_SCOPES;
    global.fetch = jest.fn();
  });

  afterAll(() => {
    process.env = originalEnv;
    global.fetch = originalFetch;
  });

  it('does not call Graph unless finance org enrichment is explicitly enabled', async () => {
    const result = await resolveFinanceUserOrgMetadata(
      [{ _id: { toString: () => 'user-1' }, email: 'user@example.com' }],
      requester,
    );

    expect(result.size).toBe(0);
    expect(getGraphApiToken).not.toHaveBeenCalled();
    expect(global.fetch).not.toHaveBeenCalled();
  });

  it('maps Graph directory fields into finance org metadata', async () => {
    process.env.ADMIN_USAGE_GRAPH_ORG_ENRICHMENT_ENABLED = 'true';
    getGraphApiToken.mockResolvedValue({ access_token: 'graph-token-1' });
    global.fetch.mockResolvedValue({
      ok: true,
      json: async () => ({
        id: 'graph-user-1',
        department: 'Finance Operations',
        jobTitle: 'Budget Analyst',
        companyName: 'Example Co',
        officeLocation: 'HQ',
      }),
    });

    const result = await resolveFinanceUserOrgMetadata(
      [
        {
          _id: { toString: () => 'user-1' },
          email: 'finance@example.com',
          openidId: 'openid-user-1',
        },
      ],
      requester,
    );

    expect(getGraphApiToken).toHaveBeenCalledWith(
      requester,
      'admin-assertion',
      'https://graph.microsoft.us/User.Read.All',
    );
    expect(String(global.fetch.mock.calls[0][0])).toContain(
      'https://graph.microsoft.us/v1.0/users/finance%40example.com',
    );
    expect(result.get('user-1')).toEqual({
      graphUserId: 'graph-user-1',
      team: 'Finance Operations',
      role: 'Budget Analyst',
      company: 'Example Co',
      officeLocation: 'HQ',
    });
  });
});
