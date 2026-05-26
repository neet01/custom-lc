jest.mock('openid-client', () => ({
  genericGrantRequest: jest.fn(),
  refreshTokenGrant: jest.fn(),
}));

jest.mock(
  '@librechat/data-schemas',
  () => ({
    logger: {
      debug: jest.fn(),
      info: jest.fn(),
      warn: jest.fn(),
      error: jest.fn(),
    },
  }),
  { virtual: true },
);

jest.mock('~/cache/getLogStores', () => jest.fn());
jest.mock('~/strategies/openidStrategy', () => ({
  getOpenIdConfig: jest.fn(() => ({ issuer: 'https://example.com' })),
}));
jest.mock('~/strategies', () => ({
  getOpenIdConfig: jest.fn(() => ({ issuer: 'https://example.com' })),
}));

const client = require('openid-client');
const getLogStores = require('~/cache/getLogStores');
const { getGraphApiToken } = require('./GraphTokenService');

describe('GraphTokenService', () => {
  const cache = {
    get: jest.fn(),
    set: jest.fn(),
  };

  const user = {
    id: 'user-1',
    openidId: 'openid-user-1',
    federatedTokens: {
      access_token: 'initial-assertion',
      refresh_token: 'refresh-token-1',
    },
  };

  beforeEach(() => {
    jest.clearAllMocks();
    getLogStores.mockReturnValue(cache);
    cache.get.mockResolvedValue(null);
    cache.set.mockResolvedValue(undefined);
    delete process.env.OPENID_SCOPE;
  });

  it('uses a cached Graph token when it is still safely valid', async () => {
    const cached = {
      access_token: 'cached-graph-token',
      expires_at: Date.now() + 10 * 60 * 1000,
      expires_in: 3600,
      scope: 'https://graph.microsoft.us/.default',
    };
    cache.get.mockResolvedValue(cached);

    const result = await getGraphApiToken(
      user,
      user.federatedTokens.access_token,
      'https://graph.microsoft.us/.default',
    );

    expect(result).toBe(cached);
    expect(client.genericGrantRequest).not.toHaveBeenCalled();
    expect(client.refreshTokenGrant).not.toHaveBeenCalled();
  });

  it('refreshes the delegated assertion before OBO when the assertion is near expiry', async () => {
    user.federatedTokens.expires_at = Math.floor(Date.now() / 1000) + 60;

    client.refreshTokenGrant.mockResolvedValue({
      access_token: 'refreshed-assertion',
      refresh_token: 'rotated-refresh-token',
      expires_in: 3600,
    });
    client.genericGrantRequest.mockResolvedValue({
      access_token: 'graph-token-1',
      expires_in: 3600,
    });

    const result = await getGraphApiToken(
      user,
      user.federatedTokens.access_token,
      'https://graph.microsoft.us/.default',
    );

    expect(client.refreshTokenGrant).toHaveBeenCalledTimes(1);
    expect(client.genericGrantRequest).toHaveBeenCalledWith(
      expect.anything(),
      'urn:ietf:params:oauth:grant-type:jwt-bearer',
      expect.objectContaining({
        assertion: 'refreshed-assertion',
        scope: 'https://graph.microsoft.us/.default',
      }),
    );
    expect(user.federatedTokens.access_token).toBe('refreshed-assertion');
    expect(user.federatedTokens.refresh_token).toBe('rotated-refresh-token');
    expect(result).toMatchObject({
      access_token: 'graph-token-1',
    });
  });

  it('refreshes and retries when Graph rejects an expired OBO assertion', async () => {
    client.genericGrantRequest
      .mockRejectedValueOnce({
        error: 'invalid_grant',
        error_description:
          'AADSTS500133: Assertion is not within its valid time range. Ensure that the access token is not expired before using it for user assertion.',
      })
      .mockResolvedValueOnce({
        access_token: 'graph-token-2',
        expires_in: 3600,
      });
    client.refreshTokenGrant.mockResolvedValue({
      access_token: 'recovered-assertion',
      expires_in: 3600,
    });

    const result = await getGraphApiToken(
      user,
      user.federatedTokens.access_token,
      'https://graph.microsoft.us/.default',
    );

    expect(client.refreshTokenGrant).toHaveBeenCalledTimes(1);
    expect(client.genericGrantRequest).toHaveBeenCalledTimes(2);
    expect(client.genericGrantRequest).toHaveBeenLastCalledWith(
      expect.anything(),
      'urn:ietf:params:oauth:grant-type:jwt-bearer',
      expect.objectContaining({
        assertion: 'recovered-assertion',
      }),
    );
    expect(result).toMatchObject({
      access_token: 'graph-token-2',
    });
  });
});
