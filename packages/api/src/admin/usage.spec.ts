import type { Response } from 'express';
import { createAdminUsageHandlers } from './usage';
import type { ServerRequest } from '~/types/http';

jest.mock('@librechat/data-schemas', () => ({
  ...jest.requireActual('@librechat/data-schemas'),
  logger: { error: jest.fn(), warn: jest.fn(), info: jest.fn(), debug: jest.fn() },
}));

describe('createAdminUsageHandlers', () => {
  afterEach(() => {
    delete process.env.USAGE_TRACKING_ENABLED;
  });

  function createReqRes(query: Record<string, string> = {}) {
    const req = { query } as unknown as ServerRequest;
    const json = jest.fn();
    const status = jest.fn().mockReturnValue({ json });
    const res = { status, json } as unknown as Response;
    return { req, res, status, json };
  }

  it('returns usage records when tracking is enabled', async () => {
    process.env.USAGE_TRACKING_ENABLED = 'true';
    const handlers = createAdminUsageHandlers({
      findUsageRecords: jest.fn().mockResolvedValue([
        {
          _id: { toString: () => 'usage-1' },
          user: { toString: () => 'user-123' },
          conversationId: 'convo-123',
          inputTokens: 100,
          outputTokens: 25,
          totalTokens: 125,
          createdAt: new Date('2026-04-13T12:00:00.000Z'),
          updatedAt: new Date('2026-04-13T12:00:00.000Z'),
        },
      ]),
      countUsageRecords: jest.fn().mockResolvedValue(1),
      summarizeUsageByUser: jest.fn(),
      summarizeUsageOverview: jest.fn(),
      findUsers: jest.fn(),
    });
    const { req, res, status, json } = createReqRes({ user_id: '507f1f77bcf86cd799439011' });

    await handlers.listUsage(req, res);

    expect(status).toHaveBeenCalledWith(200);
    expect(json).toHaveBeenCalledWith({
      usage: [
        expect.objectContaining({
          id: 'usage-1',
          userId: 'user-123',
          conversationId: 'convo-123',
          inputTokens: 100,
          outputTokens: 25,
          totalTokens: 125,
        }),
      ],
      total: 1,
      limit: 50,
      offset: 0,
    });
  });

  it('returns 503 when usage tracking is disabled', async () => {
    const handlers = createAdminUsageHandlers({
      findUsageRecords: jest.fn(),
      countUsageRecords: jest.fn(),
      summarizeUsageByUser: jest.fn(),
      summarizeUsageOverview: jest.fn(),
      findUsers: jest.fn(),
    });
    const { req, res, status, json } = createReqRes();

    await handlers.listUsage(req, res);

    expect(status).toHaveBeenCalledWith(503);
    expect(json).toHaveBeenCalledWith({ error: 'Usage tracking is disabled' });
  });

  it('returns 400 for invalid user ids', async () => {
    process.env.USAGE_TRACKING_ENABLED = 'true';
    const handlers = createAdminUsageHandlers({
      findUsageRecords: jest.fn(),
      countUsageRecords: jest.fn(),
      summarizeUsageByUser: jest.fn(),
      summarizeUsageOverview: jest.fn(),
      findUsers: jest.fn(),
    });
    const { req, res, status, json } = createReqRes({ user_id: 'bad-id' });

    await handlers.listUsage(req, res);

    expect(status).toHaveBeenCalledWith(400);
    expect(json).toHaveBeenCalledWith({ error: 'Invalid user ID format' });
  });

  it('returns a usage summary with user details', async () => {
    process.env.USAGE_TRACKING_ENABLED = 'true';
    const summarizeUsageOverview = jest.fn().mockResolvedValue({
      requestCount: 3,
      inputTokens: 400,
      outputTokens: 80,
      totalTokens: 480,
      cacheCreationTokens: 0,
      cacheReadTokens: 0,
      avgLatencyMs: 123.4,
      activeUsers: 1,
      firstSeenAt: new Date('2026-04-12T12:00:00.000Z'),
      lastSeenAt: new Date('2026-04-13T12:00:00.000Z'),
    });
    const summarizeUsageByUser = jest.fn().mockResolvedValue([
      {
        userId: '507f1f77bcf86cd799439011',
        requestCount: 3,
        inputTokens: 400,
        outputTokens: 80,
        totalTokens: 480,
        cacheCreationTokens: 0,
        cacheReadTokens: 0,
        avgLatencyMs: 123.4,
        firstSeenAt: new Date('2026-04-12T12:00:00.000Z'),
        lastSeenAt: new Date('2026-04-13T12:00:00.000Z'),
      },
    ]);
    const handlers = createAdminUsageHandlers({
      findUsageRecords: jest.fn(),
      countUsageRecords: jest.fn(),
      summarizeUsageByUser,
      summarizeUsageOverview,
      findUsers: jest.fn().mockResolvedValue([
        {
          _id: { toString: () => '507f1f77bcf86cd799439011' },
          name: 'Admin User',
          username: 'admin',
          email: 'admin@example.com',
          avatar: '',
          role: 'ADMIN',
          provider: 'local',
        },
      ]),
    });
    const { req, res, status, json } = createReqRes({ days: '7' });

    await handlers.getUsageSummary(req, res);

    expect(summarizeUsageOverview).toHaveBeenCalled();
    expect(summarizeUsageByUser).toHaveBeenCalledWith(
      expect.objectContaining({
        createdAt: expect.any(Object),
      }),
      { limit: 50, offset: 0 },
    );
    expect(status).toHaveBeenCalledWith(200);
    expect(json).toHaveBeenCalledWith({
      overview: expect.objectContaining({
        requestCount: 3,
        totalTokens: 480,
        activeUsers: 1,
        avgLatencyMs: 123.4,
      }),
      users: [
        expect.objectContaining({
          userId: '507f1f77bcf86cd799439011',
          name: 'Admin User',
          email: 'admin@example.com',
          totalTokens: 480,
          requestCount: 3,
        }),
      ],
      total: 1,
      limit: 50,
      offset: 0,
      days: 7,
    });
  });
});
