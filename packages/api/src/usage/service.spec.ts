import { createUsageRecord, persistUsageRecords, USAGE_TRACKING_ENABLED } from './service';

describe('usage service', () => {
  afterEach(() => {
    delete process.env[USAGE_TRACKING_ENABLED];
  });

  it('creates normalized usage records with cache tokens folded into input totals', () => {
    const record = createUsageRecord({
      user: 'user-123',
      conversationId: 'convo-123',
      inputTokens: 100,
      outputTokens: 40,
      cacheCreationTokens: 20,
      cacheReadTokens: 10,
      source: 'agent',
    });

    expect(record).toEqual(
      expect.objectContaining({
        user: 'user-123',
        conversationId: 'convo-123',
        inputTokens: 130,
        outputTokens: 40,
        totalTokens: 170,
        cacheCreationTokens: 20,
        cacheReadTokens: 10,
        source: 'agent',
      }),
    );
  });

  it('persists usage records when tracking is enabled', async () => {
    process.env[USAGE_TRACKING_ENABLED] = 'true';
    const createUsageRecords = jest.fn().mockResolvedValue(undefined);

    await persistUsageRecords(
      { createUsageRecords },
      [
        {
          user: 'user-123',
          conversationId: 'convo-123',
          inputTokens: 20,
          outputTokens: 5,
        },
      ],
    );

    expect(createUsageRecords).toHaveBeenCalledWith([
      expect.objectContaining({
        user: 'user-123',
        conversationId: 'convo-123',
        inputTokens: 20,
        outputTokens: 5,
        totalTokens: 25,
      }),
    ]);
  });

  it('skips persistence when tracking is disabled', async () => {
    const createUsageRecords = jest.fn().mockResolvedValue(undefined);

    await persistUsageRecords(
      { createUsageRecords },
      [
        {
          user: 'user-123',
          conversationId: 'convo-123',
          inputTokens: 20,
          outputTokens: 5,
        },
      ],
    );

    expect(createUsageRecords).not.toHaveBeenCalled();
  });
});
