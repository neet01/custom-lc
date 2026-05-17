jest.mock('@librechat/api', () => ({
  countFormattedMessageTokens: jest.fn((message) => {
    const content = typeof message?.content === 'string' ? message.content : JSON.stringify(message);
    return Math.max(1, Math.ceil(String(content || '').length / 8));
  }),
}));

const {
  compactPayloadToTarget,
  isPromptOverflowError,
} = require('~/server/controllers/agents/compaction');

describe('agent prompt compaction helper', () => {
  test('compacts older payload messages into a synthetic system summary', () => {
    const payload = Array.from({ length: 30 }, (_, index) => ({
      role: index % 2 === 0 ? 'user' : 'assistant',
      content: `Message ${index} ` + 'x'.repeat(4000),
    }));

    const result = compactPayloadToTarget({
      payload,
      maxContextTokens: 1200,
      encoding: 'claude',
      initialSummary: null,
      mode: 'overflow',
    });

    expect(result.compacted).toBe(true);
    expect(result.payload.length).toBeLessThan(payload.length);
    expect(result.payload[0].role).toBe('system');
    expect(typeof result.payload[0].content).toBe('string');
    expect(result.payload[0].content).toContain('Conversation summary of earlier turns');
    expect(result.initialSummary).toEqual(
      expect.objectContaining({
        text: expect.any(String),
        tokenCount: expect.any(Number),
      }),
    );
  });

  test('retains the most recent user message after compaction', () => {
    const payload = [
      ...Array.from({ length: 20 }, (_, index) => ({
        role: index % 2 === 0 ? 'user' : 'assistant',
        content: `Earlier ${index} ` + 'y'.repeat(3500),
      })),
      { role: 'user', content: 'Latest user request that must remain verbatim' },
    ];

    const result = compactPayloadToTarget({
      payload,
      maxContextTokens: 900,
      encoding: 'claude',
      initialSummary: null,
      mode: 'overflow',
    });

    expect(result.compacted).toBe(true);
    const lastMessage = result.payload[result.payload.length - 1];
    expect(lastMessage.role).toBe('user');
    expect(lastMessage.content).toBe('Latest user request that must remain verbatim');
  });

  test('detects provider overflow error messages', () => {
    expect(isPromptOverflowError(new Error('Input is too long for requested model.'))).toBe(true);
    expect(
      isPromptOverflowError(new Error('prompt is too long: 223784 tokens > 204660 maximum')),
    ).toBe(true);
    expect(isPromptOverflowError(new Error('socket hang up'))).toBe(false);
  });
});
