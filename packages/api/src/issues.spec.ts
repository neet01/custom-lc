import { createIssueHandlers } from './issues';

describe('createIssueHandlers', () => {
  it('creates issue reports from authenticated requests', async () => {
    const createIssueReport = jest.fn().mockResolvedValue({
      _id: { toString: () => 'issue-1' },
      user: { toString: () => 'user-1' },
      conversationId: 'convo-1',
      messageId: 'msg-1',
      category: 'faulty_mcp_tool',
      status: 'open',
      description: 'Bad tool result',
      createdAt: new Date('2026-04-17T12:00:00.000Z'),
      updatedAt: new Date('2026-04-17T12:00:00.000Z'),
    });

    const handlers = createIssueHandlers({
      createIssueReport,
    });

    const req = {
      user: { id: 'user-1' },
      body: {
        conversationId: 'convo-1',
        messageId: 'msg-1',
        category: 'faulty_mcp_tool',
        description: 'Bad tool result',
      },
    } as any;
    const res = {
      status: jest.fn().mockReturnThis(),
      json: jest.fn(),
    } as any;

    await handlers.reportIssue(req, res);

    expect(createIssueReport).toHaveBeenCalledWith(
      expect.objectContaining({
        user: 'user-1',
        conversationId: 'convo-1',
        messageId: 'msg-1',
        category: 'faulty_mcp_tool',
      }),
    );
    expect(res.status).toHaveBeenCalledWith(201);
  });

  it('validates required fields', async () => {
    const handlers = createIssueHandlers({
      createIssueReport: jest.fn(),
    });

    const req = {
      user: { id: 'user-1' },
      body: {
        category: 'other',
      },
    } as any;
    const res = {
      status: jest.fn().mockReturnThis(),
      json: jest.fn(),
    } as any;

    await handlers.reportIssue(req, res);

    expect(res.status).toHaveBeenCalledWith(400);
  });
});
