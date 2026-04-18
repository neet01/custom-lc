import { createAdminIssuesHandlers } from './issues';

describe('createAdminIssuesHandlers', () => {
  it('lists issues with reporter metadata', async () => {
    const handlers = createAdminIssuesHandlers({
      findIssueReports: jest.fn().mockResolvedValue([
        {
          _id: { toString: () => 'issue-1' },
          user: { toString: () => '507f191e810c19729de860ea' },
          conversationId: 'convo-1',
          messageId: 'msg-1',
          category: 'bad_file_transformation',
          status: 'open',
          description: 'Spreadsheet dropped the wrong column',
          createdAt: new Date('2026-04-17T12:00:00.000Z'),
          updatedAt: new Date('2026-04-17T12:00:00.000Z'),
        },
      ]),
      countIssueReports: jest.fn().mockResolvedValue(1),
      findUsers: jest.fn().mockResolvedValue([
        {
          _id: { toString: () => '507f191e810c19729de860ea' },
          name: 'Praneet Kotah',
          username: 'praneet.kotah',
          email: 'praneet.kotah@hermeus.com',
          avatar: '',
          role: 'ADMIN',
        },
      ]),
    });

    const req = {
      query: { status: 'open' },
    } as any;
    const res = {
      status: jest.fn().mockReturnThis(),
      json: jest.fn(),
    } as any;

    await handlers.listIssues(req, res);

    expect(res.status).toHaveBeenCalledWith(200);
    expect(res.json).toHaveBeenCalledWith(
      expect.objectContaining({
        issues: [
          expect.objectContaining({
            id: 'issue-1',
            reporterEmail: 'praneet.kotah@hermeus.com',
            category: 'bad_file_transformation',
            status: 'open',
          }),
        ],
        total: 1,
      }),
    );
  });
});
