const request = require('supertest');
const express = require('express');

jest.mock('~/server/middleware', () => ({
  requireJwtAuth: (req, _res, next) => {
    req.user = { id: '507f191e810c19729de860ea', role: 'ADMIN' };
    next();
  },
}));

jest.mock('~/server/middleware/roles/capabilities', () => ({
  requireCapability: () => (_req, _res, next) => next(),
}));

jest.mock('~/models', () => ({
  findIssueReports: jest.fn(),
  countIssueReports: jest.fn(),
  findUsers: jest.fn(),
}));

jest.mock('@librechat/api', () => ({
  createAdminIssuesHandlers: ({ findIssueReports, countIssueReports, findUsers }) => ({
    listIssues: async (_req, res) => {
      const issues = await findIssueReports();
      const total = await countIssueReports();
      const users = await findUsers();
      const reporterEmail = users[0]?.email || '';

      return res.status(200).json({
        issues: issues.map((issue) => ({
          id: issue._id.toString(),
          reporterEmail,
          category: issue.category,
          status: issue.status,
        })),
        total,
      });
    },
  }),
}), { virtual: true });

describe('admin issues route', () => {
  let app;
  let db;

  beforeEach(() => {
    jest.resetModules();
    db = require('~/models');
    const router = require('../admin/issues');
    app = express();
    app.use(express.json());
    app.use('/api/admin/issues', router);
  });

  it('lists issue reports', async () => {
    db.findIssueReports.mockResolvedValue([
      {
        _id: { toString: () => 'issue-1' },
        user: { toString: () => '507f191e810c19729de860ea' },
        conversationId: 'convo-1',
        messageId: 'msg-1',
        category: 'bad_file_transformation',
        status: 'open',
        description: 'Spreadsheet removed the wrong column',
        model: 'claude-3-7-sonnet',
        endpoint: 'bedrock',
        createdAt: new Date('2026-04-17T12:00:00.000Z'),
        updatedAt: new Date('2026-04-17T12:00:00.000Z'),
      },
    ]);
    db.countIssueReports.mockResolvedValue(1);
    db.findUsers.mockResolvedValue([
      {
        _id: { toString: () => '507f191e810c19729de860ea' },
        name: 'Praneet Kotah',
        username: 'praneet.kotah',
        email: 'praneet.kotah@hermeus.com',
        avatar: '',
        role: 'ADMIN',
      },
    ]);

    const res = await request(app).get('/api/admin/issues?status=open').expect(200);

    expect(res.body.total).toBe(1);
    expect(res.body.issues[0]).toMatchObject({
      id: 'issue-1',
      reporterEmail: 'praneet.kotah@hermeus.com',
      category: 'bad_file_transformation',
      status: 'open',
    });
  });
});
