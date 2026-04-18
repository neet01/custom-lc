const request = require('supertest');
const express = require('express');

jest.mock('~/server/middleware', () => ({
  requireJwtAuth: (req, _res, next) => {
    req.user = { id: '507f191e810c19729de860ea', role: 'ADMIN' };
    next();
  },
}));

jest.mock('~/models', () => ({
  createIssueReport: jest.fn(),
}));

jest.mock('@librechat/api', () => ({
  createIssueHandlers: ({ createIssueReport }) => ({
    reportIssue: async (req, res) => {
      if (!req.body?.conversationId || !req.body?.messageId || !req.body?.category) {
        return res.status(400).json({ error: 'conversationId, messageId, and category are required' });
      }

      const issue = await createIssueReport({
        user: req.user.id,
        conversationId: req.body.conversationId,
        messageId: req.body.messageId,
        category: req.body.category,
      });

      return res.status(201).json({
        issue: {
          id: issue._id.toString(),
          category: issue.category,
          status: issue.status,
        },
      });
    },
  }),
}), { virtual: true });

describe('issues route', () => {
  let app;
  let db;

  beforeEach(() => {
    jest.resetModules();
    db = require('~/models');
    const router = require('../issues');
    app = express();
    app.use(express.json());
    app.use('/api/issues', router);
  });

  it('creates an issue report', async () => {
    db.createIssueReport.mockResolvedValue({
      _id: { toString: () => 'issue-1' },
      user: { toString: () => '507f191e810c19729de860ea' },
      conversationId: 'convo-1',
      messageId: 'msg-1',
      category: 'faulty_mcp_tool',
      status: 'open',
      description: 'Tool returned the wrong ticket list',
      model: 'claude-3-7-sonnet',
      endpoint: 'bedrock',
      createdAt: new Date('2026-04-17T12:00:00.000Z'),
      updatedAt: new Date('2026-04-17T12:00:00.000Z'),
    });

    const res = await request(app)
      .post('/api/issues')
      .send({
        conversationId: 'convo-1',
        messageId: 'msg-1',
        category: 'faulty_mcp_tool',
        description: 'Tool returned the wrong ticket list',
        model: 'claude-3-7-sonnet',
        endpoint: 'bedrock',
      })
      .expect(201);

    expect(db.createIssueReport).toHaveBeenCalledWith(
      expect.objectContaining({
        user: '507f191e810c19729de860ea',
        conversationId: 'convo-1',
        messageId: 'msg-1',
        category: 'faulty_mcp_tool',
      }),
    );
    expect(res.body.issue).toMatchObject({
      id: 'issue-1',
      category: 'faulty_mcp_tool',
      status: 'open',
    });
  });

  it('validates required fields', async () => {
    await request(app).post('/api/issues').send({ category: 'other' }).expect(400);
  });
});
