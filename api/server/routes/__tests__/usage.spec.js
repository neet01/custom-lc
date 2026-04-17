const express = require('express');
const request = require('supertest');
const mongoose = require('mongoose');
const { MongoMemoryServer } = require('mongodb-memory-server');
const { createModels, createMethods } = require('@librechat/data-schemas');
const { SystemRoles } = require('librechat-data-provider');

jest.mock('~/server/middleware', () => ({
  requireJwtAuth: (_req, _res, next) => next(),
}));

jest.mock('~/server/middleware/roles/capabilities', () => ({
  requireCapability: () => (_req, _res, next) => next(),
}));

let mongoServer;
let db;

beforeAll(async () => {
  process.env.USAGE_TRACKING_ENABLED = 'true';
  mongoServer = await MongoMemoryServer.create();
  await mongoose.connect(mongoServer.getUri());
  createModels(mongoose);
  db = createMethods(mongoose);
});

afterAll(async () => {
  delete process.env.USAGE_TRACKING_ENABLED;
  await mongoose.disconnect();
  await mongoServer.stop();
});

afterEach(async () => {
  const Usage = mongoose.models.Usage;
  const User = mongoose.models.User;
  await Usage.deleteMany({});
  await User.deleteMany({});
});

function createApp(user) {
  const router = require('../admin/usage');
  const app = express();
  app.use(express.json());
  app.use((req, _res, next) => {
    req.user = user;
    next();
  });
  app.use('/api/admin/usage', router);
  return app;
}

describe('Admin Usage Routes — Integration', () => {
  const adminUserId = new mongoose.Types.ObjectId();
  const adminUser = {
    _id: adminUserId,
    id: adminUserId.toString(),
    role: SystemRoles.ADMIN,
  };

  it('GET / returns persisted usage records', async () => {
    const targetUserId = new mongoose.Types.ObjectId();
    await db.createUsageRecords([
      {
        user: targetUserId.toString(),
        conversationId: 'convo-123',
        messageId: 'msg-123',
        requestId: 'msg-123',
        model: 'gpt-4o-mini',
        provider: 'openai',
        endpoint: 'agents',
        context: 'message',
        source: 'agent',
        inputTokens: 120,
        outputTokens: 30,
        totalTokens: 150,
        latencyMs: 250,
      },
    ]);

    const app = createApp(adminUser);
    const res = await request(app)
      .get(`/api/admin/usage?user_id=${targetUserId.toString()}`)
      .expect(200);

    expect(res.body.total).toBe(1);
    expect(res.body.usage).toEqual([
      expect.objectContaining({
        userId: targetUserId.toString(),
        conversationId: 'convo-123',
        messageId: 'msg-123',
        model: 'gpt-4o-mini',
        provider: 'openai',
        endpoint: 'agents',
        context: 'message',
        source: 'agent',
        inputTokens: 120,
        outputTokens: 30,
        totalTokens: 150,
        latencyMs: 250,
      }),
    ]);
  });

  it('GET /summary returns aggregated usage by user', async () => {
    const targetUserId = new mongoose.Types.ObjectId();
    await mongoose.models.User.create({
      _id: targetUserId,
      name: 'Usage Admin',
      username: 'usage-admin',
      email: 'usage@example.com',
      avatar: '',
      role: SystemRoles.ADMIN,
      provider: 'local',
    });

    await db.createUsageRecords([
      {
        user: targetUserId.toString(),
        conversationId: 'convo-1',
        inputTokens: 100,
        outputTokens: 20,
        totalTokens: 120,
        latencyMs: 100,
      },
      {
        user: targetUserId.toString(),
        conversationId: 'convo-2',
        inputTokens: 50,
        outputTokens: 10,
        totalTokens: 60,
        latencyMs: 200,
      },
    ]);

    const app = createApp(adminUser);
    const res = await request(app).get('/api/admin/usage/summary?days=30').expect(200);

    expect(res.body.overview).toEqual(
      expect.objectContaining({
        requestCount: 2,
        inputTokens: 150,
        outputTokens: 30,
        totalTokens: 180,
        activeUsers: 1,
      }),
    );
    expect(res.body.users).toEqual([
      expect.objectContaining({
        userId: targetUserId.toString(),
        name: 'Usage Admin',
        email: 'usage@example.com',
        requestCount: 2,
        totalTokens: 180,
        avgLatencyMs: 150,
      }),
    ]);
  });
});
