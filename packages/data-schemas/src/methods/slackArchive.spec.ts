import mongoose from 'mongoose';
import { MongoMemoryServer } from 'mongodb-memory-server';
import { createModels } from '~/models';
import { createSlackArchiveMethods } from './slackArchive';

jest.mock('~/config/winston', () => ({
  error: jest.fn(),
  warn: jest.fn(),
  info: jest.fn(),
  debug: jest.fn(),
}));

let mongoServer: MongoMemoryServer;
let methods: ReturnType<typeof createSlackArchiveMethods>;

beforeAll(async () => {
  mongoServer = await MongoMemoryServer.create();
  Object.assign(mongoose.models, createModels(mongoose));
  methods = createSlackArchiveMethods(mongoose);
  await mongoose.connect(mongoServer.getUri());
});

afterAll(async () => {
  await mongoose.disconnect();
  await mongoServer.stop();
});

beforeEach(async () => {
  await mongoose.models.SlackArchiveSyncLease.deleteMany({});
});

describe('Slack archive sync leases', () => {
  it('acquires a new lease without conflicting update operators', async () => {
    const userId = new mongoose.Types.ObjectId();
    const leaseExpiresAt = new Date(Date.now() + 5 * 60 * 1000);

    const lease = await methods.acquireSlackArchiveSyncLease({
      leaseKey: `slack:user:${userId.toString()}`,
      leaseType: 'user',
      ownerToken: 'owner-token-1',
      user: userId.toString(),
      leaseExpiresAt,
      lastHeartbeatAt: new Date(),
    });

    expect(lease).toBeTruthy();
    expect(lease?.leaseKey).toBe(`slack:user:${userId.toString()}`);
    expect(lease?.ownerToken).toBe('owner-token-1');
    expect(lease?.leaseExpiresAt.toISOString()).toBe(leaseExpiresAt.toISOString());
  });

  it('does not steal an active lease owned by another token', async () => {
    const userId = new mongoose.Types.ObjectId();
    const leaseKey = `slack:user:${userId.toString()}`;

    await methods.acquireSlackArchiveSyncLease({
      leaseKey,
      leaseType: 'user',
      ownerToken: 'owner-token-1',
      user: userId.toString(),
      leaseExpiresAt: new Date(Date.now() + 5 * 60 * 1000),
    });

    const lease = await methods.acquireSlackArchiveSyncLease({
      leaseKey,
      leaseType: 'user',
      ownerToken: 'owner-token-2',
      user: userId.toString(),
      leaseExpiresAt: new Date(Date.now() + 5 * 60 * 1000),
    });

    expect(lease).toBeNull();
  });
});
