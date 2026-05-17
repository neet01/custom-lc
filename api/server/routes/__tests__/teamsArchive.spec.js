jest.mock(
  '@librechat/data-schemas',
  () => ({
    logger: {
      error: jest.fn(),
      info: jest.fn(),
    },
  }),
  { virtual: true },
);

jest.mock('~/server/middleware/requireJwtAuth', () => (req, _res, next) => next());
jest.mock('~/server/middleware', () => ({
  requireJwtAuth: (req, _res, next) => next(),
}));

jest.mock('~/server/services/TeamsArchiveService', () => ({
  TeamsArchiveSyncCancelledError: class TeamsArchiveSyncCancelledError extends Error {},
  getStatus: jest.fn(),
  getSyncStartAvailability: jest.fn(),
  syncUserArchive: jest.fn(),
  cancelRunningSync: jest.fn(),
  deleteUserArchive: jest.fn(),
  listConversations: jest.fn(),
  listConversationMessages: jest.fn(),
  searchMessages: jest.fn(),
}));

const TeamsArchiveService = require('~/server/services/TeamsArchiveService');
const router = require('../teamsArchive');

function findRouteHandler(method, path) {
  const layer = router.stack.find(
    (entry) => entry.route && entry.route.path === path && entry.route.methods[method],
  );
  if (!layer) {
    throw new Error(`Route ${method.toUpperCase()} ${path} not found`);
  }
  return layer.route.stack[layer.route.stack.length - 1].handle;
}

function createResponseRecorder() {
  return {
    statusCode: 200,
    body: undefined,
    status(code) {
      this.statusCode = code;
      return this;
    },
    json(payload) {
      this.body = payload;
      return this;
    },
  };
}

async function invoke(method, path, { body = {}, query = {}, params = {}, user } = {}) {
  const handler = findRouteHandler(method, path);
  const req = {
    body,
    query,
    params,
    user: user || { id: 'user-1' },
  };
  const res = createResponseRecorder();
  await handler(req, res);
  return res;
}

describe('Teams archive routes', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  it('GET /status returns the Teams archive status payload', async () => {
    TeamsArchiveService.getStatus.mockResolvedValue({
      enabled: true,
      conversationCount: 12,
    });

    const response = await invoke('get', '/status');

    expect(TeamsArchiveService.getStatus).toHaveBeenCalledWith({ id: 'user-1' });
    expect(response.statusCode).toBe(200);
    expect(response.body).toEqual({
      enabled: true,
      conversationCount: 12,
    });
  });

  it('POST /sync async returns 202 with already-running metadata', async () => {
    TeamsArchiveService.getSyncStartAvailability.mockResolvedValue({
      allowed: false,
      reason: 'already_running',
      syncJob: { id: 'sync-1' },
      message: 'Sync already running',
    });

    const response = await invoke('post', '/sync', {
      query: { async: 'true' },
    });

    expect(response.statusCode).toBe(202);
    expect(response.body).toMatchObject({
      accepted: true,
      status: 'running',
      alreadyRunning: true,
      syncJob: { id: 'sync-1' },
    });
    expect(TeamsArchiveService.syncUserArchive).not.toHaveBeenCalled();
  });

  it('POST /sync async starts a background sync when allowed', async () => {
    TeamsArchiveService.getSyncStartAvailability.mockResolvedValue({
      allowed: true,
    });
    TeamsArchiveService.syncUserArchive.mockResolvedValue({
      syncJob: { id: 'sync-2' },
    });

    const response = await invoke('post', '/sync', {
      query: { async: 'true' },
      body: { chatLimit: 100 },
    });

    expect(response.statusCode).toBe(202);
    expect(response.body).toMatchObject({
      accepted: true,
      status: 'running',
      mode: 'chats',
    });
    expect(TeamsArchiveService.getSyncStartAvailability).toHaveBeenCalledWith({ id: 'user-1' });
    expect(TeamsArchiveService.syncUserArchive).toHaveBeenCalledWith(
      { id: 'user-1' },
      { chatLimit: 100 },
    );
  });

  it('POST /sync runs inline sync when async is not requested', async () => {
    TeamsArchiveService.syncUserArchive.mockResolvedValue({
      syncJob: { id: 'sync-inline' },
      mode: 'chats',
    });

    const response = await invoke('post', '/sync', {
      body: { messagesPerChat: 1000 },
    });

    expect(response.statusCode).toBe(200);
    expect(TeamsArchiveService.syncUserArchive).toHaveBeenCalledWith(
      { id: 'user-1' },
      { messagesPerChat: 1000 },
    );
    expect(response.body).toMatchObject({
      syncJob: { id: 'sync-inline' },
      mode: 'chats',
    });
  });

  it('POST /reset requires confirm=true', async () => {
    const response = await invoke('post', '/reset', {
      body: {},
    });

    expect(response.statusCode).toBe(400);
    expect(response.body).toEqual({
      message: 'Teams archive reset requires confirm=true.',
    });
    expect(TeamsArchiveService.deleteUserArchive).not.toHaveBeenCalled();
  });

  it('POST /reset calls archive deletion when confirmed', async () => {
    TeamsArchiveService.deleteUserArchive.mockResolvedValue({
      deleted: true,
      counts: { conversations: 10 },
    });

    const response = await invoke('post', '/reset', {
      body: { confirm: true },
    });

    expect(response.statusCode).toBe(200);
    expect(TeamsArchiveService.deleteUserArchive).toHaveBeenCalledWith({ id: 'user-1' });
    expect(response.body).toEqual({
      deleted: true,
      counts: { conversations: 10 },
    });
  });

  it('POST /cancel returns cancel status', async () => {
    TeamsArchiveService.cancelRunningSync.mockResolvedValue({
      cancelled: true,
    });

    const response = await invoke('post', '/cancel');

    expect(response.statusCode).toBe(200);
    expect(TeamsArchiveService.cancelRunningSync).toHaveBeenCalledWith({ id: 'user-1' });
    expect(response.body).toEqual({ cancelled: true });
  });

  it('maps TeamsArchiveServiceError into an HTTP error payload', async () => {
    const error = new Error('Sync not allowed');
    error.name = 'TeamsArchiveServiceError';
    error.status = 409;
    error.details = { reason: 'concurrency_limit' };
    TeamsArchiveService.cancelRunningSync.mockRejectedValue(error);

    const response = await invoke('post', '/cancel');

    expect(response.statusCode).toBe(409);
    expect(response.body).toEqual({
      message: 'Sync not allowed',
      details: { reason: 'concurrency_limit' },
    });
  });
});
