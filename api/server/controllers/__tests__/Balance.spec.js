jest.mock('~/models', () => ({
  findBalanceByUser: jest.fn(),
  upsertBalanceFields: jest.fn(),
}));

jest.mock('~/server/services/Config', () => ({
  getAppConfig: jest.fn(),
}));

jest.mock('@librechat/api', () => ({
  getBalanceConfig: jest.fn(),
}));

const balanceController = require('../Balance');
const { findBalanceByUser, upsertBalanceFields } = require('~/models');
const { getAppConfig } = require('~/server/services/Config');
const { getBalanceConfig } = require('@librechat/api');

describe('balanceController', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  function createReqRes(overrides = {}) {
    const req = {
      user: {
        id: 'user-123',
        role: 'ADMIN',
        tenantId: 'tenant-1',
      },
      ...overrides,
    };
    const json = jest.fn();
    const status = jest.fn().mockReturnValue({ json });
    const res = { status, json };
    return { req, res, status, json };
  }

  it('returns an existing balance record', async () => {
    findBalanceByUser.mockResolvedValue({
      _id: 'balance-1',
      tokenCredits: 90000,
      autoRefillEnabled: false,
    });

    const { req, res, status, json } = createReqRes();
    await balanceController(req, res);

    expect(status).toHaveBeenCalledWith(200);
    expect(json).toHaveBeenCalledWith({
      tokenCredits: 90000,
      autoRefillEnabled: false,
    });
    expect(upsertBalanceFields).not.toHaveBeenCalled();
  });

  it('initializes a balance record from config when missing', async () => {
    findBalanceByUser.mockResolvedValue(null);
    getAppConfig.mockResolvedValue({ balance: { enabled: true, startBalance: 100000 } });
    getBalanceConfig.mockReturnValue({ enabled: true, startBalance: 100000 });
    upsertBalanceFields.mockResolvedValue({
      _id: 'balance-2',
      user: 'user-123',
      tokenCredits: 100000,
      autoRefillEnabled: false,
    });

    const { req, res, status, json } = createReqRes();
    await balanceController(req, res);

    expect(getAppConfig).toHaveBeenCalledWith({ role: 'ADMIN', tenantId: 'tenant-1' });
    expect(upsertBalanceFields).toHaveBeenCalledWith('user-123', {
      user: 'user-123',
      tokenCredits: 100000,
    });
    expect(status).toHaveBeenCalledWith(200);
    expect(json).toHaveBeenCalledWith({
      tokenCredits: 100000,
      autoRefillEnabled: false,
    });
  });
});
