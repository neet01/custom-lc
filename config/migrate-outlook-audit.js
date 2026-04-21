require('dotenv').config();

const mongoose = require('mongoose');
const { logger } = require('@librechat/data-schemas');
const { createModels } = require('@librechat/data-schemas');
const { connectDb } = require('../api/db/connect');

async function migrateOutlookAudit() {
  await connectDb();
  createModels(mongoose);

  const OutlookAudit = mongoose.models.OutlookAudit;
  if (!OutlookAudit) {
    throw new Error('OutlookAudit model failed to initialize');
  }

  await OutlookAudit.createCollection().catch((error) => {
    if (error?.codeName !== 'NamespaceExists') {
      throw error;
    }
  });
  await OutlookAudit.syncIndexes();

  logger.info('[migrate-outlook-audit] OutlookAudit collection and indexes are ready');
}

if (require.main === module) {
  migrateOutlookAudit()
    .then(async () => {
      await mongoose.disconnect();
      process.exit(0);
    })
    .catch(async (error) => {
      logger.error('[migrate-outlook-audit] Migration failed', error);
      await mongoose.disconnect().catch(() => undefined);
      process.exit(1);
    });
}

module.exports = { migrateOutlookAudit };
