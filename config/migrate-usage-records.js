require('dotenv').config();

const mongoose = require('mongoose');
const { logger } = require('@librechat/data-schemas');
const { createModels } = require('@librechat/data-schemas');
const { connectDb } = require('../api/db/connect');

async function migrateUsageRecords() {
  await connectDb();
  createModels(mongoose);

  const Usage = mongoose.models.Usage;
  if (!Usage) {
    throw new Error('Usage model failed to initialize');
  }

  await Usage.createCollection().catch((error) => {
    if (error?.codeName !== 'NamespaceExists') {
      throw error;
    }
  });
  await Usage.syncIndexes();

  logger.info('[migrate-usage-records] Usage collection and indexes are ready');
}

if (require.main === module) {
  migrateUsageRecords()
    .then(async () => {
      await mongoose.disconnect();
      process.exit(0);
    })
    .catch(async (error) => {
      logger.error('[migrate-usage-records] Migration failed', error);
      await mongoose.disconnect().catch(() => undefined);
      process.exit(1);
    });
}

module.exports = { migrateUsageRecords };
