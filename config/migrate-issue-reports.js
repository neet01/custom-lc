require('dotenv').config();

const mongoose = require('mongoose');
const { logger } = require('@librechat/data-schemas');
const { createModels } = require('@librechat/data-schemas');
const { connectDb } = require('../api/db/connect');

async function migrateIssueReports() {
  await connectDb();
  createModels(mongoose);

  const IssueReport = mongoose.models.IssueReport;
  if (!IssueReport) {
    throw new Error('IssueReport model failed to initialize');
  }

  await IssueReport.createCollection().catch((error) => {
    if (error?.codeName !== 'NamespaceExists') {
      throw error;
    }
  });
  await IssueReport.syncIndexes();

  logger.info('[migrate-issue-reports] IssueReport collection and indexes are ready');
}

if (require.main === module) {
  migrateIssueReports()
    .then(async () => {
      await mongoose.disconnect();
      process.exit(0);
    })
    .catch(async (error) => {
      logger.error('[migrate-issue-reports] Migration failed', error);
      await mongoose.disconnect().catch(() => undefined);
      process.exit(1);
    });
}

module.exports = { migrateIssueReports };
