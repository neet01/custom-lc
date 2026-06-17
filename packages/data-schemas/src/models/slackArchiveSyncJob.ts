import slackArchiveSyncJobSchema from '~/schema/slackArchiveSyncJob';
import { applyTenantIsolation } from '~/models/plugins/tenantIsolation';
import type { ISlackArchiveSyncJob } from '~/schema/slackArchiveSyncJob';

export function createSlackArchiveSyncJobModel(mongoose: typeof import('mongoose')) {
  applyTenantIsolation(slackArchiveSyncJobSchema);
  return (
    mongoose.models.SlackArchiveSyncJob ||
    mongoose.model<ISlackArchiveSyncJob>('SlackArchiveSyncJob', slackArchiveSyncJobSchema)
  );
}
