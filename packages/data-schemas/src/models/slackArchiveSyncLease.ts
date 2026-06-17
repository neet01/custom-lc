import slackArchiveSyncLeaseSchema from '~/schema/slackArchiveSyncLease';
import { applyTenantIsolation } from '~/models/plugins/tenantIsolation';
import type { ISlackArchiveSyncLease } from '~/schema/slackArchiveSyncLease';

export function createSlackArchiveSyncLeaseModel(mongoose: typeof import('mongoose')) {
  applyTenantIsolation(slackArchiveSyncLeaseSchema);
  return (
    mongoose.models.SlackArchiveSyncLease ||
    mongoose.model<ISlackArchiveSyncLease>('SlackArchiveSyncLease', slackArchiveSyncLeaseSchema)
  );
}
