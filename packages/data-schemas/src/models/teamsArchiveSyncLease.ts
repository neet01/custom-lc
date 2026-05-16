import teamsArchiveSyncLeaseSchema from '~/schema/teamsArchiveSyncLease';
import { applyTenantIsolation } from './plugins/tenantIsolation';
import type { ITeamsArchiveSyncLease } from '~/schema/teamsArchiveSyncLease';

export function createTeamsArchiveSyncLeaseModel(mongoose: typeof import('mongoose')) {
  applyTenantIsolation(teamsArchiveSyncLeaseSchema);
  return (
    mongoose.models.TeamsArchiveSyncLease ||
    mongoose.model<ITeamsArchiveSyncLease>('TeamsArchiveSyncLease', teamsArchiveSyncLeaseSchema)
  );
}
