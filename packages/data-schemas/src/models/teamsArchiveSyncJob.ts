import teamsArchiveSyncJobSchema from '~/schema/teamsArchiveSyncJob';
import { applyTenantIsolation } from '~/models/plugins/tenantIsolation';
import type { ITeamsArchiveSyncJob } from '~/schema/teamsArchiveSyncJob';

export function createTeamsArchiveSyncJobModel(mongoose: typeof import('mongoose')) {
  applyTenantIsolation(teamsArchiveSyncJobSchema);
  return (
    mongoose.models.TeamsArchiveSyncJob ||
    mongoose.model<ITeamsArchiveSyncJob>('TeamsArchiveSyncJob', teamsArchiveSyncJobSchema)
  );
}
