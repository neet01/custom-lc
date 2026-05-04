import teamsArchiveMessageSchema from '~/schema/teamsArchiveMessage';
import { applyTenantIsolation } from '~/models/plugins/tenantIsolation';
import type { ITeamsArchiveMessage } from '~/schema/teamsArchiveMessage';

export function createTeamsArchiveMessageModel(mongoose: typeof import('mongoose')) {
  applyTenantIsolation(teamsArchiveMessageSchema);
  return (
    mongoose.models.TeamsArchiveMessage ||
    mongoose.model<ITeamsArchiveMessage>('TeamsArchiveMessage', teamsArchiveMessageSchema)
  );
}
