import teamsArchiveConversationSchema from '~/schema/teamsArchiveConversation';
import { applyTenantIsolation } from '~/models/plugins/tenantIsolation';
import type { ITeamsArchiveConversation } from '~/schema/teamsArchiveConversation';

export function createTeamsArchiveConversationModel(mongoose: typeof import('mongoose')) {
  applyTenantIsolation(teamsArchiveConversationSchema);
  return (
    mongoose.models.TeamsArchiveConversation ||
    mongoose.model<ITeamsArchiveConversation>(
      'TeamsArchiveConversation',
      teamsArchiveConversationSchema,
    )
  );
}
