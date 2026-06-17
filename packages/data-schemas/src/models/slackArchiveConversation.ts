import slackArchiveConversationSchema from '~/schema/slackArchiveConversation';
import { applyTenantIsolation } from '~/models/plugins/tenantIsolation';
import type { ISlackArchiveConversation } from '~/schema/slackArchiveConversation';

export function createSlackArchiveConversationModel(mongoose: typeof import('mongoose')) {
  applyTenantIsolation(slackArchiveConversationSchema);
  return (
    mongoose.models.SlackArchiveConversation ||
    mongoose.model<ISlackArchiveConversation>('SlackArchiveConversation', slackArchiveConversationSchema)
  );
}
