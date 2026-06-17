import slackArchiveMessageSchema from '~/schema/slackArchiveMessage';
import { applyTenantIsolation } from '~/models/plugins/tenantIsolation';
import type { ISlackArchiveMessage } from '~/schema/slackArchiveMessage';

export function createSlackArchiveMessageModel(mongoose: typeof import('mongoose')) {
  applyTenantIsolation(slackArchiveMessageSchema);
  return (
    mongoose.models.SlackArchiveMessage ||
    mongoose.model<ISlackArchiveMessage>('SlackArchiveMessage', slackArchiveMessageSchema)
  );
}
