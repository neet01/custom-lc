import slackIdentityLinkSchema from '~/schema/slackIdentityLink';
import { applyTenantIsolation } from '~/models/plugins/tenantIsolation';
import type { ISlackIdentityLink } from '~/schema/slackIdentityLink';

export function createSlackIdentityLinkModel(mongoose: typeof import('mongoose')) {
  applyTenantIsolation(slackIdentityLinkSchema);
  return (
    mongoose.models.SlackIdentityLink ||
    mongoose.model<ISlackIdentityLink>('SlackIdentityLink', slackIdentityLinkSchema)
  );
}
