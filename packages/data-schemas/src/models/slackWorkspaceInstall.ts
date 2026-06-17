import slackWorkspaceInstallSchema from '~/schema/slackWorkspaceInstall';
import { applyTenantIsolation } from '~/models/plugins/tenantIsolation';
import type { ISlackWorkspaceInstall } from '~/schema/slackWorkspaceInstall';

export function createSlackWorkspaceInstallModel(mongoose: typeof import('mongoose')) {
  applyTenantIsolation(slackWorkspaceInstallSchema);
  return (
    mongoose.models.SlackWorkspaceInstall ||
    mongoose.model<ISlackWorkspaceInstall>('SlackWorkspaceInstall', slackWorkspaceInstallSchema)
  );
}
