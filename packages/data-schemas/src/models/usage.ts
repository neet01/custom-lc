import usageSchema from '~/schema/usage';
import { applyTenantIsolation } from '~/models/plugins/tenantIsolation';
import type { IUsage } from '~/schema/usage';

export function createUsageModel(mongoose: typeof import('mongoose')) {
  applyTenantIsolation(usageSchema);
  return mongoose.models.Usage || mongoose.model<IUsage>('Usage', usageSchema);
}
