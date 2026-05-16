import enterpriseMemoryJobSchema from '~/schema/enterpriseMemoryJob';
import { applyTenantIsolation } from '~/models/plugins/tenantIsolation';
import type { IEnterpriseMemoryJob } from '~/schema/enterpriseMemoryJob';

export function createEnterpriseMemoryJobModel(mongoose: typeof import('mongoose')) {
  applyTenantIsolation(enterpriseMemoryJobSchema);
  return (
    mongoose.models.EnterpriseMemoryJob ||
    mongoose.model<IEnterpriseMemoryJob>('EnterpriseMemoryJob', enterpriseMemoryJobSchema)
  );
}
