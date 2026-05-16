import enterpriseMemoryEntitySchema from '~/schema/enterpriseMemoryEntity';
import { applyTenantIsolation } from '~/models/plugins/tenantIsolation';
import type { IEnterpriseMemoryEntity } from '~/schema/enterpriseMemoryEntity';

export function createEnterpriseMemoryEntityModel(mongoose: typeof import('mongoose')) {
  applyTenantIsolation(enterpriseMemoryEntitySchema);
  return (
    mongoose.models.EnterpriseMemoryEntity ||
    mongoose.model<IEnterpriseMemoryEntity>('EnterpriseMemoryEntity', enterpriseMemoryEntitySchema)
  );
}
