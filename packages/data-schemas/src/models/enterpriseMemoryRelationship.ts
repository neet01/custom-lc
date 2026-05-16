import enterpriseMemoryRelationshipSchema from '~/schema/enterpriseMemoryRelationship';
import { applyTenantIsolation } from '~/models/plugins/tenantIsolation';
import type { IEnterpriseMemoryRelationship } from '~/schema/enterpriseMemoryRelationship';

export function createEnterpriseMemoryRelationshipModel(mongoose: typeof import('mongoose')) {
  applyTenantIsolation(enterpriseMemoryRelationshipSchema);
  return (
    mongoose.models.EnterpriseMemoryRelationship ||
    mongoose.model<IEnterpriseMemoryRelationship>(
      'EnterpriseMemoryRelationship',
      enterpriseMemoryRelationshipSchema,
    )
  );
}
