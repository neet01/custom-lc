import enterpriseMemoryChunkSchema from '~/schema/enterpriseMemoryChunk';
import { applyTenantIsolation } from '~/models/plugins/tenantIsolation';
import type { IEnterpriseMemoryChunk } from '~/schema/enterpriseMemoryChunk';

export function createEnterpriseMemoryChunkModel(mongoose: typeof import('mongoose')) {
  applyTenantIsolation(enterpriseMemoryChunkSchema);
  return (
    mongoose.models.EnterpriseMemoryChunk ||
    mongoose.model<IEnterpriseMemoryChunk>('EnterpriseMemoryChunk', enterpriseMemoryChunkSchema)
  );
}
