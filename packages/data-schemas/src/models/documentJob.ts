import documentJobSchema from '~/schema/documentJob';
import { applyTenantIsolation } from '~/models/plugins/tenantIsolation';
import type { IDocumentJob } from '~/schema/documentJob';

export function createDocumentJobModel(mongoose: typeof import('mongoose')) {
  applyTenantIsolation(documentJobSchema);
  return mongoose.models.DocumentJob || mongoose.model<IDocumentJob>('DocumentJob', documentJobSchema);
}
