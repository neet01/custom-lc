import documentSchema from '~/schema/document';
import { applyTenantIsolation } from '~/models/plugins/tenantIsolation';
import type { IDocumentRecord } from '~/schema/document';

export function createDocumentModel(mongoose: typeof import('mongoose')) {
  applyTenantIsolation(documentSchema);
  return mongoose.models.Document || mongoose.model<IDocumentRecord>('Document', documentSchema);
}
