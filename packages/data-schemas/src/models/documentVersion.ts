import documentVersionSchema from '~/schema/documentVersion';
import { applyTenantIsolation } from '~/models/plugins/tenantIsolation';
import type { IDocumentVersion } from '~/schema/documentVersion';

export function createDocumentVersionModel(mongoose: typeof import('mongoose')) {
  applyTenantIsolation(documentVersionSchema);
  return (
    mongoose.models.DocumentVersion ||
    mongoose.model<IDocumentVersion>('DocumentVersion', documentVersionSchema)
  );
}
