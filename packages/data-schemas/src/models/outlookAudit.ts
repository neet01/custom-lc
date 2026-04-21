import outlookAuditSchema from '~/schema/outlookAudit';
import { applyTenantIsolation } from '~/models/plugins/tenantIsolation';
import type { IOutlookAudit } from '~/schema/outlookAudit';

export function createOutlookAuditModel(mongoose: typeof import('mongoose')) {
  applyTenantIsolation(outlookAuditSchema);
  return (
    mongoose.models.OutlookAudit ||
    mongoose.model<IOutlookAudit>('OutlookAudit', outlookAuditSchema)
  );
}
