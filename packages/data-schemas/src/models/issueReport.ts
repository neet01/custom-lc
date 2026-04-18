import issueReportSchema from '~/schema/issueReport';
import { applyTenantIsolation } from '~/models/plugins/tenantIsolation';
import type { IIssueReport } from '~/schema/issueReport';

export function createIssueReportModel(mongoose: typeof import('mongoose')) {
  applyTenantIsolation(issueReportSchema);
  return (
    mongoose.models.IssueReport || mongoose.model<IIssueReport>('IssueReport', issueReportSchema)
  );
}
