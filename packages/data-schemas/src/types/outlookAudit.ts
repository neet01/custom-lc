import type { OutlookAuditAction, OutlookAuditStatus } from '~/schema/outlookAudit';

export type { OutlookAuditAction, OutlookAuditStatus };

export interface OutlookAuditData {
  user: string;
  action: OutlookAuditAction;
  status: OutlookAuditStatus;
  graphMessageId?: string;
  graphConversationId?: string;
  graphDraftId?: string;
  errorCode?: string;
  errorMessage?: string;
  metadata?: Record<string, unknown>;
}
