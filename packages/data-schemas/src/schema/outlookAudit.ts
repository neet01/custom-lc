import mongoose, { Schema, Document, Types } from 'mongoose';

export type OutlookAuditAction =
  | 'mailbox_listed'
  | 'calendar_viewed'
  | 'calendar_event_created'
  | 'calendar_event_updated'
  | 'calendar_event_deleted'
  | 'message_viewed'
  | 'attachment_downloaded'
  | 'message_deleted'
  | 'message_analyzed'
  | 'draft_created'
  | 'meeting_slots_proposed'
  | 'meeting_created';

export type OutlookAuditStatus = 'success' | 'failure';

export interface IOutlookAudit extends Document {
  user: Types.ObjectId;
  action: OutlookAuditAction;
  status: OutlookAuditStatus;
  graphMessageId?: string;
  graphConversationId?: string;
  graphDraftId?: string;
  errorCode?: string;
  errorMessage?: string;
  metadata?: Record<string, unknown>;
  tenantId?: string;
  createdAt?: Date;
  updatedAt?: Date;
}

const outlookAuditSchema = new Schema<IOutlookAudit>(
  {
    user: {
      type: mongoose.Schema.Types.ObjectId,
      ref: 'User',
      required: true,
      index: true,
    },
    action: {
      type: String,
      enum: [
        'mailbox_listed',
        'calendar_viewed',
        'calendar_event_created',
        'calendar_event_updated',
        'calendar_event_deleted',
        'message_viewed',
        'attachment_downloaded',
        'message_deleted',
        'message_analyzed',
        'draft_created',
        'meeting_slots_proposed',
        'meeting_created',
      ],
      required: true,
      index: true,
    },
    status: {
      type: String,
      enum: ['success', 'failure'],
      required: true,
      index: true,
    },
    graphMessageId: {
      type: String,
      index: true,
    },
    graphConversationId: {
      type: String,
      index: true,
    },
    graphDraftId: {
      type: String,
      index: true,
    },
    errorCode: {
      type: String,
      maxlength: 120,
    },
    errorMessage: {
      type: String,
      maxlength: 500,
    },
    metadata: {
      type: mongoose.Schema.Types.Mixed,
      default: undefined,
    },
    tenantId: {
      type: String,
      index: true,
    },
  },
  {
    timestamps: true,
  },
);

outlookAuditSchema.index({ user: 1, createdAt: -1 });
outlookAuditSchema.index({ action: 1, createdAt: -1 });
outlookAuditSchema.index({ status: 1, createdAt: -1 });
outlookAuditSchema.index({ tenantId: 1, createdAt: -1 });

export default outlookAuditSchema;
