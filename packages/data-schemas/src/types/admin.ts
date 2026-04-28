import type { PrincipalType, PrincipalModel, TCustomConfig } from 'librechat-data-provider';
import type { SystemCapabilities } from '~/admin/capabilities';

/* ── Capability types ───────────────────────────────────────────────── */

/** Base capabilities derived from the SystemCapabilities constant. */
export type BaseSystemCapability = (typeof SystemCapabilities)[keyof typeof SystemCapabilities];

/** Principal types that can receive config overrides. */
export type ConfigAssignTarget = 'user' | 'group' | 'role';

/** Top-level keys of the configSchema from librechat.yaml. */
export type ConfigSection = string & keyof TCustomConfig;

/** Section-level config capabilities derived from configSchema keys. */
type ConfigSectionCapability = `manage:configs:${ConfigSection}` | `read:configs:${ConfigSection}`;

/** Principal-scoped config assignment capabilities. */
type ConfigAssignCapability = `assign:configs:${ConfigAssignTarget}`;

/**
 * Union of all valid capability strings:
 * - Base capabilities from SystemCapabilities
 * - Section-level config capabilities (manage:configs:<section>, read:configs:<section>)
 * - Config assignment capabilities (assign:configs:<user|group|role>)
 */
export type SystemCapability =
  | BaseSystemCapability
  | ConfigSectionCapability
  | ConfigAssignCapability;

/** UI grouping of capabilities for the admin panel's capability editor. */
export type CapabilityCategory = {
  key: string;
  labelKey: string;
  capabilities: BaseSystemCapability[];
};

/* ── Admin API response types ───────────────────────────────────────── */

/** Config document as returned by the admin API (no Mongoose internals). */
export type AdminConfig = {
  _id: string;
  principalType: PrincipalType;
  principalId: string;
  principalModel: PrincipalModel;
  priority: number;
  overrides: Partial<TCustomConfig>;
  isActive: boolean;
  configVersion: number;
  tenantId?: string;
  createdAt?: string;
  updatedAt?: string;
};

export type AdminConfigListResponse = {
  configs: AdminConfig[];
};

export type AdminConfigResponse = {
  config: AdminConfig;
};

export type AdminConfigDeleteResponse = {
  success: boolean;
};

/** Audit action types for grant changes. */
export type AuditAction = 'grant_assigned' | 'grant_removed';

/** SystemGrant document as returned by the admin API. */
export type AdminSystemGrant = {
  id: string;
  principalType: PrincipalType;
  principalId: string;
  capability: string;
  grantedBy?: string;
  grantedAt: string;
  expiresAt?: string;
};

/** Audit log entry for grant changes as returned by the admin API. */
export type AdminAuditLogEntry = {
  id: string;
  action: AuditAction;
  actorId: string;
  actorName: string;
  targetPrincipalType: PrincipalType;
  targetPrincipalId: string;
  targetName: string;
  capability: string;
  timestamp: string;
};

/** Group as returned by the admin API. */
export type AdminGroup = {
  id: string;
  name: string;
  description: string;
  memberCount: number;
  topMembers: { name: string }[];
  isActive: boolean;
};

/** Member entry as returned by the admin API for group/role membership lists. */
export type AdminMember = {
  userId: string;
  name: string;
  email: string;
  avatarUrl?: string;
  joinedAt?: string;
};

/** Full user info returned by the admin user list endpoint. */
export type AdminUserListItem = {
  id: string;
  name: string;
  username: string;
  email: string;
  avatar: string;
  role: string;
  provider: string;
  tokenCredits: number;
  createdAt?: string;
  updatedAt?: string;
};

/** Minimal user info returned by user search endpoints. */
export type AdminUserSearchResult = {
  id: string;
  name: string;
  email: string;
  username?: string;
  avatarUrl?: string;
};

export type AdminUsageListItem = {
  id: string;
  userId: string;
  conversationId: string;
  messageId?: string;
  requestId?: string;
  sessionId?: string;
  model?: string;
  provider?: string;
  endpoint?: string;
  context?: string;
  source?: string;
  inputTokens: number;
  outputTokens: number;
  totalTokens: number;
  cacheCreationTokens?: number;
  cacheReadTokens?: number;
  latencyMs?: number;
  createdAt?: string;
  updatedAt?: string;
};

export type AdminUsageOverview = {
  requestCount: number;
  inputTokens: number;
  outputTokens: number;
  totalTokens: number;
  cacheCreationTokens: number;
  cacheReadTokens: number;
  avgLatencyMs: number | null;
  activeUsers: number;
  firstSeenAt?: string;
  lastSeenAt?: string;
  windowStart?: string;
  windowEnd?: string;
};

export type AdminUsageSummaryItem = {
  userId: string;
  name: string;
  username: string;
  email: string;
  avatar: string;
  role: string;
  provider: string;
  requestCount: number;
  inputTokens: number;
  outputTokens: number;
  totalTokens: number;
  cacheCreationTokens: number;
  cacheReadTokens: number;
  avgLatencyMs: number | null;
  firstSeenAt?: string;
  lastSeenAt?: string;
};

export type AdminIssueReportItem = {
  id: string;
  userId: string;
  reporterName: string;
  reporterEmail: string;
  reporterAvatar: string;
  reporterRole: string;
  conversationId: string;
  messageId: string;
  category:
    | 'bad_response'
    | 'faulty_mcp_tool'
    | 'bad_file_transformation'
    | 'timeout_or_error'
    | 'auth_or_permissions'
    | 'other';
  status: 'open' | 'triaged' | 'resolved';
  description?: string;
  model?: string;
  endpoint?: string;
  messagePreview?: string;
  error?: boolean;
  fileIds?: string[];
  toolName?: string;
  mcpServer?: string;
  createdAt?: string;
  updatedAt?: string;
};
