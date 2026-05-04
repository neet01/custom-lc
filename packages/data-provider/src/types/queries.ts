import type { InfiniteData } from '@tanstack/react-query';
import type * as p from '../accessPermissions';
import type * as a from '../types/agents';
import type * as s from '../schemas';
import type * as t from '../types';

export type Conversation = {
  id: string;
  createdAt: number;
  participants: string[];
  lastMessage: string;
  conversations: s.TConversation[];
};

export type ConversationListParams = {
  cursor?: string;
  isArchived?: boolean;
  sortBy?: 'title' | 'createdAt' | 'updatedAt';
  sortDirection?: 'asc' | 'desc';
  tags?: string[];
  search?: string;
};

export type MinimalConversation = Pick<
  s.TConversation,
  'conversationId' | 'endpoint' | 'title' | 'createdAt' | 'updatedAt' | 'user'
>;

export type ConversationListResponse = {
  conversations: MinimalConversation[];
  nextCursor: string | null;
};

export type ConversationData = InfiniteData<ConversationListResponse>;
export type ConversationUpdater = (
  data: ConversationData,
  conversation: s.TConversation,
) => ConversationData;

/* Messages */
export type MessagesListParams = {
  cursor?: string | null;
  sortBy?: 'endpoint' | 'createdAt' | 'updatedAt';
  sortDirection?: 'asc' | 'desc';
  pageSize?: number;
  conversationId?: string;
  messageId?: string;
  search?: string;
};

export type MessagesListResponse = {
  messages: s.TMessage[];
  nextCursor: string | null;
};

/* Shared Links */
export type SharedMessagesResponse = Omit<s.TSharedLink, 'messages'> & {
  messages: s.TMessage[];
};

export interface SharedLinksListParams {
  pageSize: number;
  isPublic: boolean;
  sortBy: 'title' | 'createdAt';
  sortDirection: 'asc' | 'desc';
  search?: string;
  cursor?: string;
}

export type SharedLinkItem = {
  shareId: string;
  title: string;
  isPublic: boolean;
  createdAt: Date;
  conversationId: string;
};

export interface SharedLinksResponse {
  links: SharedLinkItem[];
  nextCursor: string | null;
  hasNextPage: boolean;
}

export interface SharedLinkQueryData {
  pages: SharedLinksResponse[];
  pageParams: (string | null)[];
}

export type AllPromptGroupsFilterRequest = {
  category: string;
  pageNumber: string;
  pageSize: string | number;
  before?: string | null;
  after?: string | null;
  order?: 'asc' | 'desc';
  name?: string;
  author?: string;
};

export type AllPromptGroupsResponse = t.TPromptGroup[];

export type ConversationTagsResponse = s.TConversationTag[];

/* MCP Types */
export type MCPTool = {
  name: string;
  pluginKey: string;
  description: string;
};

export type MCPServer = {
  name: string;
  icon: string;
  authenticated: boolean;
  authConfig: s.TPluginAuthConfig[];
  tools: MCPTool[];
};

export type MCPServersResponse = {
  servers: Record<string, MCPServer>;
};

export type VerifyToolAuthParams = { toolId: string };
export type VerifyToolAuthResponse = {
  authenticated: boolean;
  message?: string | s.AuthType;
  authTypes?: [string, s.AuthType][];
};

export type GetToolCallParams = { conversationId: string };
export type ToolCallResults = a.ToolCallResult[];

/* Memories */
export type TUserMemory = {
  key: string;
  value: string;
  updated_at: string;
  tokenCount?: number;
};

export type MemoriesResponse = {
  memories: TUserMemory[];
  totalTokens: number;
  tokenLimit: number | null;
  usagePercentage: number | null;
};

export type PrincipalSearchParams = {
  q: string;
  limit?: number;
  types?: Array<p.PrincipalType.USER | p.PrincipalType.GROUP | p.PrincipalType.ROLE>;
};

export type PrincipalSearchResponse = {
  query: string;
  limit: number;
  types?: Array<p.PrincipalType.USER | p.PrincipalType.GROUP | p.PrincipalType.ROLE>;
  results: p.TPrincipalSearchResult[];
  count: number;
  sources: {
    local: number;
    entra: number;
  };
};

export type AccessRole = {
  accessRoleId: p.AccessRoleIds;
  name: string;
  description: string;
  permBits: number;
};

export type AccessRolesResponse = AccessRole[];

export type ListRolesResponse = {
  roles: Array<{ _id?: string; name: string; description?: string }>;
  total: number;
  limit: number;
  offset?: number;
};

export interface MCPServerStatus {
  requiresOAuth: boolean;
  connectionState: 'disconnected' | 'connecting' | 'connected' | 'error';
}

export interface MCPConnectionStatusResponse {
  success: boolean;
  connectionStatus: Record<string, MCPServerStatus>;
}

export interface MCPServerConnectionStatusResponse {
  success: boolean;
  serverName: string;
  requiresOAuth: boolean;
  connectionStatus: 'disconnected' | 'connecting' | 'connected' | 'error';
}

export interface MCPAuthValuesResponse {
  success: boolean;
  serverName: string;
  authValueFlags: Record<string, boolean>;
}

/**
 * User Favorites — pinned agents, models, and model specs.
 * Exactly one variant should be set per entry; exclusivity is enforced
 * server-side in FavoritesController. Shape is loose for state-update ergonomics.
 */
export type TUserFavorite = {
  agentId?: string;
  model?: string;
  endpoint?: string;
  spec?: string;
};

/* SharePoint Graph API Token */
export type GraphTokenParams = {
  scopes: string;
};

export type GraphTokenResponse = {
  access_token: string;
  token_type: string;
  expires_in: number;
  scope: string;
};

export type AdminUsersListParams = {
  limit?: number;
  offset?: number;
};

export type AdminUsageListParams = {
  user_id?: string;
  conversation_id?: string;
  context?: string;
  source?: string;
  limit?: number;
  offset?: number;
};

export type AdminUsageSummaryParams = AdminUsageListParams & {
  days?: number;
};

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

export type AdminUsersListResponse = {
  users: AdminUserListItem[];
  total: number;
  limit: number;
  offset?: number;
};

export type AdminUpdateUserBalanceRequest = {
  tokenCredits: number;
};

export type AdminUpdateUserBalanceResponse = {
  user: AdminUserListItem;
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

export type AdminUsageListResponse = {
  usage: AdminUsageListItem[];
  total: number;
  limit: number;
  offset?: number;
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

export type AdminUsageSummaryResponse = {
  overview: AdminUsageOverview;
  users: AdminUsageSummaryItem[];
  total: number;
  limit: number;
  offset?: number;
  days: number;
};

export type IssueReportCategory =
  | 'bad_response'
  | 'faulty_mcp_tool'
  | 'bad_file_transformation'
  | 'timeout_or_error'
  | 'auth_or_permissions'
  | 'other';

export type IssueReportStatus = 'open' | 'triaged' | 'resolved';

export type IssueReportCreateRequest = {
  conversationId: string;
  messageId: string;
  category: IssueReportCategory;
  description?: string;
  model?: string;
  endpoint?: string;
  messagePreview?: string;
  error?: boolean;
  fileIds?: string[];
  toolName?: string;
  mcpServer?: string;
};

export type IssueReportCreateResponse = {
  issue: {
    id: string;
    userId: string;
    conversationId: string;
    messageId: string;
    category: IssueReportCategory;
    status: IssueReportStatus;
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
};

export type AdminIssuesListParams = {
  user_id?: string;
  conversation_id?: string;
  category?: IssueReportCategory;
  status?: IssueReportStatus;
  limit?: number;
  offset?: number;
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
  category: IssueReportCategory;
  status: IssueReportStatus;
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

export type AdminIssuesListResponse = {
  issues: AdminIssueReportItem[];
  total: number;
  limit: number;
  offset?: number;
};

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

export type AdminOutlookAuditListParams = {
  user_id?: string;
  action?: OutlookAuditAction;
  status?: OutlookAuditStatus;
  message_id?: string;
  limit?: number;
  offset?: number;
};

export type AdminOutlookAuditItem = {
  id: string;
  userId: string;
  actorName: string;
  actorEmail: string;
  actorAvatar: string;
  actorRole: string;
  action: OutlookAuditAction;
  status: OutlookAuditStatus;
  graphMessageId?: string;
  graphConversationId?: string;
  graphDraftId?: string;
  errorCode?: string;
  errorMessage?: string;
  metadata?: Record<string, unknown>;
  createdAt?: string;
  updatedAt?: string;
};

export type AdminOutlookAuditListResponse = {
  audits: AdminOutlookAuditItem[];
  total: number;
  limit: number;
  offset?: number;
};
