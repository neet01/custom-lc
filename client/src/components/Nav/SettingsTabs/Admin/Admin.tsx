import React, { useMemo, useState } from 'react';
import { Spinner } from '@librechat/client';
import { SystemRoles } from 'librechat-data-provider';
import type {
  AdminIssueReportItem,
  AdminOutlookAuditItem,
  AdminUsageListItem,
  AdminUsageSummaryItem,
  AdminUserListItem,
} from 'librechat-data-provider';
import {
  useAdminIssuesQuery,
  useAdminOutlookAuditQuery,
  useAdminUsageQuery,
  useAdminUsageSummaryQuery,
  useAdminUsersQuery,
} from '~/data-provider';
import { useAuthContext } from '~/hooks';
import { formatDate } from '~/utils';

const DAY_OPTIONS = [7, 30, 90];

type DashboardUserRow = AdminUserListItem &
  Pick<
    AdminUsageSummaryItem,
    | 'requestCount'
    | 'inputTokens'
    | 'outputTokens'
    | 'totalTokens'
    | 'cacheCreationTokens'
    | 'cacheReadTokens'
    | 'avgLatencyMs'
    | 'firstSeenAt'
    | 'lastSeenAt'
  >;

function formatNumber(value: number | null | undefined) {
  return new Intl.NumberFormat().format(value ?? 0);
}

function formatLatency(value: number | null | undefined) {
  if (value == null) {
    return 'n/a';
  }

  return `${Math.round(value)} ms`;
}

function formatAuditAction(action: string) {
  return action
    .split('_')
    .map((part) => part.charAt(0).toUpperCase() + part.slice(1))
    .join(' ');
}

function MetricCard({
  label,
  value,
  detail,
}: {
  label: string;
  value: string;
  detail?: string;
}) {
  return (
    <div className="rounded-2xl border border-border-medium bg-surface-secondary p-4">
      <div className="text-xs uppercase tracking-[0.18em] text-text-secondary">{label}</div>
      <div className="mt-2 text-2xl font-semibold text-text-primary">{value}</div>
      {detail ? <div className="mt-1 text-xs text-text-secondary">{detail}</div> : null}
    </div>
  );
}

function TableShell({
  title,
  description,
  children,
}: {
  title: string;
  description?: string;
  children: React.ReactNode;
}) {
  return (
    <section className="rounded-2xl border border-border-medium bg-surface-primary p-4 shadow-sm">
      <div className="mb-3">
        <h3 className="text-sm font-semibold text-text-primary">{title}</h3>
        {description ? <p className="mt-1 text-xs text-text-secondary">{description}</p> : null}
      </div>
      {children}
    </section>
  );
}

function Admin() {
  const { user } = useAuthContext();
  const [days, setDays] = useState(30);
  const isAdmin = user?.role === SystemRoles.ADMIN;

  const usersQuery = useAdminUsersQuery(
    { limit: 200 },
    {
      enabled: isAdmin,
    },
  );
  const summaryQuery = useAdminUsageSummaryQuery(
    { days, limit: 200 },
    {
      enabled: isAdmin,
    },
  );
  const recentUsageQuery = useAdminUsageQuery(
    { limit: 15 },
    {
      enabled: isAdmin,
    },
  );
  const issuesQuery = useAdminIssuesQuery(
    { limit: 15, status: 'open' },
    {
      enabled: isAdmin,
    },
  );
  const outlookAuditQuery = useAdminOutlookAuditQuery(
    { limit: 25 },
    {
      enabled: isAdmin,
    },
  );

  const summaryByUser = useMemo(() => {
    const usage = summaryQuery.data?.users ?? [];
    return new Map(usage.map((item) => [item.userId, item]));
  }, [summaryQuery.data?.users]);

  const userRows = useMemo<DashboardUserRow[]>(() => {
    const users = usersQuery.data?.users ?? [];
    const rows = users.map((adminUser) => {
      const usage = summaryByUser.get(adminUser.id);
      return {
        ...adminUser,
        requestCount: usage?.requestCount ?? 0,
        inputTokens: usage?.inputTokens ?? 0,
        outputTokens: usage?.outputTokens ?? 0,
        totalTokens: usage?.totalTokens ?? 0,
        cacheCreationTokens: usage?.cacheCreationTokens ?? 0,
        cacheReadTokens: usage?.cacheReadTokens ?? 0,
        avgLatencyMs: usage?.avgLatencyMs ?? null,
        firstSeenAt: usage?.firstSeenAt,
        lastSeenAt: usage?.lastSeenAt,
      };
    });

    return rows.sort((a, b) => {
      if (b.totalTokens !== a.totalTokens) {
        return b.totalTokens - a.totalTokens;
      }
      return a.email.localeCompare(b.email);
    });
  }, [summaryByUser, usersQuery.data?.users]);

  const isLoading =
    usersQuery.isLoading ||
    summaryQuery.isLoading ||
    recentUsageQuery.isLoading ||
    issuesQuery.isLoading ||
    outlookAuditQuery.isLoading;
  const hasError =
    usersQuery.isError ||
    summaryQuery.isError ||
    recentUsageQuery.isError ||
    issuesQuery.isError ||
    outlookAuditQuery.isError;
  const overview = summaryQuery.data?.overview;
  const recentUsage = recentUsageQuery.data?.usage ?? [];
  const openIssues = issuesQuery.data?.issues ?? [];
  const outlookAudits = outlookAuditQuery.data?.audits ?? [];

  if (!isAdmin) {
    return (
      <div className="rounded-2xl border border-border-medium bg-surface-secondary p-4 text-sm text-text-secondary">
        Admin access is required to view usage analytics.
      </div>
    );
  }

  return (
    <div className="flex flex-col gap-4 p-1 text-sm text-text-primary">
      <div className="rounded-2xl border border-border-medium bg-surface-secondary p-4">
        <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
          <div>
            <h2 className="text-base font-semibold text-text-primary">Admin usage dashboard</h2>
            <p className="mt-1 text-xs text-text-secondary">
              Track request volume, token usage, and recent model activity across the workspace.
            </p>
          </div>
          <label className="flex items-center gap-2 text-xs text-text-secondary">
            <span>Time window</span>
            <select
              value={days}
              onChange={(event) => setDays(Number(event.target.value))}
              className="rounded-lg border border-border-medium bg-surface-primary px-3 py-2 text-sm text-text-primary"
            >
              {DAY_OPTIONS.map((option) => (
                <option key={option} value={option}>
                  Last {option} days
                </option>
              ))}
            </select>
          </label>
        </div>
      </div>

      {isLoading ? (
        <div className="flex min-h-48 items-center justify-center rounded-2xl border border-border-medium bg-surface-primary">
          <Spinner className="size-6" />
        </div>
      ) : null}

      {!isLoading && hasError ? (
        <div className="rounded-2xl border border-red-300 bg-red-50 p-4 text-sm text-red-700 dark:border-red-900 dark:bg-red-950/30 dark:text-red-300">
          Unable to load admin usage data. Confirm usage tracking is enabled and the current user
          has admin permissions.
        </div>
      ) : null}

      {!isLoading && !hasError ? (
        <>
          <div className="grid gap-3 md:grid-cols-2 xl:grid-cols-4">
            <MetricCard
              label="Total tokens"
              value={formatNumber(overview?.totalTokens)}
              detail={`${formatNumber(overview?.inputTokens)} input / ${formatNumber(overview?.outputTokens)} output`}
            />
            <MetricCard
              label="Requests"
              value={formatNumber(overview?.requestCount)}
              detail={`Across ${formatNumber(overview?.activeUsers)} active users`}
            />
            <MetricCard
              label="Average latency"
              value={formatLatency(overview?.avgLatencyMs)}
              detail="Average over requests with recorded latency"
            />
            <MetricCard
              label="Window"
              value={summaryQuery.data?.days ? `${summaryQuery.data.days} days` : `${days} days`}
              detail={
                overview?.windowStart && overview?.windowEnd
                  ? `${formatDate(overview.windowStart)} to ${formatDate(overview.windowEnd)}`
                  : 'Current reporting period'
              }
            />
          </div>

          <TableShell
            title="Users"
            description="All users are listed below. Token and request totals reflect the selected reporting window."
          >
            <div className="overflow-x-auto">
              <table className="min-w-full divide-y divide-border-medium text-left">
                <thead>
                  <tr className="text-xs uppercase tracking-wide text-text-secondary">
                    <th className="py-2 pr-4 font-medium">User</th>
                    <th className="py-2 pr-4 font-medium">Role</th>
                    <th className="py-2 pr-4 font-medium">Requests</th>
                    <th className="py-2 pr-4 font-medium">Tokens</th>
                    <th className="py-2 pr-4 font-medium">Input</th>
                    <th className="py-2 pr-4 font-medium">Output</th>
                    <th className="py-2 pr-4 font-medium">Last activity</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-border-light">
                  {userRows.map((row) => (
                    <tr key={row.id} className="align-top">
                      <td className="py-3 pr-4">
                        <div className="font-medium text-text-primary">
                          {row.name || row.username || row.email || row.id}
                        </div>
                        <div className="text-xs text-text-secondary">{row.email || row.username}</div>
                      </td>
                      <td className="py-3 pr-4 text-text-secondary">{row.role}</td>
                      <td className="py-3 pr-4">{formatNumber(row.requestCount)}</td>
                      <td className="py-3 pr-4">{formatNumber(row.totalTokens)}</td>
                      <td className="py-3 pr-4">{formatNumber(row.inputTokens)}</td>
                      <td className="py-3 pr-4">{formatNumber(row.outputTokens)}</td>
                      <td className="py-3 pr-4 text-text-secondary">
                        {row.lastSeenAt ? formatDate(row.lastSeenAt) : 'No usage in window'}
                      </td>
                    </tr>
                  ))}
                  {userRows.length === 0 ? (
                    <tr>
                      <td colSpan={7} className="py-6 text-center text-text-secondary">
                        No users were returned by the admin API.
                      </td>
                    </tr>
                  ) : null}
                </tbody>
              </table>
            </div>
          </TableShell>

          <TableShell
            title="Recent requests"
            description="Latest tracked model requests, including request source, model, and token totals."
          >
            <div className="overflow-x-auto">
              <table className="min-w-full divide-y divide-border-medium text-left">
                <thead>
                  <tr className="text-xs uppercase tracking-wide text-text-secondary">
                    <th className="py-2 pr-4 font-medium">Time</th>
                    <th className="py-2 pr-4 font-medium">User</th>
                    <th className="py-2 pr-4 font-medium">Model</th>
                    <th className="py-2 pr-4 font-medium">Context</th>
                    <th className="py-2 pr-4 font-medium">Source</th>
                    <th className="py-2 pr-4 font-medium">Tokens</th>
                    <th className="py-2 pr-4 font-medium">Latency</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-border-light">
                  {recentUsage.map((record: AdminUsageListItem) => {
                    const matchedUser = usersQuery.data?.users.find((item) => item.id === record.userId);
                    return (
                      <tr key={record.id}>
                        <td className="py-3 pr-4 text-text-secondary">
                          {record.createdAt ? formatDate(record.createdAt) : 'n/a'}
                        </td>
                        <td className="py-3 pr-4">
                          {matchedUser?.email || matchedUser?.name || record.userId}
                        </td>
                        <td className="py-3 pr-4">{record.model || record.provider || 'n/a'}</td>
                        <td className="py-3 pr-4">{record.context || record.endpoint || 'n/a'}</td>
                        <td className="py-3 pr-4">{record.source || 'system'}</td>
                        <td className="py-3 pr-4">{formatNumber(record.totalTokens)}</td>
                        <td className="py-3 pr-4 text-text-secondary">
                          {formatLatency(record.latencyMs)}
                        </td>
                      </tr>
                    );
                  })}
                  {recentUsage.length === 0 ? (
                    <tr>
                      <td colSpan={7} className="py-6 text-center text-text-secondary">
                        No usage records have been captured yet.
                      </td>
                    </tr>
                  ) : null}
                </tbody>
              </table>
            </div>
          </TableShell>

          <TableShell
            title="Outlook AI audit trail"
            description="Metadata-only trace of AI Inbox views, analyses, and draft creation. Email bodies are not stored here."
          >
            <div className="overflow-x-auto">
              <table className="min-w-full divide-y divide-border-medium text-left">
                <thead>
                  <tr className="text-xs uppercase tracking-wide text-text-secondary">
                    <th className="py-2 pr-4 font-medium">Time</th>
                    <th className="py-2 pr-4 font-medium">User</th>
                    <th className="py-2 pr-4 font-medium">Action</th>
                    <th className="py-2 pr-4 font-medium">Status</th>
                    <th className="py-2 pr-4 font-medium">Message</th>
                    <th className="py-2 pr-4 font-medium">Draft</th>
                    <th className="py-2 pr-4 font-medium">Details</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-border-light">
                  {outlookAudits.map((audit: AdminOutlookAuditItem) => (
                    <tr key={audit.id} className="align-top">
                      <td className="py-3 pr-4 text-text-secondary">
                        {audit.createdAt ? formatDate(audit.createdAt) : 'n/a'}
                      </td>
                      <td className="py-3 pr-4">
                        <div className="font-medium text-text-primary">
                          {audit.actorName || audit.actorEmail || audit.userId}
                        </div>
                        <div className="text-xs text-text-secondary">{audit.actorEmail}</div>
                      </td>
                      <td className="py-3 pr-4">{formatAuditAction(audit.action)}</td>
                      <td className="py-3 pr-4">
                        <span
                          className={
                            audit.status === 'success'
                              ? 'rounded-full bg-green-500/10 px-2 py-1 text-xs font-medium text-green-700 dark:text-green-300'
                              : 'rounded-full bg-red-500/10 px-2 py-1 text-xs font-medium text-red-700 dark:text-red-300'
                          }
                        >
                          {audit.status}
                        </span>
                      </td>
                      <td className="max-w-48 truncate py-3 pr-4 text-text-secondary">
                        {audit.graphMessageId || 'n/a'}
                      </td>
                      <td className="max-w-48 truncate py-3 pr-4 text-text-secondary">
                        {audit.graphDraftId || 'n/a'}
                      </td>
                      <td className="py-3 pr-4 text-text-secondary">
                        {audit.errorMessage ||
                          (audit.metadata?.analysisMode
                            ? `mode: ${String(audit.metadata.analysisMode)}`
                            : audit.metadata?.folder
                              ? `folder: ${String(audit.metadata.folder)}`
                              : 'metadata only')}
                      </td>
                    </tr>
                  ))}
                  {outlookAudits.length === 0 ? (
                    <tr>
                      <td colSpan={7} className="py-6 text-center text-text-secondary">
                        No Outlook AI audit records have been captured yet.
                      </td>
                    </tr>
                  ) : null}
                </tbody>
              </table>
            </div>
          </TableShell>

          <TableShell
            title="Reported issues"
            description="Open user reports for bad responses, MCP failures, and file transformation problems."
          >
            <div className="overflow-x-auto">
              <table className="min-w-full divide-y divide-border-medium text-left">
                <thead>
                  <tr className="text-xs uppercase tracking-wide text-text-secondary">
                    <th className="py-2 pr-4 font-medium">Reporter</th>
                    <th className="py-2 pr-4 font-medium">Category</th>
                    <th className="py-2 pr-4 font-medium">Context</th>
                    <th className="py-2 pr-4 font-medium">Notes</th>
                    <th className="py-2 pr-4 font-medium">Created</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-border-light">
                  {openIssues.map((issue: AdminIssueReportItem) => (
                    <tr key={issue.id} className="align-top">
                      <td className="py-3 pr-4">
                        <div className="font-medium text-text-primary">
                          {issue.reporterName || issue.reporterEmail || issue.userId}
                        </div>
                        <div className="text-xs text-text-secondary">{issue.reporterEmail}</div>
                      </td>
                      <td className="py-3 pr-4 text-text-secondary">{issue.category}</td>
                      <td className="py-3 pr-4 text-text-secondary">
                        <div>{issue.model || issue.endpoint || 'General chat'}</div>
                        <div className="mt-1 text-xs">
                          {issue.mcpServer || issue.toolName || issue.error
                            ? 'Execution issue'
                            : 'Response issue'}
                        </div>
                      </td>
                      <td className="py-3 pr-4">
                        <div className="max-w-md text-text-primary">
                          {issue.description || issue.messagePreview || 'No notes provided'}
                        </div>
                      </td>
                      <td className="py-3 pr-4 text-text-secondary">
                        {issue.createdAt ? formatDate(issue.createdAt) : 'n/a'}
                      </td>
                    </tr>
                  ))}
                  {openIssues.length === 0 ? (
                    <tr>
                      <td colSpan={5} className="py-6 text-center text-text-secondary">
                        No open issue reports yet.
                      </td>
                    </tr>
                  ) : null}
                </tbody>
              </table>
            </div>
          </TableShell>
        </>
      ) : null}
    </div>
  );
}

export default React.memo(Admin);
