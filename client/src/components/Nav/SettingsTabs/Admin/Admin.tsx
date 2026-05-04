import React, { useMemo, useState } from 'react';
import { Spinner, useToastContext } from '@librechat/client';
import { BarChart3, Mail, ShieldAlert, Users } from 'lucide-react';
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
  useAdminUpdateUserBalanceMutation,
  useAdminUsageQuery,
  useAdminUsageSummaryQuery,
  useAdminUsersQuery,
} from '~/data-provider';
import { useAuthContext } from '~/hooks';
import { cn, formatDate } from '~/utils';

const DAY_OPTIONS = [7, 30, 90];
const PAGE_SIZE = 25;

type AdminTab = 'usage-users' | 'recent-requests' | 'users' | 'outlook-audit' | 'issues';

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

function MetricCard({ label, value, detail }: { label: string; value: string; detail?: string }) {
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
  className,
}: {
  title: string;
  description?: string;
  children: React.ReactNode;
  className?: string;
}) {
  return (
    <section
      className={cn(
        'flex min-h-0 flex-col rounded-2xl border border-border-medium bg-surface-primary p-4 shadow-sm',
        className,
      )}
    >
      <div className="mb-3">
        <h3 className="text-sm font-semibold text-text-primary">{title}</h3>
        {description ? <p className="mt-1 text-xs text-text-secondary">{description}</p> : null}
      </div>
      <div className="flex min-h-0 flex-1 flex-col">{children}</div>
    </section>
  );
}

function TabButton({
  active,
  icon: Icon,
  label,
  onClick,
}: {
  active: boolean;
  icon: React.ComponentType<{ className?: string }>;
  label: string;
  onClick: () => void;
}) {
  return (
    <button
      type="button"
      className={cn(
        'inline-flex items-center gap-2 rounded-xl border px-3 py-2 text-sm transition-colors',
        active
          ? 'border-[#f5d000]/40 bg-[#f5d000]/10 text-text-primary'
          : 'border-border-medium bg-surface-primary text-text-secondary hover:bg-surface-hover hover:text-text-primary',
      )}
      onClick={onClick}
    >
      <Icon className="h-4 w-4" />
      {label}
    </button>
  );
}

function PaginationControls({
  total,
  limit,
  offset,
  onChange,
}: {
  total: number;
  limit: number;
  offset: number;
  onChange: (nextOffset: number) => void;
}) {
  const currentPage = Math.floor(offset / limit) + 1;
  const totalPages = Math.max(1, Math.ceil(total / limit));
  const canGoBack = offset > 0;
  const canGoForward = offset + limit < total;

  return (
    <div className="mt-4 flex flex-col gap-3 border-t border-border-light pt-3 text-xs text-text-secondary sm:flex-row sm:items-center sm:justify-between">
      <div>
        Showing {Math.min(offset + 1, total || 1)}-{Math.min(offset + limit, total)} of{' '}
        {formatNumber(total)}
      </div>
      <div className="flex items-center gap-2">
        <span>
          Page {currentPage} / {totalPages}
        </span>
        <button
          type="button"
          className="rounded-lg border border-border-medium px-2.5 py-1 font-medium text-text-primary disabled:cursor-not-allowed disabled:opacity-40"
          onClick={() => onChange(Math.max(0, offset - limit))}
          disabled={!canGoBack}
        >
          Previous
        </button>
        <button
          type="button"
          className="rounded-lg border border-border-medium px-2.5 py-1 font-medium text-text-primary disabled:cursor-not-allowed disabled:opacity-40"
          onClick={() => onChange(offset + limit)}
          disabled={!canGoForward}
        >
          Next
        </button>
      </div>
    </div>
  );
}

function EmptyRow({ colSpan, message }: { colSpan: number; message: string }) {
  return (
    <tr>
      <td colSpan={colSpan} className="py-6 text-center text-text-secondary">
        {message}
      </td>
    </tr>
  );
}

function LoadingState({ workspaceMode = false }: { workspaceMode?: boolean }) {
  return (
    <div
      className={cn(
        'flex items-center justify-center rounded-2xl border border-border-medium bg-surface-primary',
        workspaceMode ? 'min-h-[50vh]' : 'min-h-48',
      )}
    >
      <Spinner className="size-6" />
    </div>
  );
}

function ErrorState() {
  return (
    <div className="rounded-2xl border border-red-300 bg-red-50 p-4 text-sm text-red-700 dark:border-red-900 dark:bg-red-950/30 dark:text-red-300">
      Unable to load admin usage data. Confirm usage tracking is enabled and the current user has
      admin permissions.
    </div>
  );
}

function AccessDenied() {
  return (
    <div className="rounded-2xl border border-border-medium bg-surface-secondary p-4 text-sm text-text-secondary">
      Admin access is required to view usage analytics.
    </div>
  );
}

function Admin({ workspaceMode = false }: { workspaceMode?: boolean }) {
  const { user } = useAuthContext();
  const { showToast } = useToastContext();
  const [days, setDays] = useState(30);
  const [activeTab, setActiveTab] = useState<AdminTab>('usage-users');
  const [summaryOffset, setSummaryOffset] = useState(0);
  const [recentUsageOffset, setRecentUsageOffset] = useState(0);
  const [usersOffset, setUsersOffset] = useState(0);
  const [issuesOffset, setIssuesOffset] = useState(0);
  const [outlookAuditOffset, setOutlookAuditOffset] = useState(0);
  const [balanceDrafts, setBalanceDrafts] = useState<Record<string, string>>({});
  const isAdmin = user?.role === SystemRoles.ADMIN;
  const updateUserBalance = useAdminUpdateUserBalanceMutation();

  const usersQuery = useAdminUsersQuery(
    { limit: PAGE_SIZE, offset: usersOffset },
    {
      enabled: isAdmin,
      keepPreviousData: true,
    },
  );
  const summaryQuery = useAdminUsageSummaryQuery(
    { days, limit: PAGE_SIZE, offset: summaryOffset },
    {
      enabled: isAdmin,
      keepPreviousData: true,
    },
  );
  const recentUsageQuery = useAdminUsageQuery(
    { limit: PAGE_SIZE, offset: recentUsageOffset },
    {
      enabled: isAdmin,
      keepPreviousData: true,
    },
  );
  const issuesQuery = useAdminIssuesQuery(
    { limit: PAGE_SIZE, offset: issuesOffset, status: 'open' },
    {
      enabled: isAdmin,
      keepPreviousData: true,
    },
  );
  const outlookAuditQuery = useAdminOutlookAuditQuery(
    { limit: PAGE_SIZE, offset: outlookAuditOffset },
    {
      enabled: isAdmin,
      keepPreviousData: true,
    },
  );

  const userLookup = useMemo(() => {
    const lookup = new Map<
      string,
      Pick<AdminUserListItem, 'name' | 'email' | 'username'> &
        Partial<Pick<AdminUsageSummaryItem, 'name' | 'email' | 'username'>>
    >();

    for (const row of usersQuery.data?.users ?? []) {
      lookup.set(row.id, row);
    }

    for (const row of summaryQuery.data?.users ?? []) {
      if (!lookup.has(row.userId)) {
        lookup.set(row.userId, {
          name: row.name,
          email: row.email,
          username: row.username,
        });
      }
    }

    return lookup;
  }, [usersQuery.data?.users, summaryQuery.data?.users]);

  const isInitialLoading =
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

  if (!isAdmin) {
    return <AccessDenied />;
  }

  const overview = summaryQuery.data?.overview;
  const usageUsers = summaryQuery.data?.users ?? [];
  const recentUsage = recentUsageQuery.data?.usage ?? [];
  const openIssues = issuesQuery.data?.issues ?? [];
  const outlookAudits = outlookAuditQuery.data?.audits ?? [];
  const directoryUsers = usersQuery.data?.users ?? [];

  const handleBalanceSave = async (row: AdminUserListItem) => {
    const rawValue = balanceDrafts[row.id] ?? String(row.tokenCredits ?? 0);
    const trimmed = rawValue.trim();
    const parsed = Number(trimmed);

    if (!trimmed || !Number.isInteger(parsed) || parsed < 0) {
      showToast({
        status: 'error',
        message: 'Balance must be a non-negative integer.',
      });
      return;
    }

    if (parsed === row.tokenCredits) {
      setBalanceDrafts((current) => {
        const next = { ...current };
        delete next[row.id];
        return next;
      });
      return;
    }

    try {
      await updateUserBalance.mutateAsync({ userId: row.id, tokenCredits: parsed });
      setBalanceDrafts((current) => {
        const next = { ...current };
        delete next[row.id];
        return next;
      });
      showToast({
        status: 'success',
        message: 'User balance updated.',
      });
    } catch (error) {
      const message =
        error instanceof Error && error.message ? error.message : 'Failed to update user balance.';
      showToast({
        status: 'error',
        message,
      });
    }
  };

  return (
    <div
      data-tour="admin-reporting-root"
      className={cn(
        'flex min-h-0 flex-col gap-4 text-sm text-text-primary',
        workspaceMode ? 'h-full overflow-y-auto p-6' : 'p-1',
      )}
    >
      <div className="rounded-2xl border border-border-medium bg-surface-secondary p-4">
        <div className="flex flex-col gap-3 xl:flex-row xl:items-center xl:justify-between">
          <div>
            <h2
              className={cn(
                'font-semibold text-text-primary',
                workspaceMode ? 'text-xl' : 'text-base',
              )}
            >
              Admin reporting
            </h2>
            <p className="mt-1 text-xs text-text-secondary">
              Workspace-wide usage, user activity, Outlook audit events, and user-reported issues.
            </p>
          </div>
          <label className="flex items-center gap-2 text-xs text-text-secondary">
            <span>Time window</span>
            <select
              value={days}
              onChange={(event) => {
                const next = Number(event.target.value);
                setDays(next);
                setSummaryOffset(0);
              }}
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

      {isInitialLoading ? <LoadingState workspaceMode={workspaceMode} /> : null}
      {!isInitialLoading && hasError ? <ErrorState /> : null}

      {!isInitialLoading && !hasError ? (
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

          <div className="flex flex-wrap gap-2" data-tour="admin-reporting-tabs">
            <TabButton
              active={activeTab === 'usage-users'}
              icon={BarChart3}
              label="Usage by user"
              onClick={() => setActiveTab('usage-users')}
            />
            <TabButton
              active={activeTab === 'recent-requests'}
              icon={BarChart3}
              label="Recent requests"
              onClick={() => setActiveTab('recent-requests')}
            />
            <TabButton
              active={activeTab === 'users'}
              icon={Users}
              label="User directory"
              onClick={() => setActiveTab('users')}
            />
            <TabButton
              active={activeTab === 'outlook-audit'}
              icon={Mail}
              label="Outlook audit"
              onClick={() => setActiveTab('outlook-audit')}
            />
            <TabButton
              active={activeTab === 'issues'}
              icon={ShieldAlert}
              label="Reported issues"
              onClick={() => setActiveTab('issues')}
            />
          </div>

          {activeTab === 'usage-users' ? (
            <TableShell
              title="Usage by user"
              description="Token and request totals for users active in the selected reporting window."
              className={workspaceMode ? 'min-h-[55vh] overflow-hidden' : undefined}
            >
              <div className="min-h-0 flex-1 overflow-auto">
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
                    {usageUsers.map((row: AdminUsageSummaryItem) => (
                      <tr key={row.userId} className="align-top">
                        <td className="py-3 pr-4">
                          <div className="font-medium text-text-primary">
                            {row.name || row.username || row.email || row.userId}
                          </div>
                          <div className="text-xs text-text-secondary">
                            {row.email || row.username}
                          </div>
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
                    {usageUsers.length === 0 ? (
                      <EmptyRow
                        colSpan={7}
                        message="No active users were returned for this time window."
                      />
                    ) : null}
                  </tbody>
                </table>
              </div>
              <PaginationControls
                total={summaryQuery.data?.total ?? 0}
                limit={summaryQuery.data?.limit ?? PAGE_SIZE}
                offset={summaryQuery.data?.offset ?? summaryOffset}
                onChange={setSummaryOffset}
              />
            </TableShell>
          ) : null}

          {activeTab === 'recent-requests' ? (
            <TableShell
              title="Recent requests"
              description="Latest tracked model requests, including request source, model, and token totals."
              className={workspaceMode ? 'min-h-[55vh] overflow-hidden' : undefined}
            >
              <div className="min-h-0 flex-1 overflow-auto">
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
                      const matchedUser = userLookup.get(record.userId);
                      return (
                        <tr key={record.id}>
                          <td className="py-3 pr-4 text-text-secondary">
                            {record.createdAt ? formatDate(record.createdAt) : 'n/a'}
                          </td>
                          <td className="py-3 pr-4">
                            {matchedUser?.email || matchedUser?.name || record.userId}
                          </td>
                          <td className="py-3 pr-4">{record.model || record.provider || 'n/a'}</td>
                          <td className="py-3 pr-4">
                            {record.context || record.endpoint || 'n/a'}
                          </td>
                          <td className="py-3 pr-4">{record.source || 'system'}</td>
                          <td className="py-3 pr-4">{formatNumber(record.totalTokens)}</td>
                          <td className="py-3 pr-4 text-text-secondary">
                            {formatLatency(record.latencyMs)}
                          </td>
                        </tr>
                      );
                    })}
                    {recentUsage.length === 0 ? (
                      <EmptyRow colSpan={7} message="No usage records have been captured yet." />
                    ) : null}
                  </tbody>
                </table>
              </div>
              <PaginationControls
                total={recentUsageQuery.data?.total ?? 0}
                limit={recentUsageQuery.data?.limit ?? PAGE_SIZE}
                offset={recentUsageQuery.data?.offset ?? recentUsageOffset}
                onChange={setRecentUsageOffset}
              />
            </TableShell>
          ) : null}

          {activeTab === 'users' ? (
            <TableShell
              title="User directory"
              description="All users known to LibreChat, independent of current activity in the reporting window."
              className={workspaceMode ? 'min-h-[55vh] overflow-hidden' : undefined}
            >
              <div className="min-h-0 flex-1 overflow-auto">
                <table className="min-w-full divide-y divide-border-medium text-left">
                  <thead>
                    <tr className="text-xs uppercase tracking-wide text-text-secondary">
                      <th className="py-2 pr-4 font-medium">User</th>
                      <th className="py-2 pr-4 font-medium">Balance</th>
                      <th className="py-2 pr-4 font-medium">Role</th>
                      <th className="py-2 pr-4 font-medium">Provider</th>
                      <th className="py-2 pr-4 font-medium">Created</th>
                      <th className="py-2 pr-4 font-medium">Updated</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-border-light">
                    {directoryUsers.map((row: AdminUserListItem) => (
                      <tr key={row.id}>
                        <td className="py-3 pr-4">
                          <div className="font-medium text-text-primary">
                            {row.name || row.username || row.email || row.id}
                          </div>
                          <div className="text-xs text-text-secondary">
                            {row.email || row.username}
                          </div>
                        </td>
                        <td className="py-3 pr-4">
                          <div className="flex min-w-[13rem] items-center gap-2">
                            <input
                              type="number"
                              min={0}
                              step={1}
                              value={balanceDrafts[row.id] ?? String(row.tokenCredits ?? 0)}
                              onChange={(event) =>
                                setBalanceDrafts((current) => ({
                                  ...current,
                                  [row.id]: event.target.value,
                                }))
                              }
                              className="w-28 rounded-lg border border-border-medium bg-surface-secondary px-3 py-2 text-sm text-text-primary [color-scheme:light] dark:[color-scheme:dark]"
                            />
                            <button
                              type="button"
                              className="rounded-lg border border-border-medium px-3 py-2 text-xs font-medium text-text-primary disabled:cursor-not-allowed disabled:opacity-40"
                              onClick={() => void handleBalanceSave(row)}
                              disabled={
                                updateUserBalance.isLoading ||
                                !Number.isInteger(
                                  Number(balanceDrafts[row.id] ?? row.tokenCredits),
                                ) ||
                                Number(balanceDrafts[row.id] ?? row.tokenCredits) < 0 ||
                                Number(balanceDrafts[row.id] ?? row.tokenCredits) ===
                                  row.tokenCredits
                              }
                            >
                              {updateUserBalance.isLoading &&
                              updateUserBalance.variables?.userId === row.id
                                ? 'Saving...'
                                : 'Save'}
                            </button>
                          </div>
                          <div className="mt-1 text-xs text-text-secondary">
                            {formatNumber(row.tokenCredits)} credits
                          </div>
                        </td>
                        <td className="py-3 pr-4 text-text-secondary">{row.role}</td>
                        <td className="py-3 pr-4 text-text-secondary">{row.provider}</td>
                        <td className="py-3 pr-4 text-text-secondary">
                          {row.createdAt ? formatDate(row.createdAt) : 'n/a'}
                        </td>
                        <td className="py-3 pr-4 text-text-secondary">
                          {row.updatedAt ? formatDate(row.updatedAt) : 'n/a'}
                        </td>
                      </tr>
                    ))}
                    {directoryUsers.length === 0 ? (
                      <EmptyRow colSpan={6} message="No users were returned by the admin API." />
                    ) : null}
                  </tbody>
                </table>
              </div>
              <PaginationControls
                total={usersQuery.data?.total ?? 0}
                limit={usersQuery.data?.limit ?? PAGE_SIZE}
                offset={usersQuery.data?.offset ?? usersOffset}
                onChange={setUsersOffset}
              />
            </TableShell>
          ) : null}

          {activeTab === 'outlook-audit' ? (
            <TableShell
              title="Outlook AI audit trail"
              description="Metadata-only trace of AI Inbox views, analyses, and draft creation. Email bodies are not stored here."
              className={workspaceMode ? 'min-h-[55vh] overflow-hidden' : undefined}
            >
              <div className="min-h-0 flex-1 overflow-auto">
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
                      <EmptyRow
                        colSpan={7}
                        message="No Outlook AI audit records have been captured yet."
                      />
                    ) : null}
                  </tbody>
                </table>
              </div>
              <PaginationControls
                total={outlookAuditQuery.data?.total ?? 0}
                limit={outlookAuditQuery.data?.limit ?? PAGE_SIZE}
                offset={outlookAuditQuery.data?.offset ?? outlookAuditOffset}
                onChange={setOutlookAuditOffset}
              />
            </TableShell>
          ) : null}

          {activeTab === 'issues' ? (
            <TableShell
              title="Reported issues"
              description="Open user reports for bad responses, MCP failures, and file transformation problems."
              className={workspaceMode ? 'min-h-[55vh] overflow-hidden' : undefined}
            >
              <div className="min-h-0 flex-1 overflow-auto">
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
                      <EmptyRow colSpan={5} message="No open issue reports yet." />
                    ) : null}
                  </tbody>
                </table>
              </div>
              <PaginationControls
                total={issuesQuery.data?.total ?? 0}
                limit={issuesQuery.data?.limit ?? PAGE_SIZE}
                offset={issuesQuery.data?.offset ?? issuesOffset}
                onChange={setIssuesOffset}
              />
            </TableShell>
          ) : null}
        </>
      ) : null}
    </div>
  );
}

export default React.memo(Admin);
