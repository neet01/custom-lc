import React, { useMemo, useState } from 'react';
import { Spinner } from '@librechat/client';
import { RefreshCw } from 'lucide-react';
import type {
  ArchiveDiagnosticsConversation,
  ArchiveDiagnosticsCountItem,
  ArchiveDiagnosticsJob,
  ArchiveDiagnosticsParams,
  ArchiveDiagnosticsSeverity,
  ArchiveDiagnosticsSource,
} from 'librechat-data-provider';
import { useAdminArchiveDiagnosticsQuery } from '~/data-provider';
import { cn, formatDateTime } from '~/utils';

const PAGE_SIZE = 50;

const SOURCE_OPTIONS: Array<{ value: ArchiveDiagnosticsSource; label: string }> = [
  { value: 'slack', label: 'Slack' },
  { value: 'teams', label: 'Teams' },
];

const TYPE_OPTIONS: Record<ArchiveDiagnosticsSource, string[]> = {
  slack: ['public_channel', 'private_channel', 'im', 'mpim'],
  teams: ['oneOnOne', 'group', 'meeting'],
};

const STATUS_OPTIONS = ['complete', 'failed', 'pending', 'running', 'deferred_failed'];

function formatNumber(value: number | null | undefined) {
  return new Intl.NumberFormat().format(value ?? 0);
}

function getBrowserTimeZone() {
  try {
    return Intl.DateTimeFormat().resolvedOptions().timeZone || undefined;
  } catch {
    return undefined;
  }
}

function formatAdminDateTime(value: string | null | undefined) {
  if (!value) {
    return 'n/a';
  }

  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return 'n/a';
  }

  try {
    return new Intl.DateTimeFormat(undefined, {
      timeZone: getBrowserTimeZone(),
      year: 'numeric',
      month: 'numeric',
      day: 'numeric',
      hour: 'numeric',
      minute: '2-digit',
      second: '2-digit',
      hour12: true,
      timeZoneName: 'short',
    }).format(date);
  } catch {
    return formatDateTime(value);
  }
}

function titleCase(value: string | null | undefined) {
  return String(value || 'unknown')
    .replace(/[_-]+/g, ' ')
    .replace(/\b\w/g, (letter) => letter.toUpperCase());
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

function HealthBadge({ severity, state }: { severity: ArchiveDiagnosticsSeverity; state: string }) {
  return (
    <span
      className={cn(
        'inline-flex rounded-full px-2.5 py-1 text-xs font-medium',
        severity === 'ok'
          ? 'bg-green-500/10 text-green-700 dark:text-green-300'
          : severity === 'warning'
            ? 'bg-amber-500/10 text-amber-700 dark:text-amber-300'
            : 'bg-red-500/10 text-red-700 dark:text-red-300',
      )}
    >
      {titleCase(state)}
    </span>
  );
}

function BreakdownList({
  title,
  items,
  emptyLabel = 'No data',
}: {
  title: string;
  items: ArchiveDiagnosticsCountItem[];
  emptyLabel?: string;
}) {
  return (
    <div className="rounded-2xl border border-border-medium bg-surface-primary p-4">
      <h4 className="text-sm font-semibold text-text-primary">{title}</h4>
      <div className="mt-3 flex flex-wrap gap-2">
        {items.length > 0 ? (
          items.map((item) => (
            <span
              key={item.key}
              className="rounded-full border border-border-light bg-surface-secondary px-3 py-1 text-xs text-text-secondary"
            >
              <span className="font-medium text-text-primary">{item.key}</span>{' '}
              {formatNumber(item.count)}
            </span>
          ))
        ) : (
          <span className="text-xs text-text-secondary">{emptyLabel}</span>
        )}
      </div>
    </div>
  );
}

function JobCard({ title, job }: { title: string; job: ArchiveDiagnosticsJob | null }) {
  return (
    <div className="rounded-2xl border border-border-medium bg-surface-primary p-4">
      <h4 className="text-sm font-semibold text-text-primary">{title}</h4>
      {job ? (
        <div className="mt-3 space-y-1 text-xs text-text-secondary">
          <div>
            Status:{' '}
            <span className="font-medium text-text-primary">
              {job.status || job.phase || 'unknown'}
            </span>
          </div>
          <div>Created: {formatAdminDateTime(job.createdAt)}</div>
          <div>Completed: {formatAdminDateTime(job.completedAt)}</div>
          {job.errorMessage ? (
            <div className="text-red-700 dark:text-red-300">Error: {job.errorMessage}</div>
          ) : null}
        </div>
      ) : (
        <p className="mt-3 text-xs text-text-secondary">No job record found for this filter.</p>
      )}
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

function ArchiveRow({ row }: { row: ArchiveDiagnosticsConversation }) {
  return (
    <tr className="align-top">
      <td className="py-3 pr-4">
        <HealthBadge severity={row.health.severity} state={row.health.state} />
        <div className="mt-2 max-w-xs text-xs text-text-secondary">{row.health.reason}</div>
      </td>
      <td className="py-3 pr-4">
        <div className="font-medium text-text-primary">
          {row.displayName || row.sourceConversationId}
        </div>
        <div className="mt-1 max-w-xs truncate text-xs text-text-secondary">
          {row.sourceConversationId}
        </div>
        <div className="mt-1 max-w-xs truncate text-xs text-text-secondary">User: {row.userId}</div>
      </td>
      <td className="py-3 pr-4 text-text-secondary">{row.type}</td>
      <td className="py-3 pr-4">
        <div className="text-text-primary">{row.syncStatus}</div>
        {row.syncError ? (
          <div className="mt-1 max-w-xs text-xs text-red-700 dark:text-red-300">
            {row.syncError}
          </div>
        ) : null}
      </td>
      <td className="py-3 pr-4">
        <div>{formatNumber(row.messageCount)} stored</div>
        <div className="mt-1 text-xs text-text-secondary">
          {formatNumber(row.meaningfulMessageCount)} meaningful /{' '}
          {formatNumber(row.skippedMessageCount)} skipped
        </div>
      </td>
      <td className="py-3 pr-4">
        <div>{formatNumber(row.chunkCount)} total</div>
        <div className="mt-1 text-xs text-text-secondary">
          {formatNumber(row.messageChunkCount)} message / {formatNumber(row.windowChunkCount)}{' '}
          window
        </div>
      </td>
      <td className="py-3 pr-4 text-text-secondary">
        <div>Message: {formatAdminDateTime(row.lastMeaningfulMessageAt || row.lastMessageAt)}</div>
        <div className="mt-1">Chunk: {formatAdminDateTime(row.latestChunkAt)}</div>
      </td>
    </tr>
  );
}

export default function ArchiveDiagnostics() {
  const [source, setSource] = useState<ArchiveDiagnosticsSource>('slack');
  const [queryDraft, setQueryDraft] = useState('');
  const [userIdDraft, setUserIdDraft] = useState('');
  const [typeDraft, setTypeDraft] = useState('');
  const [statusDraft, setStatusDraft] = useState('');
  const [filters, setFilters] = useState<ArchiveDiagnosticsParams>({
    source: 'slack',
    limit: PAGE_SIZE,
    offset: 0,
  });

  const params = useMemo<ArchiveDiagnosticsParams>(
    () => ({
      source,
      q: String(filters.q || '').trim() || undefined,
      userId: String(filters.userId || '').trim() || undefined,
      type: String(filters.type || '').trim() || undefined,
      status: String(filters.status || '').trim() || undefined,
      limit: PAGE_SIZE,
      offset: filters.offset ?? 0,
    }),
    [filters, source],
  );

  const diagnosticsQuery = useAdminArchiveDiagnosticsQuery(params, {
    keepPreviousData: true,
  });

  const data = diagnosticsQuery.data;
  const conversations = data?.conversations ?? [];
  const total = data?.summary.filteredConversations ?? 0;
  const offset = data?.filters.offset ?? params.offset ?? 0;
  const hasSearch = Boolean(String(params.q || '').trim());

  const handleApplyFilters = (event: React.FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    setFilters({
      source,
      q: queryDraft.trim() || undefined,
      userId: userIdDraft.trim() || undefined,
      type: typeDraft || undefined,
      status: statusDraft || undefined,
      limit: PAGE_SIZE,
      offset: 0,
    });
  };

  const handleSourceChange = (nextSource: ArchiveDiagnosticsSource) => {
    setSource(nextSource);
    setTypeDraft('');
    setFilters({
      source: nextSource,
      q: queryDraft.trim() || undefined,
      userId: userIdDraft.trim() || undefined,
      status: statusDraft || undefined,
      limit: PAGE_SIZE,
      offset: 0,
    });
  };

  return (
    <section className="min-h-[55vh] overflow-visible rounded-2xl border border-border-medium bg-surface-primary p-4 shadow-sm">
      <div className="mb-4 flex flex-col gap-3 xl:flex-row xl:items-start xl:justify-between">
        <div>
          <h3 className="text-sm font-semibold text-text-primary">Archive indexing diagnostics</h3>
          <p className="mt-1 max-w-3xl text-xs text-text-secondary">
            Metadata-only visibility into archive discovery, message capture, and memory projection.
            Use this to answer whether a Slack channel or Teams chat was stored and indexed.
          </p>
        </div>
        <button
          type="button"
          onClick={() => void diagnosticsQuery.refetch()}
          disabled={diagnosticsQuery.isFetching}
          className="inline-flex items-center gap-2 rounded-xl border border-border-medium bg-surface-secondary px-3 py-2 text-sm font-medium text-text-primary transition-colors hover:bg-surface-hover disabled:cursor-not-allowed disabled:opacity-50"
        >
          <RefreshCw
            className={cn('h-4 w-4', diagnosticsQuery.isFetching ? 'animate-spin' : undefined)}
          />
          Refresh
        </button>
      </div>

      <form
        onSubmit={handleApplyFilters}
        className="mb-4 grid gap-3 rounded-2xl border border-border-light bg-surface-secondary p-3 lg:grid-cols-[1fr_1fr_1fr_1fr_auto]"
      >
        <label className="text-xs font-medium text-text-secondary">
          Source
          <select
            value={source}
            onChange={(event) => handleSourceChange(event.target.value as ArchiveDiagnosticsSource)}
            className="mt-1 w-full rounded-lg border border-border-medium bg-surface-primary px-3 py-2 text-sm text-text-primary"
          >
            {SOURCE_OPTIONS.map((option) => (
              <option key={option.value} value={option.value}>
                {option.label}
              </option>
            ))}
          </select>
        </label>
        <label className="text-xs font-medium text-text-secondary">
          Channel/chat search
          <input
            value={queryDraft}
            onChange={(event) => setQueryDraft(event.target.value)}
            placeholder="name, topic, purpose, or source ID"
            className="mt-1 w-full rounded-lg border border-border-medium bg-surface-primary px-3 py-2 text-sm text-text-primary"
          />
        </label>
        <label className="text-xs font-medium text-text-secondary">
          User ID
          <input
            value={userIdDraft}
            onChange={(event) => setUserIdDraft(event.target.value)}
            placeholder="optional LibreChat user id"
            className="mt-1 w-full rounded-lg border border-border-medium bg-surface-primary px-3 py-2 text-sm text-text-primary"
          />
        </label>
        <label className="text-xs font-medium text-text-secondary">
          Type
          <select
            value={typeDraft}
            onChange={(event) => setTypeDraft(event.target.value)}
            className="mt-1 w-full rounded-lg border border-border-medium bg-surface-primary px-3 py-2 text-sm text-text-primary"
          >
            <option value="">Any type</option>
            {TYPE_OPTIONS[source].map((option) => (
              <option key={option} value={option}>
                {option}
              </option>
            ))}
          </select>
        </label>
        <div className="grid grid-cols-[1fr_auto] gap-2 lg:grid-cols-1">
          <label className="text-xs font-medium text-text-secondary">
            Status
            <select
              value={statusDraft}
              onChange={(event) => setStatusDraft(event.target.value)}
              className="mt-1 w-full rounded-lg border border-border-medium bg-surface-primary px-3 py-2 text-sm text-text-primary"
            >
              <option value="">Any status</option>
              {STATUS_OPTIONS.map((option) => (
                <option key={option} value={option}>
                  {option}
                </option>
              ))}
            </select>
          </label>
          <button
            type="submit"
            className="self-end rounded-xl border border-border-medium bg-surface-primary px-4 py-2 text-sm font-medium text-text-primary transition-colors hover:bg-surface-hover"
          >
            Apply
          </button>
        </div>
      </form>

      {diagnosticsQuery.isLoading ? (
        <div className="flex min-h-48 items-center justify-center rounded-2xl border border-border-medium bg-surface-secondary">
          <Spinner className="size-6" />
        </div>
      ) : null}

      {!diagnosticsQuery.isLoading && diagnosticsQuery.isError ? (
        <div className="rounded-2xl border border-red-300 bg-red-50 p-4 text-sm text-red-700 dark:border-red-900 dark:bg-red-950/30 dark:text-red-300">
          Unable to load archive diagnostics. Confirm the current user has admin reporting
          permissions and the diagnostics route is available in this deployment.
        </div>
      ) : null}

      {!diagnosticsQuery.isLoading && !diagnosticsQuery.isError && data ? (
        <>
          <div className="grid gap-3 md:grid-cols-2 xl:grid-cols-4">
            <MetricCard
              label="Conversations"
              value={formatNumber(data.summary.filteredConversations)}
              detail={`${formatNumber(data.summary.totalConversations)} total for selected source/user`}
            />
            <MetricCard
              label="Messages"
              value={formatNumber(data.summary.totalMessages)}
              detail="Archived message records, not raw text shown here"
            />
            <MetricCard
              label="Chunks"
              value={formatNumber(data.summary.totalChunks)}
              detail="Enterprise memory chunks for this archive source"
            />
            <MetricCard
              label="Visible row health"
              value={`${formatNumber(data.summary.healthyConversationCount)} ok`}
              detail={`${formatNumber(data.summary.warningConversationCount)} warning / ${formatNumber(data.summary.errorConversationCount)} error`}
            />
          </div>

          <div className="mt-3 grid gap-3 lg:grid-cols-2">
            <JobCard title="Latest sync job" job={data.latestSync} />
            <JobCard title="Latest projection job" job={data.latestProjection} />
          </div>

          <div className="mt-3 grid gap-3 lg:grid-cols-2 xl:grid-cols-3">
            <BreakdownList title="Conversation types" items={data.breakdowns.conversationsByType} />
            <BreakdownList
              title="Conversation sync statuses"
              items={data.breakdowns.conversationsByStatus}
            />
            <BreakdownList title="Chunk record types" items={data.breakdowns.chunksByRecordType} />
            <BreakdownList title="Chunk types" items={data.breakdowns.chunksByChunkType} />
            <BreakdownList
              title="Skipped message reasons"
              items={data.breakdowns.skippedMessageReasons}
              emptyLabel="No skipped message reasons recorded"
            />
          </div>

          {hasSearch && total === 0 ? (
            <div className="mt-4 rounded-2xl border border-amber-300 bg-amber-50 p-4 text-sm text-amber-800 dark:border-amber-900 dark:bg-amber-950/30 dark:text-amber-200">
              No archived conversation matched this query. That usually means the conversation was
              not discovered or stored for the selected source/user, or the search does not match
              its name, topic, purpose, or source ID. Check the latest sync job and type/status
              breakdowns before debugging retrieval.
            </div>
          ) : null}

          <div className="mt-4 w-full overflow-x-auto overflow-y-visible rounded-2xl border border-border-medium">
            <table className="min-w-[72rem] divide-y divide-border-medium text-left">
              <thead className="bg-surface-secondary">
                <tr className="text-xs uppercase tracking-wide text-text-secondary">
                  <th className="py-2 pl-4 pr-4 font-medium">Health</th>
                  <th className="py-2 pr-4 font-medium">Conversation</th>
                  <th className="py-2 pr-4 font-medium">Type</th>
                  <th className="py-2 pr-4 font-medium">Sync</th>
                  <th className="py-2 pr-4 font-medium">Messages</th>
                  <th className="py-2 pr-4 font-medium">Chunks</th>
                  <th className="py-2 pr-4 font-medium">Freshness</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-border-light">
                {conversations.map((row) => (
                  <ArchiveRow key={row.id || row.sourceConversationId} row={row} />
                ))}
                {conversations.length === 0 ? (
                  <EmptyRow colSpan={7} message="No archive conversations matched these filters." />
                ) : null}
              </tbody>
            </table>
          </div>

          <PaginationControls
            total={total}
            limit={data.filters.limit || PAGE_SIZE}
            offset={offset}
            onChange={(nextOffset) =>
              setFilters((current) => ({
                ...current,
                offset: nextOffset,
              }))
            }
          />
        </>
      ) : null}
    </section>
  );
}
