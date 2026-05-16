import React from 'react';
import { useQueryClient } from '@tanstack/react-query';
import { Spinner, useToastContext } from '@librechat/client';
import { QueryKeys } from 'librechat-data-provider';
import {
  useCancelTeamsArchiveSyncMutation,
  useSyncTeamsArchiveMutation,
  useTeamsArchiveStatusQuery,
} from '~/data-provider';

function formatTimestamp(value?: string) {
  if (!value) {
    return 'Never';
  }

  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return 'Unknown';
  }

  return date.toLocaleString();
}

function getStatusTone(status?: string | null) {
  if (status === 'running') {
    return 'text-text-primary';
  }

  if (status === 'success') {
    return 'text-text-primary';
  }

  if (status === 'failure') {
    return 'text-rose-700 dark:text-rose-300';
  }

  if (status === 'cancelled') {
    return 'text-amber-700 dark:text-amber-300';
  }

  return 'text-text-secondary';
}

function getStatusLabel(status?: string | null) {
  if (status === 'running') {
    return 'Syncing';
  }

  if (status === 'success') {
    return 'Synced';
  }

  if (status === 'failure') {
    return 'Sync failed';
  }

  if (status === 'cancelled') {
    return 'Sync cancelled';
  }

  return 'Not synced';
}

function getProjectionTone(status?: string | null) {
  if (status === 'success') {
    return 'bg-emerald-500';
  }

  if (status === 'failure') {
    return 'bg-rose-500';
  }

  if (status === 'running' || status === 'pending') {
    return 'bg-amber-400';
  }

  return 'bg-zinc-400 dark:bg-zinc-500';
}

function getProjectionLabel(status?: string | null) {
  if (status === 'success') {
    return 'Active';
  }

  if (status === 'failure') {
    return 'Error';
  }

  if (status === 'running' || status === 'pending') {
    return 'Indexing';
  }

  return 'Unavailable';
}

export default function TeamsArchiveStatus() {
  const queryClient = useQueryClient();
  const { showToast } = useToastContext();
  const { data, isLoading, isFetching } = useTeamsArchiveStatusQuery({
    refetchInterval: (data) => (data?.latestSync?.status === 'running' ? 4000 : false),
  });
  const syncMutation = useSyncTeamsArchiveMutation();
  const cancelMutation = useCancelTeamsArchiveSyncMutation();

  const syncStatus = data?.latestSync?.status ?? null;
  const isSyncing = syncMutation.isLoading || cancelMutation.isLoading || syncStatus === 'running';

  const handleSync = async () => {
    try {
      await syncMutation.mutateAsync({ async: true });
      await queryClient.invalidateQueries([QueryKeys.teamsArchiveStatus]);
      showToast({
        message: 'Teams archive sync started in the background.',
        status: 'success',
      });
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Teams archive sync failed.';
      showToast({
        message,
        status: 'error',
      });
    }
  };

  const handleCancel = async () => {
    try {
      const result = await cancelMutation.mutateAsync();
      await queryClient.invalidateQueries([QueryKeys.teamsArchiveStatus]);
      showToast({
        message: result.message,
        status: result.cancelled ? 'success' : 'error',
      });
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Failed to cancel Teams archive sync.';
      showToast({
        message,
        status: 'error',
      });
    }
  };

  return (
    <div className="relative overflow-hidden rounded-[1.75rem] border border-white/40 bg-gradient-to-br from-white/85 via-white/70 to-white/45 p-4 shadow-[0_20px_60px_-24px_rgba(15,23,42,0.35)] backdrop-blur-xl dark:border-white/10 dark:from-zinc-900/85 dark:via-zinc-900/70 dark:to-neutral-950/55">
      <div className="pointer-events-none absolute inset-x-8 top-0 h-20 rounded-full bg-[#f5d000]/15 blur-3xl" />
      <div className="relative flex flex-col gap-4">
        <div className="flex flex-col gap-3 sm:flex-row sm:items-start sm:justify-between">
          <div className="text-sm font-semibold text-text-primary">Teams Archive</div>
          <div className="flex items-center gap-2">
            <button
              type="button"
              className="inline-flex items-center justify-center rounded-2xl border border-white/50 bg-white/70 px-4 py-2 text-sm font-medium text-text-primary shadow-sm backdrop-blur transition-colors hover:bg-white/90 disabled:cursor-not-allowed disabled:opacity-60 dark:border-white/10 dark:bg-zinc-800/70 dark:hover:bg-zinc-800/90"
              onClick={handleSync}
              disabled={isSyncing}
            >
              {syncMutation.isLoading ? (
                <>
                  <Spinner className="mr-2 h-4 w-4" />
                  Starting…
                </>
              ) : syncStatus === 'running' ? (
                'Syncing…'
              ) : (
                'Sync now'
              )}
            </button>
            {syncStatus === 'running' ? (
              <button
                type="button"
                className="inline-flex min-w-[7.5rem] items-center justify-center whitespace-nowrap rounded-2xl border border-rose-300/60 bg-rose-50/80 px-3 py-1.5 text-xs font-medium text-rose-700 shadow-sm backdrop-blur transition-colors hover:bg-rose-100/90 disabled:cursor-not-allowed disabled:opacity-60 dark:border-rose-500/20 dark:bg-rose-950/40 dark:text-rose-200 dark:hover:bg-rose-950/55"
                onClick={handleCancel}
                disabled={cancelMutation.isLoading}
              >
                {cancelMutation.isLoading ? (
                  <>
                    <Spinner className="mr-1.5 h-3.5 w-3.5" />
                    Cancelling…
                  </>
                ) : (
                  'Cancel sync'
                )}
              </button>
            ) : null}
          </div>
        </div>

        {isSyncing ? (
          <div className="overflow-hidden rounded-2xl border border-white/35 bg-white/45 px-4 py-3 backdrop-blur dark:border-white/10 dark:bg-zinc-900/40">
            <div className="flex items-center justify-between gap-3">
              <div>
                <div className="text-sm font-semibold text-text-primary">Indexing archive</div>
                <div className="mt-1 text-xs text-text-secondary">
                  Cortex is pulling Teams history in the background. Progress is indeterminate
                  because the final chat/message count is not known up front.
                </div>
              </div>
              <div className="shrink-0 text-xs font-medium text-text-secondary">Running</div>
            </div>
            <div className="mt-3 h-2 overflow-hidden rounded-full bg-black/5 dark:bg-white/10">
              <div className="relative h-full w-2/5 animate-pulse rounded-full bg-[#f5d000]/70 shadow-[0_0_20px_rgba(245,208,0,0.35)]" />
            </div>
          </div>
        ) : null}

        <div className="grid gap-3 sm:grid-cols-3">
          <div className="rounded-2xl border border-white/40 bg-white/55 px-4 py-3 backdrop-blur dark:border-white/10 dark:bg-zinc-900/45">
            <div className="text-[11px] font-medium uppercase tracking-[0.14em] text-text-secondary">
              Status
            </div>
            <div className={`mt-2 text-sm font-semibold ${getStatusTone(syncStatus)}`}>
              {isLoading ? 'Loading…' : getStatusLabel(syncStatus)}
            </div>
            <div className="mt-1 text-xs text-text-secondary">
              {syncStatus === 'running'
                ? 'Archive refresh is in progress.'
                : data?.latestSync?.errorMessage || 'Background sync status for Teams chat history.'}
            </div>
          </div>

          <div className="rounded-2xl border border-white/40 bg-white/55 px-4 py-3 backdrop-blur dark:border-white/10 dark:bg-zinc-900/45">
            <div className="text-[11px] font-medium uppercase tracking-[0.14em] text-text-secondary">
              Coverage
            </div>
            <div className="mt-2 text-sm font-semibold text-text-primary">
              {(data?.conversationCount ?? 0).toLocaleString()} chats
            </div>
            <div className="mt-1 text-xs text-text-secondary">
              {(data?.messageCount ?? 0).toLocaleString()} archived messages
            </div>
          </div>

          <div className="rounded-2xl border border-white/40 bg-white/55 px-4 py-3 backdrop-blur dark:border-white/10 dark:bg-zinc-900/45">
            <div className="text-[11px] font-medium uppercase tracking-[0.14em] text-text-secondary">
              Last Sync
            </div>
            <div className="mt-2 text-sm font-semibold text-text-primary">
              {formatTimestamp(data?.latestSync?.completedAt || data?.latestSync?.startedAt)}
            </div>
            <div className="mt-1 text-xs text-text-secondary">
              {data?.latestSync?.mode ? `Mode: ${data.latestSync.mode}` : 'No completed sync yet'}
            </div>
          </div>
        </div>

        <div className="rounded-2xl border border-white/35 bg-white/45 px-4 py-3 text-xs text-text-secondary backdrop-blur dark:border-white/10 dark:bg-zinc-950/35">
          <div className="flex flex-wrap items-center justify-between gap-3">
            <div className="flex items-center gap-3">
              <span className="text-[11px] font-medium uppercase tracking-[0.14em] text-text-secondary">
                Memory Projection
              </span>
              <span className="inline-flex items-center gap-2">
                <span className={`h-2.5 w-2.5 rounded-full ${getProjectionTone(data?.latestProjection?.status)}`} />
                <span className="font-medium text-text-primary">
                  {getProjectionLabel(data?.latestProjection?.status)}
                </span>
              </span>
            </div>
            <div className="flex flex-wrap items-center gap-x-4 gap-y-1">
              {data?.latestProjection?.errorMessage ? (
                <span className="text-rose-700 dark:text-rose-300">
                  {data.latestProjection.errorMessage}
                </span>
              ) : null}
              {isFetching && !isLoading ? (
                <span className="inline-flex items-center text-text-secondary">
                  <Spinner className="mr-1 h-3.5 w-3.5" />
                  Refreshing status
                </span>
              ) : null}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
