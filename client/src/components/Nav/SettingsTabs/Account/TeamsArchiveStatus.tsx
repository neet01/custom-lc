import React from 'react';
import { useQueryClient } from '@tanstack/react-query';
import { Spinner, useToastContext } from '@librechat/client';
import { QueryKeys } from 'librechat-data-provider';
import { useSyncTeamsArchiveMutation, useTeamsArchiveStatusQuery } from '~/data-provider';

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
    return 'text-sky-700 dark:text-sky-300';
  }

  if (status === 'success') {
    return 'text-emerald-700 dark:text-emerald-300';
  }

  if (status === 'failure') {
    return 'text-rose-700 dark:text-rose-300';
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

  return 'Not synced';
}

export default function TeamsArchiveStatus() {
  const queryClient = useQueryClient();
  const { showToast } = useToastContext();
  const { data, isLoading, isFetching } = useTeamsArchiveStatusQuery({
    refetchInterval: (data) => (data?.latestSync?.status === 'running' ? 4000 : false),
  });
  const syncMutation = useSyncTeamsArchiveMutation();

  const syncStatus = data?.latestSync?.status ?? null;
  const isSyncing = syncMutation.isLoading || syncStatus === 'running';

  const handleSync = async () => {
    try {
      await syncMutation.mutateAsync({});
      await queryClient.invalidateQueries([QueryKeys.teamsArchiveStatus]);
      showToast({
        message: 'Teams archive sync started.',
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

  return (
    <div className="relative overflow-hidden rounded-[1.75rem] border border-white/40 bg-gradient-to-br from-white/85 via-white/70 to-white/45 p-4 shadow-[0_20px_60px_-24px_rgba(15,23,42,0.35)] backdrop-blur-xl dark:border-white/10 dark:from-slate-900/85 dark:via-slate-900/70 dark:to-slate-950/55">
      <div className="pointer-events-none absolute inset-x-8 top-0 h-20 rounded-full bg-[#f5d000]/15 blur-3xl" />
      <div className="relative flex flex-col gap-4">
        <div className="flex flex-col gap-3 sm:flex-row sm:items-start sm:justify-between">
          <div>
            <div className="text-sm font-semibold text-text-primary">Teams Archive</div>
            <p className="mt-1 max-w-2xl text-xs leading-5 text-text-secondary">
              Sync and inspect your archived Teams chats. This powers historical retrieval and the
              enterprise memory layer for Microsoft communications.
            </p>
          </div>
          <button
            type="button"
            className="inline-flex items-center justify-center rounded-2xl border border-white/50 bg-white/70 px-4 py-2 text-sm font-medium text-text-primary shadow-sm backdrop-blur transition-colors hover:bg-white/90 disabled:cursor-not-allowed disabled:opacity-60 dark:border-white/10 dark:bg-slate-800/70 dark:hover:bg-slate-800/90"
            onClick={handleSync}
            disabled={isSyncing}
          >
            {isSyncing ? (
              <>
                <Spinner className="mr-2 h-4 w-4" />
                Syncing…
              </>
            ) : (
              'Sync now'
            )}
          </button>
        </div>

        <div className="grid gap-3 sm:grid-cols-3">
          <div className="rounded-2xl border border-white/40 bg-white/55 px-4 py-3 backdrop-blur dark:border-white/10 dark:bg-slate-800/45">
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

          <div className="rounded-2xl border border-white/40 bg-white/55 px-4 py-3 backdrop-blur dark:border-white/10 dark:bg-slate-800/45">
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

          <div className="rounded-2xl border border-white/40 bg-white/55 px-4 py-3 backdrop-blur dark:border-white/10 dark:bg-slate-800/45">
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

        <div className="rounded-2xl border border-white/35 bg-white/45 px-4 py-3 text-xs text-text-secondary backdrop-blur dark:border-white/10 dark:bg-slate-900/35">
          <div className="flex flex-wrap items-center gap-x-4 gap-y-1">
            <span>Graph base: {data?.graphBaseUrl || 'Unavailable'}</span>
            <span>Scopes: {data?.graphScopes || 'Unavailable'}</span>
            {isFetching && !isLoading ? (
              <span className="inline-flex items-center text-sky-700 dark:text-sky-300">
                <Spinner className="mr-1 h-3.5 w-3.5" />
                Refreshing status
              </span>
            ) : null}
          </div>
        </div>
      </div>
    </div>
  );
}
