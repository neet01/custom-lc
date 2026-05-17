import React from 'react';
import * as Ariakit from '@ariakit/react';
import { useQueryClient } from '@tanstack/react-query';
import { DropdownPopup, Spinner, useToastContext } from '@librechat/client';
import { QueryKeys } from 'librechat-data-provider';
import { Ellipsis, Trash2, XCircle, ShieldAlert } from 'lucide-react';
import {
  useCancelTeamsArchiveSyncMutation,
  useResetTeamsArchiveMutation,
  useSyncTeamsArchiveMutation,
  useTeamsArchiveStatusQuery,
} from '~/data-provider';
import type * as t from '~/common';

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

function formatPhase(value?: string | null) {
  switch (value) {
    case 'discovering':
    case 'discovering_chats':
      return 'Discovering chats';
    case 'syncing':
    case 'syncing_messages':
      return 'Syncing messages';
    case 'paused':
      return 'Resume needed';
    case 'complete':
      return 'Complete';
    case 'failed':
    case 'failure':
      return 'Failed';
    case 'cancelled':
      return 'Cancelled';
    default:
      return 'Idle';
  }
}

function getStatusTone(status?: string | null) {
  if (status === 'running') {
    return 'text-text-primary';
  }

  if (status === 'paused') {
    return 'text-amber-700 dark:text-amber-300';
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
  if (status === 'paused') {
    return 'Resume needed';
  }

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

function getPhaseTone(value?: string | null) {
  if (value === 'complete' || value === 'success') {
    return 'text-emerald-700 dark:text-emerald-300';
  }

  if (value === 'failed' || value === 'failure' || value === 'cancelled') {
    return 'text-rose-700 dark:text-rose-300';
  }

  if (
    value === 'discovering' ||
    value === 'discovering_chats' ||
    value === 'syncing' ||
    value === 'syncing_messages' ||
    value === 'paused'
  ) {
    return 'text-amber-700 dark:text-amber-300';
  }

  return 'text-text-primary';
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

function isProjectionActive(status?: string | null) {
  return status === 'running' || status === 'pending';
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
  const { data, isLoading } = useTeamsArchiveStatusQuery({
    refetchInterval: (data) =>
      data?.latestSync?.status === 'running' ||
      data?.backfillState?.status === 'discovering' ||
      data?.backfillState?.status === 'syncing' ||
      data?.latestProjection?.status === 'running' ||
      data?.latestProjection?.status === 'pending'
        ? 4000
        : false,
  });
  const syncMutation = useSyncTeamsArchiveMutation();
  const cancelMutation = useCancelTeamsArchiveSyncMutation();
  const resetMutation = useResetTeamsArchiveMutation();
  const [confirmReset, setConfirmReset] = React.useState(false);
  const [isMenuOpen, setIsMenuOpen] = React.useState(false);

  const syncStatus = data?.latestSync?.status ?? null;
  const backfillState = data?.backfillState;
  const isBackfillActive =
    backfillState?.status === 'discovering' || backfillState?.status === 'syncing';
  const isSyncing =
    syncMutation.isLoading || cancelMutation.isLoading || syncStatus === 'running' || isBackfillActive;
  const isMutating = syncMutation.isLoading || cancelMutation.isLoading || resetMutation.isLoading;
  const discoveredChats = backfillState?.discoveredChatCount ?? data?.conversationCount ?? 0;
  const completedChats = backfillState?.completedChatCount ?? 0;
  const runningChats = backfillState?.runningChatCount ?? 0;
  const pendingChats = backfillState?.pendingChatCount ?? 0;
  const failedChats = backfillState?.failedChatCount ?? 0;
  const totalMessages = backfillState?.totalMessageCount ?? data?.messageCount ?? 0;
  const projectionCoverage = data?.projectionCoverage;
  const processedChats = completedChats + failedChats;
  const hasBackfillBacklog = backfillState?.status === 'paused' && pendingChats > 0;
  const determinateProgress =
    backfillState?.discoveryComplete && discoveredChats > 0
      ? Math.max(2, Math.min(100, Math.round((processedChats / discoveredChats) * 100)))
      : null;
  const activePhase = formatPhase(
    data?.latestSync?.phase ||
      (isBackfillActive ? backfillState?.status : undefined) ||
      'idle',
  );
  const latestPhaseValue = data?.latestSync?.phase || backfillState?.status;
  const phaseLabel = formatPhase(latestPhaseValue);
  const statusLabel = isSyncing ? 'Syncing' : getStatusLabel(syncStatus);
  const statusDetail =
    isSyncing
      ? `${activePhase}. ${completedChats.toLocaleString()} complete, ${pendingChats.toLocaleString()} pending${failedChats > 0 ? `, ${failedChats.toLocaleString()} failed` : ''}.`
      : hasBackfillBacklog
        ? `Latest run completed. ${pendingChats.toLocaleString()} chats remain pending for the next sync run${failedChats > 0 ? `, and ${failedChats.toLocaleString()} failed` : ''}.`
      : backfillState?.errorMessage ||
        data?.latestSync?.errorMessage ||
        'Background sync status for Teams chat history.';

  React.useEffect(() => {
    if (isSyncing && confirmReset) {
      setConfirmReset(false);
    }
  }, [confirmReset, isSyncing]);

  React.useEffect(() => {
    if (!isMenuOpen && confirmReset) {
      setConfirmReset(false);
    }
  }, [confirmReset, isMenuOpen]);

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

  const handleReset = async () => {
    if (!confirmReset) {
      setConfirmReset(true);
      return;
    }

    try {
      const result = await resetMutation.mutateAsync();
      setConfirmReset(false);
      setIsMenuOpen(false);
      await queryClient.invalidateQueries([QueryKeys.teamsArchiveStatus]);
      showToast({
        message: result.message,
        status: 'success',
      });
    } catch (error) {
      const message =
        error instanceof Error ? error.message : 'Failed to clear archived Teams data.';
      showToast({
        message,
        status: 'error',
      });
    }
  };

  const actionMenuItems = React.useMemo<t.MenuItemProps[]>(() => {
    const items: t.MenuItemProps[] = [];

    if (syncStatus === 'running') {
      items.push({
        id: 'cancel-sync',
        label: cancelMutation.isLoading ? 'Cancelling…' : 'Cancel sync',
        icon: cancelMutation.isLoading ? <Spinner className="size-4" /> : <XCircle className="size-4" />,
        disabled: cancelMutation.isLoading || resetMutation.isLoading,
        onClick: () => {
          void handleCancel();
        },
      });
    }

    items.push({
      id: 'delete-archive',
      label: confirmReset ? 'Confirm delete archive' : 'Delete archived Teams data',
      icon: confirmReset ? <ShieldAlert className="size-4" /> : <Trash2 className="size-4" />,
      disabled: isSyncing || isMutating,
      hideOnClick: false,
      separate: syncStatus === 'running',
      onClick: () => {
        void handleReset();
      },
    });

    if (confirmReset) {
      items.push({
        id: 'keep-archive',
        label: 'Keep archive data',
        icon: <XCircle className="size-4" />,
        hideOnClick: false,
        disabled: resetMutation.isLoading,
        onClick: () => {
          setConfirmReset(false);
          setIsMenuOpen(false);
        },
      });
    }

    return items;
  }, [
    cancelMutation.isLoading,
    confirmReset,
    isMutating,
    isSyncing,
    resetMutation.isLoading,
    syncStatus,
  ]);

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
              disabled={isSyncing || resetMutation.isLoading}
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
            <DropdownPopup
              portal={true}
              menuId="teams-archive-actions"
              focusLoop={true}
              isOpen={isMenuOpen}
              setIsOpen={setIsMenuOpen}
              unmountOnHide={true}
              className="z-[125]"
              trigger={
                <Ariakit.MenuButton
                  aria-label="Teams archive actions"
                  aria-expanded={isMenuOpen}
                  disabled={isMutating}
                  className="inline-flex h-10 w-10 items-center justify-center rounded-2xl border border-white/50 bg-white/70 text-text-primary shadow-sm backdrop-blur transition-colors hover:bg-white/90 disabled:cursor-not-allowed disabled:opacity-60 dark:border-white/10 dark:bg-zinc-800/70 dark:hover:bg-zinc-800/90"
                >
                  <Ellipsis className="size-4" aria-hidden="true" />
                </Ariakit.MenuButton>
              }
              items={actionMenuItems}
            />
          </div>
        </div>

        {isSyncing ? (
          <div className="overflow-hidden rounded-2xl border border-white/35 bg-white/45 px-4 py-3 backdrop-blur dark:border-white/10 dark:bg-zinc-900/40">
            <div className="flex items-center justify-between gap-3">
              <div>
                <div className="text-sm font-semibold text-text-primary">Indexing archive</div>
                <div className="mt-1 text-xs text-text-secondary">
                  {activePhase}. {completedChats.toLocaleString()} complete,{' '}
                  {pendingChats.toLocaleString()} pending
                  {failedChats > 0 ? `, ${failedChats.toLocaleString()} failed` : ''}.
                </div>
              </div>
              <div className="shrink-0 text-right text-xs font-medium text-text-secondary">
                <div>{activePhase}</div>
                <div className="mt-1">
                  {data?.activeSyncs ?? 0}/{data?.maxConcurrentSyncs ?? 0} active slots
                </div>
              </div>
            </div>
            <div className="mt-3 h-2 overflow-hidden rounded-full bg-black/5 dark:bg-white/10">
              {determinateProgress !== null ? (
                <div
                  className="relative h-full rounded-full bg-[#f5d000]/70 shadow-[0_0_20px_rgba(245,208,0,0.35)] transition-[width] duration-500"
                  style={{ width: `${determinateProgress}%` }}
                />
              ) : (
                <div className="relative h-full w-2/5 animate-pulse rounded-full bg-[#f5d000]/70 shadow-[0_0_20px_rgba(245,208,0,0.35)]" />
              )}
            </div>
          </div>
        ) : null}

        <div className="grid gap-3 sm:grid-cols-3">
          <div className="rounded-2xl border border-white/40 bg-white/55 px-4 py-3 backdrop-blur dark:border-white/10 dark:bg-zinc-900/45">
            <div className="text-[11px] font-medium uppercase tracking-[0.14em] text-text-secondary">
              Status
            </div>
              <div className={`mt-2 text-sm font-semibold ${getStatusTone(backfillState?.status === 'paused' ? 'paused' : syncStatus)}`}>
                {isLoading ? 'Loading…' : statusLabel}
              </div>
            <div className="mt-1 text-xs text-text-secondary">
              {statusDetail}
            </div>
          </div>

          <div className="rounded-2xl border border-white/40 bg-white/55 px-4 py-3 backdrop-blur dark:border-white/10 dark:bg-zinc-900/45">
            <div className="text-[11px] font-medium uppercase tracking-[0.14em] text-text-secondary">
              Coverage
            </div>
            <div className="mt-2 text-sm font-semibold text-text-primary">
              {discoveredChats.toLocaleString()} discovered chats
            </div>
            <div className="mt-1 text-xs text-text-secondary">
              {totalMessages.toLocaleString()} archived messages
            </div>
          </div>

          <div className="rounded-2xl border border-white/40 bg-white/55 px-4 py-3 backdrop-blur dark:border-white/10 dark:bg-zinc-900/45">
            <div className="text-[11px] font-medium uppercase tracking-[0.14em] text-text-secondary">
              Progress
            </div>
              <div className="mt-2 text-sm font-semibold text-text-primary">
              {isSyncing
                ? `${completedChats.toLocaleString()} complete`
                : hasBackfillBacklog
                  ? `${pendingChats.toLocaleString()} remaining`
                  : 'No pending chats'}
              </div>
            <div className="mt-1 text-xs text-text-secondary">
              {isSyncing
                ? `${pendingChats.toLocaleString()} pending${failedChats > 0 ? `, ${failedChats.toLocaleString()} failed` : ''}`
                : hasBackfillBacklog
                  ? `Start another sync run to continue the backfill${failedChats > 0 ? `, ${failedChats.toLocaleString()} failed` : ''}.`
                  : failedChats > 0
                    ? `${failedChats.toLocaleString()} failed chats need review`
                    : 'Archive backlog is clear.'}
            </div>
          </div>
        </div>

        <div className="rounded-2xl border border-white/35 bg-white/45 px-4 py-3 text-xs text-text-secondary backdrop-blur dark:border-white/10 dark:bg-zinc-950/35">
          <div className="flex flex-wrap items-center justify-between gap-3">
            <div>
              <span className="text-[11px] font-medium uppercase tracking-[0.14em] text-text-secondary">
                Last Sync
              </span>
              <div className="mt-1 text-sm font-semibold text-text-primary">
                {formatTimestamp(data?.latestSync?.completedAt || data?.latestSync?.startedAt)}
              </div>
            </div>
            <div className="text-right">
              <div className="text-[11px] font-medium uppercase tracking-[0.14em] text-text-secondary">
                Latest Phase
              </div>
              <div className={`mt-1 text-sm font-semibold ${getPhaseTone(latestPhaseValue)}`}>
                {phaseLabel}
              </div>
            </div>
          </div>
        </div>

        <div className="rounded-2xl border border-white/35 bg-white/45 px-4 py-3 text-xs text-text-secondary backdrop-blur dark:border-white/10 dark:bg-zinc-950/35">
          <div className="flex flex-wrap items-center justify-between gap-3">
            <div>
              <div className="flex items-center gap-3">
                <span className="text-[11px] font-medium uppercase tracking-[0.14em] text-text-secondary">
                  Memory Projection
                </span>
                <span className="inline-flex items-center gap-2">
                  {isProjectionActive(data?.latestProjection?.status) ? (
                    <Spinner className="h-3.5 w-3.5 text-amber-500" />
                  ) : (
                    <span
                      className={`h-2.5 w-2.5 rounded-full ${getProjectionTone(
                        data?.latestProjection?.status,
                      )}`}
                    />
                  )}
                  <span className="font-medium text-text-primary">
                    {getProjectionLabel(data?.latestProjection?.status)}
                  </span>
                </span>
              </div>
              {projectionCoverage ? (
                <div className="mt-1 text-xs text-text-secondary">
                  {projectionCoverage.indexedConversationCount.toLocaleString()} /{' '}
                  {projectionCoverage.totalConversationCount.toLocaleString()} chats indexed
                  {projectionCoverage.totalConversationCount > 0
                    ? ` (${projectionCoverage.coveragePercent.toFixed(1)}%)`
                    : ''}
                </div>
              ) : null}
            </div>
            <div className="flex flex-wrap items-center gap-x-4 gap-y-1">
              {data?.latestProjection?.errorMessage ? (
                <span className="text-rose-700 dark:text-rose-300">
                  {data.latestProjection.errorMessage}
                </span>
              ) : null}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
