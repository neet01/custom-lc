import React from 'react';
import { useQueryClient } from '@tanstack/react-query';
import {
  HoverCard,
  HoverCardContent,
  HoverCardPortal,
  HoverCardTrigger,
  Spinner,
  useToastContext,
} from '@librechat/client';
import { QueryKeys } from 'librechat-data-provider';
import {
  useCancelTeamsArchiveSyncMutation,
  useResetTeamsArchiveMutation,
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
  if (status === 'success') {
    return {
      badge: 'border-emerald-200 bg-emerald-50 text-emerald-700 dark:border-emerald-900/60 dark:bg-emerald-950/40 dark:text-emerald-300',
      dot: 'bg-emerald-500',
    };
  }

  if (status === 'partial' || status === 'paused' || status === 'cancelled') {
    return {
      badge: 'border-amber-200 bg-amber-50 text-amber-700 dark:border-amber-900/60 dark:bg-amber-950/40 dark:text-amber-300',
      dot: 'bg-amber-500',
    };
  }

  if (status === 'failure') {
    return {
      badge: 'border-rose-200 bg-rose-50 text-rose-700 dark:border-rose-900/60 dark:bg-rose-950/40 dark:text-rose-300',
      dot: 'bg-rose-500',
    };
  }

  if (status === 'running') {
    return {
      badge: 'border-sky-200 bg-sky-50 text-sky-700 dark:border-sky-900/60 dark:bg-sky-950/40 dark:text-sky-300',
      dot: 'bg-sky-500',
    };
  }

  return {
    badge: 'border-border-light bg-surface-secondary text-text-secondary',
    dot: 'bg-zinc-400 dark:bg-zinc-500',
  };
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

  if (status === 'partial') {
    return 'Partially synced';
  }

  if (status === 'failure') {
    return 'Sync failed';
  }

  if (status === 'cancelled') {
    return 'Sync cancelled';
  }

  return 'Ready to sync';
}

function actionButtonClassName(emphasis: 'primary' | 'secondary' = 'primary') {
  if (emphasis === 'secondary') {
    return 'inline-flex items-center justify-center rounded-xl border border-border-light bg-surface-primary px-3 py-2 text-sm font-medium text-text-primary transition-colors hover:bg-surface-hover disabled:cursor-not-allowed disabled:opacity-60';
  }

  return 'inline-flex items-center justify-center rounded-xl border border-border-light bg-surface-secondary px-3 py-2 text-sm font-medium text-text-primary transition-colors hover:bg-surface-hover disabled:cursor-not-allowed disabled:opacity-60';
}

export default function TeamsArchiveStatus() {
  const queryClient = useQueryClient();
  const { showToast } = useToastContext();
  const { data, isLoading } = useTeamsArchiveStatusQuery({
    refetchInterval: (statusData) =>
      statusData?.latestSync?.status === 'running' ||
      statusData?.backfillState?.status === 'discovering' ||
      statusData?.backfillState?.status === 'syncing' ||
      statusData?.latestProjection?.status === 'running' ||
      statusData?.latestProjection?.status === 'pending'
        ? 4000
        : false,
  });
  const syncMutation = useSyncTeamsArchiveMutation();
  const cancelMutation = useCancelTeamsArchiveSyncMutation();
  const resetMutation = useResetTeamsArchiveMutation();

  const syncStatus = data?.latestSync?.status ?? null;
  const backfillState = data?.backfillState;
  const backfillStatus = backfillState?.status;
  const isBackfillActive = backfillStatus === 'discovering' || backfillStatus === 'syncing';
  const isSyncing = syncStatus === 'running' || isBackfillActive;
  const isBusy = syncMutation.isLoading || cancelMutation.isLoading || resetMutation.isLoading;

  const discoveredChats = backfillState?.discoveredChatCount ?? data?.conversationCount ?? 0;
  const completedChats = backfillState?.completedChatCount ?? 0;
  const pendingChats = backfillState?.pendingChatCount ?? 0;
  const failedChats = backfillState?.failedChatCount ?? 0;
  const totalMessages = backfillState?.totalMessageCount ?? data?.messageCount ?? 0;
  const processedChats = completedChats + failedChats;
  const hasArchive = (data?.conversationCount ?? 0) > 0 || (data?.messageCount ?? 0) > 0;
  const hasBackfillBacklog = backfillStatus === 'paused' && pendingChats > 0;
  const latestPhaseValue = data?.latestSync?.phase || backfillStatus;
  const activePhase = formatPhase(latestPhaseValue);
  const effectiveStatus =
    isSyncing ? 'running' : hasBackfillBacklog ? 'paused' : syncStatus ?? 'idle';
  const tone = getStatusTone(effectiveStatus);
  const statusLabel = getStatusLabel(effectiveStatus);
  const progressSummary = isSyncing
    ? `${completedChats.toLocaleString()} complete, ${pendingChats.toLocaleString()} pending${failedChats > 0 ? `, ${failedChats.toLocaleString()} failed` : ''}`
    : hasBackfillBacklog
      ? `${pendingChats.toLocaleString()} chats remain for the next run`
      : 'Archive backlog is clear';

  const handlePrimaryAction = async () => {
    if (isSyncing) {
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
      return;
    }

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

  const handleDelete = async () => {
    if (!window.confirm('Delete archived Teams data for the current user?')) {
      return;
    }

    try {
      const result = await resetMutation.mutateAsync();
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

  const projectionCoverage = data?.projectionCoverage;
  const projectionStatus = data?.latestProjection?.status ?? 'idle';

  return (
    <div className="rounded-2xl border border-border-medium bg-surface-primary p-4 shadow-sm">
      <div className="flex flex-col gap-3 sm:flex-row sm:items-start sm:justify-between">
        <div className="min-w-0">
          <div className="text-sm font-semibold text-text-primary">Teams archive</div>
          <div className="mt-2 flex flex-wrap items-center gap-2">
            <div
              className={`inline-flex items-center gap-2 rounded-full border px-3 py-1 text-xs font-medium ${tone.badge}`}
            >
              <span className={`h-2.5 w-2.5 rounded-full ${tone.dot}`} />
              {isLoading ? 'Loading…' : statusLabel}
            </div>
            <HoverCard openDelay={1500}>
              <HoverCardTrigger asChild>
                <button
                  type="button"
                  className="text-xs text-text-secondary underline decoration-dotted underline-offset-4 hover:text-text-primary"
                >
                  Details
                </button>
              </HoverCardTrigger>
              <HoverCardPortal>
                <HoverCardContent side="bottom" align="start" className="z-[140] w-80">
                  <div className="space-y-3 text-sm">
                    <div className="font-medium text-text-primary">{progressSummary}</div>
                    <div className="space-y-1 text-text-secondary">
                      <div>
                        Phase:{' '}
                        <span className="font-medium text-text-primary">{activePhase}</span>
                      </div>
                      <div>
                        Coverage:{' '}
                        <span className="font-medium text-text-primary">
                          {discoveredChats.toLocaleString()} chats, {totalMessages.toLocaleString()}{' '}
                          messages
                        </span>
                      </div>
                      <div>
                        Progress:{' '}
                        <span className="font-medium text-text-primary">
                          {processedChats.toLocaleString()} processed,{' '}
                          {pendingChats.toLocaleString()} pending
                          {failedChats > 0 ? `, ${failedChats.toLocaleString()} failed` : ''}
                        </span>
                      </div>
                      <div>
                        Last sync:{' '}
                        <span className="font-medium text-text-primary">
                          {formatTimestamp(
                            data?.latestSync?.completedAt || data?.latestSync?.startedAt,
                          )}
                        </span>
                      </div>
                      <div>
                        Projection:{' '}
                        <span className="font-medium text-text-primary">
                          {projectionStatus}
                          {projectionCoverage?.totalConversationCount
                            ? `, ${projectionCoverage.indexedConversationCount.toLocaleString()}/${projectionCoverage.totalConversationCount.toLocaleString()} chats projected`
                            : ''}
                        </span>
                      </div>
                    </div>
                  </div>
                </HoverCardContent>
              </HoverCardPortal>
            </HoverCard>
          </div>
        </div>

        <div className="flex flex-wrap items-center gap-2">
          <button
            type="button"
            className={actionButtonClassName('primary')}
            onClick={() => void handlePrimaryAction()}
            disabled={isBusy}
          >
            {syncMutation.isLoading || cancelMutation.isLoading ? (
              <>
                <Spinner className="mr-2 h-4 w-4" />
                {isSyncing ? 'Cancelling…' : 'Starting…'}
              </>
            ) : isSyncing ? (
              'Cancel sync'
            ) : (
              'Sync archive'
            )}
          </button>

          {hasArchive ? (
            <button
              type="button"
              className={actionButtonClassName('secondary')}
              onClick={() => void handleDelete()}
              disabled={isBusy || isSyncing}
            >
              {resetMutation.isLoading ? (
                <>
                  <Spinner className="mr-2 h-4 w-4" />
                  Deleting…
                </>
              ) : (
                'Delete archive'
              )}
            </button>
          ) : null}
        </div>
      </div>
    </div>
  );
}
