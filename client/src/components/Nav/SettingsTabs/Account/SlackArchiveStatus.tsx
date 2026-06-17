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
import { dataService, QueryKeys } from 'librechat-data-provider';
import {
  useCancelSlackArchiveSyncMutation,
  useResetSlackArchiveMutation,
  useSlackArchiveStatusQuery,
  useSyncSlackArchiveMutation,
} from '~/data-provider';

function formatTimestamp(value?: string | null) {
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
  if (status === 'success') {
    return {
      badge: 'border-emerald-200 bg-emerald-50 text-emerald-700 dark:border-emerald-900/60 dark:bg-emerald-950/40 dark:text-emerald-300',
      dot: 'bg-emerald-500',
    };
  }

  if (status === 'partial' || status === 'cancelled') {
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

function getStatusLabel({
  installConfigured,
  connected,
  syncStatus,
}: {
  installConfigured?: boolean;
  connected?: boolean;
  syncStatus?: string | null;
}) {
  if (!installConfigured) {
    return 'OAuth not configured';
  }

  if (!connected) {
    return 'GovSlack not connected';
  }

  if (syncStatus === 'running') {
    return 'Syncing';
  }

  if (syncStatus === 'success') {
    return 'Synced';
  }

  if (syncStatus === 'partial') {
    return 'Partially synced';
  }

  if (syncStatus === 'failure') {
    return 'Sync failed';
  }

  if (syncStatus === 'cancelled') {
    return 'Sync cancelled';
  }

  return 'Ready to sync';
}

function getPrimaryActionLabel({
  installConfigured,
  connected,
  isRedirecting,
  isSyncing,
}: {
  installConfigured?: boolean;
  connected?: boolean;
  isRedirecting: boolean;
  isSyncing: boolean;
}) {
  if (!installConfigured) {
    return 'Config needed';
  }

  if (!connected) {
    return isRedirecting ? 'Opening GovSlack…' : 'Connect GovSlack';
  }

  if (isSyncing) {
    return 'Cancel sync';
  }

  return 'Sync archive';
}

function getDetailsSummary({
  installConfigured,
  connected,
  latestError,
  hasArchive,
}: {
  installConfigured?: boolean;
  connected?: boolean;
  latestError?: string;
  hasArchive: boolean;
}) {
  if (!installConfigured) {
    return 'Add the GovSlack client ID, client secret, redirect URI, and state secret before install.';
  }

  if (!connected) {
    return 'Install the GovSlack app and link the current user before sync can begin.';
  }

  if (latestError) {
    return latestError;
  }

  if (hasArchive) {
    return 'Read-only GovSlack archive data is available for search and summaries.';
  }

  return 'GovSlack is connected and ready to ingest workspace history for the current user.';
}

function actionButtonClassName(emphasis: 'primary' | 'secondary' = 'primary') {
  if (emphasis === 'secondary') {
    return 'inline-flex items-center justify-center rounded-xl border border-border-light bg-surface-primary px-3 py-2 text-sm font-medium text-text-primary transition-colors hover:bg-surface-hover disabled:cursor-not-allowed disabled:opacity-60';
  }

  return 'inline-flex items-center justify-center rounded-xl border border-border-light bg-surface-secondary px-3 py-2 text-sm font-medium text-text-primary transition-colors hover:bg-surface-hover disabled:cursor-not-allowed disabled:opacity-60';
}

export default function SlackArchiveStatus() {
  const queryClient = useQueryClient();
  const { showToast } = useToastContext();
  const { data, isLoading } = useSlackArchiveStatusQuery({
    refetchInterval: (statusData) => (statusData?.latestSync?.status === 'running' ? 4000 : false),
  });
  const syncMutation = useSyncSlackArchiveMutation();
  const cancelMutation = useCancelSlackArchiveSyncMutation();
  const resetMutation = useResetSlackArchiveMutation();
  const [isRedirecting, setIsRedirecting] = React.useState(false);

  const syncStatus = data?.latestSync?.status ?? null;
  const isSyncing = syncStatus === 'running';
  const isBusy =
    syncMutation.isLoading || cancelMutation.isLoading || resetMutation.isLoading || isRedirecting;
  const hasArchive = (data?.conversationCount ?? 0) > 0 || (data?.messageCount ?? 0) > 0;
  const statusLabel = getStatusLabel({
    installConfigured: data?.oauth.installConfigured,
    connected: data?.oauth.connected,
    syncStatus,
  });
  const tone = getStatusTone(
    !data?.oauth.installConfigured || !data?.oauth.connected ? 'failure' : syncStatus,
  );

  React.useEffect(() => {
    if (data?.oauth.connected) {
      setIsRedirecting(false);
    }
  }, [data?.oauth.connected]);

  const invalidateStatus = React.useCallback(async () => {
    await queryClient.invalidateQueries([QueryKeys.slackArchiveStatus]);
  }, [queryClient]);

  const handleConnect = React.useCallback(async () => {
    if (!data?.oauth.installConfigured) {
      showToast({
        message:
          'GovSlack OAuth is not configured yet. Add the Slack client ID, client secret, redirect URI, and state secret first.',
        status: 'error',
      });
      return;
    }

    try {
      setIsRedirecting(true);
      const result = await dataService.getSlackArchiveInstallUrl({
        returnTo: window.location.href,
      });
      window.location.assign(result.installUrl);
    } catch (error) {
      setIsRedirecting(false);
      const message =
        error instanceof Error ? error.message : 'Failed to start GovSlack connection.';
      showToast({
        message,
        status: 'error',
      });
    }
  }, [data?.oauth.installConfigured, showToast]);

  const handlePrimaryAction = async () => {
    if (!data?.oauth.connected) {
      await handleConnect();
      return;
    }

    if (isSyncing) {
      try {
        const result = await cancelMutation.mutateAsync();
        await invalidateStatus();
        showToast({
          message: result.message,
          status: result.cancelled ? 'success' : 'error',
        });
      } catch (error) {
        const message =
          error instanceof Error ? error.message : 'Failed to cancel Slack archive sync.';
        showToast({
          message,
          status: 'error',
        });
      }
      return;
    }

    try {
      await syncMutation.mutateAsync({ async: true });
      await invalidateStatus();
      showToast({
        message: 'Slack archive sync started in the background.',
        status: 'success',
      });
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Slack archive sync failed.';
      showToast({
        message,
        status: 'error',
      });
    }
  };

  const handleDelete = async () => {
    if (!window.confirm('Delete archived Slack data for the current user?')) {
      return;
    }

    try {
      const result = await resetMutation.mutateAsync();
      await invalidateStatus();
      showToast({
        message: result.message,
        status: 'success',
      });
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Failed to clear archived Slack data.';
      showToast({
        message,
        status: 'error',
      });
    }
  };

  const detailsSummary = getDetailsSummary({
    installConfigured: data?.oauth.installConfigured,
    connected: data?.oauth.connected,
    latestError: data?.latestSync?.errorMessage,
    hasArchive,
  });

  return (
    <div className="rounded-2xl border border-border-medium bg-surface-primary p-4 shadow-sm">
      <div className="flex flex-col gap-3 sm:flex-row sm:items-start sm:justify-between">
        <div className="min-w-0">
          <div className="text-sm font-semibold text-text-primary">Slack archive</div>
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
                    <div className="font-medium text-text-primary">{detailsSummary}</div>
                    <div className="space-y-1 text-text-secondary">
                      <div>
                        Connection:{' '}
                        <span className="font-medium text-text-primary">
                          {data?.oauth.connected ? 'Connected' : 'Not connected'}
                        </span>
                      </div>
                      <div>
                        Team:{' '}
                        <span className="font-medium text-text-primary">
                          {data?.oauth.teamId || 'Unavailable'}
                        </span>
                      </div>
                      <div>
                        Coverage:{' '}
                        <span className="font-medium text-text-primary">
                          {(data?.conversationCount ?? 0).toLocaleString()} conversations,{' '}
                          {(data?.messageCount ?? 0).toLocaleString()} messages
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
                        Scope:{' '}
                        <span className="font-medium text-text-primary">
                          Channels, DMs, threads
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
            {syncMutation.isLoading || cancelMutation.isLoading || isRedirecting ? (
              <>
                <Spinner className="mr-2 h-4 w-4" />
                {getPrimaryActionLabel({
                  installConfigured: data?.oauth.installConfigured,
                  connected: data?.oauth.connected,
                  isRedirecting,
                  isSyncing,
                })}
              </>
            ) : (
              getPrimaryActionLabel({
                installConfigured: data?.oauth.installConfigured,
                connected: data?.oauth.connected,
                isRedirecting,
                isSyncing,
              })
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
