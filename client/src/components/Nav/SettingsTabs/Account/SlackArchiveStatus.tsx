import React from 'react';
import { useQueryClient } from '@tanstack/react-query';
import { apiBaseUrl, QueryKeys } from 'librechat-data-provider';
import { Spinner, useToastContext } from '@librechat/client';
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
  if (status === 'running') {
    return 'text-text-primary';
  }

  if (status === 'success') {
    return 'text-emerald-700 dark:text-emerald-300';
  }

  if (status === 'partial') {
    return 'text-amber-700 dark:text-amber-300';
  }

  if (status === 'failure') {
    return 'text-rose-700 dark:text-rose-300';
  }

  if (status === 'cancelled') {
    return 'text-amber-700 dark:text-amber-300';
  }

  return 'text-text-secondary';
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

  return 'Ready for archive sync';
}

function getStatusDetail({
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
    return 'Set the GovSlack client ID, client secret, redirect URI, and state secret before install.';
  }

  if (!connected) {
    return 'Install the GovSlack app for this workspace to store archive tokens for the current user.';
  }

  if (latestError) {
    return latestError;
  }

  if (hasArchive) {
    return 'Read-only GovSlack archive data is available for search, summaries, and follow-up retrieval.';
  }

  return 'GovSlack is connected and ready to ingest the current user workspace history.';
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
    return 'Syncing…';
  }

  return 'Sync now';
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
  const isSyncing = syncMutation.isLoading || cancelMutation.isLoading || syncStatus === 'running';
  const hasArchive = (data?.conversationCount ?? 0) > 0 || (data?.messageCount ?? 0) > 0;
  const statusLabel = getStatusLabel({
    installConfigured: data?.oauth.installConfigured,
    connected: data?.oauth.connected,
    syncStatus,
  });
  const statusDetail = getStatusDetail({
    installConfigured: data?.oauth.installConfigured,
    connected: data?.oauth.connected,
    latestError: data?.latestSync?.errorMessage,
    hasArchive,
  });

  React.useEffect(() => {
    if (data?.oauth.connected) {
      setIsRedirecting(false);
    }
  }, [data?.oauth.connected]);

  const invalidateStatus = React.useCallback(async () => {
    await queryClient.invalidateQueries([QueryKeys.slackArchiveStatus]);
  }, [queryClient]);

  const handleConnect = React.useCallback(() => {
    if (!data?.oauth.installConfigured) {
      showToast({
        message:
          'GovSlack OAuth is not configured yet. Add the Slack client ID, client secret, redirect URI, and state secret first.',
        status: 'error',
      });
      return;
    }

    setIsRedirecting(true);
    const returnTo = encodeURIComponent(window.location.href);
    const url = `${apiBaseUrl()}/api/slack-archive/oauth/start?redirect=true&returnTo=${returnTo}`;
    window.location.assign(url);
  }, [data?.oauth.installConfigured, showToast]);

  const handleSync = async () => {
    if (!data?.oauth.connected) {
      handleConnect();
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

  const handleCancel = async () => {
    try {
      const result = await cancelMutation.mutateAsync();
      await invalidateStatus();
      showToast({
        message: result.message,
        status: result.cancelled ? 'success' : 'error',
      });
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Failed to cancel Slack archive sync.';
      showToast({
        message,
        status: 'error',
      });
    }
  };

  const handleReset = async () => {
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

  return (
    <div className="relative overflow-hidden rounded-[1.75rem] border border-white/40 bg-gradient-to-br from-white/85 via-white/70 to-white/45 p-4 shadow-[0_20px_60px_-24px_rgba(15,23,42,0.35)] backdrop-blur-xl dark:border-white/10 dark:from-zinc-900/85 dark:via-zinc-900/70 dark:to-neutral-950/55">
      <div className="pointer-events-none absolute inset-x-8 top-0 h-20 rounded-full bg-sky-400/10 blur-3xl" />
      <div className="relative flex flex-col gap-4">
        <div className="flex flex-col gap-3 sm:flex-row sm:items-start sm:justify-between">
          <div>
            <div className="text-sm font-semibold text-text-primary">Slack Archive</div>
            <div className="mt-1 text-xs uppercase tracking-[0.14em] text-text-secondary">
              GovSlack
            </div>
          </div>
          <div className="flex flex-wrap items-center gap-2">
            <button
              type="button"
              className="inline-flex items-center justify-center rounded-2xl border border-white/50 bg-white/70 px-4 py-2 text-sm font-medium text-text-primary shadow-sm backdrop-blur transition-colors hover:bg-white/90 disabled:cursor-not-allowed disabled:opacity-60 dark:border-white/10 dark:bg-zinc-800/70 dark:hover:bg-zinc-800/90"
              onClick={data?.oauth.connected ? handleSync : handleConnect}
              disabled={isRedirecting || syncMutation.isLoading || cancelMutation.isLoading}
            >
              {syncMutation.isLoading || isRedirecting ? (
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
            {isSyncing ? (
              <button
                type="button"
                className="inline-flex items-center justify-center rounded-2xl border border-white/50 bg-white/70 px-4 py-2 text-sm font-medium text-text-primary shadow-sm backdrop-blur transition-colors hover:bg-white/90 disabled:cursor-not-allowed disabled:opacity-60 dark:border-white/10 dark:bg-zinc-800/70 dark:hover:bg-zinc-800/90"
                onClick={handleCancel}
                disabled={cancelMutation.isLoading || resetMutation.isLoading}
              >
                {cancelMutation.isLoading ? (
                  <>
                    <Spinner className="mr-2 h-4 w-4" />
                    Cancelling…
                  </>
                ) : (
                  'Cancel sync'
                )}
              </button>
            ) : hasArchive ? (
              <button
                type="button"
                className="inline-flex items-center justify-center rounded-2xl border border-white/50 bg-white/70 px-4 py-2 text-sm font-medium text-text-primary shadow-sm backdrop-blur transition-colors hover:bg-white/90 disabled:cursor-not-allowed disabled:opacity-60 dark:border-white/10 dark:bg-zinc-800/70 dark:hover:bg-zinc-800/90"
                onClick={handleReset}
                disabled={resetMutation.isLoading}
              >
                {resetMutation.isLoading ? (
                  <>
                    <Spinner className="mr-2 h-4 w-4" />
                    Resetting…
                  </>
                ) : (
                  'Reset archive'
                )}
              </button>
            ) : null}
          </div>
        </div>

        <div className="grid gap-3 sm:grid-cols-3">
          <div className="rounded-2xl border border-white/40 bg-white/55 px-4 py-3 backdrop-blur dark:border-white/10 dark:bg-zinc-900/45">
            <div className="text-[11px] font-medium uppercase tracking-[0.14em] text-text-secondary">
              Status
            </div>
            <div className={`mt-2 text-sm font-semibold ${getStatusTone(syncStatus)}`}>
              {isLoading ? 'Loading…' : statusLabel}
            </div>
            <div className="mt-1 text-xs text-text-secondary">{statusDetail}</div>
          </div>

          <div className="rounded-2xl border border-white/40 bg-white/55 px-4 py-3 backdrop-blur dark:border-white/10 dark:bg-zinc-900/45">
            <div className="text-[11px] font-medium uppercase tracking-[0.14em] text-text-secondary">
              Connection
            </div>
            <div className="mt-2 text-sm font-semibold text-text-primary">
              {data?.oauth.connected ? 'Connected' : 'Not connected'}
            </div>
            <div className="mt-1 text-xs text-text-secondary">
              {data?.oauth.teamId
                ? `Team ${data.oauth.teamId}`
                : 'Workspace install and user link are required before archive sync.'}
            </div>
          </div>

          <div className="rounded-2xl border border-white/40 bg-white/55 px-4 py-3 backdrop-blur dark:border-white/10 dark:bg-zinc-900/45">
            <div className="text-[11px] font-medium uppercase tracking-[0.14em] text-text-secondary">
              Coverage
            </div>
            <div className="mt-2 text-sm font-semibold text-text-primary">
              {(data?.conversationCount ?? 0).toLocaleString()} conversations
            </div>
            <div className="mt-1 text-xs text-text-secondary">
              {(data?.messageCount ?? 0).toLocaleString()} archived messages
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
                Scope
              </div>
              <div className="mt-1 text-sm font-semibold text-text-primary">Channels, DMs, threads</div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
