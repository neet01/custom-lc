import { useEffect, useMemo, useState, startTransition } from 'react';
import { useQueryClient } from '@tanstack/react-query';
import { RefreshCw, Sparkles, Trash2 } from 'lucide-react';
import type {
  OutlookAnalyzeResponse,
  OutlookDraftResponse,
  OutlookMessage,
} from 'librechat-data-provider';
import { QueryKeys } from 'librechat-data-provider';
import {
  useAnalyzeOutlookMessageMutation,
  useCreateOutlookDraftMutation,
  useDeleteOutlookMessageMutation,
  useOutlookMessageQuery,
  useOutlookMessagesQuery,
  useOutlookStatusQuery,
} from '~/data-provider';
import { cn } from '~/utils';

type InboxView = 'focused' | 'other' | 'all';

function formatSender(message?: OutlookMessage) {
  if (!message?.from) {
    return 'Unknown sender';
  }
  return message.from.name || message.from.address || 'Unknown sender';
}

function formatDate(value?: string) {
  if (!value) {
    return '';
  }
  return new Intl.DateTimeFormat(undefined, {
    month: 'short',
    day: 'numeric',
    hour: 'numeric',
    minute: '2-digit',
  }).format(new Date(value));
}

function EmptyState({ title, description }: { title: string; description: string }) {
  return (
    <div className="m-4 rounded-2xl border border-border-light bg-surface-secondary p-4 text-sm">
      <div className="font-medium text-text-primary">{title}</div>
      <div className="mt-1 text-text-secondary">{description}</div>
    </div>
  );
}

function ViewTabs({
  active,
  onChange,
}: {
  active: InboxView;
  onChange: (view: InboxView) => void;
}) {
  const tabs: Array<{ id: InboxView; label: string }> = [
    { id: 'focused', label: 'Focused' },
    { id: 'other', label: 'Other' },
    { id: 'all', label: 'All' },
  ];

  return (
    <div className="flex rounded-xl border border-border-light bg-surface-secondary p-1">
      {tabs.map((tab) => (
        <button
          key={tab.id}
          type="button"
          className={cn(
            'flex-1 rounded-lg px-3 py-1.5 text-xs font-semibold transition-colors',
            active === tab.id
              ? 'bg-surface-primary text-text-primary shadow-sm'
              : 'text-text-secondary hover:bg-surface-hover hover:text-text-primary',
          )}
          onClick={() => onChange(tab.id)}
        >
          {tab.label}
        </button>
      ))}
    </div>
  );
}

function InsightsCard({ analysis }: { analysis?: OutlookAnalyzeResponse | null }) {
  if (!analysis) {
    return null;
  }

  const { insights } = analysis;
  return (
    <div className="rounded-2xl border border-blue-500/20 bg-blue-500/5 p-4">
      <div className="flex items-center gap-2 text-xs font-semibold uppercase tracking-wide text-blue-600 dark:text-blue-300">
        <Sparkles className="h-3.5 w-3.5" aria-hidden="true" />
        AI Inbox Insights
      </div>
      <p className="mt-2 text-sm leading-6 text-text-primary">{insights.summary}</p>
      <div className="mt-3 grid gap-3 sm:grid-cols-2">
        <div>
          <div className="text-xs font-semibold text-text-primary">Suggested actions</div>
          <ul className="mt-1 list-disc space-y-1 pl-4 text-xs leading-5 text-text-secondary">
            {insights.suggestedActions.map((action) => (
              <li key={action}>{action}</li>
            ))}
          </ul>
        </div>
        <div>
          <div className="text-xs font-semibold text-text-primary">Signals</div>
          <ul className="mt-1 list-disc space-y-1 pl-4 text-xs leading-5 text-text-secondary">
            {insights.riskSignals.map((signal) => (
              <li key={signal}>{signal}</li>
            ))}
          </ul>
        </div>
      </div>
      {insights.calendarSignals != null && insights.calendarSignals.length > 0 && (
        <div className="mt-3 rounded-xl border border-amber-500/20 bg-amber-500/5 p-3">
          <div className="text-xs font-semibold text-text-primary">Calendar context</div>
          <ul className="mt-1 list-disc space-y-1 pl-4 text-xs leading-5 text-text-secondary">
            {insights.calendarSignals.map((signal) => (
              <li key={signal}>{signal}</li>
            ))}
          </ul>
        </div>
      )}
      {insights.mode === 'local-extractive' && (
        <p className="mt-3 text-[11px] leading-4 text-text-secondary">
          This first-pass analysis is local and extractive. It does not send email content through a
          model until model-backed analysis is explicitly wired in.
        </p>
      )}
    </div>
  );
}

export default function OutlookPanel() {
  const queryClient = useQueryClient();
  const [selectedId, setSelectedId] = useState<string | undefined>();
  const [inboxView, setInboxView] = useState<InboxView>('focused');
  const [analysis, setAnalysis] = useState<OutlookAnalyzeResponse | null>(null);
  const [draftResult, setDraftResult] = useState<OutlookDraftResponse | null>(null);
  const [draftInstructions, setDraftInstructions] = useState('');
  const [statusMessage, setStatusMessage] = useState('');

  const { data: status, isLoading: statusLoading } = useOutlookStatusQuery();
  const mailboxEnabled = Boolean(status?.enabled && status?.connected);

  const {
    data: messageList,
    isLoading: messagesLoading,
    refetch,
  } = useOutlookMessagesQuery(
    { folder: 'inbox', inboxView, limit: 50 },
    { enabled: mailboxEnabled },
  );

  const messages = useMemo(() => messageList?.messages ?? [], [messageList?.messages]);

  const { data: selectedMessage, isLoading: messageLoading } = useOutlookMessageQuery(selectedId, {
    enabled: mailboxEnabled && Boolean(selectedId),
  });

  const analyzeMutation = useAnalyzeOutlookMessageMutation();
  const draftMutation = useCreateOutlookDraftMutation();
  const deleteMutation = useDeleteOutlookMessageMutation();

  useEffect(() => {
    if (messages.length === 0) {
      startTransition(() => setSelectedId(undefined));
      return;
    }
    if (!selectedId || !messages.some((message) => message.id === selectedId)) {
      startTransition(() => setSelectedId(messages[0].id));
    }
  }, [messages, selectedId]);

  useEffect(() => {
    setAnalysis(null);
    setDraftResult(null);
    setDraftInstructions('');
    setStatusMessage('');
  }, [selectedId, inboxView]);

  const handleAnalyze = async () => {
    if (!selectedId) {
      return;
    }
    const result = await analyzeMutation.mutateAsync(selectedId);
    setAnalysis(result);
  };

  const handleDraft = async () => {
    if (!selectedId) {
      return;
    }
    const result = await draftMutation.mutateAsync({
      messageId: selectedId,
      payload: {
        instructions: draftInstructions,
        tone: 'professional',
      },
    });
    setDraftResult(result);
  };

  const handleDelete = async () => {
    if (!selectedId) {
      return;
    }
    const confirmed = window.confirm('Move this email to Deleted Items?');
    if (!confirmed) {
      return;
    }

    const currentIndex = messages.findIndex((message) => message.id === selectedId);
    const nextMessage = messages[currentIndex + 1] ?? messages[currentIndex - 1];
    const result = await deleteMutation.mutateAsync(selectedId);
    setStatusMessage(result.message);
    setSelectedId(nextMessage?.id);
    queryClient.removeQueries([QueryKeys.outlookMessage, selectedId]);
    await refetch();
  };

  if (statusLoading) {
    return <EmptyState title="Loading Outlook" description="Checking mailbox configuration..." />;
  }

  if (!status?.enabled) {
    return (
      <EmptyState
        title="Outlook AI Inbox is disabled"
        description="Set OUTLOOK_AI_ENABLED=true and configure delegated Graph scopes to enable it."
      />
    );
  }

  if (!status.connected) {
    return (
      <EmptyState
        title="Connect Outlook"
        description="Sign in with Entra ID, enable OPENID_REUSE_TOKENS, and consent to the configured Microsoft Graph mail scopes."
      />
    );
  }

  return (
    <div className="flex h-full min-h-0 flex-col bg-surface-primary text-text-primary">
      <div className="border-b border-border-light px-4 py-3">
        <div className="flex items-start justify-between gap-3">
          <div>
            <h2 className="text-base font-semibold">AI Inbox</h2>
            <p className="text-xs text-text-secondary">
              Delegated Outlook access via Microsoft Graph
            </p>
          </div>
          <button
            type="button"
            className="inline-flex items-center gap-1.5 rounded-lg border border-border-light px-3 py-1.5 text-xs font-medium hover:bg-surface-hover disabled:opacity-60"
            onClick={() => refetch()}
            disabled={messagesLoading}
          >
            <RefreshCw className={cn('h-3.5 w-3.5', messagesLoading && 'animate-spin')} />
            Refresh
          </button>
        </div>
        <div className="mt-3">
          <ViewTabs active={inboxView} onChange={setInboxView} />
        </div>
      </div>

      <div className="grid min-h-0 flex-1 grid-cols-1 md:grid-cols-[minmax(240px,34%)_minmax(0,1fr)]">
        <div className="min-h-0 overflow-y-auto border-b border-border-light md:border-b-0 md:border-r">
          {messagesLoading && (
            <EmptyState title="Loading messages" description="Fetching recent inbox metadata..." />
          )}
          {!messagesLoading && messages.length === 0 && (
            <EmptyState
              title="No messages found"
              description={`Your ${inboxView === 'all' ? 'inbox' : inboxView} query returned no mail.`}
            />
          )}
          {messages.map((message) => (
            <button
              key={message.id}
              type="button"
              className={cn(
                'block w-full border-b border-border-light px-4 py-3 text-left transition-colors hover:bg-surface-hover',
                selectedId === message.id && 'bg-surface-active-alt',
              )}
              onClick={() => setSelectedId(message.id)}
            >
              <div className="flex items-start justify-between gap-2">
                <div className="min-w-0">
                  <div className="truncate text-sm font-semibold">{message.subject}</div>
                  <div className="truncate text-xs text-text-secondary">{formatSender(message)}</div>
                </div>
                <div className="whitespace-nowrap text-[11px] text-text-secondary">
                  {formatDate(message.receivedDateTime)}
                </div>
              </div>
              <p className="mt-1 line-clamp-2 text-xs leading-5 text-text-secondary">
                {message.bodyPreview}
              </p>
              {message.inferenceClassification && (
                <span className="mt-2 inline-flex rounded-full bg-surface-tertiary px-2 py-0.5 text-[10px] font-semibold uppercase tracking-wide text-text-secondary">
                  {message.inferenceClassification}
                </span>
              )}
            </button>
          ))}
        </div>

        <div className="flex min-h-0 flex-col overflow-hidden">
          {!selectedId && (
            <EmptyState title="Select an email" description="Choose a message to inspect or draft against." />
          )}

          {selectedId && messageLoading && (
            <EmptyState title="Loading email" description="Fetching the selected message body..." />
          )}

          {selectedMessage && (
            <>
              <div className="border-b border-border-light px-5 py-4">
                <div className="flex flex-wrap items-start justify-between gap-3">
                  <div className="min-w-0 flex-1">
                    <h3 className="break-words text-lg font-semibold leading-6">
                      {selectedMessage.subject}
                    </h3>
                    <div className="mt-1 text-xs text-text-secondary">
                      From {formatSender(selectedMessage)}
                      {selectedMessage.receivedDateTime
                        ? ` • ${formatDate(selectedMessage.receivedDateTime)}`
                        : ''}
                    </div>
                    {selectedMessage.webLink && (
                      <a
                        className="mt-2 inline-block text-xs font-medium text-blue-600 hover:underline dark:text-blue-300"
                        href={selectedMessage.webLink}
                        target="_blank"
                        rel="noreferrer"
                      >
                        Open in Outlook
                      </a>
                    )}
                  </div>
                  <button
                    type="button"
                    className="inline-flex items-center gap-1.5 rounded-lg border border-red-500/30 px-3 py-2 text-xs font-semibold text-red-600 hover:bg-red-500/10 disabled:opacity-60 dark:text-red-300"
                    onClick={handleDelete}
                    disabled={deleteMutation.isLoading}
                  >
                    <Trash2 className="h-3.5 w-3.5" />
                    {deleteMutation.isLoading ? 'Deleting...' : 'Delete'}
                  </button>
                </div>
                {statusMessage && (
                  <div className="mt-3 rounded-xl border border-green-500/20 bg-green-500/5 px-3 py-2 text-xs text-green-700 dark:text-green-300">
                    {statusMessage}
                  </div>
                )}
              </div>

              <div className="min-h-0 flex-1 overflow-y-auto px-5 py-4">
                <article className="rounded-2xl border border-border-light bg-surface-secondary p-4">
                  <pre className="max-h-[42vh] overflow-y-auto whitespace-pre-wrap break-words font-sans text-sm leading-6 text-text-primary">
                    {selectedMessage.body || selectedMessage.bodyPreview || 'No body text available.'}
                  </pre>
                </article>
              </div>

              <div className="border-t border-border-light bg-surface-primary-alt px-5 py-4">
                <div className="rounded-2xl border border-border-light bg-surface-primary p-4 shadow-sm">
                  <div className="flex flex-wrap items-center gap-2">
                    <button
                      type="button"
                      className="rounded-lg bg-blue-600 px-3 py-2 text-xs font-semibold text-white hover:bg-blue-700 disabled:opacity-60"
                      onClick={handleAnalyze}
                      disabled={analyzeMutation.isLoading}
                    >
                      {analyzeMutation.isLoading ? 'Analyzing...' : 'Analyze email'}
                    </button>
                    <button
                      type="button"
                      className="rounded-lg border border-border-light px-3 py-2 text-xs font-semibold hover:bg-surface-hover disabled:opacity-60"
                      onClick={handleDraft}
                      disabled={draftMutation.isLoading}
                    >
                      {draftMutation.isLoading ? 'Creating draft...' : 'Create reply draft'}
                    </button>
                  </div>

                  <textarea
                    className="mt-3 max-h-32 min-h-20 w-full resize-y rounded-xl border border-border-light bg-surface-primary p-3 text-sm outline-none focus:border-blue-500"
                    placeholder="Optional drafting guidance, e.g. ask for budget owner and due date..."
                    value={draftInstructions}
                    onChange={(event) => setDraftInstructions(event.target.value)}
                  />

                  {analyzeMutation.error != null && (
                    <p className="mt-2 text-xs text-red-500">Unable to analyze this email.</p>
                  )}
                  {draftMutation.error != null && (
                    <p className="mt-2 text-xs text-red-500">Unable to create a draft reply.</p>
                  )}
                  {deleteMutation.error != null && (
                    <p className="mt-2 text-xs text-red-500">Unable to delete this email.</p>
                  )}

                  <div className="mt-3 space-y-3">
                    <InsightsCard analysis={analysis} />

                    {draftResult && (
                      <div className="rounded-2xl border border-green-500/20 bg-green-500/5 p-3 text-sm">
                        <div className="font-semibold text-green-700 dark:text-green-300">
                          {draftResult.message}
                        </div>
                        {draftResult.bodyPreview && (
                          <p className="mt-2 max-h-24 overflow-y-auto text-xs leading-5 text-text-secondary">
                            {draftResult.bodyPreview}
                          </p>
                        )}
                        {draftResult.webLink && (
                          <a
                            className="mt-2 inline-block text-xs font-medium text-green-700 hover:underline dark:text-green-300"
                            href={draftResult.webLink}
                            target="_blank"
                            rel="noreferrer"
                          >
                            Open draft
                          </a>
                        )}
                      </div>
                    )}
                  </div>
                </div>
              </div>
            </>
          )}
        </div>
      </div>
    </div>
  );
}
