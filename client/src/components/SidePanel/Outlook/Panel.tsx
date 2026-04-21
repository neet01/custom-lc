import { useEffect, useMemo, useState, startTransition } from 'react';
import type {
  OutlookAnalyzeResponse,
  OutlookDraftResponse,
  OutlookMessage,
} from 'librechat-data-provider';
import {
  useAnalyzeOutlookMessageMutation,
  useCreateOutlookDraftMutation,
  useOutlookMessageQuery,
  useOutlookMessagesQuery,
  useOutlookStatusQuery,
} from '~/data-provider';
import { cn } from '~/utils';

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

function InsightsCard({ analysis }: { analysis?: OutlookAnalyzeResponse | null }) {
  if (!analysis) {
    return null;
  }

  const { insights } = analysis;
  return (
    <div className="rounded-2xl border border-blue-500/20 bg-blue-500/5 p-3">
      <div className="text-xs font-semibold uppercase tracking-wide text-blue-600 dark:text-blue-300">
        AI Inbox Insights
      </div>
      <p className="mt-2 text-sm text-text-primary">{insights.summary}</p>
      <div className="mt-3 text-xs font-semibold text-text-primary">Suggested actions</div>
      <ul className="mt-1 list-disc space-y-1 pl-4 text-xs text-text-secondary">
        {insights.suggestedActions.map((action) => (
          <li key={action}>{action}</li>
        ))}
      </ul>
      <div className="mt-3 text-xs font-semibold text-text-primary">Signals</div>
      <ul className="mt-1 list-disc space-y-1 pl-4 text-xs text-text-secondary">
        {insights.riskSignals.map((signal) => (
          <li key={signal}>{signal}</li>
        ))}
      </ul>
      {insights.mode === 'local-extractive' && (
        <p className="mt-3 text-[11px] text-text-secondary">
          This first-pass analysis is local and extractive. It does not send email content through a
          model until model-backed analysis is explicitly wired in.
        </p>
      )}
    </div>
  );
}

export default function OutlookPanel() {
  const [selectedId, setSelectedId] = useState<string | undefined>();
  const [analysis, setAnalysis] = useState<OutlookAnalyzeResponse | null>(null);
  const [draftResult, setDraftResult] = useState<OutlookDraftResponse | null>(null);
  const [draftInstructions, setDraftInstructions] = useState('');

  const { data: status, isLoading: statusLoading } = useOutlookStatusQuery();
  const mailboxEnabled = Boolean(status?.enabled && status?.connected);

  const {
    data: messageList,
    isLoading: messagesLoading,
    refetch,
  } = useOutlookMessagesQuery(
    { folder: 'inbox', limit: 25 },
    { enabled: mailboxEnabled },
  );

  const messages = useMemo(() => messageList?.messages ?? [], [messageList?.messages]);

  const { data: selectedMessage, isLoading: messageLoading } = useOutlookMessageQuery(selectedId, {
    enabled: mailboxEnabled && Boolean(selectedId),
  });

  const analyzeMutation = useAnalyzeOutlookMessageMutation();
  const draftMutation = useCreateOutlookDraftMutation();

  useEffect(() => {
    if (!selectedId && messages.length > 0) {
      startTransition(() => setSelectedId(messages[0].id));
    }
  }, [messages, selectedId]);

  useEffect(() => {
    setAnalysis(null);
    setDraftResult(null);
    setDraftInstructions('');
  }, [selectedId]);

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
        <div className="flex items-center justify-between gap-2">
          <div>
            <h2 className="text-base font-semibold">AI Inbox</h2>
            <p className="text-xs text-text-secondary">Delegated Outlook access via Microsoft Graph</p>
          </div>
          <button
            type="button"
            className="rounded-lg border border-border-light px-3 py-1.5 text-xs font-medium hover:bg-surface-hover"
            onClick={() => refetch()}
          >
            Refresh
          </button>
        </div>
      </div>

      <div className="grid min-h-0 flex-1 grid-rows-[minmax(180px,40%)_1fr]">
        <div className="min-h-0 overflow-y-auto border-b border-border-light">
          {messagesLoading && (
            <EmptyState title="Loading messages" description="Fetching recent inbox metadata..." />
          )}
          {!messagesLoading && messages.length === 0 && (
            <EmptyState title="No messages found" description="Your recent inbox query returned no mail." />
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
                  <div className="truncate text-sm font-medium">{message.subject}</div>
                  <div className="truncate text-xs text-text-secondary">{formatSender(message)}</div>
                </div>
                <div className="whitespace-nowrap text-[11px] text-text-secondary">
                  {formatDate(message.receivedDateTime)}
                </div>
              </div>
              <p className="mt-1 line-clamp-2 text-xs text-text-secondary">{message.bodyPreview}</p>
            </button>
          ))}
        </div>

        <div className="min-h-0 overflow-y-auto p-4">
          {!selectedId && (
            <EmptyState title="Select an email" description="Choose a message to inspect or draft against." />
          )}

          {selectedId && messageLoading && (
            <EmptyState title="Loading email" description="Fetching the selected message body..." />
          )}

          {selectedMessage && (
            <div className="space-y-4">
              <div>
                <h3 className="text-lg font-semibold">{selectedMessage.subject}</h3>
                <div className="mt-1 text-xs text-text-secondary">
                  From {formatSender(selectedMessage)}
                  {selectedMessage.receivedDateTime ? ` • ${formatDate(selectedMessage.receivedDateTime)}` : ''}
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

              <div className="rounded-2xl border border-border-light bg-surface-secondary p-3">
                <pre className="max-h-56 whitespace-pre-wrap break-words font-sans text-sm text-text-primary">
                  {selectedMessage.body || selectedMessage.bodyPreview || 'No body text available.'}
                </pre>
              </div>

              <div className="flex flex-wrap gap-2">
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
                className="min-h-20 w-full rounded-xl border border-border-light bg-surface-primary p-3 text-sm outline-none focus:border-blue-500"
                placeholder="Optional drafting guidance, e.g. ask for budget owner and due date..."
                value={draftInstructions}
                onChange={(event) => setDraftInstructions(event.target.value)}
              />

              {analyzeMutation.error != null && (
                <p className="text-xs text-red-500">Unable to analyze this email.</p>
              )}
              {draftMutation.error != null && (
                <p className="text-xs text-red-500">Unable to create a draft reply.</p>
              )}

              <InsightsCard analysis={analysis} />

              {draftResult && (
                <div className="rounded-2xl border border-green-500/20 bg-green-500/5 p-3 text-sm">
                  <div className="font-semibold text-green-700 dark:text-green-300">
                    {draftResult.message}
                  </div>
                  {draftResult.bodyPreview && (
                    <p className="mt-2 text-xs text-text-secondary">{draftResult.bodyPreview}</p>
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
          )}
        </div>
      </div>
    </div>
  );
}
