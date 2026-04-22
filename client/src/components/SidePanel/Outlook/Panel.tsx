import { useEffect, useMemo, useState, startTransition } from 'react';
import { useQueryClient } from '@tanstack/react-query';
import DOMPurify from 'dompurify';
import { CalendarDays, CalendarPlus, Mail, RefreshCw, Sparkles, Trash2 } from 'lucide-react';
import type {
  OutlookAnalyzeResponse,
  OutlookCreateMeetingResponse,
  OutlookDraftResponse,
  OutlookMeetingSlotsResponse,
  OutlookMeetingSlot,
  OutlookMessage,
} from 'librechat-data-provider';
import { QueryKeys } from 'librechat-data-provider';
import {
  useAnalyzeOutlookMessageMutation,
  useCreateOutlookDraftMutation,
  useCreateOutlookMeetingMutation,
  useDeleteOutlookMessageMutation,
  useOutlookMessageQuery,
  useOutlookMessagesQuery,
  useOutlookStatusQuery,
  useProposeOutlookMeetingSlotsMutation,
} from '~/data-provider';
import { cn } from '~/utils';

type InboxView = 'focused' | 'other' | 'all';

type OutlookConversation = {
  id: string;
  latest: OutlookMessage;
  messages: OutlookMessage[];
};

const OUTLOOK_ANALYSIS_CACHE_KEY = 'cortex.outlook.analysisByMessage';

function loadCachedAnalysis() {
  if (typeof window === 'undefined') {
    return {};
  }

  try {
    const raw = window.sessionStorage.getItem(OUTLOOK_ANALYSIS_CACHE_KEY);
    if (!raw) {
      return {};
    }
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === 'object'
      ? (parsed as Record<string, OutlookAnalyzeResponse>)
      : {};
  } catch {
    return {};
  }
}

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

function formatMeetingDateTime(value?: { dateTime: string; timeZone?: string }) {
  if (!value?.dateTime) {
    return '';
  }
  const date = new Date(value.dateTime);
  if (Number.isNaN(date.getTime())) {
    return `${value.dateTime} ${value.timeZone || ''}`.trim();
  }
  return new Intl.DateTimeFormat(undefined, {
    weekday: 'short',
    month: 'short',
    day: 'numeric',
    hour: 'numeric',
    minute: '2-digit',
    timeZone: value.timeZone === 'UTC' ? 'UTC' : undefined,
  }).format(date);
}

function EmptyState({ title, description }: { title: string; description: string }) {
  return (
    <div className="m-4 rounded-2xl border border-border-light bg-surface-secondary p-4 text-sm">
      <div className="font-medium text-text-primary">{title}</div>
      <div className="mt-1 text-text-secondary">{description}</div>
    </div>
  );
}

function getMessageTimestamp(message: OutlookMessage) {
  return new Date(message.receivedDateTime || message.sentDateTime || 0).getTime();
}

function groupMessagesByConversation(messages: OutlookMessage[]): OutlookConversation[] {
  const groups = new Map<string, OutlookMessage[]>();
  for (const message of messages) {
    const key = message.conversationId || message.id;
    groups.set(key, [...(groups.get(key) ?? []), message]);
  }

  return Array.from(groups.entries())
    .map(([id, groupMessages]) => {
      const sorted = [...groupMessages].sort(
        (a, b) => getMessageTimestamp(b) - getMessageTimestamp(a),
      );
      return {
        id,
        latest: sorted[0],
        messages: sorted,
      };
    })
    .sort((a, b) => getMessageTimestamp(b.latest) - getMessageTimestamp(a.latest));
}

function getThreadMessages(message: OutlookMessage): OutlookMessage[] {
  if (Array.isArray(message.thread) && message.thread.length > 0) {
    return [...message.thread].sort((a, b) => getMessageTimestamp(a) - getMessageTimestamp(b));
  }
  return [message];
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

function sanitizeEmailHtml(html: string, loadRemoteImages: boolean) {
  const sanitizer = DOMPurify();
  const sanitized = sanitizer.sanitize(html, {
    ADD_ATTR: ['target', 'rel', 'referrerpolicy', 'loading'],
    FORBID_TAGS: [
      'base',
      'button',
      'embed',
      'form',
      'iframe',
      'input',
      'link',
      'meta',
      'object',
      'script',
      'select',
      'textarea',
    ],
  });

  if (typeof window === 'undefined') {
    return { html: sanitized, blockedImageCount: 0 };
  }

  const document = new DOMParser().parseFromString(sanitized, 'text/html');
  let blockedImageCount = 0;

  document.querySelectorAll('a[href]').forEach((link) => {
    link.setAttribute('target', '_blank');
    link.setAttribute('rel', 'noreferrer noopener');
  });

  document.querySelectorAll('img[src]').forEach((image) => {
    const src = image.getAttribute('src') || '';
    const isRemoteImage = /^https?:\/\//i.test(src);
    const isInlineImage = /^cid:/i.test(src);

    if ((isRemoteImage && !loadRemoteImages) || isInlineImage) {
      blockedImageCount += 1;
      const placeholder = document.createElement('span');
      placeholder.className =
        'cortex-email-image-placeholder';
      placeholder.textContent = isInlineImage ? 'Inline image unavailable' : 'Remote image blocked';
      image.replaceWith(placeholder);
      return;
    }

    image.setAttribute('loading', 'lazy');
    image.setAttribute('referrerpolicy', 'no-referrer');
    image.removeAttribute('width');
    image.removeAttribute('height');
  });

  return {
    html: document.body.innerHTML,
    blockedImageCount,
  };
}

function EmailBody({ message }: { message: OutlookMessage }) {
  const [loadRemoteImages, setLoadRemoteImages] = useState(false);
  const sanitized = useMemo(
    () => sanitizeEmailHtml(message.bodyHtml || '', loadRemoteImages),
    [message.bodyHtml, loadRemoteImages],
  );

  if (!message.bodyHtml) {
    return (
      <pre className="max-h-[32vh] overflow-y-auto whitespace-pre-wrap break-words font-sans text-sm leading-6 text-text-primary">
        {message.body || message.bodyPreview || 'No body text available.'}
      </pre>
    );
  }

  return (
    <div className="space-y-3">
      {sanitized.blockedImageCount > 0 && (
        <div className="flex flex-wrap items-center justify-between gap-2 rounded-xl border border-amber-500/20 bg-amber-500/5 px-3 py-2 text-xs text-amber-800 dark:text-amber-200">
          <span>
            {sanitized.blockedImageCount} image{sanitized.blockedImageCount === 1 ? '' : 's'}{' '}
            blocked for privacy.
          </span>
          {!loadRemoteImages && (
            <button
              type="button"
              className="rounded-lg border border-amber-500/30 px-2 py-1 font-semibold hover:bg-amber-500/10"
              onClick={() => setLoadRemoteImages(true)}
            >
              Load remote images
            </button>
          )}
        </div>
      )}
      <div
        className="max-h-[46vh] overflow-y-auto rounded-xl bg-white p-4 text-sm leading-6 text-gray-950 shadow-inner dark:bg-white dark:text-gray-950 [&_.cortex-email-image-placeholder]:my-2 [&_.cortex-email-image-placeholder]:inline-block [&_.cortex-email-image-placeholder]:rounded-lg [&_.cortex-email-image-placeholder]:border [&_.cortex-email-image-placeholder]:border-dashed [&_.cortex-email-image-placeholder]:border-gray-300 [&_.cortex-email-image-placeholder]:bg-gray-50 [&_.cortex-email-image-placeholder]:px-3 [&_.cortex-email-image-placeholder]:py-2 [&_.cortex-email-image-placeholder]:text-xs [&_.cortex-email-image-placeholder]:text-gray-500 [&_a]:text-blue-700 [&_a]:underline [&_blockquote]:border-l-4 [&_blockquote]:border-gray-300 [&_blockquote]:pl-3 [&_img]:h-auto [&_img]:max-w-full [&_table]:max-w-full [&_table]:border-collapse"
        dangerouslySetInnerHTML={{ __html: sanitized.html }}
      />
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
      {insights.identitySignals != null && insights.identitySignals.length > 0 && (
        <div className="mt-3 rounded-xl border border-cyan-500/20 bg-cyan-500/5 p-3">
          <div className="text-xs font-semibold text-text-primary">Identity context</div>
          <ul className="mt-1 list-disc space-y-1 pl-4 text-xs leading-5 text-text-secondary">
            {insights.identitySignals.map((signal) => (
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

function MeetingSchedulerCard({
  slots,
  result,
  onCreate,
  isCreating,
}: {
  slots?: OutlookMeetingSlotsResponse | null;
  result?: OutlookCreateMeetingResponse | null;
  onCreate: (slot: OutlookMeetingSlot) => void;
  isCreating: boolean;
}) {
  if (!slots && !result) {
    return null;
  }

  return (
    <div className="rounded-2xl border border-amber-500/20 bg-amber-500/5 p-3 text-sm">
      <div className="flex items-center gap-2 text-xs font-semibold uppercase tracking-wide text-amber-700 dark:text-amber-300">
        <CalendarPlus className="h-3.5 w-3.5" aria-hidden="true" />
        Meeting scheduler
      </div>

      {slots && (
        <>
          <div className="mt-2 text-xs text-text-secondary">
            Proposed {slots.suggestions.length} slot(s) for {slots.attendees.length} attendee(s).
            Pick one to create a Teams meeting on your calendar and prepare a reply draft. Invites
            are not sent automatically.
          </div>
          {slots.suggestions.length === 0 && (
            <p className="mt-2 text-xs text-red-500">
              No meeting slots were found
              {slots.emptySuggestionsReason ? `: ${slots.emptySuggestionsReason}` : '.'}
            </p>
          )}
          <div className="mt-3 space-y-2">
            {slots.suggestions.map((slot) => (
              <div
                key={slot.id}
                className="rounded-xl border border-border-light bg-surface-primary p-3"
              >
                <div className="text-sm font-semibold text-text-primary">
                  {formatMeetingDateTime(slot.start)}
                </div>
                <div className="mt-1 text-xs text-text-secondary">
                  Ends {formatMeetingDateTime(slot.end)}
                  {slot.confidence != null ? ` • Confidence ${Math.round(slot.confidence)}%` : ''}
                </div>
                {slot.suggestionReason && (
                  <div className="mt-1 text-xs text-text-secondary">{slot.suggestionReason}</div>
                )}
                <button
                  type="button"
                  className="mt-2 rounded-lg bg-amber-600 px-3 py-1.5 text-xs font-semibold text-white hover:bg-amber-700 disabled:opacity-60"
                  onClick={() => onCreate(slot)}
                  disabled={isCreating}
                >
                  {isCreating ? 'Preparing...' : 'Prepare Teams draft'}
                </button>
              </div>
            ))}
          </div>
        </>
      )}

      {result && (
        <div className="mt-3 rounded-xl border border-green-500/20 bg-green-500/5 p-3">
          <div className="font-semibold text-green-700 dark:text-green-300">{result.message}</div>
          {result.event?.onlineMeeting?.joinUrl && (
            <a
              className="mt-2 inline-block text-xs font-medium text-green-700 hover:underline dark:text-green-300"
              href={result.event.onlineMeeting.joinUrl}
              target="_blank"
              rel="noreferrer"
            >
              Open Teams meeting
            </a>
          )}
          {result.draft?.webLink && (
            <a
              className="ml-3 mt-2 inline-block text-xs font-medium text-green-700 hover:underline dark:text-green-300"
              href={result.draft.webLink}
              target="_blank"
              rel="noreferrer"
            >
              Open reply draft
            </a>
          )}
        </div>
      )}
    </div>
  );
}

export default function OutlookPanel() {
  const queryClient = useQueryClient();
  const [selectedId, setSelectedId] = useState<string | undefined>();
  const [inboxView, setInboxView] = useState<InboxView>('focused');
  const [analysisByMessage, setAnalysisByMessage] = useState<
    Record<string, OutlookAnalyzeResponse>
  >(loadCachedAnalysis);
  const [draftResultByMessage, setDraftResultByMessage] = useState<
    Record<string, OutlookDraftResponse>
  >({});
  const [meetingSlotsByMessage, setMeetingSlotsByMessage] = useState<
    Record<string, OutlookMeetingSlotsResponse>
  >({});
  const [meetingResultByMessage, setMeetingResultByMessage] = useState<
    Record<string, OutlookCreateMeetingResponse>
  >({});
  const [draftInstructions, setDraftInstructions] = useState('');
  const [statusMessage, setStatusMessage] = useState('');

  const { data: status, isLoading: statusLoading } = useOutlookStatusQuery();
  const mailboxEnabled = Boolean(status?.enabled && status?.connected);

  const {
    data: messageList,
    isLoading: messagesLoading,
    refetch,
  } = useOutlookMessagesQuery(
    { folder: 'inbox', inboxView, limit: 100 },
    { enabled: mailboxEnabled },
  );

  const messages = useMemo(() => messageList?.messages ?? [], [messageList?.messages]);
  const conversations = useMemo(() => groupMessagesByConversation(messages), [messages]);
  const analysis = selectedId ? analysisByMessage[selectedId] : null;
  const draftResult = selectedId ? draftResultByMessage[selectedId] : null;
  const meetingSlots = selectedId ? meetingSlotsByMessage[selectedId] : null;
  const meetingResult = selectedId ? meetingResultByMessage[selectedId] : null;

  const { data: selectedMessage, isLoading: messageLoading } = useOutlookMessageQuery(selectedId, {
    enabled: mailboxEnabled && Boolean(selectedId),
  });

  const analyzeMutation = useAnalyzeOutlookMessageMutation();
  const draftMutation = useCreateOutlookDraftMutation();
  const deleteMutation = useDeleteOutlookMessageMutation();
  const meetingSlotsMutation = useProposeOutlookMeetingSlotsMutation();
  const createMeetingMutation = useCreateOutlookMeetingMutation();

  useEffect(() => {
    if (conversations.length === 0) {
      startTransition(() => setSelectedId(undefined));
      return;
    }
    if (
      !selectedId ||
      !conversations.some((conversation) => conversation.latest.id === selectedId)
    ) {
      startTransition(() => setSelectedId(conversations[0].latest.id));
    }
  }, [conversations, selectedId]);

  useEffect(() => {
    setStatusMessage('');
  }, [selectedId, inboxView]);

  useEffect(() => {
    try {
      window.sessionStorage.setItem(OUTLOOK_ANALYSIS_CACHE_KEY, JSON.stringify(analysisByMessage));
    } catch {
      // Best-effort cache only; Outlook remains usable if storage is blocked or quota-limited.
    }
  }, [analysisByMessage]);

  const handleAnalyze = async () => {
    if (!selectedId) {
      return;
    }
    const result = await analyzeMutation.mutateAsync(selectedId);
    setAnalysisByMessage((current) => ({ ...current, [selectedId]: result }));
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
    setDraftResultByMessage((current) => ({ ...current, [selectedId]: result }));
  };

  const handleFindMeetingSlots = async () => {
    if (!selectedId) {
      return;
    }
    const result = await meetingSlotsMutation.mutateAsync({
      messageId: selectedId,
      payload: {
        durationMinutes: 30,
        maxCandidates: 5,
      },
    });
    setMeetingSlotsByMessage((current) => ({ ...current, [selectedId]: result }));
    setMeetingResultByMessage((current) => {
      const next = { ...current };
      delete next[selectedId];
      return next;
    });
  };

  const handleCreateMeeting = async (slot: OutlookMeetingSlot) => {
    if (!selectedId || !meetingSlots) {
      return;
    }
    const confirmed = window.confirm(
      'Create a Teams meeting on your calendar and prepare a reply draft? This will not send invites automatically.',
    );
    if (!confirmed) {
      return;
    }
    const result = await createMeetingMutation.mutateAsync({
      messageId: selectedId,
      payload: {
        slot: {
          start: slot.start,
          end: slot.end,
        },
        subject: meetingSlots.subject,
        attendees: meetingSlots.attendees,
        instructions: draftInstructions,
        createReplyDraft: true,
        sendInvites: false,
      },
    });
    setMeetingResultByMessage((current) => ({ ...current, [selectedId]: result }));
  };

  const handleDelete = async () => {
    if (!selectedId) {
      return;
    }
    const confirmed = window.confirm('Move this email to Deleted Items?');
    if (!confirmed) {
      return;
    }

    const currentIndex = conversations.findIndex(
      (conversation) => conversation.latest.id === selectedId,
    );
    const nextConversation = conversations[currentIndex + 1] ?? conversations[currentIndex - 1];
    const result = await deleteMutation.mutateAsync(selectedId);
    setStatusMessage(result.message);
    setSelectedId(nextConversation?.latest.id);
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
        <div className="mt-2 flex items-center gap-1.5 text-[11px] text-text-secondary">
          <CalendarDays className="h-3.5 w-3.5" aria-hidden="true" />
          {status.calendarContextEnabled
            ? 'Calendar context is used during Analyze/Draft when an email looks scheduling-related.'
            : 'Calendar context is off. Set OUTLOOK_AI_INCLUDE_CALENDAR=true to include scheduling context.'}
        </div>
      </div>

      <div className="grid min-h-0 flex-1 grid-cols-1 md:grid-cols-[minmax(240px,34%)_minmax(0,1fr)]">
        <div className="min-h-0 overflow-y-auto border-b border-border-light md:border-b-0 md:border-r">
          {messagesLoading && (
            <EmptyState title="Loading messages" description="Fetching recent inbox metadata..." />
          )}
          {!messagesLoading && conversations.length === 0 && (
            <EmptyState
              title="No messages found"
              description={`Your ${inboxView === 'all' ? 'inbox' : inboxView} query returned no mail.`}
            />
          )}
          {conversations.map((conversation) => {
            const message = conversation.latest;
            const threadCount = message.threadMessageCount || conversation.messages.length;
            return (
              <button
                key={conversation.id}
                type="button"
                className={cn(
                  'block w-full border-b border-border-light px-3 py-2 text-left transition-colors hover:bg-surface-hover',
                  selectedId === message.id && 'bg-surface-active-alt',
                )}
                onClick={() => setSelectedId(message.id)}
              >
                <div className="flex items-start justify-between gap-2">
                  <div className="min-w-0">
                    <div className="flex min-w-0 items-center gap-1.5">
                      <div className="truncate text-sm font-semibold">{message.subject}</div>
                      {threadCount > 1 && (
                        <span className="inline-flex shrink-0 items-center gap-1 rounded-full bg-blue-500/10 px-1.5 py-0.5 text-[10px] font-semibold text-blue-700 dark:text-blue-300">
                          <Mail className="h-3 w-3" aria-hidden="true" />
                          {threadCount}
                        </span>
                      )}
                    </div>
                    <div className="truncate text-[11px] text-text-secondary">
                      {formatSender(message)}
                    </div>
                  </div>
                  <div className="whitespace-nowrap text-[11px] text-text-secondary">
                    {formatDate(message.receivedDateTime)}
                  </div>
                </div>
                <p className="mt-0.5 line-clamp-1 text-xs leading-4 text-text-secondary">
                  {message.bodyPreview}
                </p>
                {message.inferenceClassification && (
                  <span className="mt-1 inline-flex rounded-full bg-surface-tertiary px-1.5 py-0.5 text-[9px] font-semibold uppercase tracking-wide text-text-secondary">
                    {message.inferenceClassification}
                  </span>
                )}
              </button>
            );
          })}
        </div>

        <div className="flex min-h-0 flex-col overflow-hidden">
          {!selectedId && (
            <EmptyState
              title="Select an email"
              description="Choose a message to inspect or draft against."
            />
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
                <div className="space-y-3">
                  {getThreadMessages(selectedMessage).map((threadMessage) => (
                    <article
                      key={threadMessage.id}
                      className={cn(
                        'rounded-2xl border border-border-light bg-surface-secondary p-4',
                        threadMessage.id === selectedMessage.id &&
                          'border-blue-500/30 bg-blue-500/5',
                      )}
                    >
                      <div className="mb-2 flex flex-wrap items-center justify-between gap-2 border-b border-border-light pb-2">
                        <div className="min-w-0">
                          <div className="truncate text-sm font-semibold">
                            {formatSender(threadMessage)}
                          </div>
                          <div className="text-[11px] text-text-secondary">
                            {formatDate(
                              threadMessage.receivedDateTime || threadMessage.sentDateTime,
                            )}
                          </div>
                        </div>
                        {threadMessage.id === selectedMessage.id && (
                          <span className="rounded-full bg-blue-500/10 px-2 py-0.5 text-[10px] font-semibold uppercase tracking-wide text-blue-700 dark:text-blue-300">
                            selected
                          </span>
                        )}
                      </div>
                      <EmailBody message={threadMessage} />
                    </article>
                  ))}
                </div>
              </div>

              <div className="max-h-[48vh] shrink-0 overflow-y-auto border-t border-border-light bg-surface-primary-alt px-5 py-0">
                <div className="my-4 rounded-2xl border border-border-light bg-surface-primary p-4 shadow-sm">
                  <div className="sticky top-0 z-[1] -mx-4 -mt-4 border-b border-border-light bg-surface-primary px-4 py-4 shadow-sm">
                    <div className="flex flex-wrap items-center gap-2">
                      <button
                        type="button"
                        className="rounded-lg bg-blue-600 px-3 py-2 text-xs font-semibold text-white hover:bg-blue-700 disabled:opacity-60"
                        onClick={handleAnalyze}
                        disabled={analyzeMutation.isLoading}
                      >
                        {analyzeMutation.isLoading
                          ? 'Analyzing...'
                          : analysis
                            ? 'Refresh analysis'
                            : 'Analyze email'}
                      </button>
                      <button
                        type="button"
                        className="rounded-lg border border-border-light px-3 py-2 text-xs font-semibold hover:bg-surface-hover disabled:opacity-60"
                        onClick={handleDraft}
                        disabled={draftMutation.isLoading}
                      >
                        {draftMutation.isLoading ? 'Creating draft...' : 'Create reply draft'}
                      </button>
                      <button
                        type="button"
                        className="inline-flex items-center gap-1.5 rounded-lg border border-amber-500/30 px-3 py-2 text-xs font-semibold text-amber-700 hover:bg-amber-500/10 disabled:opacity-60 dark:text-amber-300"
                        onClick={handleFindMeetingSlots}
                        disabled={
                          meetingSlotsMutation.isLoading || !status.meetingSchedulingEnabled
                        }
                      >
                        <CalendarPlus className="h-3.5 w-3.5" aria-hidden="true" />
                        {meetingSlotsMutation.isLoading
                          ? 'Finding times...'
                          : meetingSlots
                            ? 'Refresh meeting times'
                            : 'Find meeting times'}
                      </button>
                    </div>

                    <textarea
                      className="mt-3 max-h-32 min-h-20 w-full resize-y rounded-xl border border-border-light bg-surface-primary p-3 text-sm outline-none focus:border-blue-500"
                      placeholder="Optional drafting guidance, e.g. ask for budget owner and due date..."
                      value={draftInstructions}
                      onChange={(event) => setDraftInstructions(event.target.value)}
                    />
                  </div>

                  {analyzeMutation.error != null && (
                    <p className="mt-2 text-xs text-red-500">Unable to analyze this email.</p>
                  )}
                  {draftMutation.error != null && (
                    <p className="mt-2 text-xs text-red-500">Unable to create a draft reply.</p>
                  )}
                  {deleteMutation.error != null && (
                    <p className="mt-2 text-xs text-red-500">Unable to delete this email.</p>
                  )}
                  {meetingSlotsMutation.error != null && (
                    <p className="mt-2 text-xs text-red-500">Unable to find meeting times.</p>
                  )}
                  {createMeetingMutation.error != null && (
                    <p className="mt-2 text-xs text-red-500">Unable to create Teams meeting.</p>
                  )}
                  {!status.meetingSchedulingEnabled && (
                    <p className="mt-2 text-xs text-text-secondary">
                      Meeting scheduling is disabled. Set OUTLOOK_AI_ENABLE_MEETING_SCHEDULING=true.
                    </p>
                  )}

                  <div className="mt-3 space-y-3">
                    <InsightsCard analysis={analysis} />
                    <MeetingSchedulerCard
                      slots={meetingSlots}
                      result={meetingResult}
                      onCreate={handleCreateMeeting}
                      isCreating={createMeetingMutation.isLoading}
                    />

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
