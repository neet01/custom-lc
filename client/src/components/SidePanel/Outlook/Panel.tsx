import { useCallback, useEffect, useMemo, useRef, useState, startTransition } from 'react';
import type { ComponentType, ReactNode } from 'react';
import { useQueryClient } from '@tanstack/react-query';
import DOMPurify from 'dompurify';
import { useToastContext } from '@librechat/client';
import { AnimatePresence, motion } from 'framer-motion';
import {
  CalendarDays,
  CalendarPlus,
  CheckCircle2,
  ChevronDown,
  Loader2,
  Mail,
  RefreshCw,
  Sparkles,
  Trash2,
} from 'lucide-react';
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
const OUTLOOK_DENSITY_KEY = 'cortex.outlook.listDensity';
const DELETE_UNDO_WINDOW_MS = 8000;

type DensityMode = 'comfortable' | 'compact';

type PendingDeleteBatch = {
  id: string;
  label: string;
  messageIds: string[];
  expiresAt: number;
};

function loadDensityMode(): DensityMode {
  if (typeof window === 'undefined') {
    return 'comfortable';
  }
  const value = window.localStorage.getItem(OUTLOOK_DENSITY_KEY);
  return value === 'compact' ? 'compact' : 'comfortable';
}

function useProgressiveText(value?: string) {
  const [displayValue, setDisplayValue] = useState(value || '');

  useEffect(() => {
    const target = String(value || '');
    if (!target) {
      setDisplayValue('');
      return;
    }

    let cursor = 0;
    const step = Math.max(6, Math.ceil(target.length / 70));
    setDisplayValue('');
    const intervalId = window.setInterval(() => {
      cursor = Math.min(target.length, cursor + step);
      setDisplayValue(target.slice(0, cursor));
      if (cursor >= target.length) {
        window.clearInterval(intervalId);
      }
    }, 16);

    return () => {
      window.clearInterval(intervalId);
    };
  }, [value]);

  return displayValue;
}

function ActionButton({
  label,
  loadingLabel,
  successLabel,
  onClick,
  isLoading,
  isSuccess,
  disabled,
  className,
  icon: Icon,
}: {
  label: string;
  loadingLabel: string;
  successLabel?: string;
  onClick: () => void;
  isLoading?: boolean;
  isSuccess?: boolean;
  disabled?: boolean;
  className?: string;
  icon?: ComponentType<{ className?: string; 'aria-hidden'?: boolean }>;
}) {
  return (
    <button
      type="button"
      className={cn(
        'inline-flex items-center gap-1.5 rounded-lg px-3 py-2 text-xs font-semibold transition-all duration-150 active:scale-[0.98] disabled:cursor-not-allowed disabled:opacity-60',
        className,
      )}
      onClick={onClick}
      disabled={disabled || isLoading}
    >
      {isLoading ? <Loader2 className="h-3.5 w-3.5 animate-spin" aria-hidden="true" /> : null}
      {!isLoading && isSuccess ? <CheckCircle2 className="h-3.5 w-3.5" aria-hidden="true" /> : null}
      {!isLoading && !isSuccess && Icon ? (
        <Icon className="h-3.5 w-3.5" aria-hidden="true" />
      ) : null}
      {isLoading ? loadingLabel : isSuccess ? successLabel || label : label}
    </button>
  );
}

function CollapsiblePanel({
  title,
  defaultOpen = true,
  children,
}: {
  title: string;
  defaultOpen?: boolean;
  children: ReactNode;
}) {
  const [open, setOpen] = useState(defaultOpen);

  return (
    <section className="rounded-2xl border border-border-light bg-surface-primary shadow-sm">
      <button
        type="button"
        className="flex w-full items-center justify-between px-3 py-2 text-left text-xs font-semibold uppercase tracking-wide text-text-secondary transition-colors hover:bg-surface-hover"
        onClick={() => setOpen((current) => !current)}
      >
        <span>{title}</span>
        <ChevronDown
          className={cn(
            'h-3.5 w-3.5 transition-transform duration-150',
            open ? 'rotate-0' : '-rotate-90',
          )}
          aria-hidden="true"
        />
      </button>
      <div
        className={cn(
          'overflow-hidden transition-[max-height,opacity] duration-200',
          open ? 'max-h-[1200px] opacity-100' : 'max-h-0 opacity-0',
        )}
      >
        <div className="border-t border-border-light px-3 py-3">{children}</div>
      </div>
    </section>
  );
}

function MessageListSkeleton({ density }: { density: DensityMode }) {
  return (
    <div className="space-y-2 px-3 py-3">
      {Array.from({ length: density === 'compact' ? 10 : 7 }).map((_, index) => (
        <div
          key={index}
          className={cn(
            'animate-pulse rounded-xl border border-border-light bg-surface-secondary p-3',
            density === 'compact' ? 'h-16' : 'h-20',
          )}
        />
      ))}
    </div>
  );
}

function MessageDetailSkeleton() {
  return (
    <div className="animate-pulse space-y-3 px-5 py-4">
      <div className="h-6 w-2/3 rounded bg-surface-secondary" />
      <div className="h-4 w-1/3 rounded bg-surface-secondary" />
      <div className="h-36 rounded-2xl bg-surface-secondary" />
      <div className="h-24 rounded-2xl bg-surface-secondary" />
    </div>
  );
}

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
      placeholder.className = 'cortex-email-image-placeholder';
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
  const progressiveSummary = useProgressiveText(insights.summary);
  return (
    <div className="rounded-2xl border border-blue-500/20 bg-blue-500/5 p-4">
      <div className="flex items-center gap-2 text-xs font-semibold uppercase tracking-wide text-blue-600 dark:text-blue-300">
        <Sparkles className="h-3.5 w-3.5" aria-hidden="true" />
        AI Inbox Insights
      </div>
      <p className="mt-2 text-sm leading-6 text-text-primary">{progressiveSummary}</p>
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
            Pick one to prepare a Teams meeting draft with attendees on your calendar. Invites are
            not sent automatically.
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
                {slot.confidenceReason && (
                  <div className="mt-1 text-xs text-amber-700 dark:text-amber-300">
                    {slot.confidenceReason}
                  </div>
                )}
                <ActionButton
                  label="Prepare meeting draft"
                  loadingLabel="Preparing..."
                  className="mt-2 bg-amber-600 text-white hover:bg-amber-700"
                  onClick={() => onCreate(slot)}
                  isLoading={isCreating}
                  icon={CalendarPlus}
                />
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
          {(result.meetingDraft?.webLink || result.event?.webLink) && (
            <a
              className="ml-3 mt-2 inline-block text-xs font-medium text-green-700 hover:underline dark:text-green-300"
              href={result.meetingDraft?.webLink || result.event?.webLink}
              target="_blank"
              rel="noreferrer"
            >
              Open meeting draft
            </a>
          )}
        </div>
      )}
    </div>
  );
}

function DraftResultCard({ draftResult }: { draftResult: OutlookDraftResponse }) {
  const progressiveDraftPreview = useProgressiveText(draftResult.bodyPreview || '');

  return (
    <div className="rounded-2xl border border-green-500/20 bg-green-500/5 p-3 text-sm">
      <div className="font-semibold text-green-700 dark:text-green-300">{draftResult.message}</div>
      {draftResult.bodyPreview && (
        <p className="mt-2 max-h-24 overflow-y-auto text-xs leading-5 text-text-secondary">
          {progressiveDraftPreview}
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
  );
}

export default function OutlookPanel() {
  const queryClient = useQueryClient();
  const { showToast } = useToastContext();
  const [selectedId, setSelectedId] = useState<string | undefined>();
  const [inboxView, setInboxView] = useState<InboxView>('focused');
  const [densityMode, setDensityMode] = useState<DensityMode>(loadDensityMode);
  const [actionSuccess, setActionSuccess] = useState<Record<string, boolean>>({});
  const [analysisByMessage, setAnalysisByMessage] =
    useState<Record<string, OutlookAnalyzeResponse>>(loadCachedAnalysis);
  const [draftResultByMessage, setDraftResultByMessage] = useState<
    Record<string, OutlookDraftResponse>
  >({});
  const [meetingSlotsByMessage, setMeetingSlotsByMessage] = useState<
    Record<string, OutlookMeetingSlotsResponse>
  >({});
  const [meetingResultByMessage, setMeetingResultByMessage] = useState<
    Record<string, OutlookCreateMeetingResponse>
  >({});
  const [selectedDeleteIds, setSelectedDeleteIds] = useState<string[]>([]);
  const [pendingDeleteBatches, setPendingDeleteBatches] = useState<PendingDeleteBatch[]>([]);
  const [optimisticallyHiddenIds, setOptimisticallyHiddenIds] = useState<string[]>([]);
  const [nowMs, setNowMs] = useState(Date.now());
  const [actionRailScrolled, setActionRailScrolled] = useState(false);
  const [draftInstructions, setDraftInstructions] = useState('');
  const [statusMessage, setStatusMessage] = useState('');
  const successTimerRef = useRef<Record<string, number>>({});
  const deleteTimerRef = useRef<Record<string, number>>({});
  const pendingDeleteRef = useRef<PendingDeleteBatch[]>([]);

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
  const hiddenIdSet = useMemo(() => new Set(optimisticallyHiddenIds), [optimisticallyHiddenIds]);
  const visibleConversations = useMemo(
    () => conversations.filter((conversation) => !hiddenIdSet.has(conversation.latest.id)),
    [conversations, hiddenIdSet],
  );
  const visibleConversationIds = useMemo(
    () => visibleConversations.map((conversation) => conversation.latest.id),
    [visibleConversations],
  );
  const selectedDeleteIdSet = useMemo(() => new Set(selectedDeleteIds), [selectedDeleteIds]);
  const allVisibleSelected =
    visibleConversationIds.length > 0 &&
    visibleConversationIds.every((messageId) => selectedDeleteIdSet.has(messageId));
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
    if (visibleConversations.length === 0) {
      startTransition(() => setSelectedId(undefined));
      return;
    }
    if (
      !selectedId ||
      !visibleConversations.some((conversation) => conversation.latest.id === selectedId)
    ) {
      startTransition(() => setSelectedId(visibleConversations[0].latest.id));
    }
  }, [visibleConversations, selectedId]);

  useEffect(() => {
    setStatusMessage('');
    setActionRailScrolled(false);
  }, [inboxView, selectedId]);

  useEffect(() => {
    window.localStorage.setItem(OUTLOOK_DENSITY_KEY, densityMode);
  }, [densityMode]);

  useEffect(() => {
    const visibleIds = new Set(visibleConversationIds);
    setSelectedDeleteIds((current) => current.filter((messageId) => visibleIds.has(messageId)));
  }, [visibleConversationIds]);

  useEffect(() => {
    if (pendingDeleteBatches.length === 0) {
      return;
    }
    const intervalId = window.setInterval(() => setNowMs(Date.now()), 250);
    return () => window.clearInterval(intervalId);
  }, [pendingDeleteBatches.length]);

  useEffect(() => {
    pendingDeleteRef.current = pendingDeleteBatches;
  }, [pendingDeleteBatches]);

  useEffect(() => {
    return () => {
      Object.values(successTimerRef.current).forEach((timerId) => window.clearTimeout(timerId));
      Object.values(deleteTimerRef.current).forEach((timerId) => window.clearTimeout(timerId));
    };
  }, []);

  useEffect(() => {
    try {
      window.sessionStorage.setItem(OUTLOOK_ANALYSIS_CACHE_KEY, JSON.stringify(analysisByMessage));
    } catch {
      // Best-effort cache only; Outlook remains usable if storage is blocked or quota-limited.
    }
  }, [analysisByMessage]);

  const markActionSuccess = useCallback((key: string) => {
    if (successTimerRef.current[key]) {
      window.clearTimeout(successTimerRef.current[key]);
    }
    setActionSuccess((current) => ({ ...current, [key]: true }));
    successTimerRef.current[key] = window.setTimeout(() => {
      setActionSuccess((current) => ({ ...current, [key]: false }));
      delete successTimerRef.current[key];
    }, 1600);
  }, []);

  const upsertPendingBatches = useCallback((nextBatches: PendingDeleteBatch[]) => {
    setPendingDeleteBatches(nextBatches);
    setOptimisticallyHiddenIds(
      Array.from(new Set(nextBatches.flatMap((batch) => batch.messageIds))),
    );
  }, []);

  const deleteMessages = useCallback(
    async (messageIds: string[]) => {
      const deleted: string[] = [];
      const failed: string[] = [];
      for (const messageId of messageIds) {
        try {
          await deleteMutation.mutateAsync(messageId);
          deleted.push(messageId);
          queryClient.removeQueries([QueryKeys.outlookMessage, messageId]);
        } catch {
          failed.push(messageId);
        }
      }
      return { deleted, failed };
    },
    [deleteMutation, queryClient],
  );

  const finalizeDeleteBatch = useCallback(
    async (batchId: string) => {
      const batch = pendingDeleteRef.current.find((candidate) => candidate.id === batchId);
      if (!batch) {
        return;
      }

      const remainingBatches = pendingDeleteRef.current.filter(
        (candidate) => candidate.id !== batchId,
      );
      delete deleteTimerRef.current[batchId];
      upsertPendingBatches(remainingBatches);

      const { deleted, failed } = await deleteMessages(batch.messageIds);
      await refetch();

      if (failed.length > 0 && deleted.length > 0) {
        showToast({
          message: `Moved ${deleted.length} email(s). ${failed.length} email(s) failed.`,
          severity: 'warning',
          duration: 5000,
        });
        return;
      }
      if (failed.length > 0) {
        showToast({
          message: `Unable to delete ${failed.length} email(s).`,
          severity: 'error',
          duration: 5000,
        });
        return;
      }
      showToast({
        message: `Moved ${deleted.length} email(s) to Deleted Items.`,
        severity: 'success',
      });
    },
    [deleteMessages, refetch, showToast, upsertPendingBatches],
  );

  const queueDeleteBatch = useCallback(
    (messageIds: string[], label: string) => {
      if (messageIds.length === 0) {
        return;
      }

      const batchId = `${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;
      const expiresAt = Date.now() + DELETE_UNDO_WINDOW_MS;
      const nextBatch: PendingDeleteBatch = {
        id: batchId,
        label,
        messageIds,
        expiresAt,
      };

      const nextBatches = [...pendingDeleteRef.current, nextBatch];
      upsertPendingBatches(nextBatches);
      setSelectedDeleteIds((current) =>
        current.filter((messageId) => !messageIds.includes(messageId)),
      );
      setSelectedId((current) => (current && messageIds.includes(current) ? undefined : current));
      setStatusMessage(`Queued ${messageIds.length} email(s) for delete. Undo available.`);
      markActionSuccess('delete');
      showToast({
        message: `Queued ${messageIds.length} email(s) for delete. Undo available for 8s.`,
        severity: 'info',
        duration: 2500,
      });

      deleteTimerRef.current[batchId] = window.setTimeout(() => {
        finalizeDeleteBatch(batchId).catch(() => {
          showToast({
            message: 'Delete queue processing failed. Please retry.',
            severity: 'error',
            duration: 5000,
          });
        });
      }, DELETE_UNDO_WINDOW_MS);
    },
    [finalizeDeleteBatch, markActionSuccess, showToast, upsertPendingBatches],
  );

  const undoDeleteBatch = useCallback(
    (batchId: string) => {
      if (deleteTimerRef.current[batchId]) {
        window.clearTimeout(deleteTimerRef.current[batchId]);
        delete deleteTimerRef.current[batchId];
      }
      const nextBatches = pendingDeleteRef.current.filter((batch) => batch.id !== batchId);
      upsertPendingBatches(nextBatches);
      showToast({
        message: 'Delete undone.',
        severity: 'success',
      });
    },
    [showToast, upsertPendingBatches],
  );

  const handleAnalyze = async () => {
    if (!selectedId) {
      return;
    }
    const result = await analyzeMutation.mutateAsync(selectedId);
    setAnalysisByMessage((current) => ({ ...current, [selectedId]: result }));
    markActionSuccess('analyze');
    showToast({ message: 'Email analysis updated.', severity: 'success' });
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
    markActionSuccess('draft');
    showToast({ message: 'Reply draft refreshed.', severity: 'success' });
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
    markActionSuccess('meetingSlots');
    showToast({ message: `Found ${result.suggestions.length} time slots.`, severity: 'success' });
  };

  const handleCreateMeeting = async (slot: OutlookMeetingSlot) => {
    if (!selectedId || !meetingSlots) {
      return;
    }
    const confirmed = window.confirm(
      'Prepare an Outlook meeting draft from this slot? You can review it before sending invites.',
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
        createReplyDraft: false,
        sendInvites: false,
      },
    });
    setMeetingResultByMessage((current) => ({ ...current, [selectedId]: result }));
    markActionSuccess('meetingCreate');
    showToast({ message: 'Meeting draft prepared.', severity: 'success' });
  };

  const handleDelete = async () => {
    if (!selectedId) {
      return;
    }
    queueDeleteBatch([selectedId], 'Single message');
  };

  const toggleDeleteSelection = (messageId: string, checked: boolean) => {
    setSelectedDeleteIds((current) => {
      if (checked) {
        if (current.includes(messageId)) {
          return current;
        }
        return [...current, messageId];
      }
      return current.filter((id) => id !== messageId);
    });
  };

  const toggleSelectVisible = () => {
    if (allVisibleSelected) {
      const visibleSet = new Set(visibleConversationIds);
      setSelectedDeleteIds((current) => current.filter((messageId) => !visibleSet.has(messageId)));
      return;
    }
    setSelectedDeleteIds((current) => {
      const next = new Set(current);
      for (const messageId of visibleConversationIds) {
        next.add(messageId);
      }
      return Array.from(next);
    });
  };

  const handleBulkDelete = async () => {
    if (selectedDeleteIds.length === 0) {
      return;
    }
    queueDeleteBatch(selectedDeleteIds, 'Bulk delete');
  };

  const handleRefresh = async () => {
    await refetch();
    markActionSuccess('refresh');
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
    <div className="relative flex h-full min-h-0 flex-col bg-surface-primary text-text-primary">
      <div className="border-b border-border-light px-4 py-3">
        <div className="flex items-start justify-between gap-3">
          <div>
            <h2 className="text-base font-semibold">AI Inbox</h2>
            <p className="text-xs text-text-secondary">
              Delegated Outlook access via Microsoft Graph
            </p>
          </div>
          <ActionButton
            label="Refresh"
            loadingLabel="Refreshing..."
            successLabel="Updated"
            className="border border-border-light hover:bg-surface-hover"
            icon={RefreshCw}
            onClick={handleRefresh}
            isLoading={messagesLoading}
            isSuccess={actionSuccess.refresh}
          />
        </div>
        <div className="mt-3">
          <ViewTabs active={inboxView} onChange={setInboxView} />
        </div>
        <div className="mt-2 flex flex-wrap items-center gap-2">
          <div className="inline-flex rounded-lg border border-border-light bg-surface-secondary p-0.5">
            <button
              type="button"
              className={cn(
                'rounded-md px-2 py-1 text-[11px] font-semibold transition-colors',
                densityMode === 'comfortable'
                  ? 'bg-surface-primary text-text-primary shadow-sm'
                  : 'text-text-secondary hover:bg-surface-hover',
              )}
              onClick={() => setDensityMode('comfortable')}
            >
              Comfortable
            </button>
            <button
              type="button"
              className={cn(
                'rounded-md px-2 py-1 text-[11px] font-semibold transition-colors',
                densityMode === 'compact'
                  ? 'bg-surface-primary text-text-primary shadow-sm'
                  : 'text-text-secondary hover:bg-surface-hover',
              )}
              onClick={() => setDensityMode('compact')}
            >
              Compact
            </button>
          </div>
          <button
            type="button"
            className="rounded-lg border border-border-light px-2.5 py-1 text-[11px] font-semibold hover:bg-surface-hover disabled:opacity-60"
            onClick={toggleSelectVisible}
            disabled={visibleConversationIds.length === 0}
          >
            {allVisibleSelected ? 'Clear visible selection' : 'Select visible'}
          </button>
          {selectedDeleteIds.length > 0 && (
            <ActionButton
              label={`Delete selected (${selectedDeleteIds.length})`}
              loadingLabel="Deleting selected..."
              successLabel="Queued"
              className="inline-flex items-center gap-1.5 rounded-lg border border-red-500/30 px-2.5 py-1 text-[11px] font-semibold text-red-600 hover:bg-red-500/10 disabled:opacity-60 dark:text-red-300"
              onClick={handleBulkDelete}
              icon={Trash2}
              isSuccess={actionSuccess.delete}
            />
          )}
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
          {messagesLoading && <MessageListSkeleton density={densityMode} />}
          {!messagesLoading && visibleConversations.length === 0 && (
            <EmptyState
              title="No messages found"
              description={`Your ${inboxView === 'all' ? 'inbox' : inboxView} query returned no mail.`}
            />
          )}
          {visibleConversations.map((conversation) => {
            const message = conversation.latest;
            const threadCount = message.threadMessageCount || conversation.messages.length;
            return (
              <motion.div
                layout
                initial={{ opacity: 0, y: 4 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -4 }}
                key={conversation.id}
                className={cn(
                  'flex items-start gap-2 border-b border-border-light px-2 transition-colors duration-150',
                  densityMode === 'compact' ? 'py-1.5' : 'py-2',
                  selectedId === message.id && 'bg-surface-active-alt',
                )}
              >
                <div className="pt-1">
                  <input
                    type="checkbox"
                    aria-label={`Select ${message.subject}`}
                    className="h-4 w-4 rounded border-border-light"
                    checked={selectedDeleteIdSet.has(message.id)}
                    onChange={(event) => toggleDeleteSelection(message.id, event.target.checked)}
                  />
                </div>
                <button
                  type="button"
                  className={cn(
                    'min-w-0 flex-1 rounded-lg px-1 text-left transition-colors hover:bg-surface-hover',
                    densityMode === 'compact' ? 'py-0.5' : 'py-1',
                  )}
                  onClick={() => setSelectedId(message.id)}
                >
                  <div className="flex items-start justify-between gap-2">
                    <div className="min-w-0">
                      <div className="flex min-w-0 items-center gap-1.5">
                        <div
                          className={cn(
                            'truncate font-semibold',
                            densityMode === 'compact' ? 'text-[13px]' : 'text-sm',
                          )}
                        >
                          {message.subject}
                        </div>
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
              </motion.div>
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

          {selectedId && messageLoading && <MessageDetailSkeleton />}

          <AnimatePresence mode="wait">
            {selectedMessage && (
              <motion.div
                key={selectedMessage.id}
                initial={{ opacity: 0, y: 8 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -8 }}
                transition={{ duration: 0.18, ease: 'easeOut' }}
                className="flex min-h-0 flex-1 flex-col overflow-hidden"
              >
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
                    <ActionButton
                      label="Delete"
                      loadingLabel="Deleting..."
                      successLabel="Queued"
                      className="border border-red-500/30 text-red-600 hover:bg-red-500/10 dark:text-red-300"
                      onClick={handleDelete}
                      icon={Trash2}
                      isSuccess={actionSuccess.delete}
                    />
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
                      <motion.article
                        layout
                        initial={{ opacity: 0, y: 4 }}
                        animate={{ opacity: 1, y: 0 }}
                        transition={{ duration: 0.16, ease: 'easeOut' }}
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
                      </motion.article>
                    ))}
                  </div>
                </div>

                <div
                  className="max-h-[48vh] shrink-0 overflow-y-auto border-t border-border-light bg-surface-primary-alt px-5 py-0"
                  onScroll={(event) => setActionRailScrolled(event.currentTarget.scrollTop > 4)}
                >
                  <div className="my-4 rounded-2xl border border-border-light bg-surface-primary p-4 shadow-sm">
                    <div
                      className={cn(
                        'sticky top-0 z-[2] -mx-4 -mt-4 border-b border-border-light bg-surface-primary px-4 py-4',
                        actionRailScrolled && 'shadow-sm',
                      )}
                    >
                      <div className="flex flex-wrap items-center gap-2">
                        <ActionButton
                          label={analysis ? 'Refresh analysis' : 'Analyze email'}
                          loadingLabel="Analyzing..."
                          successLabel="Done"
                          className="bg-blue-600 text-white hover:bg-blue-700"
                          onClick={handleAnalyze}
                          isLoading={analyzeMutation.isLoading}
                          isSuccess={actionSuccess.analyze}
                          icon={Sparkles}
                        />
                        <ActionButton
                          label="Create reply draft"
                          loadingLabel="Creating draft..."
                          successLabel="Draft ready"
                          className="border border-border-light hover:bg-surface-hover"
                          onClick={handleDraft}
                          isLoading={draftMutation.isLoading}
                          isSuccess={actionSuccess.draft}
                        />
                        <ActionButton
                          label={meetingSlots ? 'Refresh meeting times' : 'Find meeting times'}
                          loadingLabel="Finding times..."
                          successLabel="Times ready"
                          className="border border-amber-500/30 text-amber-700 hover:bg-amber-500/10 dark:text-amber-300"
                          onClick={handleFindMeetingSlots}
                          isLoading={meetingSlotsMutation.isLoading}
                          isSuccess={actionSuccess.meetingSlots}
                          disabled={!status.meetingSchedulingEnabled}
                          icon={CalendarPlus}
                        />
                      </div>

                      <textarea
                        className="mt-3 max-h-32 min-h-20 w-full resize-y rounded-xl border border-border-light bg-surface-primary p-3 text-sm outline-none focus:border-blue-500"
                        placeholder="Optional drafting guidance, e.g. ask for budget owner and due date..."
                        value={draftInstructions}
                        onChange={(event) => setDraftInstructions(event.target.value)}
                      />
                      <div className="mt-2 h-3 bg-surface-primary" />
                    </div>

                    {analyzeMutation.error != null && (
                      <div className="mt-2 flex items-center justify-between rounded-lg border border-red-500/20 bg-red-500/5 px-3 py-2 text-xs text-red-600">
                        <span>Unable to analyze this email.</span>
                        <button type="button" className="underline" onClick={handleAnalyze}>
                          Retry
                        </button>
                      </div>
                    )}
                    {draftMutation.error != null && (
                      <div className="mt-2 flex items-center justify-between rounded-lg border border-red-500/20 bg-red-500/5 px-3 py-2 text-xs text-red-600">
                        <span>Unable to create a draft reply.</span>
                        <button type="button" className="underline" onClick={handleDraft}>
                          Retry
                        </button>
                      </div>
                    )}
                    {deleteMutation.error != null && (
                      <p className="mt-2 text-xs text-red-500">Unable to delete this email.</p>
                    )}
                    {meetingSlotsMutation.error != null && (
                      <div className="mt-2 flex items-center justify-between rounded-lg border border-red-500/20 bg-red-500/5 px-3 py-2 text-xs text-red-600">
                        <span>Unable to find meeting times.</span>
                        <button
                          type="button"
                          className="underline"
                          onClick={handleFindMeetingSlots}
                        >
                          Retry
                        </button>
                      </div>
                    )}
                    {createMeetingMutation.error != null && (
                      <p className="mt-2 text-xs text-red-500">Unable to create Teams meeting.</p>
                    )}
                    {!status.meetingSchedulingEnabled && (
                      <p className="mt-2 text-xs text-text-secondary">
                        Meeting scheduling is disabled. Set
                        OUTLOOK_AI_ENABLE_MEETING_SCHEDULING=true.
                      </p>
                    )}

                    <div className="mt-3 space-y-3">
                      <CollapsiblePanel title="AI Inbox insights" defaultOpen>
                        {analyzeMutation.isLoading && !analysis && (
                          <div className="h-24 animate-pulse rounded-xl bg-surface-secondary" />
                        )}
                        <InsightsCard analysis={analysis} />
                      </CollapsiblePanel>
                      <CollapsiblePanel title="Meeting scheduler" defaultOpen>
                        {meetingSlotsMutation.isLoading && !meetingSlots && (
                          <div className="h-16 animate-pulse rounded-xl bg-surface-secondary" />
                        )}
                        <MeetingSchedulerCard
                          slots={meetingSlots}
                          result={meetingResult}
                          onCreate={handleCreateMeeting}
                          isCreating={createMeetingMutation.isLoading}
                        />
                      </CollapsiblePanel>

                      {draftResult && (
                        <CollapsiblePanel title="Reply draft" defaultOpen>
                          <DraftResultCard draftResult={draftResult} />
                        </CollapsiblePanel>
                      )}
                      {draftMutation.isLoading && !draftResult && (
                        <CollapsiblePanel title="Reply draft" defaultOpen>
                          <div className="h-16 animate-pulse rounded-xl bg-surface-secondary" />
                        </CollapsiblePanel>
                      )}
                    </div>
                  </div>
                </div>
              </motion.div>
            )}
          </AnimatePresence>
        </div>
      </div>
      {pendingDeleteBatches.length > 0 && (
        <div className="pointer-events-none absolute bottom-4 right-4 z-20 flex w-[340px] max-w-[calc(100%-2rem)] flex-col gap-2">
          {pendingDeleteBatches.map((batch) => {
            const secondsLeft = Math.max(1, Math.ceil((batch.expiresAt - nowMs) / 1000));
            return (
              <div
                key={batch.id}
                className="pointer-events-auto rounded-xl border border-amber-500/30 bg-surface-primary p-3 shadow-lg"
              >
                <div className="text-xs font-semibold text-text-primary">
                  {batch.messageIds.length} email(s) queued for delete
                </div>
                <div className="mt-1 text-[11px] text-text-secondary">
                  Undo available for {secondsLeft}s.
                </div>
                <div className="mt-2 flex items-center justify-between">
                  <span className="text-[11px] text-text-secondary">{batch.label}</span>
                  <button
                    type="button"
                    className="rounded-md border border-border-light px-2 py-1 text-[11px] font-semibold transition-colors hover:bg-surface-hover"
                    onClick={() => undoDeleteBatch(batch.id)}
                  >
                    Undo
                  </button>
                </div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}
