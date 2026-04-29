import {
  useCallback,
  useDeferredValue,
  useEffect,
  useMemo,
  useRef,
  useState,
  startTransition,
} from 'react';
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
  MessageSquareText,
  Minimize2,
  RefreshCw,
  Search,
  Sparkles,
  Trash2,
  X,
} from 'lucide-react';
import type {
  OutlookAnalyzeResponse,
  OutlookAnalyzeSelectionResponse,
  OutlookBrief,
  OutlookCalendarEvent,
  OutlookCalendarEventMutationRequest,
  OutlookCalendarResponse,
  OutlookCreateMeetingResponse,
  OutlookDailyBriefResponse,
  OutlookDraftResponse,
  OutlookMeetingSlotsResponse,
  OutlookMeetingSlot,
  OutlookMessage,
  OutlookMessagesResponse,
} from 'librechat-data-provider';
import { QueryKeys } from 'librechat-data-provider';
import {
  useAnalyzeOutlookMessageMutation,
  useAnalyzeSelectedOutlookMessagesMutation,
  useCreateOutlookCalendarEventMutation,
  useCreateOutlookDraftMutation,
  useCreateOutlookMeetingMutation,
  useDeleteOutlookCalendarEventMutation,
  useDeleteOutlookMessageMutation,
  useOutlookCalendarQuery,
  useOutlookDailyBriefMutation,
  useOutlookMessageQuery,
  useOutlookMessagesQuery,
  useOutlookStatusQuery,
  useUpdateOutlookCalendarEventMutation,
  useProposeOutlookMeetingSlotsMutation,
  useUpdateOutlookMessageReadStateMutation,
} from '~/data-provider';
import { cn } from '~/utils';

type InboxView = 'focused' | 'other' | 'all';
type OutlookWorkspaceTab = 'inbox' | 'calendar';
type CalendarViewMode = 'day' | 'week';
type CalendarEditorMode = 'create' | 'edit' | null;

type CalendarEventFormState = {
  subject: string;
  location: string;
  startDate: string;
  startTime: string;
  endDate: string;
  endTime: string;
  attendees: string;
  body: string;
  isOnlineMeeting: boolean;
};

type OutlookConversation = {
  id: string;
  latest: OutlookMessage;
  messages: OutlookMessage[];
};

const OUTLOOK_ANALYSIS_CACHE_KEY = 'cortex.outlook.analysisByMessage';
const OUTLOOK_DENSITY_KEY = 'cortex.outlook.listDensity';
const OUTLOOK_ASSISTANT_PANEL_SIZE_KEY = 'cortex.outlook.assistantPanelSize';
const DELETE_UNDO_WINDOW_MS = 8000;
const MAILBOX_REFRESH_INTERVAL_MS = 15000;
const ASSISTANT_PANEL_DEFAULT_WIDTH = 420;
const ASSISTANT_PANEL_DEFAULT_HEIGHT = 640;
const ASSISTANT_PANEL_MIN_WIDTH = 360;
const ASSISTANT_PANEL_MIN_HEIGHT = 420;

type DensityMode = 'comfortable' | 'compact';

type PendingDeleteBatch = {
  id: string;
  label: string;
  messageIds: string[];
  expiresAt: number;
};

type AssistantPanelSize = {
  width: number;
  height: number;
};

function loadDensityMode(): DensityMode {
  if (typeof window === 'undefined') {
    return 'comfortable';
  }
  const value = window.localStorage.getItem(OUTLOOK_DENSITY_KEY);
  return value === 'compact' ? 'compact' : 'comfortable';
}

function loadAssistantPanelSize(): AssistantPanelSize {
  if (typeof window === 'undefined') {
    return {
      width: ASSISTANT_PANEL_DEFAULT_WIDTH,
      height: ASSISTANT_PANEL_DEFAULT_HEIGHT,
    };
  }

  try {
    const raw = window.localStorage.getItem(OUTLOOK_ASSISTANT_PANEL_SIZE_KEY);
    if (!raw) {
      return {
        width: ASSISTANT_PANEL_DEFAULT_WIDTH,
        height: ASSISTANT_PANEL_DEFAULT_HEIGHT,
      };
    }

    const parsed = JSON.parse(raw);
    const width = Number(parsed?.width);
    const height = Number(parsed?.height);
    return {
      width: Number.isFinite(width) ? width : ASSISTANT_PANEL_DEFAULT_WIDTH,
      height: Number.isFinite(height) ? height : ASSISTANT_PANEL_DEFAULT_HEIGHT,
    };
  } catch {
    return {
      width: ASSISTANT_PANEL_DEFAULT_WIDTH,
      height: ASSISTANT_PANEL_DEFAULT_HEIGHT,
    };
  }
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

function getDraftTimestamp(message: OutlookMessage) {
  return new Date(
    message.lastModifiedDateTime ||
      message.createdDateTime ||
      message.receivedDateTime ||
      message.sentDateTime ||
      0,
  ).getTime();
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

function getDraftReplies(message: OutlookMessage): OutlookMessage[] {
  if (Array.isArray(message.draftReplies) && message.draftReplies.length > 0) {
    return [...message.draftReplies].sort((a, b) => getDraftTimestamp(b) - getDraftTimestamp(a));
  }
  return [];
}

function formatRecipients(recipients?: OutlookMessage['toRecipients']) {
  if (!Array.isArray(recipients) || recipients.length === 0) {
    return '';
  }
  return recipients
    .map((recipient) => recipient?.name || recipient?.address)
    .filter(Boolean)
    .join(', ');
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

function WorkspaceTabs({
  active,
  onChange,
}: {
  active: OutlookWorkspaceTab;
  onChange: (tab: OutlookWorkspaceTab) => void;
}) {
  const tabs: Array<{
    id: OutlookWorkspaceTab;
    label: string;
    icon: ComponentType<{ className?: string }>;
  }> = [
    { id: 'inbox', label: 'Inbox', icon: Mail },
    { id: 'calendar', label: 'Calendar', icon: CalendarDays },
  ];

  return (
    <div className="inline-flex rounded-xl border border-border-light bg-surface-secondary p-1">
      {tabs.map((tab) => (
        <button
          key={tab.id}
          type="button"
          className={cn(
            'inline-flex items-center gap-2 rounded-lg px-3 py-1.5 text-xs font-semibold transition-colors',
            active === tab.id
              ? 'bg-surface-primary text-text-primary shadow-sm'
              : 'text-text-secondary hover:bg-surface-hover hover:text-text-primary',
          )}
          onClick={() => onChange(tab.id)}
        >
          <tab.icon className="h-3.5 w-3.5" aria-hidden="true" />
          {tab.label}
        </button>
      ))}
    </div>
  );
}

function CalendarModeTabs({
  active,
  onChange,
}: {
  active: CalendarViewMode;
  onChange: (view: CalendarViewMode) => void;
}) {
  return (
    <div className="inline-flex rounded-lg border border-border-light bg-surface-primary p-0.5">
      {(['day', 'week'] as CalendarViewMode[]).map((view) => (
        <button
          key={view}
          type="button"
          className={cn(
            'rounded-md px-2.5 py-1 text-[11px] font-semibold capitalize transition-colors',
            active === view
              ? 'bg-surface-primary-alt text-text-primary shadow-sm'
              : 'text-text-secondary hover:bg-surface-hover hover:text-text-primary',
          )}
          onClick={() => onChange(view)}
        >
          {view}
        </button>
      ))}
    </div>
  );
}

function startOfLocalDay(date: Date) {
  const next = new Date(date);
  next.setHours(0, 0, 0, 0);
  return next;
}

function addDays(date: Date, days: number) {
  const next = new Date(date);
  next.setDate(next.getDate() + days);
  return next;
}

function startOfLocalWeek(date: Date) {
  const next = startOfLocalDay(date);
  const dayOffset = (next.getDay() + 6) % 7;
  return addDays(next, -dayOffset);
}

function toDateInputValue(date: Date) {
  const year = date.getFullYear();
  const month = `${date.getMonth() + 1}`.padStart(2, '0');
  const day = `${date.getDate()}`.padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function fromDateInputValue(value?: string) {
  if (!value) {
    return startOfLocalDay(new Date());
  }
  const parsed = new Date(`${value}T00:00:00`);
  return Number.isNaN(parsed.getTime()) ? startOfLocalDay(new Date()) : parsed;
}

function buildCalendarWindow(dateValue: string, view: CalendarViewMode) {
  const anchor = fromDateInputValue(dateValue);
  const start = view === 'week' ? startOfLocalWeek(anchor) : startOfLocalDay(anchor);
  const end = addDays(start, view === 'week' ? 7 : 1);
  return { start, end };
}

function formatCalendarHeaderDate(date: Date, view: CalendarViewMode) {
  return new Intl.DateTimeFormat(undefined, {
    weekday: view === 'week' ? 'short' : 'long',
    month: 'short',
    day: 'numeric',
  }).format(date);
}

function formatCalendarTimeRange(event: OutlookCalendarEvent) {
  if (!event.start?.dateTime || !event.end?.dateTime) {
    return 'Time unavailable';
  }
  if (event.isAllDay) {
    return 'All day';
  }

  const start = new Date(event.start.dateTime);
  const end = new Date(event.end.dateTime);
  if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime())) {
    return `${event.start.dateTime} - ${event.end.dateTime}`;
  }

  const formatter = new Intl.DateTimeFormat(undefined, {
    hour: 'numeric',
    minute: '2-digit',
  });
  return `${formatter.format(start)} - ${formatter.format(end)}`;
}

function buildCalendarBuckets(
  calendarData: OutlookCalendarResponse | undefined,
  view: CalendarViewMode,
) {
  if (!calendarData) {
    return [];
  }

  const start = new Date(calendarData.startDateTime);
  const bucketCount = view === 'week' ? 7 : 1;
  const eventMap = new Map<string, OutlookCalendarEvent[]>();

  for (const event of calendarData.events ?? []) {
    const eventStart = event.start?.dateTime ? new Date(event.start.dateTime) : null;
    const key = eventStart && !Number.isNaN(eventStart.getTime())
      ? toDateInputValue(eventStart)
      : '';
    if (!key) {
      continue;
    }
    eventMap.set(key, [...(eventMap.get(key) ?? []), event]);
  }

  return Array.from({ length: bucketCount }).map((_, index) => {
    const date = addDays(start, index);
    const key = toDateInputValue(date);
    const events = [...(eventMap.get(key) ?? [])].sort((a, b) => {
      const first = new Date(a.start?.dateTime || 0).getTime();
      const second = new Date(b.start?.dateTime || 0).getTime();
      return first - second;
    });

    return {
      key,
      date,
      label: formatCalendarHeaderDate(date, view),
      events,
    };
  });
}

function formatWorkingHours(workingHours?: OutlookCalendarResponse['workingHours']) {
  if (!workingHours?.startTime || !workingHours?.endTime) {
    return 'Working hours unavailable';
  }

  const days = Array.isArray(workingHours.daysOfWeek) ? workingHours.daysOfWeek : [];
  const shortDays = days.slice(0, 5).join(', ');
  return `${shortDays || 'Configured days'} • ${workingHours.startTime}-${workingHours.endTime} ${workingHours.timeZone || ''}`.trim();
}

function toLocalInputParts(value?: { dateTime?: string }) {
  const parsed = value?.dateTime ? new Date(value.dateTime) : null;
  if (!parsed || Number.isNaN(parsed.getTime())) {
    return {
      date: toDateInputValue(new Date()),
      time: '09:00',
    };
  }

  const hours = `${parsed.getHours()}`.padStart(2, '0');
  const minutes = `${parsed.getMinutes()}`.padStart(2, '0');
  return {
    date: toDateInputValue(parsed),
    time: `${hours}:${minutes}`,
  };
}

function serializeCalendarAttendees(event?: OutlookCalendarEvent) {
  return (event?.attendees ?? [])
    .map((attendee) => {
      if (!attendee?.address) {
        return '';
      }
      return attendee.name && attendee.name !== attendee.address
        ? `${attendee.name} <${attendee.address}>`
        : attendee.address;
    })
    .filter(Boolean)
    .join(', ');
}

function buildCalendarFormState(event?: OutlookCalendarEvent): CalendarEventFormState {
  const start = toLocalInputParts(event?.start);
  const end = toLocalInputParts(event?.end);
  return {
    subject: event?.subject || '',
    location: event?.location || '',
    startDate: start.date,
    startTime: start.time,
    endDate: end.date,
    endTime: end.time,
    attendees: serializeCalendarAttendees(event),
    body: '',
    isOnlineMeeting: Boolean(event?.isOnlineMeeting),
  };
}

function parseCalendarAttendeesInput(value: string) {
  return value
    .split(/[\n,;]+/)
    .map((part) => part.trim())
    .filter(Boolean)
    .map((part) => {
      const match = part.match(/^(.*?)<([^>]+)>$/);
      if (match) {
        return {
          name: match[1].trim(),
          address: match[2].trim(),
        };
      }
      return {
        name: '',
        address: part,
      };
    })
    .filter((attendee) => attendee.address.includes('@'));
}

function toCalendarDateTime(date: string, time: string) {
  const safeDate = String(date || '').trim();
  const safeTime = String(time || '').trim() || '09:00';
  return new Date(`${safeDate}T${safeTime}:00`);
}

function buildCalendarMutationPayload(
  form: CalendarEventFormState,
): OutlookCalendarEventMutationRequest {
  const start = toCalendarDateTime(form.startDate, form.startTime);
  const end = toCalendarDateTime(form.endDate, form.endTime);

  return {
    subject: form.subject.trim(),
    location: form.location.trim(),
    attendees: parseCalendarAttendeesInput(form.attendees),
    body: form.body.trim(),
    isOnlineMeeting: form.isOnlineMeeting,
    start: {
      dateTime: start.toISOString(),
      timeZone: 'UTC',
    },
    end: {
      dateTime: end.toISOString(),
      timeZone: 'UTC',
    },
  };
}

function CalendarWorkspace({
  calendarData,
  isLoading,
  viewMode,
  selectedEventId,
  onSelectEvent,
  editorMode,
  form,
  onFormChange,
  onStartCreate,
  onStartEdit,
  onCancelEdit,
  onSubmit,
  onDelete,
  isSubmitting,
  isDeleting,
  mutationError,
}: {
  calendarData?: OutlookCalendarResponse;
  isLoading: boolean;
  viewMode: CalendarViewMode;
  selectedEventId?: string;
  onSelectEvent: (eventId: string) => void;
  editorMode: CalendarEditorMode;
  form: CalendarEventFormState;
  onFormChange: (field: keyof CalendarEventFormState, value: string | boolean) => void;
  onStartCreate: () => void;
  onStartEdit: () => void;
  onCancelEdit: () => void;
  onSubmit: () => void;
  onDelete: () => void;
  isSubmitting: boolean;
  isDeleting: boolean;
  mutationError?: string;
}) {
  const buckets = useMemo(
    () => buildCalendarBuckets(calendarData, viewMode),
    [calendarData, viewMode],
  );
  const selectedEvent = useMemo(
    () => calendarData?.events.find((event) => event.id === selectedEventId) ?? calendarData?.events[0],
    [calendarData?.events, selectedEventId],
  );

  if (isLoading) {
    return (
      <div className="flex min-h-0 flex-1 gap-4 overflow-hidden px-4 py-4">
        <div className="min-h-0 flex-1 space-y-3 overflow-hidden">
          {Array.from({ length: viewMode === 'week' ? 3 : 5 }).map((_, index) => (
            <div
              key={index}
              className="h-28 animate-pulse rounded-2xl border border-border-light bg-surface-secondary"
            />
          ))}
        </div>
        <div className="hidden w-[360px] animate-pulse rounded-2xl border border-border-light bg-surface-secondary xl:block" />
      </div>
    );
  }

  if (!calendarData || calendarData.events.length === 0) {
    return (
      <EmptyState
        title="No calendar events"
        description="No events were returned for the selected date range."
      />
    );
  }

  return (
    <div className="flex min-h-0 flex-1 overflow-hidden">
      <div className="min-h-0 flex-1 overflow-auto px-4 py-4">
        <div
          className={cn(
            'gap-3',
            viewMode === 'week'
              ? 'grid min-w-[980px] grid-cols-7'
              : 'grid grid-cols-1',
          )}
        >
          {buckets.map((bucket) => (
            <section
              key={bucket.key}
              className="flex min-h-[420px] flex-col rounded-2xl border border-border-light bg-surface-secondary"
            >
              <div className="border-b border-border-light px-3 py-3">
                <div className="text-sm font-semibold text-text-primary">{bucket.label}</div>
                <div className="mt-0.5 text-[11px] text-text-secondary">
                  {bucket.events.length} event{bucket.events.length === 1 ? '' : 's'}
                </div>
              </div>
              <div className="min-h-0 flex-1 space-y-2 overflow-y-auto px-3 py-3">
                {bucket.events.length === 0 ? (
                  <div className="rounded-xl border border-dashed border-border-light bg-surface-primary px-3 py-4 text-xs text-text-secondary">
                    No events for this day.
                  </div>
                ) : (
                  bucket.events.map((event) => {
                    const isSelected = event.id === selectedEvent?.id;
                    return (
                      <button
                        key={event.id}
                        type="button"
                        className={cn(
                          'w-full rounded-xl border px-3 py-3 text-left transition-colors',
                          isSelected
                            ? 'border-[#f5d000]/40 bg-[#f5d000]/10'
                            : 'border-border-light bg-surface-primary hover:bg-surface-hover',
                        )}
                        onClick={() => onSelectEvent(event.id)}
                      >
                        <div className="flex items-start justify-between gap-2">
                          <div className="min-w-0">
                            <div className="line-clamp-2 text-sm font-semibold text-text-primary">
                              {event.subject}
                            </div>
                            <div className="mt-1 text-[11px] font-medium text-text-secondary">
                              {formatCalendarTimeRange(event)}
                            </div>
                          </div>
                          {event.isOnlineMeeting ? (
                            <span className="rounded-full bg-blue-500/10 px-2 py-0.5 text-[10px] font-semibold uppercase tracking-wide text-blue-700 dark:text-blue-300">
                              Teams
                            </span>
                          ) : null}
                        </div>
                        {event.location ? (
                          <div className="mt-2 line-clamp-1 text-xs text-text-secondary">
                            {event.location}
                          </div>
                        ) : null}
                        {event.attendees != null && event.attendees.length > 0 ? (
                          <div className="mt-2 text-[11px] text-text-secondary">
                            {event.attendees.length} attendee{event.attendees.length === 1 ? '' : 's'}
                          </div>
                        ) : null}
                      </button>
                    );
                  })
                )}
              </div>
            </section>
          ))}
        </div>
      </div>

      <aside className="hidden w-[360px] flex-shrink-0 border-l border-border-light bg-surface-primary xl:flex xl:flex-col">
        {editorMode ? (
          <div className="min-h-0 flex-1 overflow-y-auto px-5 py-4">
            <div className="text-xs font-semibold uppercase tracking-wide text-[#b88a00] dark:text-[#f5d000]">
              {editorMode === 'create' ? 'New calendar event' : 'Edit calendar event'}
            </div>
            <div className="mt-4 space-y-3">
              <div>
                <label className="text-[11px] font-semibold uppercase tracking-wide text-text-secondary">
                  Subject
                </label>
                <input
                  type="text"
                  className="mt-1 h-10 w-full rounded-xl border border-border-light bg-surface-secondary px-3 text-sm outline-none focus:border-blue-500"
                  value={form.subject}
                  onChange={(event) => onFormChange('subject', event.target.value)}
                />
              </div>
              <div>
                <label className="text-[11px] font-semibold uppercase tracking-wide text-text-secondary">
                  Location
                </label>
                <input
                  type="text"
                  className="mt-1 h-10 w-full rounded-xl border border-border-light bg-surface-secondary px-3 text-sm outline-none focus:border-blue-500"
                  value={form.location}
                  onChange={(event) => onFormChange('location', event.target.value)}
                />
              </div>
              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="text-[11px] font-semibold uppercase tracking-wide text-text-secondary">
                    Start date
                  </label>
                  <input
                    type="date"
                    className="mt-1 h-10 w-full rounded-xl border border-border-light bg-surface-secondary px-3 text-sm outline-none focus:border-blue-500"
                    value={form.startDate}
                    onChange={(event) => onFormChange('startDate', event.target.value)}
                  />
                </div>
                <div>
                  <label className="text-[11px] font-semibold uppercase tracking-wide text-text-secondary">
                    Start time
                  </label>
                  <input
                    type="time"
                    className="mt-1 h-10 w-full rounded-xl border border-border-light bg-surface-secondary px-3 text-sm outline-none focus:border-blue-500"
                    value={form.startTime}
                    onChange={(event) => onFormChange('startTime', event.target.value)}
                  />
                </div>
                <div>
                  <label className="text-[11px] font-semibold uppercase tracking-wide text-text-secondary">
                    End date
                  </label>
                  <input
                    type="date"
                    className="mt-1 h-10 w-full rounded-xl border border-border-light bg-surface-secondary px-3 text-sm outline-none focus:border-blue-500"
                    value={form.endDate}
                    onChange={(event) => onFormChange('endDate', event.target.value)}
                  />
                </div>
                <div>
                  <label className="text-[11px] font-semibold uppercase tracking-wide text-text-secondary">
                    End time
                  </label>
                  <input
                    type="time"
                    className="mt-1 h-10 w-full rounded-xl border border-border-light bg-surface-secondary px-3 text-sm outline-none focus:border-blue-500"
                    value={form.endTime}
                    onChange={(event) => onFormChange('endTime', event.target.value)}
                  />
                </div>
              </div>
              <div>
                <label className="text-[11px] font-semibold uppercase tracking-wide text-text-secondary">
                  Attendees
                </label>
                <textarea
                  className="mt-1 min-h-24 w-full rounded-xl border border-border-light bg-surface-secondary px-3 py-2 text-sm outline-none focus:border-blue-500"
                  value={form.attendees}
                  onChange={(event) => onFormChange('attendees', event.target.value)}
                  placeholder="name@company.com, vendor@example.com"
                />
              </div>
              {editorMode === 'create' && (
                <div>
                  <label className="text-[11px] font-semibold uppercase tracking-wide text-text-secondary">
                    Notes
                  </label>
                  <textarea
                    className="mt-1 min-h-24 w-full rounded-xl border border-border-light bg-surface-secondary px-3 py-2 text-sm outline-none focus:border-blue-500"
                    value={form.body}
                    onChange={(event) => onFormChange('body', event.target.value)}
                    placeholder="Optional meeting description"
                  />
                </div>
              )}
              <label className="flex items-center gap-2 text-sm text-text-primary">
                <input
                  type="checkbox"
                  className="h-4 w-4 rounded border-border-light accent-[#f5d000]"
                  checked={form.isOnlineMeeting}
                  onChange={(event) => onFormChange('isOnlineMeeting', event.target.checked)}
                />
                Create Teams meeting link
              </label>
              {mutationError ? (
                <div className="rounded-xl border border-red-500/20 bg-red-500/5 px-3 py-2 text-xs text-red-600">
                  {mutationError}
                </div>
              ) : null}
              <div className="flex flex-wrap gap-2 pt-2">
                <ActionButton
                  label={editorMode === 'create' ? 'Create event' : 'Save changes'}
                  loadingLabel={editorMode === 'create' ? 'Creating...' : 'Saving...'}
                  className="bg-blue-600 text-white hover:bg-blue-700"
                  onClick={onSubmit}
                  isLoading={isSubmitting}
                />
                <ActionButton
                  label="Cancel"
                  loadingLabel="Cancel"
                  className="border border-border-light hover:bg-surface-hover"
                  onClick={onCancelEdit}
                />
              </div>
            </div>
          </div>
        ) : selectedEvent ? (
          <div className="min-h-0 flex-1 overflow-y-auto px-5 py-4">
            <div className="flex items-center justify-between gap-2">
              <div className="text-xs font-semibold uppercase tracking-wide text-[#b88a00] dark:text-[#f5d000]">
                Calendar event
              </div>
              <button
                type="button"
                className="rounded-lg border border-border-light px-2 py-1 text-[11px] font-semibold hover:bg-surface-hover"
                onClick={onStartCreate}
              >
                New event
              </button>
            </div>
            <h3 className="mt-2 text-lg font-semibold text-text-primary">{selectedEvent.subject}</h3>
            <div className="mt-2 text-sm text-text-secondary">
              {formatCalendarTimeRange(selectedEvent)}
            </div>
            {selectedEvent.location ? (
              <div className="mt-3 rounded-xl border border-border-light bg-surface-secondary px-3 py-2 text-sm text-text-primary">
                {selectedEvent.location}
              </div>
            ) : null}
            <div className="mt-4 space-y-3 text-sm">
              {selectedEvent.organizer?.address ? (
                <div>
                  <div className="text-[11px] font-semibold uppercase tracking-wide text-text-secondary">
                    Organizer
                  </div>
                  <div className="mt-1 text-text-primary">
                    {selectedEvent.organizer.name || selectedEvent.organizer.address}
                  </div>
                </div>
              ) : null}
              {selectedEvent.attendees != null && selectedEvent.attendees.length > 0 ? (
                <div>
                  <div className="text-[11px] font-semibold uppercase tracking-wide text-text-secondary">
                    Attendees
                  </div>
                  <div className="mt-1 space-y-1">
                    {selectedEvent.attendees.slice(0, 12).map((attendee) => (
                      <div key={`${attendee.address}-${attendee.response || 'none'}`} className="text-text-primary">
                        {attendee.name || attendee.address}
                        {attendee.response ? (
                          <span className="ml-1 text-xs text-text-secondary">
                            ({attendee.response})
                          </span>
                        ) : null}
                      </div>
                    ))}
                  </div>
                </div>
              ) : null}
              {selectedEvent.bodyPreview ? (
                <div>
                  <div className="text-[11px] font-semibold uppercase tracking-wide text-text-secondary">
                    Preview
                  </div>
                  <p className="mt-1 whitespace-pre-wrap text-text-primary">
                    {selectedEvent.bodyPreview}
                  </p>
                </div>
              ) : null}
            </div>
            {selectedEvent.webLink ? (
              <a
                className="mt-4 inline-block text-sm font-medium text-blue-600 hover:underline dark:text-blue-300"
                href={selectedEvent.webLink}
                target="_blank"
                rel="noreferrer"
              >
                Open in Outlook
              </a>
            ) : null}
            {mutationError ? (
              <div className="mt-4 rounded-xl border border-red-500/20 bg-red-500/5 px-3 py-2 text-xs text-red-600">
                {mutationError}
              </div>
            ) : null}
            <div className="mt-4 flex flex-wrap gap-2">
              <ActionButton
                label="Edit event"
                loadingLabel="Opening..."
                className="border border-border-light hover:bg-surface-hover"
                onClick={onStartEdit}
              />
              <ActionButton
                label="Cancel event"
                loadingLabel="Cancelling..."
                className="border border-red-500/30 text-red-600 hover:bg-red-500/10 dark:text-red-300"
                onClick={onDelete}
                isLoading={isDeleting}
              />
            </div>
          </div>
        ) : (
          <div className="px-5 py-4 text-sm text-text-secondary">
            <div>Select an event to inspect it.</div>
            <button
              type="button"
              className="mt-3 rounded-lg border border-border-light px-2.5 py-1.5 text-[11px] font-semibold hover:bg-surface-hover"
              onClick={onStartCreate}
            >
              New event
            </button>
          </div>
        )}
      </aside>
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

function normalizeReadableEmailText(value?: string) {
  return String(value || '')
    .replace(/\r/g, '')
    .replace(/\u00a0/g, ' ')
    .replace(/[ \t]+$/gm, '')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function isQuotedHeaderLine(line: string) {
  return /^(from|sent|to|cc|subject|date):\s+/i.test(line.trim());
}

function isQuoteBoundary(line: string) {
  const trimmed = line.trim();
  return (
    /^-{2,}\s*original message\s*-{2,}$/i.test(trimmed) ||
    /^_{5,}$/.test(trimmed) ||
    /^on .+ wrote:$/i.test(trimmed) ||
    /^from:\s.+/i.test(trimmed) ||
    /^>/.test(trimmed)
  );
}

function stripQuotedHistory(value?: string) {
  const normalized = normalizeReadableEmailText(value);
  if (!normalized) {
    return '';
  }

  const lines = normalized.split('\n');
  const kept: string[] = [];

  for (let index = 0; index < lines.length; index += 1) {
    const line = lines[index];
    const nextLines = lines.slice(index, index + 5);
    const quotedHeaderCount = nextLines.filter(isQuotedHeaderLine).length;

    if (isQuoteBoundary(line) || quotedHeaderCount >= 2) {
      break;
    }

    kept.push(line);
  }

  return normalizeReadableEmailText(
    kept
      .join('\n')
      .replace(/\n\s*Get Outlook for .+$/i, '')
      .replace(/\n\s*Sent from my .+$/i, ''),
  );
}

function getReadableEmailBody(message: OutlookMessage) {
  const source = message.body || message.bodyPreview || '';
  const stripped = stripQuotedHistory(source);
  const fallback = normalizeReadableEmailText(source);
  const text = stripped || fallback || 'No body text available.';

  return {
    text,
    wasCleaned: Boolean(stripped && fallback && stripped.length < fallback.length * 0.9),
  };
}

function ReadableEmailBody({ text }: { text: string }) {
  const blocks = text
    .split(/\n{2,}/)
    .map((block) => block.trim())
    .filter(Boolean);

  return (
    <div className="max-h-[46vh] space-y-3 overflow-y-auto rounded-xl border border-border-light bg-surface-primary p-4 text-sm leading-6 text-text-primary">
      {blocks.map((block, index) => (
        <p key={`${index}-${block.slice(0, 24)}`} className="whitespace-pre-wrap break-words">
          {block}
        </p>
      ))}
    </div>
  );
}

function EmailBody({ message }: { message: OutlookMessage }) {
  const [loadRemoteImages, setLoadRemoteImages] = useState(false);
  const [showSimplified, setShowSimplified] = useState(false);
  const sanitized = useMemo(
    () => sanitizeEmailHtml(message.bodyHtml || '', loadRemoteImages),
    [message.bodyHtml, loadRemoteImages],
  );
  const readable = useMemo(
    () => getReadableEmailBody(message),
    [message.body, message.bodyPreview],
  );
  const hasOriginal = Boolean(message.bodyHtml || message.body);

  if (!hasOriginal) {
    return (
      <pre className="max-h-[32vh] overflow-y-auto whitespace-pre-wrap break-words font-sans text-sm leading-6 text-text-primary">
        {message.bodyPreview || 'No body text available.'}
      </pre>
    );
  }

  if (showSimplified) {
    return (
      <div className="space-y-2">
        <div className="flex flex-wrap items-center justify-between gap-2">
          <div className="text-[11px] font-semibold uppercase tracking-wide text-text-secondary">
            Simplified view{readable.wasCleaned ? ' • quoted history hidden' : ''}
          </div>
          <button
            type="button"
            className="rounded-lg border border-border-light px-2.5 py-1 text-[11px] font-semibold transition-colors hover:bg-surface-hover"
            onClick={() => setShowSimplified(false)}
          >
            Display HTML view
          </button>
        </div>
        <ReadableEmailBody text={readable.text} />
      </div>
    );
  }

  return (
    <div className="space-y-3">
      <div className="flex justify-end">
        <button
          type="button"
          className="rounded-lg border border-border-light px-2.5 py-1 text-[11px] font-semibold transition-colors hover:bg-surface-hover"
          onClick={() => setShowSimplified(true)}
        >
          Display simplified view
        </button>
      </div>
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
        className="max-h-[46vh] overflow-y-auto rounded-xl border border-gray-200 bg-white p-4 text-sm leading-6 text-gray-950 shadow-inner dark:border-gray-200 dark:bg-white dark:text-gray-950 [&_*]:max-w-full [&_.cortex-email-image-placeholder]:my-2 [&_.cortex-email-image-placeholder]:inline-block [&_.cortex-email-image-placeholder]:rounded-lg [&_.cortex-email-image-placeholder]:border [&_.cortex-email-image-placeholder]:border-dashed [&_.cortex-email-image-placeholder]:border-gray-300 [&_.cortex-email-image-placeholder]:bg-gray-50 [&_.cortex-email-image-placeholder]:px-3 [&_.cortex-email-image-placeholder]:py-2 [&_.cortex-email-image-placeholder]:text-xs [&_.cortex-email-image-placeholder]:text-gray-500 [&_a]:text-blue-700 [&_a]:underline [&_blockquote]:border-l-4 [&_blockquote]:border-gray-300 [&_blockquote]:pl-3 [&_img]:h-auto [&_img]:max-w-full [&_table]:max-w-full [&_table]:border-collapse"
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

function BriefCard({ brief, metadata }: { brief?: OutlookBrief | null; metadata?: ReactNode }) {
  if (!brief) {
    return null;
  }

  const progressiveSummary = useProgressiveText(brief.summary);
  const sections = [
    { title: 'Priorities', items: brief.priorities },
    { title: 'Follow-ups', items: brief.followUps },
    { title: 'Meeting highlights', items: brief.meetingHighlights },
    { title: 'Notable emails', items: brief.notableEmails },
    { title: 'Risks', items: brief.risks },
  ].filter((section) => Array.isArray(section.items) && section.items.length > 0);

  return (
    <div className="rounded-2xl border border-blue-500/20 bg-blue-500/5 p-4">
      <div className="flex items-center gap-2 text-xs font-semibold uppercase tracking-wide text-blue-600 dark:text-blue-300">
        <Sparkles className="h-3.5 w-3.5" aria-hidden="true" />
        {brief.headline}
      </div>
      {metadata ? <div className="mt-2 text-xs text-text-secondary">{metadata}</div> : null}
      <p className="mt-2 text-sm leading-6 text-text-primary">{progressiveSummary}</p>
      <div className="mt-3 grid gap-3 sm:grid-cols-2">
        {sections.map((section) => (
          <div key={section.title}>
            <div className="text-xs font-semibold text-text-primary">{section.title}</div>
            <ul className="mt-1 list-disc space-y-1 pl-4 text-xs leading-5 text-text-secondary">
              {section.items.map((item) => (
                <li key={item}>{item}</li>
              ))}
            </ul>
          </div>
        ))}
      </div>
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
            Proposed {slots.suggestions.length} slot(s) for{' '}
            {(slots.schedulingAttendees || slots.attendees).length} internal attendee(s)
            {Array.isArray(slots.externalAttendeesExcluded) &&
            slots.externalAttendeesExcluded.length > 0
              ? ` (${slots.externalAttendeesExcluded.length} external attendee calendar${
                  slots.externalAttendeesExcluded.length === 1 ? '' : 's'
                } excluded)`
              : ''}
            . Pick one to schedule a Teams meeting from this thread. You will confirm before Outlook
            sends invites.
          </div>
          {Array.isArray(slots.availabilityNotes) && slots.availabilityNotes.length > 0 && (
            <div className="mt-2 space-y-1">
              {slots.availabilityNotes.map((note, index) => (
                <p
                  key={`${slots.messageId}-availability-note-${index}`}
                  className="text-xs text-amber-700 dark:text-amber-300"
                >
                  {note}
                </p>
              ))}
            </div>
          )}
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
                  label="Schedule Meeting"
                  loadingLabel="Scheduling..."
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
          {result.meetingNotePreview && (
            <div className="mt-3 rounded-lg border border-green-500/20 bg-green-500/10 p-2">
              <div className="text-[11px] font-semibold uppercase tracking-wide text-green-800 dark:text-green-200">
                Invite note
              </div>
              <p className="mt-1 whitespace-pre-wrap text-xs leading-5 text-text-secondary">
                {result.meetingNotePreview}
              </p>
            </div>
          )}
          {(result.meetingDraft?.webLink || result.event?.webLink) && (
            <a
              className="ml-3 mt-2 inline-block text-xs font-medium text-green-700 hover:underline dark:text-green-300"
              href={result.meetingDraft?.webLink || result.event?.webLink}
              target="_blank"
              rel="noreferrer"
            >
              Open calendar event
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
  const [workspaceTab, setWorkspaceTab] = useState<OutlookWorkspaceTab>('inbox');
  const [selectedId, setSelectedId] = useState<string | undefined>();
  const [selectedCalendarEventId, setSelectedCalendarEventId] = useState<string | undefined>();
  const [calendarEditorMode, setCalendarEditorMode] = useState<CalendarEditorMode>(null);
  const [calendarMutationError, setCalendarMutationError] = useState<string>('');
  const [calendarForm, setCalendarForm] = useState<CalendarEventFormState>(() =>
    buildCalendarFormState(),
  );
  const [inboxView, setInboxView] = useState<InboxView>('focused');
  const [calendarViewMode, setCalendarViewMode] = useState<CalendarViewMode>('week');
  const [calendarDate, setCalendarDate] = useState(() => toDateInputValue(new Date()));
  const [densityMode, setDensityMode] = useState<DensityMode>(loadDensityMode);
  const [actionSuccess, setActionSuccess] = useState<Record<string, boolean>>({});
  const [mailboxControlsOpen, setMailboxControlsOpen] = useState(false);
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
  const [selectedSummaryResult, setSelectedSummaryResult] =
    useState<OutlookAnalyzeSelectionResponse | null>(null);
  const [dailyBriefResult, setDailyBriefResult] = useState<OutlookDailyBriefResponse | null>(null);
  const [selectedDeleteIds, setSelectedDeleteIds] = useState<string[]>([]);
  const [pendingDeleteBatches, setPendingDeleteBatches] = useState<PendingDeleteBatch[]>([]);
  const [optimisticallyHiddenIds, setOptimisticallyHiddenIds] = useState<string[]>([]);
  const [nowMs, setNowMs] = useState(Date.now());
  const [assistantPanelOpen, setAssistantPanelOpen] = useState(false);
  const [assistantPanelScrolled, setAssistantPanelScrolled] = useState(false);
  const [assistantPanelSize, setAssistantPanelSize] =
    useState<AssistantPanelSize>(loadAssistantPanelSize);
  const [assistantPanelResizing, setAssistantPanelResizing] = useState(false);
  const [draftInstructions, setDraftInstructions] = useState('');
  const [searchTerm, setSearchTerm] = useState('');
  const [statusMessage, setStatusMessage] = useState('');
  const successTimerRef = useRef<Record<string, number>>({});
  const deleteTimerRef = useRef<Record<string, number>>({});
  const pendingDeleteRef = useRef<PendingDeleteBatch[]>([]);
  const assistantResizeCleanupRef = useRef<(() => void) | null>(null);

  const { data: status, isLoading: statusLoading } = useOutlookStatusQuery();
  const mailboxEnabled = Boolean(status?.enabled && status?.connected);
  const deferredSearchTerm = useDeferredValue(searchTerm);
  const normalizedSearchTerm = deferredSearchTerm.trim();
  const calendarWindow = useMemo(
    () => buildCalendarWindow(calendarDate, calendarViewMode),
    [calendarDate, calendarViewMode],
  );
  const messageListParams = useMemo(
    () => ({
      folder: 'inbox' as const,
      inboxView,
      limit: 100,
      search: normalizedSearchTerm || undefined,
    }),
    [inboxView, normalizedSearchTerm],
  );
  const calendarQueryParams = useMemo(
    () => ({
      startDateTime: calendarWindow.start.toISOString(),
      endDateTime: calendarWindow.end.toISOString(),
      view: calendarViewMode,
      limit: calendarViewMode === 'week' ? 100 : 40,
    }),
    [calendarWindow.end, calendarWindow.start, calendarViewMode],
  );

  const {
    data: messageList,
    isLoading: messagesLoading,
    refetch,
  } = useOutlookMessagesQuery(messageListParams, {
    enabled: mailboxEnabled,
    keepPreviousData: true,
    refetchInterval: mailboxEnabled ? MAILBOX_REFRESH_INTERVAL_MS : false,
    refetchIntervalInBackground: false,
  });
  const {
    data: calendarData,
    isLoading: calendarLoading,
    refetch: refetchCalendar,
  } = useOutlookCalendarQuery(calendarQueryParams, {
    enabled: mailboxEnabled && workspaceTab === 'calendar',
    keepPreviousData: true,
    refetchInterval: mailboxEnabled && workspaceTab === 'calendar' ? MAILBOX_REFRESH_INTERVAL_MS : false,
    refetchIntervalInBackground: false,
  });

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
  const inboxViewLabel = inboxView.charAt(0).toUpperCase() + inboxView.slice(1);
  const analysis = selectedId ? analysisByMessage[selectedId] : null;
  const draftResult = selectedId ? draftResultByMessage[selectedId] : null;
  const meetingSlots = selectedId ? meetingSlotsByMessage[selectedId] : null;
  const meetingResult = selectedId ? meetingResultByMessage[selectedId] : null;
  const calendarEvents = useMemo(() => calendarData?.events ?? [], [calendarData?.events]);
  const selectedCalendarEvent = useMemo(
    () => calendarEvents.find((event) => event.id === selectedCalendarEventId) ?? calendarEvents[0],
    [calendarEvents, selectedCalendarEventId],
  );

  const { data: selectedMessage, isLoading: messageLoading } = useOutlookMessageQuery(selectedId, {
    enabled: mailboxEnabled && Boolean(selectedId),
  });
  const threadMessages = useMemo(
    () => (selectedMessage ? getThreadMessages(selectedMessage) : []),
    [selectedMessage],
  );
  const draftReplies = useMemo(
    () => (selectedMessage ? getDraftReplies(selectedMessage) : []),
    [selectedMessage],
  );

  const analyzeMutation = useAnalyzeOutlookMessageMutation();
  const analyzeSelectedMutation = useAnalyzeSelectedOutlookMessagesMutation();
  const createCalendarEventMutation = useCreateOutlookCalendarEventMutation();
  const draftMutation = useCreateOutlookDraftMutation();
  const updateCalendarEventMutation = useUpdateOutlookCalendarEventMutation();
  const deleteCalendarEventMutation = useDeleteOutlookCalendarEventMutation();
  const deleteMutation = useDeleteOutlookMessageMutation();
  const dailyBriefMutation = useOutlookDailyBriefMutation();
  const updateReadStateMutation = useUpdateOutlookMessageReadStateMutation();
  const meetingSlotsMutation = useProposeOutlookMeetingSlotsMutation();
  const createMeetingMutation = useCreateOutlookMeetingMutation();

  const updateCachedReadState = useCallback(
    (messageId: string, isRead: boolean) => {
      queryClient.setQueryData<OutlookMessagesResponse>(
        [QueryKeys.outlookMessages, messageListParams],
        (current) =>
          current == null
            ? current
            : {
                ...current,
                messages: current.messages.map((message) =>
                  message.id === messageId ? { ...message, isRead } : message,
                ),
              },
      );
      queryClient.setQueryData<OutlookMessage | undefined>(
        [QueryKeys.outlookMessage, messageId],
        (current) => (current == null ? current : { ...current, isRead }),
      );
    },
    [messageListParams, queryClient],
  );

  const handleSelectMessage = useCallback(
    (message: OutlookMessage) => {
      startTransition(() => setSelectedId(message.id));

      if (message.isRead) {
        return;
      }

      updateCachedReadState(message.id, true);
      void updateReadStateMutation
        .mutateAsync({ messageId: message.id, isRead: true })
        .catch(() => {
          updateCachedReadState(message.id, false);
          showToast({
            message: 'Unable to update Outlook read state.',
            severity: 'warning',
            duration: 4000,
          });
          void refetch();
        });
    },
    [refetch, showToast, updateCachedReadState, updateReadStateMutation],
  );

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
    setAssistantPanelScrolled(false);
    setAssistantPanelOpen(false);
    setMailboxControlsOpen(false);
  }, [workspaceTab, inboxView, selectedId]);

  useEffect(() => {
    if (workspaceTab !== 'calendar') {
      return;
    }
    if (calendarEvents.length === 0) {
      setSelectedCalendarEventId(undefined);
      return;
    }
    if (!selectedCalendarEventId || !calendarEvents.some((event) => event.id === selectedCalendarEventId)) {
      setSelectedCalendarEventId(calendarEvents[0].id);
    }
  }, [calendarEvents, selectedCalendarEventId, workspaceTab]);

  useEffect(() => {
    if (workspaceTab !== 'calendar') {
      setCalendarEditorMode(null);
      setCalendarMutationError('');
    }
  }, [workspaceTab]);

  useEffect(() => {
    window.localStorage.setItem(OUTLOOK_DENSITY_KEY, densityMode);
  }, [densityMode]);

  useEffect(() => {
    try {
      window.localStorage.setItem(
        OUTLOOK_ASSISTANT_PANEL_SIZE_KEY,
        JSON.stringify(assistantPanelSize),
      );
    } catch {
      // Best-effort persistence only.
    }
  }, [assistantPanelSize]);

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
      assistantResizeCleanupRef.current?.();
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

  const handleAssistantResizeStart = useCallback(
    (event: React.MouseEvent<HTMLButtonElement>) => {
      event.preventDefault();
      event.stopPropagation();

      assistantResizeCleanupRef.current?.();
      setAssistantPanelResizing(true);

      const startX = event.clientX;
      const startY = event.clientY;
      const startWidth = assistantPanelSize.width;
      const startHeight = assistantPanelSize.height;

      const maxWidth = Math.max(
        ASSISTANT_PANEL_MIN_WIDTH,
        Math.min(window.innerWidth - 48, 760),
      );
      const maxHeight = Math.max(
        ASSISTANT_PANEL_MIN_HEIGHT,
        Math.min(window.innerHeight - 48, 860),
      );

      const handleMove = (moveEvent: MouseEvent) => {
        const nextWidth = Math.max(
          ASSISTANT_PANEL_MIN_WIDTH,
          Math.min(startWidth - (moveEvent.clientX - startX), maxWidth),
        );
        const nextHeight = Math.max(
          ASSISTANT_PANEL_MIN_HEIGHT,
          Math.min(startHeight - (moveEvent.clientY - startY), maxHeight),
        );

        setAssistantPanelSize({
          width: nextWidth,
          height: nextHeight,
        });
      };

      const cleanup = () => {
        window.removeEventListener('mousemove', handleMove);
        window.removeEventListener('mouseup', handleUp);
        document.body.style.userSelect = '';
        document.body.style.cursor = '';
        assistantResizeCleanupRef.current = null;
        setAssistantPanelResizing(false);
      };

      const handleUp = () => {
        cleanup();
      };

      assistantResizeCleanupRef.current = cleanup;
      document.body.style.userSelect = 'none';
      document.body.style.cursor = 'nwse-resize';
      window.addEventListener('mousemove', handleMove);
      window.addEventListener('mouseup', handleUp);
    },
    [assistantPanelSize.height, assistantPanelSize.width],
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
      'Schedule this Teams meeting now? This will send an Outlook invite to the selected attendees.',
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
        sendInvites: true,
      },
    });
    setMeetingResultByMessage((current) => ({ ...current, [selectedId]: result }));
    markActionSuccess('meetingCreate');
    showToast({ message: 'Meeting scheduled and invites sent.', severity: 'success' });
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

  const handleAnalyzeSelected = async () => {
    if (selectedDeleteIds.length === 0) {
      return;
    }

    const result = await analyzeSelectedMutation.mutateAsync({
      messageIds: selectedDeleteIds,
    });
    setAssistantPanelScrolled(false);
    setAssistantPanelOpen(true);
    setSelectedSummaryResult(result);
    markActionSuccess('analyzeSelected');
    showToast({
      message: `Generated summary for ${result.messageCount} selected email(s).`,
      severity: 'success',
    });
  };

  const handleDailyBrief = async () => {
    const result = await dailyBriefMutation.mutateAsync();
    setAssistantPanelScrolled(false);
    setAssistantPanelOpen(true);
    setDailyBriefResult(result);
    markActionSuccess('dailyBrief');
    showToast({
      message: 'Daily brief generated.',
      severity: 'success',
    });
  };

  const updateCalendarFormField = useCallback(
    (field: keyof CalendarEventFormState, value: string | boolean) => {
      setCalendarForm((current) => ({ ...current, [field]: value }));
    },
    [],
  );

  const beginCalendarCreate = useCallback(() => {
    setCalendarMutationError('');
    setCalendarEditorMode('create');
    setCalendarForm(buildCalendarFormState());
  }, []);

  const beginCalendarEdit = useCallback(() => {
    if (!selectedCalendarEvent) {
      return;
    }
    setCalendarMutationError('');
    setCalendarEditorMode('edit');
    setCalendarForm(buildCalendarFormState(selectedCalendarEvent));
  }, [selectedCalendarEvent]);

  const cancelCalendarEdit = useCallback(() => {
    setCalendarMutationError('');
    setCalendarEditorMode(null);
    setCalendarForm(buildCalendarFormState(selectedCalendarEvent));
  }, [selectedCalendarEvent]);

  const handleSubmitCalendarEvent = async () => {
    try {
      setCalendarMutationError('');
      const payload = buildCalendarMutationPayload(calendarForm);
      const startTime = new Date(payload.start.dateTime).getTime();
      const endTime = new Date(payload.end.dateTime).getTime();
      if (!payload.subject.trim()) {
        setCalendarMutationError('Subject is required.');
        return;
      }
      if (Number.isNaN(startTime) || Number.isNaN(endTime) || endTime <= startTime) {
        setCalendarMutationError('End time must be after start time.');
        return;
      }

      if (calendarEditorMode === 'create') {
        const result = await createCalendarEventMutation.mutateAsync(payload);
        await refetchCalendar();
        setSelectedCalendarEventId(result.event.id);
        setCalendarEditorMode(null);
        setCalendarForm(buildCalendarFormState(result.event));
        markActionSuccess('calendarSave');
        showToast({ message: 'Calendar event created.', severity: 'success' });
        return;
      }

      if (calendarEditorMode === 'edit' && selectedCalendarEventId) {
        const result = await updateCalendarEventMutation.mutateAsync({
          eventId: selectedCalendarEventId,
          payload,
        });
        await refetchCalendar();
        setSelectedCalendarEventId(result.event.id);
        setCalendarEditorMode(null);
        setCalendarForm(buildCalendarFormState(result.event));
        markActionSuccess('calendarSave');
        showToast({ message: 'Calendar event updated.', severity: 'success' });
      }
    } catch (error) {
      const nextMessage =
        error instanceof Error ? error.message : 'Unable to save the calendar event.';
      setCalendarMutationError(nextMessage);
    }
  };

  const handleDeleteCalendarEvent = async () => {
    if (!selectedCalendarEventId) {
      return;
    }
    const confirmed = window.confirm(
      'Remove this calendar event? This will update Outlook immediately.',
    );
    if (!confirmed) {
      return;
    }

    try {
      setCalendarMutationError('');
      await deleteCalendarEventMutation.mutateAsync(selectedCalendarEventId);
      await refetchCalendar();
      setSelectedCalendarEventId(undefined);
      setCalendarEditorMode(null);
      markActionSuccess('calendarDelete');
      showToast({ message: 'Calendar event removed.', severity: 'success' });
    } catch (error) {
      const nextMessage =
        error instanceof Error ? error.message : 'Unable to remove the calendar event.';
      setCalendarMutationError(nextMessage);
    }
  };

  const handleRefresh = async () => {
    if (workspaceTab === 'calendar') {
      await refetchCalendar();
      markActionSuccess('refresh');
      return;
    }

    await refetch();
    markActionSuccess('refresh');
  };

  const handleCalendarToday = () => {
    setCalendarDate(toDateInputValue(new Date()));
  };

  const calendarSummary = calendarData
    ? `${calendarData.events.length} event${calendarData.events.length === 1 ? '' : 's'} • ${formatWorkingHours(calendarData.workingHours)}`
    : 'Calendar range loading';

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
        <div className="flex items-center justify-between gap-3">
          <div className="min-w-0">
            <h2 className="text-base font-semibold">Outlook workspace</h2>
            <p className="truncate text-xs text-text-secondary">
              {workspaceTab === 'inbox'
                ? 'Inbox operations with on-demand AI actions.'
                : 'Calendar visibility first. Editing remains in Outlook for now.'}
            </p>
          </div>
          <ActionButton
            label="Refresh"
            loadingLabel="Refreshing..."
            successLabel="Updated"
            className="border border-border-light hover:bg-surface-hover"
            icon={RefreshCw}
            onClick={handleRefresh}
            isLoading={workspaceTab === 'calendar' ? calendarLoading : messagesLoading}
            isSuccess={actionSuccess.refresh}
          />
        </div>
        <div className="mt-3 flex flex-wrap items-center gap-3">
          <WorkspaceTabs active={workspaceTab} onChange={setWorkspaceTab} />
          {workspaceTab === 'calendar' && (
            <div className="flex flex-wrap items-center gap-2">
              <input
                type="date"
                className="h-9 rounded-xl border border-border-light bg-surface-secondary px-3 text-sm outline-none focus:border-blue-500"
                value={calendarDate}
                onChange={(event) => setCalendarDate(event.target.value)}
                aria-label="Select calendar date"
              />
              <button
                type="button"
                className="rounded-lg border border-border-light px-2.5 py-1.5 text-[11px] font-semibold hover:bg-surface-hover"
                onClick={handleCalendarToday}
              >
                Today
              </button>
              <CalendarModeTabs active={calendarViewMode} onChange={setCalendarViewMode} />
            </div>
          )}
        </div>

        {workspaceTab === 'inbox' ? (
          <>
            <div className="mt-2">
              <div className="relative mb-2">
                <Search
                  className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-text-secondary"
                  aria-hidden="true"
                />
                <input
                  type="search"
                  className="h-10 w-full rounded-xl border border-border-light bg-surface-secondary py-2 pl-9 pr-10 text-sm outline-none transition-colors placeholder:text-text-secondary focus:border-blue-500 focus:bg-surface-primary"
                  placeholder="Search inbox by sender, subject, or message text"
                  value={searchTerm}
                  onChange={(event) => setSearchTerm(event.target.value)}
                  aria-label="Search Outlook inbox"
                />
                {searchTerm.trim().length > 0 && (
                  <button
                    type="button"
                    className="absolute right-2 top-1/2 inline-flex h-7 w-7 -translate-y-1/2 items-center justify-center rounded-lg text-text-secondary transition-colors hover:bg-surface-hover hover:text-text-primary"
                    onClick={() => setSearchTerm('')}
                    aria-label="Clear inbox search"
                  >
                    <X className="h-4 w-4" aria-hidden="true" />
                  </button>
                )}
              </div>
              <button
                type="button"
                className="inline-flex items-center gap-2 rounded-lg border border-border-light px-2.5 py-1.5 text-[11px] font-semibold transition-colors hover:bg-surface-hover"
                onClick={() => setMailboxControlsOpen((current) => !current)}
              >
                <span>Mailbox controls</span>
                <span className="rounded bg-surface-secondary px-1.5 py-0.5 text-[10px]">
                  {inboxViewLabel}
                </span>
                <ChevronDown
                  className={cn(
                    'h-3.5 w-3.5 transition-transform duration-150',
                    mailboxControlsOpen ? 'rotate-180' : 'rotate-0',
                  )}
                  aria-hidden="true"
                />
              </button>
            </div>
            <div
              className={cn(
                'overflow-hidden transition-[max-height,opacity] duration-200',
                mailboxControlsOpen ? 'mt-2 max-h-80 opacity-100' : 'max-h-0 opacity-0',
              )}
            >
              <div className="space-y-2 rounded-xl border border-border-light bg-surface-secondary p-2">
                <ViewTabs active={inboxView} onChange={setInboxView} />
                <div className="flex flex-wrap items-center gap-2">
                  <div className="inline-flex rounded-lg border border-border-light bg-surface-primary p-0.5">
                    <button
                      type="button"
                      className={cn(
                        'rounded-md px-2 py-1 text-[11px] font-semibold transition-colors',
                        densityMode === 'comfortable'
                          ? 'bg-surface-primary-alt text-text-primary shadow-sm'
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
                          ? 'bg-surface-primary-alt text-text-primary shadow-sm'
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
                </div>
                <div className="flex items-center gap-1.5 text-[11px] text-text-secondary">
                  <CalendarDays className="h-3.5 w-3.5" aria-hidden="true" />
                  {status.calendarContextEnabled
                    ? 'Calendar context is enabled for scheduling analysis.'
                    : 'Calendar context is disabled.'}
                </div>
              </div>
            </div>
            <div className="mt-3 flex flex-wrap items-center gap-2">
              <ActionButton
                label={`Delete selected emails (${selectedDeleteIds.length})`}
                loadingLabel="Deleting selected..."
                successLabel="Queued"
                className="border border-red-500/30 text-red-600 hover:bg-red-500/10 dark:text-red-300"
                onClick={handleBulkDelete}
                icon={Trash2}
                isSuccess={actionSuccess.delete}
                disabled={selectedDeleteIds.length === 0}
              />
              <ActionButton
                label={`Analyze selected emails (${selectedDeleteIds.length})`}
                loadingLabel="Generating summary..."
                successLabel="Summary ready"
                className="bg-blue-600 text-white hover:bg-blue-700"
                onClick={handleAnalyzeSelected}
                icon={Sparkles}
                isLoading={analyzeSelectedMutation.isLoading}
                isSuccess={actionSuccess.analyzeSelected}
                disabled={selectedDeleteIds.length === 0}
              />
              <ActionButton
                label="Generate daily brief"
                loadingLabel="Building brief..."
                successLabel="Brief ready"
                className="border border-amber-500/30 text-amber-700 hover:bg-amber-500/10 dark:text-amber-300"
                onClick={handleDailyBrief}
                icon={CalendarDays}
                isLoading={dailyBriefMutation.isLoading}
                isSuccess={actionSuccess.dailyBrief}
              />
            </div>
          </>
        ) : (
          <div className="mt-3 rounded-xl border border-border-light bg-surface-secondary px-3 py-2 text-[11px] text-text-secondary">
            {calendarSummary}
          </div>
        )}
      </div>

      {workspaceTab === 'calendar' ? (
        <CalendarWorkspace
          calendarData={calendarData}
          isLoading={calendarLoading}
          viewMode={calendarViewMode}
          selectedEventId={selectedCalendarEventId}
          onSelectEvent={setSelectedCalendarEventId}
          editorMode={calendarEditorMode}
          form={calendarForm}
          onFormChange={updateCalendarFormField}
          onStartCreate={beginCalendarCreate}
          onStartEdit={beginCalendarEdit}
          onCancelEdit={cancelCalendarEdit}
          onSubmit={handleSubmitCalendarEvent}
          onDelete={handleDeleteCalendarEvent}
          isSubmitting={
            createCalendarEventMutation.isLoading || updateCalendarEventMutation.isLoading
          }
          isDeleting={deleteCalendarEventMutation.isLoading}
          mutationError={calendarMutationError}
        />
      ) : (
        <div className="grid min-h-0 flex-1 grid-cols-1 md:grid-cols-[minmax(240px,34%)_minmax(0,1fr)]">
        <div className="min-h-0 overflow-y-auto border-b border-border-light md:border-b-0 md:border-r">
          {messagesLoading && <MessageListSkeleton density={densityMode} />}
          {!messagesLoading && visibleConversations.length === 0 && (
            <EmptyState
              title={normalizedSearchTerm ? 'No matching emails found' : 'No messages found'}
              description={
                normalizedSearchTerm
                  ? 'Try a different sender, subject, or phrase.'
                  : `Your ${inboxView === 'all' ? 'inbox' : inboxView} query returned no mail.`
              }
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
                  message.isRead && selectedId !== message.id && 'opacity-80',
                  !message.isRead && selectedId !== message.id && 'bg-[#f5d000]/[0.04]',
                  selectedId === message.id && 'bg-surface-active-alt',
                )}
              >
                <div className="pt-1">
                  <input
                    type="checkbox"
                    aria-label={`Select ${message.subject}`}
                    className="h-4 w-4 rounded border-border-light accent-[#f5d000] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#f5d000]/45"
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
                  onClick={() => handleSelectMessage(message)}
                >
                  <div className="flex items-start justify-between gap-2">
                    <div className="min-w-0">
                      <div className="flex min-w-0 items-center gap-1.5">
                        {!message.isRead && (
                          <span
                            className="mt-0.5 h-2 w-2 shrink-0 rounded-full bg-[#f5d000]"
                            aria-label="Unread email"
                            title="Unread email"
                          />
                        )}
                        <div
                          className={cn(
                            'truncate font-semibold',
                            message.isRead ? 'text-text-primary/90' : 'text-text-primary',
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
                      <div
                        className={cn(
                          'truncate text-[11px] font-semibold',
                          message.isRead
                            ? 'text-text-primary/90'
                            : 'text-[#b88a00] dark:text-[#f5d000]',
                        )}
                      >
                        {formatSender(message)}
                      </div>
                    </div>
                    <div className="whitespace-nowrap text-[11px] text-text-secondary">
                      {formatDate(message.receivedDateTime)}
                    </div>
                  </div>
                  <p
                    className={cn(
                      'mt-0.5 line-clamp-1 text-xs leading-4',
                      message.isRead ? 'text-text-secondary/80' : 'text-text-secondary',
                    )}
                  >
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
                className="relative flex min-h-0 flex-1 flex-col overflow-hidden"
              >
                <div className="border-b border-border-light px-5 py-4">
                  <div className="flex flex-wrap items-start justify-between gap-3">
                    <div className="min-w-0 flex-1">
                      <h3 className="break-words text-lg font-semibold leading-6">
                        {selectedMessage.subject}
                      </h3>
                      <div className="mt-1 text-xs text-text-secondary">
                        From{' '}
                        <span className="font-semibold text-[#b88a00] dark:text-[#f5d000]">
                          {formatSender(selectedMessage)}
                        </span>
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
                    {threadMessages.map((threadMessage) => (
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
                            <div className="truncate text-sm font-semibold text-[#b88a00] dark:text-[#f5d000]">
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

                    {draftReplies.length > 0 && (
                      <section className="space-y-2 pt-1">
                        <div className="text-[11px] font-semibold uppercase tracking-wide text-text-secondary">
                          Draft replies (not sent)
                        </div>
                        {draftReplies.map((draftMessage) => {
                          const toLine = formatRecipients(draftMessage.toRecipients);
                          const ccLine = formatRecipients(draftMessage.ccRecipients);
                          return (
                            <article
                              key={draftMessage.id}
                              className="rounded-2xl border border-amber-500/30 bg-amber-500/5 p-4"
                            >
                              <div className="mb-2 flex flex-wrap items-center justify-between gap-2 border-b border-border-light pb-2">
                                <div className="min-w-0">
                                  <div className="truncate text-sm font-semibold">
                                    {draftMessage.subject || 'Draft reply'}
                                  </div>
                                  <div className="text-[11px] text-text-secondary">
                                    Updated{' '}
                                    {formatDate(
                                      draftMessage.lastModifiedDateTime ||
                                        draftMessage.createdDateTime ||
                                        draftMessage.sentDateTime ||
                                        draftMessage.receivedDateTime,
                                    )}
                                  </div>
                                </div>
                                <span className="rounded-full bg-amber-500/15 px-2 py-0.5 text-[10px] font-semibold uppercase tracking-wide text-amber-700 dark:text-amber-300">
                                  Draft
                                </span>
                              </div>
                              {toLine ? (
                                <div className="text-xs text-text-secondary">
                                  To: <span className="text-text-primary">{toLine}</span>
                                </div>
                              ) : null}
                              {ccLine ? (
                                <div className="mt-0.5 text-xs text-text-secondary">
                                  Cc: <span className="text-text-primary">{ccLine}</span>
                                </div>
                              ) : null}
                              <p className="mt-2 line-clamp-2 text-xs leading-5 text-text-secondary">
                                {draftMessage.bodyPreview || 'Draft content preview unavailable.'}
                              </p>
                              {draftMessage.webLink ? (
                                <a
                                  className="mt-2 inline-block text-xs font-medium text-blue-600 hover:underline dark:text-blue-300"
                                  href={draftMessage.webLink}
                                  target="_blank"
                                  rel="noreferrer"
                                >
                                  Open draft in Outlook
                                </a>
                              ) : null}
                            </article>
                          );
                        })}
                      </section>
                    )}
                  </div>
                </div>

                <div
                  className="pointer-events-none absolute bottom-4 right-4 z-10 flex max-w-[calc(100%-2rem)] flex-col items-end"
                  style={{ width: `${assistantPanelSize.width}px` }}
                >
                  {!assistantPanelOpen && (
                    <button
                      type="button"
                      className="pointer-events-auto inline-flex items-center gap-2 rounded-full border border-[#f5d000]/70 bg-[#f5d000] px-4 py-2 text-xs font-semibold text-black shadow-lg shadow-[#f5d000]/20 transition-colors hover:bg-[#ffe05c]"
                      onClick={() => {
                        setAssistantPanelScrolled(false);
                        setAssistantPanelOpen(true);
                      }}
                    >
                      <MessageSquareText className="h-3.5 w-3.5" aria-hidden="true" />
                      AI assistant
                    </button>
                  )}

                  {assistantPanelOpen && (
                    <div
                      className={cn(
                        'pointer-events-auto relative w-full overflow-hidden rounded-2xl border border-border-light bg-surface-primary shadow-2xl',
                        assistantPanelResizing && 'select-none',
                      )}
                      style={{ height: `${assistantPanelSize.height}px` }}
                    >
                      <button
                        type="button"
                        className="absolute left-0 top-0 z-[3] h-5 w-5 cursor-nwse-resize rounded-br-lg border-b border-r border-border-light bg-surface-secondary/90 text-text-secondary hover:bg-surface-hover"
                        onMouseDown={handleAssistantResizeStart}
                        aria-label="Resize AI assistant panel"
                        title="Drag to resize"
                      >
                        <span className="sr-only">Resize AI assistant panel</span>
                        <span className="pointer-events-none absolute left-1 top-1 h-2.5 w-2.5 border-l border-t border-current opacity-70" />
                      </button>
                      <div className="flex h-full flex-col overflow-hidden">
                        <div
                          className={cn(
                            'sticky top-0 z-[2] border-b border-border-light bg-surface-primary px-4 py-3',
                            assistantPanelScrolled && 'shadow-sm',
                          )}
                        >
                          <div className="flex items-center justify-between">
                            <div className="text-xs font-semibold uppercase tracking-wide text-[#b88a00] dark:text-[#f5d000]">
                              AI assistant
                            </div>
                            <button
                              type="button"
                              className="rounded-md border border-border-light p-1.5 transition-colors hover:bg-surface-hover"
                              onClick={() => {
                                setAssistantPanelScrolled(false);
                                setAssistantPanelOpen(false);
                              }}
                              aria-label="Minimize AI assistant panel"
                            >
                              <Minimize2 className="h-3.5 w-3.5" aria-hidden="true" />
                            </button>
                          </div>

                          <div className="mt-3 flex flex-wrap items-center gap-2">
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
                        </div>

                        <div
                          className="min-h-0 flex-1 overflow-y-auto space-y-3 px-4 pb-4 pt-3"
                          onScroll={(event) =>
                            setAssistantPanelScrolled(event.currentTarget.scrollTop > 4)
                          }
                        >
                          {analyzeMutation.error != null && (
                            <div className="mt-2 flex items-center justify-between rounded-lg border border-red-500/20 bg-red-500/5 px-3 py-2 text-xs text-red-600">
                              <span>Unable to analyze this email.</span>
                              <button type="button" className="underline" onClick={handleAnalyze}>
                                Retry
                              </button>
                            </div>
                          )}
                          {analyzeSelectedMutation.error != null && (
                            <div className="mt-2 flex items-center justify-between rounded-lg border border-red-500/20 bg-red-500/5 px-3 py-2 text-xs text-red-600">
                              <span>Unable to summarize the selected emails.</span>
                              <button
                                type="button"
                                className="underline"
                                onClick={handleAnalyzeSelected}
                              >
                                Retry
                              </button>
                            </div>
                          )}
                          {dailyBriefMutation.error != null && (
                            <div className="mt-2 flex items-center justify-between rounded-lg border border-red-500/20 bg-red-500/5 px-3 py-2 text-xs text-red-600">
                              <span>Unable to generate the daily brief.</span>
                              <button
                                type="button"
                                className="underline"
                                onClick={handleDailyBrief}
                              >
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
                            <p className="mt-2 text-xs text-red-500">
                              Unable to delete this email.
                            </p>
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
                            <p className="mt-2 text-xs text-red-500">
                              Unable to schedule Teams meeting.
                            </p>
                          )}
                          {!status.meetingSchedulingEnabled && (
                            <p className="mt-2 text-xs text-text-secondary">
                              Meeting scheduling is disabled. Set
                              OUTLOOK_AI_ENABLE_MEETING_SCHEDULING=true.
                            </p>
                          )}

                          {(analyzeSelectedMutation.isLoading ||
                            selectedSummaryResult ||
                            analyzeSelectedMutation.error) && (
                            <CollapsiblePanel title="Selected email summary" defaultOpen>
                              {analyzeSelectedMutation.isLoading && !selectedSummaryResult && (
                                <div className="h-24 animate-pulse rounded-xl bg-surface-secondary" />
                              )}
                              <BriefCard
                                brief={selectedSummaryResult?.brief}
                                metadata={
                                  selectedSummaryResult
                                    ? `${selectedSummaryResult.messageCount} selected email(s)`
                                    : null
                                }
                              />
                            </CollapsiblePanel>
                          )}

                          {(dailyBriefMutation.isLoading ||
                            dailyBriefResult ||
                            dailyBriefMutation.error) && (
                            <CollapsiblePanel title="Daily brief" defaultOpen>
                              {dailyBriefMutation.isLoading && !dailyBriefResult && (
                                <div className="h-24 animate-pulse rounded-xl bg-surface-secondary" />
                              )}
                              <BriefCard
                                brief={dailyBriefResult?.brief}
                                metadata={
                                  dailyBriefResult
                                    ? `${dailyBriefResult.emailCount} email(s) • ${dailyBriefResult.meetingCount} meeting(s) • last 24 hours`
                                    : null
                                }
                              />
                            </CollapsiblePanel>
                          )}

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
                  )}
                </div>
              </motion.div>
            )}
          </AnimatePresence>
        </div>
      </div>
      )}
      {pendingDeleteBatches.length > 0 && (
        <div className="pointer-events-none absolute bottom-4 left-4 z-20 flex w-[340px] max-w-[calc(100%-2rem)] flex-col gap-2">
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
