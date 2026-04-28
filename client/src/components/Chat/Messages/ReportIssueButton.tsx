import React, { useMemo, useState } from 'react';
import type { TConversation, TMessage } from 'librechat-data-provider';
import { Button, OGDialog, OGDialogContent, OGDialogTitle } from '@librechat/client';
import { AlertTriangle } from 'lucide-react';
import { useCreateIssueReportMutation } from '~/data-provider';

type Props = {
  conversation: TConversation | null;
  message: TMessage;
  isLast?: boolean;
  className?: string;
  alwaysVisible?: boolean;
};

const CATEGORY_OPTIONS = [
  { value: 'bad_response', label: 'Bad response' },
  { value: 'faulty_mcp_tool', label: 'Faulty MCP tool' },
  { value: 'bad_file_transformation', label: 'Bad file transformation' },
  { value: 'timeout_or_error', label: 'Timeout or error' },
  { value: 'auth_or_permissions', label: 'Auth or permissions' },
  { value: 'other', label: 'Other' },
] as const;

function extractMessagePreview(message: TMessage) {
  if (typeof message.text === 'string' && message.text.trim()) {
    return message.text.slice(0, 280);
  }

  if (typeof message.content === 'string') {
    return message.content.slice(0, 280);
  }

  if (Array.isArray(message.content)) {
    const text = message.content
      .map((part) => {
        if (part == null) {
          return '';
        }
        if (typeof part === 'string') {
          return part;
        }
        if ('text' in part && typeof part.text === 'string') {
          return part.text;
        }
        return '';
      })
      .join(' ');

    return text.slice(0, 280);
  }

  return undefined;
}

function extractFileIds(message: TMessage) {
  if (!Array.isArray(message.files)) {
    return undefined;
  }

  const fileIds = message.files
    .map((file) => {
      if (file && typeof file === 'object') {
        const candidate =
          (file as { file_id?: string; id?: string }).file_id ?? (file as { id?: string }).id;
        return typeof candidate === 'string' ? candidate : undefined;
      }
      return undefined;
    })
    .filter((value): value is string => Boolean(value));

  return fileIds.length ? fileIds : undefined;
}

export default function ReportIssueButton({
  conversation,
  message,
  isLast = false,
  className = '',
  alwaysVisible = false,
}: Props) {
  const [open, setOpen] = useState(false);
  const [category, setCategory] = useState<(typeof CATEGORY_OPTIONS)[number]['value']>(
    'bad_response',
  );
  const [description, setDescription] = useState('');
  const [submitted, setSubmitted] = useState(false);
  const mutation = useCreateIssueReportMutation();

  const messagePreview = useMemo(() => extractMessagePreview(message), [message]);
  const fileIds = useMemo(() => extractFileIds(message), [message]);

  const handleSubmit = async () => {
    if (!conversation?.conversationId || !message?.messageId) {
      return;
    }

    await mutation.mutateAsync({
      conversationId: conversation.conversationId,
      messageId: message.messageId,
      category,
      description: description.trim() || undefined,
      model: message.model ?? undefined,
      endpoint: message.endpoint ?? conversation.endpointType ?? conversation.endpoint ?? undefined,
      messagePreview,
      error: Boolean(message.error),
      fileIds,
    });

    setSubmitted(true);
    setOpen(false);
  };

  return (
    <>
      <button
        type="button"
        title="Report issue"
        aria-label="Report issue"
        onClick={() => setOpen(true)}
        className={`hover-button rounded-lg p-1.5 text-text-secondary-alt hover:bg-surface-hover hover:text-text-primary focus-visible:ring-2 focus-visible:ring-black focus-visible:outline-none dark:focus-visible:ring-white md:group-hover:visible md:group-focus-within:visible md:group-[.final-completion]:visible ${!isLast && !alwaysVisible ? 'md:opacity-0 md:group-hover:opacity-100 md:group-focus-within:opacity-100' : ''} ${submitted ? 'bg-surface-hover text-text-primary' : ''} ${className}`.trim()}
      >
        <AlertTriangle size={19} />
      </button>
      <OGDialog open={open} onOpenChange={setOpen}>
        <OGDialogContent className="w-11/12 max-w-lg">
          <OGDialogTitle className="text-lg font-semibold text-text-primary">
            Report issue
          </OGDialogTitle>
          <div className="mt-3 flex flex-col gap-3 text-sm text-text-primary">
            <label className="flex flex-col gap-1">
              <span className="text-xs uppercase tracking-[0.14em] text-text-secondary">
                Category
              </span>
              <select
                value={category}
                onChange={(event) =>
                  setCategory(event.target.value as (typeof CATEGORY_OPTIONS)[number]['value'])
                }
                className="rounded-xl border border-border-medium bg-surface-primary px-3 py-2 text-text-primary [color-scheme:light] dark:[color-scheme:dark]"
              >
                {CATEGORY_OPTIONS.map((option) => (
                  <option
                    key={option.value}
                    value={option.value}
                    className="bg-surface-primary text-text-primary"
                  >
                    {option.label}
                  </option>
                ))}
              </select>
            </label>
            <label className="flex flex-col gap-1">
              <span className="text-xs uppercase tracking-[0.14em] text-text-secondary">
                What went wrong?
              </span>
              <textarea
                value={description}
                onChange={(event) => setDescription(event.target.value)}
                rows={4}
                maxLength={1000}
                placeholder="Add a short note for admins. Include the tool name or file issue if helpful."
                className="w-full rounded-xl border border-border-medium bg-transparent p-3 text-text-primary"
              />
            </label>
            {messagePreview ? (
              <div className="rounded-xl border border-border-light bg-surface-secondary p-3 text-xs text-text-secondary">
                <div className="mb-1 font-semibold text-text-primary">Message preview</div>
                <div>{messagePreview}</div>
              </div>
            ) : null}
          </div>
          <div className="mt-4 flex justify-end gap-2">
            <Button variant="outline" onClick={() => setOpen(false)}>
              Cancel
            </Button>
            <Button
              variant="submit"
              onClick={handleSubmit}
              disabled={mutation.isLoading || (!description.trim() && category === 'other')}
            >
              {mutation.isLoading ? 'Submitting...' : 'Submit report'}
            </Button>
          </div>
        </OGDialogContent>
      </OGDialog>
    </>
  );
}
