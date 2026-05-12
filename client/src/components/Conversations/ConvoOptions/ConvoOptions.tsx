import { useState, useId, useRef, memo, useCallback, useMemo } from 'react';
import * as Ariakit from '@ariakit/react';
import { useParams, useNavigate } from 'react-router-dom';
import { QueryKeys } from 'librechat-data-provider';
import { BookmarkFilledIcon, BookmarkIcon } from '@radix-ui/react-icons';
import { useQueryClient } from '@tanstack/react-query';
import { DropdownPopup, Spinner, useToastContext } from '@librechat/client';
import { Ellipsis, Share2, CopyPlus, Archive, Pen, Trash, BookmarkPlus } from 'lucide-react';
import type { MouseEvent } from 'react';
import type { TMessage } from 'librechat-data-provider';
import {
  useDuplicateConversationMutation,
  useDeleteConversationMutation,
  useGetStartupConfig,
  useArchiveConvoMutation,
  useConversationTagsQuery,
  useTagConversationMutation,
} from '~/data-provider';
import { BookmarkEditDialog } from '~/components/Bookmarks';
import { useLocalize, useNavigateToConvo, useNewConvo, useBookmarkSuccess } from '~/hooks';
import { NotificationSeverity } from '~/common';
import { BookmarkContext, useChatContext } from '~/Providers';
import DeleteButton from './DeleteButton';
import ShareButton from './ShareButton';
import { cn } from '~/utils';

function ConvoOptions({
  conversationId,
  title,
  tags,
  retainView,
  renameHandler,
  isPopoverActive,
  setIsPopoverActive,
  isActiveConvo,
  isShiftHeld = false,
}: {
  conversationId: string | null;
  title: string | null;
  tags?: string[];
  retainView: () => void;
  renameHandler: (e: MouseEvent) => void;
  isPopoverActive: boolean;
  setIsPopoverActive: (open: boolean) => void;
  isActiveConvo: boolean;
  isShiftHeld?: boolean;
}) {
  const localize = useLocalize();
  const queryClient = useQueryClient();
  const { index } = useChatContext();
  const { data: startupConfig } = useGetStartupConfig();
  const { navigateToConvo } = useNavigateToConvo(index);
  const { showToast } = useToastContext();

  const navigate = useNavigate();
  const { conversationId: currentConvoId } = useParams();
  const { newConversation } = useNewConvo();

  const menuId = useId();
  const shareButtonRef = useRef<HTMLButtonElement>(null);
  const deleteButtonRef = useRef<HTMLButtonElement>(null);
  const newBookmarkRef = useRef<HTMLButtonElement>(null);
  const [showShareDialog, setShowShareDialog] = useState(false);
  const [showDeleteDialog, setShowDeleteDialog] = useState(false);
  const [showBookmarkDialog, setShowBookmarkDialog] = useState(false);
  const [announcement, setAnnouncement] = useState('');
  const updateConvoTags = useBookmarkSuccess(conversationId ?? '');
  const bookmarkTags = tags ?? [];

  const archiveConvoMutation = useArchiveConvoMutation();
  const tagMutation = useTagConversationMutation(conversationId ?? '', {
    onSuccess: (newTags) => {
      updateConvoTags(newTags);
    },
    onError: () => {
      showToast({
        message: localize('com_ui_bookmarks_update_error'),
        severity: NotificationSeverity.ERROR,
        showIcon: true,
      });
    },
  });
  const { data: availableBookmarks } = useConversationTagsQuery();

  const deleteMutation = useDeleteConversationMutation({
    onSuccess: () => {
      if (currentConvoId === conversationId || currentConvoId === 'new') {
        newConversation();
        navigate('/c/new', { replace: true });
      }
      retainView();
      showToast({
        message: localize('com_ui_convo_delete_success'),
        severity: NotificationSeverity.SUCCESS,
        showIcon: true,
      });
    },
    onError: () => {
      showToast({
        message: localize('com_ui_convo_delete_error'),
        severity: NotificationSeverity.ERROR,
        showIcon: true,
      });
    },
  });

  const duplicateConversation = useDuplicateConversationMutation({
    onSuccess: (data) => {
      navigateToConvo(data.conversation);
      showToast({
        message: localize('com_ui_duplication_success'),
        status: 'success',
      });
      setIsPopoverActive(false);
    },
    onMutate: () => {
      showToast({
        message: localize('com_ui_duplication_processing'),
        status: 'info',
      });
    },
    onError: () => {
      showToast({
        message: localize('com_ui_duplication_error'),
        status: 'error',
      });
    },
  });

  const isDuplicateLoading = duplicateConversation.isLoading;
  const isArchiveLoading = archiveConvoMutation.isLoading;
  const isDeleteLoading = deleteMutation.isLoading;

  const shareHandler = useCallback(() => {
    setShowShareDialog(true);
  }, []);

  const deleteHandler = useCallback(() => {
    setShowDeleteDialog(true);
  }, []);

  const handleInstantDelete = useCallback(
    (e: MouseEvent) => {
      e.stopPropagation();
      const convoId = conversationId ?? '';
      if (!convoId) {
        return;
      }
      const messages = queryClient.getQueryData<TMessage[]>([QueryKeys.messages, convoId]);
      const thread_id = messages?.[messages.length - 1]?.thread_id;
      const endpoint = messages?.[messages.length - 1]?.endpoint;
      deleteMutation.mutate({ conversationId: convoId, thread_id, endpoint, source: 'button' });
    },
    [conversationId, deleteMutation, queryClient],
  );

  const handleArchiveClick = useCallback(
    async (e?: MouseEvent) => {
      e?.stopPropagation();
      const convoId = conversationId ?? '';
      if (!convoId) {
        return;
      }

      archiveConvoMutation.mutate(
        { conversationId: convoId, isArchived: true },
        {
          onSuccess: () => {
            setAnnouncement(localize('com_ui_convo_archived'));
            setTimeout(() => {
              setAnnouncement('');
            }, 10000);
            if (currentConvoId === convoId || currentConvoId === 'new') {
              newConversation();
              navigate('/c/new', { replace: true });
            }
            retainView();
            setIsPopoverActive(false);
          },
          onError: () => {
            showToast({
              message: localize('com_ui_archive_error'),
              severity: NotificationSeverity.ERROR,
              showIcon: true,
            });
          },
        },
      );
    },
    [
      conversationId,
      currentConvoId,
      archiveConvoMutation,
      navigate,
      newConversation,
      retainView,
      setIsPopoverActive,
      showToast,
      localize,
    ],
  );

  const handleDuplicateClick = useCallback(() => {
    duplicateConversation.mutate({
      conversationId: conversationId ?? '',
    });
  }, [conversationId, duplicateConversation]);

  const handleBookmarkToggle = useCallback(
    (tag: string) => {
      const convoId = conversationId ?? '';
      if (!convoId || tag.trim().length === 0) {
        return;
      }

      const existingTags = (bookmarkTags ?? []).filter((value) =>
        (availableBookmarks ?? []).some((bookmark) => bookmark.tag === value),
      );
      const nextTags = existingTags.includes(tag)
        ? existingTags.filter((value) => value !== tag)
        : [...existingTags, tag];

      tagMutation.mutate({ tags: nextTags, tag });
    },
    [availableBookmarks, bookmarkTags, conversationId, tagMutation],
  );

  const dropdownItems = useMemo(
    () => {
      const bookmarkItems = (availableBookmarks ?? []).map((bookmark) => {
        const isSelected = bookmarkTags.includes(bookmark.tag);
        return {
          id: bookmark.tag,
          label: bookmark.tag,
          onClick: () => handleBookmarkToggle(bookmark.tag),
          hideOnClick: false,
          disabled: tagMutation.isLoading,
          ariaChecked: isSelected,
          icon: isSelected ? (
            <BookmarkFilledIcon className="size-4" />
          ) : (
            <BookmarkIcon className="size-4" />
          ),
        };
      });

      return [
        {
          label: localize('com_ui_share'),
          onClick: shareHandler,
          icon: <Share2 className="icon-sm mr-2 text-text-primary" aria-hidden="true" />,
          show: startupConfig && startupConfig.sharedLinksEnabled,
          ariaHasPopup: 'dialog' as const,
          ariaControls: 'share-conversation-dialog',
          /** NOTE: THE FOLLOWING PROPS ARE REQUIRED FOR MENU ITEMS THAT OPEN DIALOGS */
          hideOnClick: false,
          ref: shareButtonRef,
          render: (props) => <button {...props} />,
        },
        {
          label: localize('com_ui_bookmarks_new'),
          onClick: () => setShowBookmarkDialog(true),
          icon: <BookmarkPlus className="icon-sm mr-2 text-text-primary" aria-hidden="true" />,
          hideOnClick: false,
          ref: newBookmarkRef,
          render: (props) => <button {...props} />,
        },
        ...bookmarkItems,
        {
          label: localize('com_ui_rename'),
          onClick: renameHandler,
          icon: <Pen className="icon-sm mr-2 text-text-primary" aria-hidden="true" />,
        },
        {
          label: localize('com_ui_duplicate'),
          onClick: handleDuplicateClick,
          hideOnClick: false,
          icon: isDuplicateLoading ? (
            <Spinner className="size-4" />
          ) : (
            <CopyPlus className="icon-sm mr-2 text-text-primary" aria-hidden="true" />
          ),
        },
        {
          label: localize('com_ui_archive'),
          onClick: handleArchiveClick,
          hideOnClick: false,
          icon: isArchiveLoading ? (
            <Spinner className="size-4" />
          ) : (
            <Archive className="icon-sm mr-2 text-text-primary" aria-hidden="true" />
          ),
        },
        {
          label: localize('com_ui_delete'),
          onClick: deleteHandler,
          icon: <Trash className="icon-sm mr-2 text-text-primary" aria-hidden="true" />,
          ariaHasPopup: 'dialog' as const,
          ariaControls: 'delete-conversation-dialog',
          /** NOTE: THE FOLLOWING PROPS ARE REQUIRED FOR MENU ITEMS THAT OPEN DIALOGS */
          hideOnClick: false,
          ref: deleteButtonRef,
          render: (props) => <button {...props} />,
        },
      ];
    },
    [
      availableBookmarks,
      bookmarkTags,
      localize,
      shareHandler,
      startupConfig,
      renameHandler,
      deleteHandler,
      isArchiveLoading,
      isDuplicateLoading,
      handleArchiveClick,
      handleBookmarkToggle,
      handleDuplicateClick,
      tagMutation.isLoading,
    ],
  );

  const buttonClassName = cn(
    'inline-flex h-7 w-7 items-center justify-center rounded-md border-none p-0 text-sm font-medium ring-ring-primary transition-all duration-200 ease-in-out focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-offset-2 disabled:opacity-50',
    isActiveConvo === true || isPopoverActive
      ? 'opacity-100'
      : 'opacity-0 focus:opacity-100 group-focus-within:opacity-100 group-hover:opacity-100 data-[open]:opacity-100',
  );

  if (isShiftHeld && isActiveConvo && !isPopoverActive && !showShareDialog && !showDeleteDialog) {
    return (
      <div className="flex items-center gap-0.5">
        <button
          aria-label={localize('com_ui_archive')}
          className={cn(buttonClassName, 'hover:bg-surface-hover')}
          onClick={handleArchiveClick}
          disabled={isArchiveLoading}
        >
          {isArchiveLoading ? (
            <Spinner className="size-4" />
          ) : (
            <Archive className="icon-md text-text-secondary" aria-hidden={true} />
          )}
        </button>
        <button
          aria-label={localize('com_ui_delete')}
          className={cn(buttonClassName, 'hover:bg-surface-hover')}
          onClick={handleInstantDelete}
          disabled={isDeleteLoading}
        >
          {isDeleteLoading ? (
            <Spinner className="size-4" />
          ) : (
            <Trash className="icon-md text-text-secondary" aria-hidden={true} />
          )}
        </button>
      </div>
    );
  }

  return (
    <>
      <span className="sr-only" aria-live="polite" aria-atomic="true">
        {announcement}
      </span>
      <BookmarkContext.Provider value={{ bookmarks: availableBookmarks || [] }}>
        <DropdownPopup
          portal={true}
          menuId={menuId}
          focusLoop={true}
          className="z-[125]"
          unmountOnHide={true}
          isOpen={isPopoverActive}
          setIsOpen={setIsPopoverActive}
          trigger={
            <Ariakit.MenuButton
              id={`conversation-menu-${conversationId}`}
              aria-label={localize('com_nav_convo_menu_options')}
              aria-expanded={isPopoverActive}
              className={cn(
                'inline-flex h-7 w-7 items-center justify-center gap-2 rounded-md border-none p-0 text-sm font-medium ring-ring-primary transition-all duration-200 ease-in-out focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-offset-2 disabled:opacity-50',
                isActiveConvo === true || isPopoverActive
                  ? 'opacity-100'
                  : 'opacity-0 focus:opacity-100 group-focus-within:opacity-100 group-hover:opacity-100 data-[open]:opacity-100',
              )}
              onClick={(e: MouseEvent<HTMLButtonElement>) => {
                e.stopPropagation();
              }}
              onKeyDown={(e: React.KeyboardEvent<HTMLButtonElement>) => {
                if (e.key === 'Enter' || e.key === ' ') {
                  e.stopPropagation();
                }
              }}
            >
              <Ellipsis className="icon-md text-text-secondary" aria-hidden={true} />
            </Ariakit.MenuButton>
          }
          items={dropdownItems}
        />
      </BookmarkContext.Provider>
      {showShareDialog && (
        <ShareButton
          conversationId={conversationId ?? ''}
          open={showShareDialog}
          onOpenChange={setShowShareDialog}
          triggerRef={shareButtonRef}
        />
      )}
      {showDeleteDialog && (
        <DeleteButton
          title={title ?? ''}
          retainView={retainView}
          triggerRef={deleteButtonRef}
          setMenuOpen={setIsPopoverActive}
          showDeleteDialog={showDeleteDialog}
          conversationId={conversationId ?? ''}
          setShowDeleteDialog={setShowDeleteDialog}
        />
      )}
      <BookmarkEditDialog
        open={showBookmarkDialog}
        setOpen={setShowBookmarkDialog}
        tags={bookmarkTags}
        setTags={updateConvoTags}
        triggerRef={newBookmarkRef}
        conversationId={conversationId ?? ''}
        context="ConvoOptions - BookmarkEditDialog"
      />
    </>
  );
}

export default memo(ConvoOptions, (prevProps, nextProps) => {
  return (
    prevProps.conversationId === nextProps.conversationId &&
    prevProps.title === nextProps.title &&
    prevProps.isPopoverActive === nextProps.isPopoverActive &&
    prevProps.isActiveConvo === nextProps.isActiveConvo &&
    prevProps.isShiftHeld === nextProps.isShiftHeld
  );
});
