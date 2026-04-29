import { useRecoilValue } from 'recoil';
import { QueryKeys, dataService } from 'librechat-data-provider';
import { useMutation, useQuery } from '@tanstack/react-query';
import type {
  QueryObserverResult,
  UseMutationResult,
  UseQueryOptions,
} from '@tanstack/react-query';
import type {
  OutlookAnalyzeResponse,
  OutlookAnalyzeSelectionRequest,
  OutlookAnalyzeSelectionResponse,
  OutlookCalendarParams,
  OutlookCalendarResponse,
  OutlookCalendarEventMutationRequest,
  OutlookCalendarEventMutationResponse,
  OutlookCreateMeetingRequest,
  OutlookCreateMeetingResponse,
  OutlookDailyBriefResponse,
  OutlookDeleteResponse,
  OutlookDraftRequest,
  OutlookDraftResponse,
  OutlookMessage,
  OutlookMeetingSlotsRequest,
  OutlookMeetingSlotsResponse,
  OutlookMessagesParams,
  OutlookMessagesResponse,
  OutlookStatusResponse,
  OutlookUpdateReadStateResponse,
} from 'librechat-data-provider';
import store from '~/store';

export const useOutlookStatusQuery = (
  config?: UseQueryOptions<OutlookStatusResponse>,
): QueryObserverResult<OutlookStatusResponse> => {
  const queriesEnabled = useRecoilValue<boolean>(store.queriesEnabled);

  return useQuery<OutlookStatusResponse>(
    [QueryKeys.outlookStatus],
    () => dataService.getOutlookStatus(),
    {
      refetchOnWindowFocus: false,
      ...config,
      enabled: (config?.enabled ?? true) === true && queriesEnabled,
    },
  );
};

export const useOutlookMessagesQuery = (
  params: OutlookMessagesParams = {},
  config?: UseQueryOptions<OutlookMessagesResponse>,
): QueryObserverResult<OutlookMessagesResponse> => {
  const queriesEnabled = useRecoilValue<boolean>(store.queriesEnabled);

  return useQuery<OutlookMessagesResponse>(
    [QueryKeys.outlookMessages, params],
    () => dataService.getOutlookMessages(params),
    {
      refetchOnWindowFocus: false,
      ...config,
      enabled: (config?.enabled ?? true) === true && queriesEnabled,
    },
  );
};

export const useOutlookCalendarQuery = (
  params: OutlookCalendarParams = {},
  config?: UseQueryOptions<OutlookCalendarResponse>,
): QueryObserverResult<OutlookCalendarResponse> => {
  const queriesEnabled = useRecoilValue<boolean>(store.queriesEnabled);

  return useQuery<OutlookCalendarResponse>(
    [QueryKeys.outlookCalendar, params],
    () => dataService.getOutlookCalendar(params),
    {
      refetchOnWindowFocus: false,
      ...config,
      enabled: (config?.enabled ?? true) === true && queriesEnabled,
    },
  );
};

export const useCreateOutlookCalendarEventMutation = (): UseMutationResult<
  OutlookCalendarEventMutationResponse,
  unknown,
  OutlookCalendarEventMutationRequest
> => {
  return useMutation((payload: OutlookCalendarEventMutationRequest) =>
    dataService.createOutlookCalendarEvent(payload),
  );
};

export const useUpdateOutlookCalendarEventMutation = (): UseMutationResult<
  OutlookCalendarEventMutationResponse,
  unknown,
  { eventId: string; payload: OutlookCalendarEventMutationRequest }
> => {
  return useMutation(({ eventId, payload }) =>
    dataService.updateOutlookCalendarEvent(eventId, payload),
  );
};

export const useDeleteOutlookCalendarEventMutation = (): UseMutationResult<
  { eventId: string; message: string },
  unknown,
  string
> => {
  return useMutation((eventId: string) => dataService.deleteOutlookCalendarEvent(eventId));
};

export const useOutlookMessageQuery = (
  messageId?: string,
  config?: UseQueryOptions<OutlookMessage>,
): QueryObserverResult<OutlookMessage> => {
  const queriesEnabled = useRecoilValue<boolean>(store.queriesEnabled);

  return useQuery<OutlookMessage>(
    [QueryKeys.outlookMessage, messageId],
    () => dataService.getOutlookMessage(messageId ?? ''),
    {
      refetchOnWindowFocus: false,
      ...config,
      enabled: Boolean(messageId) && (config?.enabled ?? true) === true && queriesEnabled,
    },
  );
};

export const useAnalyzeOutlookMessageMutation = (): UseMutationResult<
  OutlookAnalyzeResponse,
  unknown,
  string
> => {
  return useMutation((messageId: string) => dataService.analyzeOutlookMessage(messageId));
};

export const useAnalyzeSelectedOutlookMessagesMutation = (): UseMutationResult<
  OutlookAnalyzeSelectionResponse,
  unknown,
  OutlookAnalyzeSelectionRequest
> => {
  return useMutation((payload: OutlookAnalyzeSelectionRequest) =>
    dataService.analyzeSelectedOutlookMessages(payload),
  );
};

export const useCreateOutlookDraftMutation = (): UseMutationResult<
  OutlookDraftResponse,
  unknown,
  { messageId: string; payload: OutlookDraftRequest }
> => {
  return useMutation(({ messageId, payload }) =>
    dataService.createOutlookDraft(messageId, payload),
  );
};

export const useDeleteOutlookMessageMutation = (): UseMutationResult<
  OutlookDeleteResponse,
  unknown,
  string
> => {
  return useMutation((messageId: string) => dataService.deleteOutlookMessage(messageId));
};

export const useUpdateOutlookMessageReadStateMutation = (): UseMutationResult<
  OutlookUpdateReadStateResponse,
  unknown,
  { messageId: string; isRead: boolean }
> => {
  return useMutation(({ messageId, isRead }) =>
    dataService.updateOutlookMessageReadState(messageId, { isRead }),
  );
};

export const useOutlookDailyBriefMutation = (): UseMutationResult<
  OutlookDailyBriefResponse,
  unknown,
  void
> => {
  return useMutation(() => dataService.getOutlookDailyBrief());
};

export const useProposeOutlookMeetingSlotsMutation = (): UseMutationResult<
  OutlookMeetingSlotsResponse,
  unknown,
  { messageId: string; payload: OutlookMeetingSlotsRequest }
> => {
  return useMutation(({ messageId, payload }) =>
    dataService.proposeOutlookMeetingSlots(messageId, payload),
  );
};

export const useCreateOutlookMeetingMutation = (): UseMutationResult<
  OutlookCreateMeetingResponse,
  unknown,
  { messageId: string; payload: OutlookCreateMeetingRequest }
> => {
  return useMutation(({ messageId, payload }) =>
    dataService.createOutlookMeeting(messageId, payload),
  );
};
