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
  OutlookDraftRequest,
  OutlookDraftResponse,
  OutlookMessage,
  OutlookMessagesParams,
  OutlookMessagesResponse,
  OutlookStatusResponse,
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

export const useCreateOutlookDraftMutation = (): UseMutationResult<
  OutlookDraftResponse,
  unknown,
  { messageId: string; payload: OutlookDraftRequest }
> => {
  return useMutation(({ messageId, payload }) => dataService.createOutlookDraft(messageId, payload));
};
