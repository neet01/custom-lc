import { useRecoilValue } from 'recoil';
import { QueryKeys, dataService } from 'librechat-data-provider';
import { useMutation, useQuery } from '@tanstack/react-query';
import type {
  QueryObserverResult,
  UseMutationResult,
  UseQueryOptions,
} from '@tanstack/react-query';
import type {
  SlackArchiveCancelResponse,
  SlackArchiveResetResponse,
  SlackArchiveStatusResponse,
  SlackArchiveSyncAcceptedResponse,
  SlackArchiveSyncRequest,
  SlackArchiveSyncResponse,
} from 'librechat-data-provider';
import store from '~/store';

export const useSlackArchiveStatusQuery = (
  config?: UseQueryOptions<SlackArchiveStatusResponse>,
): QueryObserverResult<SlackArchiveStatusResponse> => {
  const queriesEnabled = useRecoilValue<boolean>(store.queriesEnabled);

  return useQuery<SlackArchiveStatusResponse>(
    [QueryKeys.slackArchiveStatus],
    () => dataService.getSlackArchiveStatus(),
    {
      refetchOnWindowFocus: false,
      ...config,
      enabled: (config?.enabled ?? true) === true && queriesEnabled,
    },
  );
};

export const useSyncSlackArchiveMutation = (): UseMutationResult<
  SlackArchiveSyncResponse | SlackArchiveSyncAcceptedResponse,
  unknown,
  SlackArchiveSyncRequest
> => {
  return useMutation((payload: SlackArchiveSyncRequest) => dataService.syncSlackArchive(payload));
};

export const useCancelSlackArchiveSyncMutation = (): UseMutationResult<
  SlackArchiveCancelResponse,
  unknown,
  void
> => {
  return useMutation(() => dataService.cancelSlackArchiveSync());
};

export const useResetSlackArchiveMutation = (): UseMutationResult<
  SlackArchiveResetResponse,
  unknown,
  void
> => {
  return useMutation(() => dataService.resetSlackArchive());
};
