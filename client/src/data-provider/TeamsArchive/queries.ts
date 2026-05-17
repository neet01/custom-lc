import { useRecoilValue } from 'recoil';
import { QueryKeys, dataService } from 'librechat-data-provider';
import { useMutation, useQuery } from '@tanstack/react-query';
import type {
  QueryObserverResult,
  UseMutationResult,
  UseQueryOptions,
} from '@tanstack/react-query';
import type {
  TeamsArchiveCancelResponse,
  TeamsArchiveResetResponse,
  TeamsArchiveSyncAcceptedResponse,
  TeamsArchiveStatusResponse,
  TeamsArchiveSyncRequest,
  TeamsArchiveSyncResponse,
} from 'librechat-data-provider';
import store from '~/store';

export const useTeamsArchiveStatusQuery = (
  config?: UseQueryOptions<TeamsArchiveStatusResponse>,
): QueryObserverResult<TeamsArchiveStatusResponse> => {
  const queriesEnabled = useRecoilValue<boolean>(store.queriesEnabled);

  return useQuery<TeamsArchiveStatusResponse>(
    [QueryKeys.teamsArchiveStatus],
    () => dataService.getTeamsArchiveStatus(),
    {
      refetchOnWindowFocus: false,
      ...config,
      enabled: (config?.enabled ?? true) === true && queriesEnabled,
    },
  );
};

export const useSyncTeamsArchiveMutation = (): UseMutationResult<
  TeamsArchiveSyncResponse | TeamsArchiveSyncAcceptedResponse,
  unknown,
  TeamsArchiveSyncRequest
> => {
  return useMutation((payload: TeamsArchiveSyncRequest) => dataService.syncTeamsArchive(payload));
};

export const useCancelTeamsArchiveSyncMutation = (): UseMutationResult<
  TeamsArchiveCancelResponse,
  unknown,
  void
> => {
  return useMutation(() => dataService.cancelTeamsArchiveSync());
};

export const useResetTeamsArchiveMutation = (): UseMutationResult<
  TeamsArchiveResetResponse,
  unknown,
  void
> => {
  return useMutation(() => dataService.resetTeamsArchive());
};
