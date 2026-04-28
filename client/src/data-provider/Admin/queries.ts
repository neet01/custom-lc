import { useRecoilValue } from 'recoil';
import { QueryKeys, dataService } from 'librechat-data-provider';
import { useMutation, useQuery, useQueryClient } from '@tanstack/react-query';
import type {
  QueryObserverResult,
  UseMutationResult,
  UseQueryOptions,
} from '@tanstack/react-query';
import type t from 'librechat-data-provider';
import store from '~/store';

export const useAdminUsersQuery = (
  params: t.AdminUsersListParams = {},
  config?: UseQueryOptions<t.AdminUsersListResponse>,
): QueryObserverResult<t.AdminUsersListResponse> => {
  const queriesEnabled = useRecoilValue<boolean>(store.queriesEnabled);

  return useQuery<t.AdminUsersListResponse>(
    [QueryKeys.adminUsers, params],
    () => dataService.getAdminUsers(params),
    {
      refetchOnWindowFocus: false,
      refetchOnReconnect: false,
      refetchOnMount: false,
      ...config,
      enabled: (config?.enabled ?? true) === true && queriesEnabled,
    },
  );
};

export const useAdminUsageQuery = (
  params: t.AdminUsageListParams = {},
  config?: UseQueryOptions<t.AdminUsageListResponse>,
): QueryObserverResult<t.AdminUsageListResponse> => {
  const queriesEnabled = useRecoilValue<boolean>(store.queriesEnabled);

  return useQuery<t.AdminUsageListResponse>(
    [QueryKeys.adminUsage, params],
    () => dataService.getAdminUsage(params),
    {
      refetchOnWindowFocus: false,
      refetchOnReconnect: false,
      refetchOnMount: false,
      ...config,
      enabled: (config?.enabled ?? true) === true && queriesEnabled,
    },
  );
};

export const useAdminUsageSummaryQuery = (
  params: t.AdminUsageSummaryParams = {},
  config?: UseQueryOptions<t.AdminUsageSummaryResponse>,
): QueryObserverResult<t.AdminUsageSummaryResponse> => {
  const queriesEnabled = useRecoilValue<boolean>(store.queriesEnabled);

  return useQuery<t.AdminUsageSummaryResponse>(
    [QueryKeys.adminUsageSummary, params],
    () => dataService.getAdminUsageSummary(params),
    {
      refetchOnWindowFocus: false,
      refetchOnReconnect: false,
      refetchOnMount: false,
      ...config,
      enabled: (config?.enabled ?? true) === true && queriesEnabled,
    },
  );
};

export const useAdminIssuesQuery = (
  params: t.AdminIssuesListParams = {},
  config?: UseQueryOptions<t.AdminIssuesListResponse>,
): QueryObserverResult<t.AdminIssuesListResponse> => {
  const queriesEnabled = useRecoilValue<boolean>(store.queriesEnabled);

  return useQuery<t.AdminIssuesListResponse>(
    [QueryKeys.adminIssues, params],
    () => dataService.getAdminIssues(params),
    {
      refetchOnWindowFocus: false,
      refetchOnReconnect: false,
      refetchOnMount: false,
      ...config,
      enabled: (config?.enabled ?? true) === true && queriesEnabled,
    },
  );
};

export const useAdminOutlookAuditQuery = (
  params: t.AdminOutlookAuditListParams = {},
  config?: UseQueryOptions<t.AdminOutlookAuditListResponse>,
): QueryObserverResult<t.AdminOutlookAuditListResponse> => {
  const queriesEnabled = useRecoilValue<boolean>(store.queriesEnabled);

  return useQuery<t.AdminOutlookAuditListResponse>(
    [QueryKeys.adminOutlookAudit, params],
    () => dataService.getAdminOutlookAudit(params),
    {
      refetchOnWindowFocus: false,
      refetchOnReconnect: false,
      refetchOnMount: false,
      ...config,
      enabled: (config?.enabled ?? true) === true && queriesEnabled,
    },
  );
};

export const useAdminUpdateUserBalanceMutation = (): UseMutationResult<
  t.AdminUpdateUserBalanceResponse,
  unknown,
  { userId: string; tokenCredits: number },
  { previousUsers: [unknown[], t.AdminUsersListResponse | undefined][] }
> => {
  const queryClient = useQueryClient();

  return useMutation(
    ({ userId, tokenCredits }) =>
      dataService.updateAdminUserBalance(userId, {
        tokenCredits,
      }),
    {
      onMutate: async ({ userId, tokenCredits }) => {
        await queryClient.cancelQueries([QueryKeys.adminUsers]);
        const previousUsers = queryClient.getQueriesData<t.AdminUsersListResponse>([
          QueryKeys.adminUsers,
        ]);

        queryClient.setQueriesData<t.AdminUsersListResponse>([QueryKeys.adminUsers], (current) => {
          if (!current) {
            return current;
          }

          return {
            ...current,
            users: current.users.map((user) =>
              user.id === userId ? { ...user, tokenCredits } : user,
            ),
          };
        });

        return { previousUsers };
      },
      onError: (_error, _variables, context) => {
        for (const [queryKey, previousValue] of context?.previousUsers ?? []) {
          queryClient.setQueryData(queryKey, previousValue);
        }
      },
      onSuccess: ({ user }) => {
        queryClient.setQueriesData<t.AdminUsersListResponse>([QueryKeys.adminUsers], (current) => {
          if (!current) {
            return current;
          }

          return {
            ...current,
            users: current.users.map((existingUser) =>
              existingUser.id === user.id ? user : existingUser,
            ),
          };
        });
      },
      onSettled: () => {
        queryClient.invalidateQueries([QueryKeys.adminUsers]);
      },
    },
  );
};
