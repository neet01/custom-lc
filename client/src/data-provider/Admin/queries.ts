import { useRecoilValue } from 'recoil';
import { QueryKeys, dataService } from 'librechat-data-provider';
import { useQuery } from '@tanstack/react-query';
import type { QueryObserverResult, UseQueryOptions } from '@tanstack/react-query';
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
