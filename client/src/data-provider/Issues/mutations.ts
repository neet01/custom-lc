import { useMutation, useQueryClient } from '@tanstack/react-query';
import { QueryKeys, dataService } from 'librechat-data-provider';
import type t from 'librechat-data-provider';

export function useCreateIssueReportMutation() {
  const queryClient = useQueryClient();

  return useMutation(
    (payload: t.IssueReportCreateRequest) => dataService.createIssueReport(payload),
    {
      onSuccess: () => {
        queryClient.invalidateQueries([QueryKeys.adminIssues]);
      },
    },
  );
}
