import type { SlackWorkspaceInstallStatus } from '~/schema/slackWorkspaceInstall';

export type { SlackWorkspaceInstallStatus };

export interface SlackWorkspaceInstallData {
  installedByUser?: string;
  teamId?: string;
  teamName?: string;
  enterpriseId?: string;
  enterpriseName?: string;
  botUserId?: string;
  botAccessToken?: string;
  botScopes?: string;
  userScopes?: string;
  installPayload?: Record<string, unknown>;
  status?: SlackWorkspaceInstallStatus;
  installedAt?: Date;
  lastValidatedAt?: Date;
}
