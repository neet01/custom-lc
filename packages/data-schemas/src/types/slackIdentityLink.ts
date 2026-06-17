import type { SlackIdentityLinkStatus } from '~/schema/slackIdentityLink';

export type { SlackIdentityLinkStatus };

export interface SlackIdentityLinkData {
  user: string;
  slackUserId: string;
  teamId?: string;
  teamName?: string;
  enterpriseId?: string;
  enterpriseName?: string;
  slackEmail?: string;
  slackDisplayName?: string;
  status?: SlackIdentityLinkStatus;
  source?: 'oauth_install' | 'admin_link' | 'manual';
  linkedAt?: Date;
  lastVerifiedAt?: Date;
}
