import type { EnterpriseMemoryJobStatus } from '~/schema/enterpriseMemoryJob';
import type { EnterpriseMemoryVisibilityScope } from '~/schema/enterpriseMemoryEntity';

export type { EnterpriseMemoryJobStatus, EnterpriseMemoryVisibilityScope };

export interface EnterpriseMemoryEntityData {
  user?: string;
  tenantId?: string;
  visibilityScope?: EnterpriseMemoryVisibilityScope;
  source: string;
  entityType: string;
  canonicalKey: string;
  displayName: string;
  aliases?: string[];
  summary?: string;
  sourceRecordType?: string;
  sourceRecordId?: string;
  sourceParentRecordId?: string;
  sourceUpdatedAt?: Date;
  attributes?: Record<string, unknown>;
}

export interface EnterpriseMemoryRelationshipData {
  user?: string;
  tenantId?: string;
  visibilityScope?: EnterpriseMemoryVisibilityScope;
  source: string;
  relationshipType: string;
  fromEntityId: string;
  toEntityId: string;
  sourceRecordType?: string;
  sourceRecordId?: string;
  sourceUpdatedAt?: Date;
  attributes?: Record<string, unknown>;
}

export interface EnterpriseMemoryChunkData {
  user?: string;
  tenantId?: string;
  visibilityScope?: EnterpriseMemoryVisibilityScope;
  source: string;
  sourceRecordType: string;
  sourceRecordId: string;
  sourceParentRecordId?: string;
  parentEntityId?: string;
  entityIds?: string[];
  chunkType: string;
  title?: string;
  text: string;
  summary?: string;
  orderIndex?: number;
  sourceTimestamp?: Date;
  metadata?: Record<string, unknown>;
}

export interface EnterpriseMemoryJobData {
  user?: string;
  tenantId?: string;
  visibilityScope?: EnterpriseMemoryVisibilityScope;
  source: string;
  jobType: string;
  status: EnterpriseMemoryJobStatus;
  sourceRecordType?: string;
  sourceRecordId?: string;
  checkpoint?: Record<string, unknown>;
  stats?: Record<string, unknown>;
  errorMessage?: string;
  startedAt?: Date;
  completedAt?: Date;
  lastHeartbeatAt?: Date;
}
