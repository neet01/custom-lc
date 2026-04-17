import React from 'react';
import { InfoHoverCard, ESide, Label } from '@librechat/client';
import { useLocalize } from '~/hooks';

function clamp(value: number, min: number, max: number) {
  return Math.min(Math.max(value, min), max);
}

function formatNumber(value: number) {
  return new Intl.NumberFormat().format(Math.round(value));
}

function getUsageStats(limit?: number, remaining?: number) {
  const safeLimit = Math.max(limit ?? 0, 0);
  const rawRemaining = Math.max(remaining ?? 0, 0);

  if (safeLimit <= 0) {
    return {
      used: 0,
      remaining: rawRemaining,
      percentUsed: 0,
      percentRemaining: 0,
    };
  }

  const boundedRemaining = clamp(rawRemaining, 0, safeLimit);
  const used = clamp(safeLimit - boundedRemaining, 0, safeLimit);

  return {
    used,
    remaining: rawRemaining,
    percentUsed: clamp((used / safeLimit) * 100, 0, 100),
    percentRemaining: clamp((boundedRemaining / safeLimit) * 100, 0, 100),
  };
}

export interface TokenUsageProgressProps {
  limit?: number;
  tokenCredits?: number;
  compact?: boolean;
}

export default function TokenUsageProgress({
  limit,
  tokenCredits,
  compact = false,
}: TokenUsageProgressProps) {
  const localize = useLocalize();
  const { used, remaining, percentRemaining } = getUsageStats(limit, tokenCredits);
  const colorHue = Math.round((percentRemaining / 100) * 120);
  const leadHue = clamp(colorHue + 18, 0, 120);
  const trailHue = clamp(colorHue - 18, 0, 120);
  const fillStyle = {
    width: `${percentRemaining}%`,
    background: `linear-gradient(90deg, hsl(${leadHue} 78% 54%), hsl(${trailHue} 86% 46%))`,
  };

  if (!limit || limit <= 0) {
    return (
      <div className="flex items-center justify-between">
        <div className="flex items-center space-x-2">
          <Label className="font-light">{localize('com_nav_balance')}</Label>
          <InfoHoverCard side={ESide.Bottom} text={localize('com_nav_info_balance')} />
        </div>
        <span className="text-sm font-medium text-text-primary" role="note">
          {tokenCredits !== undefined ? formatNumber(tokenCredits) : '0'}
        </span>
      </div>
    );
  }

  return (
    <div className={compact ? 'space-y-2' : 'space-y-3'}>
      <div className="flex items-center justify-between">
        <div className="flex items-center space-x-2">
          <Label className="font-light">{localize('com_nav_balance')}</Label>
          <InfoHoverCard side={ESide.Bottom} text={localize('com_nav_info_balance')} />
        </div>
        <div className="text-right">
          <div className="text-sm font-semibold text-text-primary">{formatNumber(remaining)}</div>
          <div className="text-[11px] uppercase tracking-[0.16em] text-text-secondary">
            remaining
          </div>
        </div>
      </div>

      <div className="space-y-2">
        <div className="h-3 overflow-hidden rounded-full bg-surface-tertiary shadow-inner">
          <div
            className="h-full rounded-full transition-[width] duration-500 ease-out"
            style={fillStyle}
          />
        </div>
        <div
          className={
            compact
              ? 'grid grid-cols-3 gap-2 text-[11px] text-text-secondary'
              : 'grid grid-cols-3 gap-3 text-xs text-text-secondary'
          }
        >
          <div>
            <div className="uppercase tracking-[0.14em]">Used</div>
            <div className="mt-1 font-medium text-text-primary">{formatNumber(used)}</div>
          </div>
          <div>
            <div className="uppercase tracking-[0.14em]">Remaining</div>
            <div className="mt-1 font-medium text-text-primary">{formatNumber(remaining)}</div>
          </div>
          <div>
            <div className="uppercase tracking-[0.14em]">Limit</div>
            <div className="mt-1 font-medium text-text-primary">{formatNumber(limit)}</div>
          </div>
        </div>
      </div>
    </div>
  );
}
