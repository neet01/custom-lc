import React from 'react';
import { InfoHoverCard, ESide, Label } from '@librechat/client';
import { useLocalize } from '~/hooks';

const CREDITS_PER_USD = 1_000_000;
const BEDROCK_MODELING_NOTE =
  'Estimated budget modeled for AWS Bedrock Claude Sonnet 3.7 and Claude Sonnet 4.5. Current rates: $3 / 1M input tokens, $15 / 1M output tokens, cache write $3.75 / 1M, cache read $0.30 / 1M.';

function clamp(value: number, min: number, max: number) {
  return Math.min(Math.max(value, min), max);
}

function formatNumber(value: number) {
  return new Intl.NumberFormat().format(Math.round(value));
}

function creditsToUsd(value: number) {
  return value / CREDITS_PER_USD;
}

function formatUsd(value: number) {
  const absValue = Math.abs(value);
  const fractionDigits = absValue >= 100 ? 0 : absValue >= 1 ? 2 : 4;
  return new Intl.NumberFormat('en-US', {
    style: 'currency',
    currency: 'USD',
    minimumFractionDigits: fractionDigits,
    maximumFractionDigits: fractionDigits,
  }).format(value);
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
  const usedUsd = creditsToUsd(used);
  const remainingUsd = creditsToUsd(Math.max(remaining, 0));
  const limitUsd = creditsToUsd(Math.max(limit ?? 0, 0));
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
          <Label className="font-light">Estimated Budget</Label>
          <InfoHoverCard
            side={ESide.Bottom}
            text={`${localize('com_nav_info_balance')}. ${BEDROCK_MODELING_NOTE}`}
          />
        </div>
        <div className="text-right" role="note">
          <div className="text-sm font-semibold text-text-primary">
            {formatUsd(creditsToUsd(tokenCredits ?? 0))}
          </div>
          <div className="text-[11px] uppercase tracking-[0.16em] text-text-secondary">
            {formatNumber(tokenCredits ?? 0)} credits
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className={compact ? 'space-y-2' : 'space-y-3'}>
      <div className="flex items-center justify-between">
        <div className="flex items-center space-x-2">
          <Label className="font-light">Estimated Budget</Label>
          <InfoHoverCard
            side={ESide.Bottom}
            text={`${localize('com_nav_info_balance')}. ${BEDROCK_MODELING_NOTE}`}
          />
        </div>
        <div className="text-right">
          <div className="text-sm font-semibold text-text-primary">{formatUsd(remainingUsd)}</div>
          <div className="text-[11px] uppercase tracking-[0.16em] text-text-secondary">
            budget left
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
            <div className="uppercase tracking-[0.14em]">Spend Used</div>
            <div className="mt-1 font-medium text-text-primary">{formatUsd(usedUsd)}</div>
            <div className="mt-1 text-[11px] text-text-secondary">{formatNumber(used)} credits</div>
          </div>
          <div>
            <div className="uppercase tracking-[0.14em]">Budget Left</div>
            <div className="mt-1 font-medium text-text-primary">{formatUsd(remainingUsd)}</div>
            <div className="mt-1 text-[11px] text-text-secondary">
              {formatNumber(remaining)} credits
            </div>
          </div>
          <div>
            <div className="uppercase tracking-[0.14em]">Budget Cap</div>
            <div className="mt-1 font-medium text-text-primary">{formatUsd(limitUsd)}</div>
            <div className="mt-1 text-[11px] text-text-secondary">{formatNumber(limit)} credits</div>
          </div>
        </div>
        <div className={compact ? 'text-[10px] text-text-secondary' : 'text-[11px] text-text-secondary'}>
          {BEDROCK_MODELING_NOTE}
        </div>
      </div>
    </div>
  );
}
