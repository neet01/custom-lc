import React from 'react';
import TokenUsageProgress from './TokenUsageProgress';

interface TokenCreditsItemProps {
  tokenCredits?: number;
  tokenLimit?: number;
}

const TokenCreditsItem: React.FC<TokenCreditsItemProps> = ({ tokenCredits, tokenLimit }) => (
  <TokenUsageProgress limit={tokenLimit} tokenCredits={tokenCredits} />
);

export default TokenCreditsItem;
