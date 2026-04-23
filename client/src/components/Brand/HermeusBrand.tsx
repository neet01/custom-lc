import { cn } from '~/utils';

type HermeusBrandProps = {
  compact?: boolean;
  className?: string;
  markClassName?: string;
  ariaLabel?: string;
};

const HERMEUS_LOGO_SRC = 'assets/hermeus-logo.png';
const HERMEUS_EMBLEM_SRC = 'assets/hermeus-emblem.png';

export function HermeusMark({ className }: { className?: string }) {
  return (
    <span
      role="img"
      aria-label="Hermeus Cortex"
      className={cn('flex h-9 w-9 shrink-0 items-center justify-center', className)}
    >
      <img src={HERMEUS_EMBLEM_SRC} alt="" className="h-full w-full object-contain" />
    </span>
  );
}

export default function HermeusBrand({
  compact = false,
  className,
  markClassName,
  ariaLabel = 'Hermeus Cortex',
}: HermeusBrandProps) {
  return (
    <div
      className={cn('flex min-w-0 items-center gap-3 text-text-primary', className)}
      aria-label={ariaLabel}
    >
      {compact ? (
        <HermeusMark className={markClassName} />
      ) : (
        <span className="flex max-w-full items-center rounded-md bg-[#050505] px-3 py-2">
          <img
            src={HERMEUS_LOGO_SRC}
            alt=""
            className={cn('h-8 w-auto max-w-full object-contain', markClassName)}
          />
        </span>
      )}
    </div>
  );
}
