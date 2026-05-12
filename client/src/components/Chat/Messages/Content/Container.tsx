import { TMessage } from 'librechat-data-provider';
import { cn } from '~/utils';
import Files from './Files';

const Container = ({ children, message }: { children: React.ReactNode; message?: TMessage }) => (
  <div
    className="text-message flex min-h-[20px] flex-col items-start gap-3 overflow-visible [.text-message+&]:mt-5"
    dir="auto"
  >
    {message?.isCreatedByUser === true && <Files message={message} />}
    <div
      className={cn(
        'w-full',
        message?.isCreatedByUser === true &&
          'rounded-2xl border border-[#f5d000]/30 bg-[#f5d000]/12 px-4 py-3 shadow-[0_8px_24px_rgba(245,208,0,0.08)]',
      )}
    >
      {children}
    </div>
  </div>
);

export default Container;
