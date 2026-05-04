import React, { useState } from 'react';
import DisplayUsernameMessages from './DisplayUsernameMessages';
import DeleteAccount from './DeleteAccount';
import Avatar from './Avatar';
import EnableTwoFactorItem from './TwoFactorAuthentication';
import BackupCodesItem from './BackupCodesItem';
import { useGetStartupConfig } from '~/data-provider';
import { useAuthContext } from '~/hooks';
import { useTutorial } from '~/Providers';
import type { TutorialId } from '~/tutorials/definitions';

function TutorialLauncher({ onRequestClose }: { onRequestClose?: () => void }) {
  const { tutorials, startTutorial } = useTutorial();
  const [open, setOpen] = useState(false);

  const launchTutorial = (tutorialId: TutorialId) => {
    onRequestClose?.();
    window.setTimeout(() => startTutorial(tutorialId), 180);
  };

  return (
    <div className="rounded-2xl border border-border-medium bg-surface-secondary p-4">
      <div className="flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
        <div>
          <div className="text-sm font-semibold text-text-primary">Tutorials</div>
          <p className="mt-1 text-xs text-text-secondary">
            Launch guided walkthroughs for the major Cortex workspaces from here when needed.
          </p>
        </div>
        <button
          type="button"
          className="rounded-xl border border-border-medium px-3 py-2 text-sm font-medium text-text-primary transition-colors hover:bg-surface-hover"
          onClick={() => setOpen((current) => !current)}
        >
          {open ? 'Hide tutorials' : 'Start tutorial'}
        </button>
      </div>

      {open ? (
        <div className="mt-4 space-y-3">
          {tutorials.map((tutorial) => (
            <button
              key={tutorial.id}
              type="button"
              className="block w-full rounded-2xl border border-border-medium bg-surface-primary p-4 text-left transition-colors hover:bg-surface-hover"
              onClick={() => launchTutorial(tutorial.id)}
            >
              <div className="text-sm font-semibold text-text-primary">{tutorial.title}</div>
              <div className="mt-1 text-xs leading-5 text-text-secondary">
                {tutorial.description}
              </div>
              <div className="mt-2 text-[11px] font-medium uppercase tracking-wide text-[#b88a00] dark:text-[#f5d000]">
                {tutorial.steps.length} steps
              </div>
            </button>
          ))}
        </div>
      ) : null}
    </div>
  );
}

function Account({ onRequestClose }: { onRequestClose?: () => void }) {
  const { user } = useAuthContext();
  const { data: startupConfig } = useGetStartupConfig();

  return (
    <div className="flex flex-col gap-3 p-1 text-sm text-text-primary">
      <div className="pb-3">
        <TutorialLauncher onRequestClose={onRequestClose} />
      </div>
      <div className="pb-3">
        <DisplayUsernameMessages />
      </div>
      <div className="pb-3">
        <Avatar />
      </div>
      {user?.provider === 'local' && (
        <>
          <div className="pb-3">
            <EnableTwoFactorItem />
          </div>
          {user?.twoFactorEnabled && (
            <div className="pb-3">
              <BackupCodesItem />
            </div>
          )}
        </>
      )}
      {startupConfig?.allowAccountDeletion !== false && (
        <div className="pb-3">
          <DeleteAccount />
        </div>
      )}
    </div>
  );
}

export default React.memo(Account);
