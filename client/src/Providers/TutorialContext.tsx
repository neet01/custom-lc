import {
  createContext,
  useCallback,
  useContext,
  useEffect,
  useMemo,
  useRef,
  useState,
} from 'react';
import type { ReactNode } from 'react';
import { useSetRecoilState } from 'recoil';
import { Button } from '@librechat/client';
import { useActivePanel } from './ActivePanelContext';
import store from '~/store';
import {
  buildTutorialDefinitions,
  type TutorialDefinition,
  type TutorialId,
  type TutorialStep,
} from '~/tutorials/definitions';
import { cn } from '~/utils';

type RectLike = { top: number; left: number; width: number; height: number };

interface TutorialContextValue {
  tutorials: TutorialDefinition[];
  activeTutorial: TutorialDefinition | null;
  activeStep: TutorialStep | null;
  isOpen: boolean;
  startTutorial: (tutorialId: TutorialId) => void;
  closeTutorial: () => void;
  nextStep: () => void;
  previousStep: () => void;
}

const TutorialContext = createContext<TutorialContextValue | undefined>(undefined);

function getStepRect(element: Element | null): RectLike | null {
  if (!(element instanceof HTMLElement)) {
    return null;
  }

  const rect = element.getBoundingClientRect();
  if (rect.width <= 0 || rect.height <= 0) {
    return null;
  }

  return {
    top: Math.max(12, rect.top - 8),
    left: Math.max(12, rect.left - 8),
    width: rect.width + 16,
    height: rect.height + 16,
  };
}

function findTargetElement(target: string | undefined): HTMLElement | null {
  if (!target) {
    return null;
  }

  return document.querySelector<HTMLElement>(`[data-tour="${target}"]`);
}

function getCardPosition(rect: RectLike | null) {
  const cardWidth = 340;
  const cardHeight = 240;
  const viewportWidth = window.innerWidth;
  const viewportHeight = window.innerHeight;
  const gap = 20;

  if (!rect) {
    return {
      top: Math.max(24, viewportHeight / 2 - cardHeight / 2),
      left: Math.max(24, viewportWidth / 2 - cardWidth / 2),
    };
  }

  const rightPlacement = rect.left + rect.width + gap;
  if (rightPlacement + cardWidth <= viewportWidth - 24) {
    return {
      top: Math.min(Math.max(24, rect.top), viewportHeight - cardHeight - 24),
      left: rightPlacement,
    };
  }

  const leftPlacement = rect.left - gap - cardWidth;
  if (leftPlacement >= 24) {
    return {
      top: Math.min(Math.max(24, rect.top), viewportHeight - cardHeight - 24),
      left: leftPlacement,
    };
  }

  const belowPlacement = rect.top + rect.height + gap;
  if (belowPlacement + cardHeight <= viewportHeight - 24) {
    return {
      top: belowPlacement,
      left: Math.min(Math.max(24, rect.left), viewportWidth - cardWidth - 24),
    };
  }

  return {
    top: Math.max(24, rect.top - cardHeight - gap),
    left: Math.min(Math.max(24, rect.left), viewportWidth - cardWidth - 24),
  };
}

export function TutorialProvider({ children }: { children: ReactNode }) {
  const { setActive } = useActivePanel();
  const setSidebarExpanded = useSetRecoilState(store.sidebarExpanded);
  const [activeTutorialId, setActiveTutorialId] = useState<TutorialId | null>(null);
  const [stepIndex, setStepIndex] = useState(0);
  const [targetRect, setTargetRect] = useState<RectLike | null>(null);
  const activeStepCleanupRef = useRef<number | null>(null);

  const tutorialContext = useMemo(
    () => ({
      openPanel: (panelId: string) => {
        setActive(panelId);
      },
      expandSidebar: () => {
        setSidebarExpanded(true);
      },
      collapseSidebar: () => {
        setSidebarExpanded(false);
      },
      dispatch: (eventName: string) => {
        window.dispatchEvent(new CustomEvent(eventName));
      },
    }),
    [setActive, setSidebarExpanded],
  );

  const tutorialMap = useMemo(() => buildTutorialDefinitions(tutorialContext), [tutorialContext]);
  const tutorials = useMemo(() => Object.values(tutorialMap), [tutorialMap]);
  const activeTutorial = activeTutorialId ? tutorialMap[activeTutorialId] : null;
  const activeStep = activeTutorial?.steps[stepIndex] ?? null;

  const closeTutorial = useCallback(() => {
    setActiveTutorialId(null);
    setStepIndex(0);
    setTargetRect(null);
  }, []);

  const startTutorial = useCallback((tutorialId: TutorialId) => {
    setActiveTutorialId(tutorialId);
    setStepIndex(0);
  }, []);

  const nextStep = useCallback(() => {
    setStepIndex((current) => {
      if (!activeTutorial) {
        return current;
      }

      if (current >= activeTutorial.steps.length - 1) {
        return current;
      }

      return current + 1;
    });
  }, [activeTutorial]);

  const previousStep = useCallback(() => {
    setStepIndex((current) => Math.max(0, current - 1));
  }, []);

  useEffect(() => {
    if (!activeStep) {
      return;
    }

    activeStep.beforeEnter?.(tutorialContext);

    if (!activeStep.target) {
      setTargetRect(null);
      return;
    }

    let cancelled = false;
    let attempts = 0;

    const resolveTarget = () => {
      if (cancelled) {
        return;
      }

      const element = findTargetElement(activeStep.target);
      if (element) {
        element.scrollIntoView({
          behavior: 'smooth',
          block: 'center',
          inline: 'center',
        });

        const updateRect = () => {
          if (cancelled) {
            return;
          }
          setTargetRect(getStepRect(element));
        };

        updateRect();
        window.addEventListener('resize', updateRect);
        window.addEventListener('scroll', updateRect, true);

        return () => {
          window.removeEventListener('resize', updateRect);
          window.removeEventListener('scroll', updateRect, true);
        };
      }

      attempts += 1;
      if (attempts > 20) {
        setTargetRect(null);
        return;
      }

      activeStepCleanupRef.current = window.setTimeout(resolveTarget, 120);
    };

    const cleanupListeners = resolveTarget();

    return () => {
      cancelled = true;
      setTargetRect(null);
      if (activeStepCleanupRef.current != null) {
        window.clearTimeout(activeStepCleanupRef.current);
        activeStepCleanupRef.current = null;
      }
      cleanupListeners?.();
    };
  }, [activeStep, tutorialContext]);

  const value = useMemo<TutorialContextValue>(
    () => ({
      tutorials,
      activeTutorial,
      activeStep,
      isOpen: activeTutorial != null && activeStep != null,
      startTutorial,
      closeTutorial,
      nextStep,
      previousStep,
    }),
    [tutorials, activeTutorial, activeStep, startTutorial, closeTutorial, nextStep, previousStep],
  );

  const stepCount = activeTutorial?.steps.length ?? 0;
  const currentStepNumber = activeStep ? stepIndex + 1 : 0;
  const cardPosition =
    typeof window !== 'undefined' ? getCardPosition(targetRect) : { top: 24, left: 24 };

  return (
    <TutorialContext.Provider value={value}>
      {children}
      {activeTutorial && activeStep ? (
        <div className="pointer-events-auto fixed inset-0 z-[140]">
          <div className="absolute inset-0 bg-transparent" />
          {targetRect ? (
            <div
              className="pointer-events-none absolute rounded-2xl border border-[#f5d000]/80 shadow-[0_0_0_9999px_rgba(6,9,16,0.62)]"
              style={{
                top: `${targetRect.top}px`,
                left: `${targetRect.left}px`,
                width: `${targetRect.width}px`,
                height: `${targetRect.height}px`,
              }}
            />
          ) : (
            <div className="pointer-events-none absolute inset-0 bg-black/60" />
          )}
          <div
            className={cn(
              'absolute w-[340px] rounded-2xl border border-border-medium bg-surface-primary p-4 shadow-2xl',
            )}
            style={{
              top: `${cardPosition.top}px`,
              left: `${cardPosition.left}px`,
            }}
          >
            <div className="flex items-start justify-between gap-3">
              <div>
                <div className="text-[11px] font-semibold uppercase tracking-[0.18em] text-[#b88a00] dark:text-[#f5d000]">
                  Tutorial
                </div>
                <h3 className="mt-1 text-base font-semibold text-text-primary">
                  {activeStep.title}
                </h3>
              </div>
              <button
                type="button"
                className="rounded-lg border border-border-light px-2 py-1 text-xs text-text-secondary transition-colors hover:bg-surface-hover hover:text-text-primary"
                onClick={closeTutorial}
              >
                Close
              </button>
            </div>

            <p className="mt-3 text-sm leading-6 text-text-secondary">{activeStep.description}</p>

            <div className="mt-4 flex items-center justify-between text-xs text-text-secondary">
              <span>{activeTutorial.title}</span>
              <span>
                {currentStepNumber} / {stepCount}
              </span>
            </div>

            <div className="mt-4 flex items-center justify-between gap-2">
              <Button
                variant="outline"
                size="sm"
                className="min-w-24"
                onClick={previousStep}
                disabled={stepIndex === 0}
              >
                Back
              </Button>
              {stepIndex >= stepCount - 1 ? (
                <Button className="min-w-24" size="sm" onClick={closeTutorial}>
                  Finish
                </Button>
              ) : (
                <Button className="min-w-24" size="sm" onClick={nextStep}>
                  Next
                </Button>
              )}
            </div>
          </div>
        </div>
      ) : null}
    </TutorialContext.Provider>
  );
}

export function useTutorial() {
  const context = useContext(TutorialContext);
  if (context === undefined) {
    throw new Error('useTutorial must be used within a TutorialProvider');
  }

  return context;
}
