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
type StepPlacement = TutorialStep['placement'];
type HighlightCleanup = () => void;

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

function applyTargetHighlight(element: HTMLElement): HighlightCleanup {
  const style = element.style;
  const computedStyle = window.getComputedStyle(element);
  const previous = {
    position: style.position,
    zIndex: style.zIndex,
    outline: style.outline,
    outlineOffset: style.outlineOffset,
    boxShadow: style.boxShadow,
    transition: style.transition,
    scrollMargin: style.scrollMargin,
  };

  if (computedStyle.position === 'static') {
    style.position = 'relative';
  }
  style.zIndex = '141';
  style.outline = '3px solid rgba(245, 208, 0, 0.88)';
  style.outlineOffset = '4px';
  style.boxShadow = '0 0 24px rgba(245, 208, 0, 0.22)';
  style.transition = [style.transition, 'outline-color 160ms ease, box-shadow 160ms ease']
    .filter(Boolean)
    .join(', ');
  style.scrollMargin = '120px';

  element.setAttribute('data-tutorial-active', 'true');

  return () => {
    style.position = previous.position;
    style.zIndex = previous.zIndex;
    style.outline = previous.outline;
    style.outlineOffset = previous.outlineOffset;
    style.boxShadow = previous.boxShadow;
    style.transition = previous.transition;
    style.scrollMargin = previous.scrollMargin;
    element.removeAttribute('data-tutorial-active');
  };
}

function getCardPosition(rect: RectLike | null, placement: StepPlacement = 'auto') {
  const cardWidth = Math.min(380, Math.max(320, window.innerWidth - 32));
  const cardHeight = 280;
  const viewportWidth = window.innerWidth;
  const viewportHeight = window.innerHeight;
  const gap = 20;

  if (placement === 'center') {
    return {
      top: Math.max(24, viewportHeight / 2 - cardHeight / 2),
      left: Math.max(16, viewportWidth / 2 - cardWidth / 2),
      width: cardWidth,
    };
  }

  if (!rect) {
    return {
      top: Math.max(24, viewportHeight / 2 - cardHeight / 2),
      left: Math.max(16, viewportWidth / 2 - cardWidth / 2),
      width: cardWidth,
    };
  }

  const placements =
    placement && placement !== 'auto'
      ? [placement]
      : (['right', 'left', 'bottom', 'top'] as const);

  const rightPlacement = rect.left + rect.width + gap;
  const leftPlacement = rect.left - gap - cardWidth;
  const belowPlacement = rect.top + rect.height + gap;
  const abovePlacement = rect.top - cardHeight - gap;

  for (const nextPlacement of placements) {
    if (nextPlacement === 'right' && rightPlacement + cardWidth <= viewportWidth - 16) {
      return {
        top: Math.min(Math.max(24, rect.top), viewportHeight - cardHeight - 24),
        left: rightPlacement,
        width: cardWidth,
      };
    }

    if (nextPlacement === 'left' && leftPlacement >= 16) {
      return {
        top: Math.min(Math.max(24, rect.top), viewportHeight - cardHeight - 24),
        left: leftPlacement,
        width: cardWidth,
      };
    }

    if (nextPlacement === 'bottom' && belowPlacement + cardHeight <= viewportHeight - 24) {
      return {
        top: belowPlacement,
        left: Math.min(Math.max(16, rect.left), viewportWidth - cardWidth - 16),
        width: cardWidth,
      };
    }

    if (nextPlacement === 'top' && abovePlacement >= 24) {
      return {
        top: abovePlacement,
        left: Math.min(Math.max(16, rect.left), viewportWidth - cardWidth - 16),
        width: cardWidth,
      };
    }
  }

  return {
    top: Math.max(24, viewportHeight / 2 - cardHeight / 2),
    left: Math.max(16, viewportWidth / 2 - cardWidth / 2),
    width: cardWidth,
  };
}

export function TutorialProvider({ children }: { children: ReactNode }) {
  const { setActive } = useActivePanel();
  const setSidebarExpanded = useSetRecoilState(store.sidebarExpanded);
  const [activeTutorialId, setActiveTutorialId] = useState<TutorialId | null>(null);
  const [stepIndex, setStepIndex] = useState(0);
  const [targetRect, setTargetRect] = useState<RectLike | null>(null);
  const activeStepCleanupRef = useRef<number | null>(null);
  const targetHighlightCleanupRef = useRef<HighlightCleanup | null>(null);

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
    targetHighlightCleanupRef.current?.();
    targetHighlightCleanupRef.current = null;
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
    targetHighlightCleanupRef.current?.();
    targetHighlightCleanupRef.current = null;

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
        targetHighlightCleanupRef.current = applyTargetHighlight(element);
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
      if (activeStepCleanupRef.current != null) {
        window.clearTimeout(activeStepCleanupRef.current);
        activeStepCleanupRef.current = null;
      }
      targetHighlightCleanupRef.current?.();
      targetHighlightCleanupRef.current = null;
      setTargetRect(null);
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
    typeof window !== 'undefined'
      ? getCardPosition(targetRect, activeStep?.placement)
      : { top: 24, left: 24, width: 340 };

  return (
    <TutorialContext.Provider value={value}>
      {children}
      {activeTutorial && activeStep ? (
        <div className="pointer-events-auto fixed inset-0 z-[140]">
          <div className="pointer-events-none absolute inset-0 bg-black/60" />
          <div
            className={cn(
              'absolute z-[142] rounded-2xl border border-border-medium bg-surface-primary p-4 shadow-2xl',
            )}
            style={{
              top: `${cardPosition.top}px`,
              left: `${cardPosition.left}px`,
              width: `${cardPosition.width}px`,
              maxWidth: 'calc(100vw - 2rem)',
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

            <div className="mt-4">
              <div className="mb-2 flex items-center justify-between text-xs text-text-secondary">
                <span>{activeTutorial.title}</span>
                <span>
                  {currentStepNumber} / {stepCount}
                </span>
              </div>
              <div className="h-1.5 overflow-hidden rounded-full bg-surface-secondary">
                <div
                  className="h-full rounded-full bg-[#f5d000] transition-[width] duration-200"
                  style={{
                    width: `${stepCount > 0 ? (currentStepNumber / stepCount) * 100 : 0}%`,
                  }}
                />
              </div>
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
