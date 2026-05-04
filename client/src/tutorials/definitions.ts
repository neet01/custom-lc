export type TutorialId = 'cortex-overview' | 'outlook-workspace' | 'admin-reporting';

export interface TutorialDefinitionContext {
  openPanel: (panelId: string) => void;
  expandSidebar: () => void;
  collapseSidebar: () => void;
  dispatch: (eventName: string) => void;
}

export interface TutorialStep {
  id: string;
  title: string;
  description: string;
  target?: string;
  beforeEnter?: (context: TutorialDefinitionContext) => void;
}

export interface TutorialDefinition {
  id: TutorialId;
  title: string;
  description: string;
  steps: TutorialStep[];
}

export function buildTutorialDefinitions(
  context: TutorialDefinitionContext,
): Record<TutorialId, TutorialDefinition> {
  return {
    'cortex-overview': {
      id: 'cortex-overview',
      title: 'Cortex overview',
      description: 'A short walkthrough of the main workspace entry points in Cortex.',
      steps: [
        {
          id: 'overview-intro',
          title: 'Cortex tutorials',
          description:
            'This walkthrough highlights the main workspace surfaces. Use Settings > Account > Start tutorial any time to reopen it.',
        },
        {
          id: 'overview-outlook',
          title: 'Outlook workspace',
          description:
            'Open Outlook from the sidebar to work with email, scheduling, and calendar actions in one place.',
          target: 'sidebar-outlook',
          beforeEnter: () => {
            context.expandSidebar();
          },
        },
        {
          id: 'overview-admin',
          title: 'Admin reporting',
          description:
            'Admin reporting surfaces token usage, request history, user balances, Outlook audit events, and issue reporting.',
          target: 'sidebar-admin-reporting',
          beforeEnter: () => {
            context.expandSidebar();
          },
        },
        {
          id: 'overview-account',
          title: 'Account menu',
          description:
            'Your account menu gives access to settings, balances, file management, and this tutorial launcher.',
          target: 'sidebar-account',
          beforeEnter: () => {
            context.expandSidebar();
          },
        },
      ],
    },
    'outlook-workspace': {
      id: 'outlook-workspace',
      title: 'Outlook workspace',
      description: 'Walk through the inbox-focused Outlook workspace and its AI actions.',
      steps: [
        {
          id: 'outlook-open',
          title: 'Open Outlook',
          description:
            'The Outlook workspace lives behind this sidebar entry. It replaces the normal chat surface while it is active.',
          target: 'sidebar-outlook',
          beforeEnter: () => {
            context.expandSidebar();
          },
        },
        {
          id: 'outlook-root',
          title: 'Workspace shell',
          description:
            'This is the Outlook workspace container. It keeps inbox and calendar workflows under a single application surface.',
          target: 'outlook-workspace',
          beforeEnter: () => {
            context.openPanel('outlook');
            context.dispatch('cortex:tutorial-open-outlook-inbox');
          },
        },
        {
          id: 'outlook-tabs',
          title: 'Workspace tabs',
          description:
            'Switch between Inbox and Calendar here. The current implementation is inbox-first, with calendar mutations already available.',
          target: 'outlook-workspace-tabs',
          beforeEnter: () => {
            context.openPanel('outlook');
          },
        },
        {
          id: 'outlook-toolbar',
          title: 'Inbox controls',
          description:
            'This row contains search, mailbox controls, bulk actions, and brief/summarization entry points.',
          target: 'outlook-inbox-toolbar',
          beforeEnter: () => {
            context.openPanel('outlook');
            context.dispatch('cortex:tutorial-open-outlook-inbox');
          },
        },
        {
          id: 'outlook-list',
          title: 'Message list',
          description:
            'The left column is the working inbox list. It supports focused/other/all views, compact density, and bulk selection.',
          target: 'outlook-message-list',
          beforeEnter: () => {
            context.openPanel('outlook');
            context.dispatch('cortex:tutorial-open-outlook-inbox');
          },
        },
        {
          id: 'outlook-viewer',
          title: 'Email viewer',
          description:
            'The right column is the message and thread viewer. This is where the selected email thread, draft visibility, and delete actions live.',
          target: 'outlook-email-viewer',
          beforeEnter: () => {
            context.openPanel('outlook');
            context.dispatch('cortex:tutorial-open-outlook-inbox');
          },
        },
      ],
    },
    'admin-reporting': {
      id: 'admin-reporting',
      title: 'Admin reporting',
      description: 'Review the full-page admin analytics workspace.',
      steps: [
        {
          id: 'admin-open',
          title: 'Open admin reporting',
          description:
            'Admins can use this workspace to review token consumption, user activity, balances, audit trails, and reported issues.',
          target: 'sidebar-admin-reporting',
          beforeEnter: () => {
            context.expandSidebar();
          },
        },
        {
          id: 'admin-root',
          title: 'Reporting workspace',
          description:
            'The page header and summary cards show the current reporting window, top-level token usage, and request volume.',
          target: 'admin-reporting-root',
          beforeEnter: () => {
            context.openPanel('admin-reporting');
          },
        },
        {
          id: 'admin-tabs',
          title: 'Reporting tabs',
          description:
            'Use the tabs to move between usage-by-user, recent requests, directory management, Outlook audit history, and reported issues.',
          target: 'admin-reporting-tabs',
          beforeEnter: () => {
            context.openPanel('admin-reporting');
          },
        },
      ],
    },
  };
}
