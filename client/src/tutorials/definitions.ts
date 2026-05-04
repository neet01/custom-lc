export type TutorialId = 'chat-agents' | 'outlook-analysis';

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
  placement?: 'auto' | 'left' | 'right' | 'top' | 'bottom' | 'center';
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
    'chat-agents': {
      id: 'chat-agents',
      title: 'MCP, Jira, Confluence, and chat',
      description:
        'Connect MCP-backed enterprise tools first, then use the main chat surface with Jira and Confluence agents.',
      steps: [
        {
          id: 'chat-intro',
          title: 'Enterprise tools first',
          description:
            'Before using Jira or Confluence through Cortex, users need to open the MCP Servers panel and confirm the relevant connection is configured there.',
          placement: 'center',
          beforeEnter: () => {
            context.openPanel('conversations');
            context.collapseSidebar();
          },
        },
        {
          id: 'mcp-sidebar',
          title: 'Open MCP Servers',
          description:
            'This sidebar entry opens the MCP Servers panel. Use it first when a Jira or Confluence integration needs to be connected or reconfigured.',
          target: 'sidebar-mcp-builder',
          placement: 'right',
          beforeEnter: () => {
            context.expandSidebar();
          },
        },
        {
          id: 'mcp-panel',
          title: 'Connect Jira and Confluence here',
          description:
            'This panel is the setup surface for MCP-backed tools. If Jira or Confluence is not already connected, configure it here before returning to chat.',
          target: 'mcp-builder-panel',
          placement: 'right',
          beforeEnter: () => {
            context.expandSidebar();
            context.openPanel('mcp-builder');
          },
        },
        {
          id: 'mcp-add-server',
          title: 'Add or edit a server',
          description:
            'Use this control to add a new MCP server entry. Existing Jira or Confluence server cards in this panel can also be reviewed and updated from here.',
          target: 'mcp-builder-add-server',
          placement: 'bottom',
          beforeEnter: () => {
            context.expandSidebar();
            context.openPanel('mcp-builder');
          },
        },
        {
          id: 'chat-model-selector',
          title: 'Model and agent selector',
          description:
            'After the MCP connection is in place, return to the main chat surface and use this selector to choose the active Jira or Confluence agent.',
          target: 'chat-model-selector',
          placement: 'bottom',
          beforeEnter: () => {
            context.openPanel('conversations');
            context.collapseSidebar();
          },
        },
        {
          id: 'chat-composer',
          title: 'Chat composer',
          description:
            'This is the main composer shell. After you choose Jira or Confluence, type the request here and keep the task scoped to the agent you selected.',
          target: 'chat-composer',
          placement: 'top',
          beforeEnter: () => {
            context.openPanel('conversations');
            context.collapseSidebar();
          },
        },
        {
          id: 'chat-text-input',
          title: 'Request entry',
          description:
            'Enter the task in plain language. For example: summarize the latest Jira blockers, or search Confluence for deployment runbooks.',
          target: 'chat-text-input',
          placement: 'top',
          beforeEnter: () => {
            context.openPanel('conversations');
            context.collapseSidebar();
          },
        },
        {
          id: 'chat-send',
          title: 'Submit the request',
          description:
            'When the prompt is ready, submit it here. The current agent or model selection controls how Cortex routes and answers the request.',
          target: 'chat-send-button',
          placement: 'left',
          beforeEnter: () => {
            context.openPanel('conversations');
            context.collapseSidebar();
          },
        },
      ],
    },
    'outlook-analysis': {
      id: 'outlook-analysis',
      title: 'Outlook and email analysis',
      description: 'Work through the Outlook inbox flow, including AI analysis and reply actions.',
      steps: [
        {
          id: 'outlook-open',
          title: 'Open Outlook',
          description:
            'The Outlook workspace lives behind this sidebar entry. It replaces the normal chat surface while it is active.',
          target: 'sidebar-outlook',
          placement: 'right',
          beforeEnter: () => {
            context.expandSidebar();
          },
        },
        {
          id: 'outlook-tabs',
          title: 'Outlook workspace tabs',
          description:
            'Use these tabs to move between Inbox and Calendar. The analysis flow starts from Inbox after selecting an email thread.',
          target: 'outlook-workspace-tabs',
          placement: 'bottom',
          beforeEnter: () => {
            context.openPanel('outlook');
            context.dispatch('cortex:tutorial-open-outlook-inbox');
          },
        },
        {
          id: 'outlook-list',
          title: 'Choose an email thread',
          description:
            'Start by selecting the email or thread you want to inspect. The right-hand viewer and the AI assistant operate on the current selection.',
          target: 'outlook-example-thread',
          placement: 'right',
          beforeEnter: () => {
            context.openPanel('outlook');
            context.dispatch('cortex:tutorial-open-outlook-inbox');
          },
        },
        {
          id: 'outlook-viewer',
          title: 'Review the selected thread',
          description:
            'The message viewer shows the selected thread, any already-saved drafts for that conversation, and the raw context the AI actions will reference.',
          target: 'outlook-email-viewer',
          placement: 'left',
          beforeEnter: () => {
            context.openPanel('outlook');
            context.dispatch('cortex:tutorial-open-outlook-inbox');
          },
        },
        {
          id: 'outlook-ai-toggle',
          title: 'Open the AI assistant',
          description:
            'Use this button to open the floating AI assistant panel for the selected email thread.',
          target: 'outlook-ai-toggle',
          placement: 'left',
          beforeEnter: () => {
            context.openPanel('outlook');
            context.dispatch('cortex:tutorial-open-outlook-inbox');
            context.dispatch('cortex:tutorial-open-outlook-assistant');
          },
        },
        {
          id: 'outlook-ai-actions',
          title: 'Analysis and reply actions',
          description:
            'This panel is where users analyze the selected email, generate a reply draft, or find meeting times. Start with Analyze email when the user needs a structured readout before responding.',
          target: 'outlook-ai-actions',
          placement: 'left',
          beforeEnter: () => {
            context.openPanel('outlook');
            context.dispatch('cortex:tutorial-open-outlook-inbox');
            context.dispatch('cortex:tutorial-open-outlook-assistant');
          },
        },
        {
          id: 'outlook-analyze-button',
          title: 'Run email analysis',
          description:
            'This is the entry point for the analysis flow. It produces the AI summary and action framing for the selected thread, without requiring the user to write a prompt manually.',
          target: 'outlook-ai-analyze',
          placement: 'left',
          beforeEnter: () => {
            context.openPanel('outlook');
            context.dispatch('cortex:tutorial-open-outlook-inbox');
            context.dispatch('cortex:tutorial-open-outlook-assistant');
          },
        },
        {
          id: 'outlook-ai-results',
          title: 'Review results in-panel',
          description:
            'After analysis runs, the result stays in this scrollable assistant panel so users can review it, draft a response, or continue with scheduling actions against the same thread.',
          target: 'outlook-ai-panel',
          placement: 'left',
          beforeEnter: () => {
            context.openPanel('outlook');
            context.dispatch('cortex:tutorial-open-outlook-inbox');
            context.dispatch('cortex:tutorial-open-outlook-assistant');
          },
        },
      ],
    },
  };
}
