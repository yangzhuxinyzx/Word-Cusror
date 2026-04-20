export type AgentAttachmentScope = 'session' | 'turn' | 'document' | 'workspace'

export type AgentAttachmentKind = 'text' | 'json' | 'reference'

export interface AgentAttachmentDefinition {
  type: string
  scope: AgentAttachmentScope
  kind: AgentAttachmentKind
  description: string
  volatile?: boolean
}

export interface AgentAttachment<TPayload = unknown> {
  id: string
  type: string
  scope: AgentAttachmentScope
  kind: AgentAttachmentKind
  payload: TPayload
}

export const CORE_ATTACHMENT_DEFINITIONS: readonly AgentAttachmentDefinition[] = [
  {
    type: 'active_skill',
    scope: 'turn',
    kind: 'json',
    description: 'Active skill selected for the current turn',
    volatile: true,
  },
  {
    type: 'skill_descriptions',
    scope: 'session',
    kind: 'text',
    description: 'Visible skill descriptions and activation commands',
    volatile: true,
  },
  {
    type: 'current_document_summary',
    scope: 'document',
    kind: 'text',
    description: 'Summary of the active document for the current turn',
    volatile: true,
  },
  {
    type: 'document_structure_delta',
    scope: 'document',
    kind: 'text',
    description: 'Structured delta for the active document outline',
    volatile: true,
  },
  {
    type: 'available_tools_delta',
    scope: 'turn',
    kind: 'json',
    description: 'Tool visibility delta for the current turn',
    volatile: true,
  },
  {
    type: 'workspace_context_delta',
    scope: 'workspace',
    kind: 'text',
    description: 'Incremental workspace context surfaced to the runtime',
    volatile: true,
  },
  {
    type: 'workspace_profile',
    scope: 'workspace',
    kind: 'json',
    description: 'Structured workspace profile captured by init-style scanning',
    volatile: true,
  },
  {
    type: 'relevant_memories',
    scope: 'turn',
    kind: 'json',
    description: 'Relevant memory snippets injected for the current turn',
    volatile: true,
  },
  {
    type: 'ppt_edit_context',
    scope: 'document',
    kind: 'json',
    description: 'PPT editing context for the active slide selection',
    volatile: true,
  },
  {
    type: 'excel_sheet_context',
    scope: 'document',
    kind: 'json',
    description: 'Excel worksheet context for the active sheet',
    volatile: true,
  },
] as const
