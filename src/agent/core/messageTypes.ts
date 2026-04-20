export type AgentMessageRole = 'system' | 'user' | 'assistant' | 'tool'

export type AgentMessageKind =
  | 'text'
  | 'status'
  | 'tool_call'
  | 'tool_result'
  | 'attachment'

export type AgentMessageOrigin =
  | 'legacy_ai_context'
  | 'legacy_chat_panel'
  | 'runtime'
  | 'tool'

export interface AgentMessageMetadata {
  turnId?: string
  toolId?: string
  attachmentType?: string
  origin?: AgentMessageOrigin
}

export interface AgentRuntimeMessage {
  id: string
  role: AgentMessageRole
  kind: AgentMessageKind
  content: string
  createdAt: Date
  metadata?: AgentMessageMetadata
}

