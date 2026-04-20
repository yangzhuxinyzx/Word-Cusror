export type AgentRuntimePhase =
  | 'idle'
  | 'awaiting_model'
  | 'streaming'
  | 'executing_tools'
  | 'completed'
  | 'errored'

export interface AgentConversationState {
  phase: AgentRuntimePhase
  currentTurnId: string | null
  activeToolIds: string[]
  pendingAttachmentTypes: string[]
  lastError: string | null
}

export function createEmptyConversationState(): AgentConversationState {
  return {
    phase: 'idle',
    currentTurnId: null,
    activeToolIds: [],
    pendingAttachmentTypes: [],
    lastError: null,
  }
}

