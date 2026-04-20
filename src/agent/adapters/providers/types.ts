import type { ToolCallIR } from '../../tools/ir'
import type { MessageContent, ToolResult } from '../../core/runtimeTypes'

export type ProviderKind =
  | 'openai_compatible'
  | 'anthropic_messages'
  | 'legacy_text'

export interface ProviderCapabilityMatrix {
  provider: ProviderKind
  supportsNativeToolUse: boolean
  supportsReasoning: boolean
  supportsMultimodal: boolean
  supportsPromptCache: boolean
  supportsDeferredTools: boolean
  supportsStructuredToolSchema: boolean
}

export interface ProviderRequestEnvelope {
  model: string
  messages: Array<{ role: string; content: MessageContent }>
  tools?: unknown[]
}

export interface ProviderConversationMessage {
  role: string
  content: MessageContent
  nativePayload?: Record<string, unknown>
}

export interface ProviderToolResultBinding {
  call: ToolCallIR
  result: ToolResult
}

export interface ProviderToolCallAdapter {
  provider: ProviderKind
  capabilities: ProviderCapabilityMatrix
  toToolCalls?: (response: unknown) => ToolCallIR[]
  toAssistantConversationMessage?: (
    response: unknown,
  ) => ProviderConversationMessage | null
  fromToolResults?: (
    bindings: ProviderToolResultBinding[],
  ) => ProviderConversationMessage[]
}
