import type {
  AgentToolConcurrencyPolicy,
  AgentToolDomain,
  AgentToolExecutionContext,
  AgentToolMutation,
} from './contracts'

export type ToolCallSource = 'native' | 'legacy_bracket' | 'legacy_xml' | 'legacy_tool_use' | 'synthetic'

export type ToolCallStatus =
  | 'parsed'
  | 'validated'
  | 'scheduled'
  | 'executing'
  | 'completed'
  | 'failed'
  | 'cancelled'

export interface ToolCallIR {
  toolCallId: string
  toolName: string
  input: Record<string, unknown>
  status: ToolCallStatus
  source: ToolCallSource
  rawInput?: string
  turnId?: string
  startedAt?: string
  finishedAt?: string
  domain?: AgentToolDomain
  mutation?: AgentToolMutation
  concurrency?: AgentToolConcurrencyPolicy
  metadata?: Record<string, unknown>
}

export interface ToolProgressIR {
  toolCallId: string
  toolName: string
  status: 'queued' | 'running' | 'progress'
  message: string
  timestamp: string
  payload?: Record<string, unknown>
}

export interface ToolErrorIR {
  toolCallId: string
  toolName: string
  code:
    | 'tool_not_found'
    | 'validation_failed'
    | 'permission_denied'
    | 'execution_failed'
    | 'cancelled'
    | 'unknown'
  message: string
  timestamp: string
  details?: Record<string, unknown>
}

export interface ToolResultIR {
  toolCallId: string
  toolName: string
  success: boolean
  message: string
  timestamp: string
  payload?: Record<string, unknown>
  error?: ToolErrorIR
}

export interface ToolExecutionPipelineState {
  call: ToolCallIR
  context: AgentToolExecutionContext
  progress: ToolProgressIR[]
  result?: ToolResultIR
}

export function createToolCallIR(params: {
  toolName: string
  input: Record<string, unknown>
  source: ToolCallSource
  rawInput?: string
  turnId?: string
  metadata?: Record<string, unknown>
  domain?: AgentToolDomain
  mutation?: AgentToolMutation
  concurrency?: AgentToolConcurrencyPolicy
}): ToolCallIR {
  return {
    toolCallId: `tool-${Date.now()}-${Math.random().toString(16).slice(2)}`,
    toolName: params.toolName,
    input: { ...params.input },
    source: params.source,
    rawInput: params.rawInput,
    turnId: params.turnId,
    metadata: params.metadata ? { ...params.metadata } : undefined,
    status: 'parsed',
    startedAt: new Date().toISOString(),
    domain: params.domain,
    mutation: params.mutation,
    concurrency: params.concurrency,
  }
}

export function createToolProgressIR(params: {
  toolCallId: string
  toolName: string
  status: ToolProgressIR['status']
  message: string
  payload?: Record<string, unknown>
}): ToolProgressIR {
  return {
    toolCallId: params.toolCallId,
    toolName: params.toolName,
    status: params.status,
    message: params.message,
    timestamp: new Date().toISOString(),
    payload: params.payload ? { ...params.payload } : undefined,
  }
}

export function createToolErrorIR(params: {
  toolCallId: string
  toolName: string
  code: ToolErrorIR['code']
  message: string
  details?: Record<string, unknown>
}): ToolErrorIR {
  return {
    toolCallId: params.toolCallId,
    toolName: params.toolName,
    code: params.code,
    message: params.message,
    timestamp: new Date().toISOString(),
    details: params.details ? { ...params.details } : undefined,
  }
}

export function createToolResultIR(params: {
  toolCallId: string
  toolName: string
  success: boolean
  message: string
  payload?: Record<string, unknown>
  error?: ToolErrorIR
}): ToolResultIR {
  return {
    toolCallId: params.toolCallId,
    toolName: params.toolName,
    success: params.success,
    message: params.message,
    timestamp: new Date().toISOString(),
    payload: params.payload ? { ...params.payload } : undefined,
    error: params.error,
  }
}
