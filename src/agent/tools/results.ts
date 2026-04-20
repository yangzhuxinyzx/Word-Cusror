export interface AgentToolExecutionResult {
  tool: string
  success: boolean
  message: string
  data?: Record<string, unknown>
}

export type { ToolErrorIR, ToolProgressIR, ToolResultIR } from './ir'
