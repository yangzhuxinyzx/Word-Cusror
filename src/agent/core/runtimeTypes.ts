import type { ToolCallSource } from '../tools/ir'
import type { KnowledgeSearchResult } from '../../types'

export type MessageContent =
  | string
  | Array<
      | { type: 'text'; text: string }
      | { type: 'image_url'; image_url: { url: string } }
    >

export interface ToolResult {
  tool: string
  success: boolean
  message: string
  data?: Record<string, unknown>
}

export type AgentDebugEvent =
  | {
      type: 'turn_start'
      turnId: string
      timestamp: string
      model: string
      baseUrl: string
      userInput: string
      hasDocumentContext: boolean
      hasFilesContext: boolean
      imageCount: number
      recentMessagesCount: number
    }
  | {
      type: 'api_response_raw'
      turnId: string
      timestamp: string
      iteration: number
      stage: 'loop' | 'forced_summary'
      response: string
      rawResponse?: unknown
      hasToolCall: boolean
    }
  | {
      type: 'tool_calls_parsed'
      turnId: string
      timestamp: string
      iteration: number
      calls: Array<{
        tool: string
        args: Record<string, string>
        source?: ToolCallSource
      }>
    }
  | {
      type: 'tool_call_skipped'
      turnId: string
      timestamp: string
      iteration: number
      tool: string
      args: Record<string, string>
      reason: string
    }
  | {
      type: 'tool_result'
      turnId: string
      timestamp: string
      iteration: number
      index: number
      total: number
      tool: string
      args: Record<string, string>
      result: ToolResult
    }
  | {
      type: 'final_summary'
      turnId: string
      timestamp: string
      iteration: number
      source: 'normal' | 'forced_stop' | 'max_iterations'
      content: string
    }
  | {
      type: 'turn_complete'
      turnId: string
      timestamp: string
      totalIterations: number
      finalContent: string
      toolResults: ToolResult[]
    }
  | {
      type: 'turn_error'
      turnId: string
      timestamp: string
      iteration: number
      aborted: boolean
      name?: string
      message: string
      stack?: string
    }

export interface AgentCallbacks {
  onToolCall?: (tool: string, args: Record<string, string>) => Promise<ToolResult>
  onToolCallStart?: (tool: string) => void
  onToolCallPreview?: (tool: string, args: Record<string, string>) => void
  onToolCallSkipped?: (
    tool: string,
    args: Record<string, string>,
    reason: string,
  ) => void
  onTextChunk?: (text: string) => void
  onDebugEvent?: (event: AgentDebugEvent) => void | Promise<void>
  onContent?: (content: string) => void
  onComplete?: (
    content: string,
    toolResults: ToolResult[],
    reasoning?: string,
    meta?: { knowledgeHits?: KnowledgeSearchResult[] },
  ) => void
  onThinking?: (thinking: string) => void
  getLatestDocument?: () => string
}
