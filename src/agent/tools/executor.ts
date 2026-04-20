import type { AgentToolDefinition } from './contracts'
import {
  createToolCallIR,
  createToolErrorIR,
  createToolProgressIR,
  createToolResultIR,
  type ToolCallIR,
  type ToolExecutionPipelineState,
  type ToolProgressIR,
  type ToolResultIR,
} from './ir'
import { AgentToolRegistry } from './registry'
import type { AgentToolExecutionResult } from './results'
import { AgentToolScheduler } from './scheduler'

export type ExecutableAgentTool = AgentToolDefinition<
  Record<string, string>,
  AgentToolExecutionResult
>

export class AgentToolExecutor {
  private registry: AgentToolRegistry
  private scheduler: AgentToolScheduler
  private lastProgressEvents: ToolProgressIR[] = []
  private lastPipelineState: ToolExecutionPipelineState | null = null

  constructor(options?: {
    registry?: AgentToolRegistry
    scheduler?: AgentToolScheduler
  }) {
    this.registry = options?.registry ?? new AgentToolRegistry()
    this.scheduler = options?.scheduler ?? new AgentToolScheduler()
  }

  register(tool: ExecutableAgentTool): void {
    this.registry.register(tool)
  }

  registerMany(tools: readonly ExecutableAgentTool[]): void {
    tools.forEach((tool) => this.register(tool))
  }

  has(id: string): boolean {
    return this.registry.has(id)
  }

  get(id: string): ExecutableAgentTool | undefined {
    return this.registry.get(id) as ExecutableAgentTool | undefined
  }

  listDefinitions(): ExecutableAgentTool[] {
    return this.registry.list() as ExecutableAgentTool[]
  }

  async executeCall(
    call: ToolCallIR,
  ): Promise<{
    state: ToolExecutionPipelineState
    result: ToolResultIR
  } | null> {
    const resolved = this.registry.resolve(call.toolName)
    if (!resolved) {
      const error = createToolErrorIR({
        toolCallId: call.toolCallId,
        toolName: call.toolName,
        code: 'tool_not_found',
        message: `Tool not found: ${call.toolName}`,
      })
      const result = createToolResultIR({
        toolCallId: call.toolCallId,
        toolName: call.toolName,
        success: false,
        message: error.message,
        error,
      })
      const state = {
        call: { ...call, status: 'failed', finishedAt: result.timestamp },
        context: {},
        progress: [],
        result,
      }
      this.lastPipelineState = state
      return { state, result }
    }

    const tool = resolved.definition as ExecutableAgentTool
    const normalizeHandlerInput = (input: Record<string, unknown>): Record<string, string> =>
      Object.fromEntries(
        Object.entries(input || {}).map(([key, value]) => {
          if (value === null || value === undefined) return [key, '']
          if (typeof value === 'string') return [key, value]
          if (typeof value === 'number' || typeof value === 'boolean') {
            return [key, String(value)]
          }
          try {
            return [key, JSON.stringify(value)]
          } catch {
            return [key, String(value)]
          }
        }),
      )
    const normalizedCall =
      resolved.id === call.toolName
        ? call
        : {
            ...call,
            toolName: resolved.id,
            metadata: {
              ...(call.metadata || {}),
              requestedToolName: call.toolName,
              aliasResolvedFrom: call.toolName,
            },
          }

    const scheduled = this.scheduler.schedule(normalizedCall, tool)
    const normalizedInput = normalizeHandlerInput(
      scheduled.call.input as Record<string, unknown>,
    )
    const progress: ToolProgressIR[] = []
    const pushProgress = (status: ToolProgressIR['status'], message: string) => {
      const event = createToolProgressIR({
        toolCallId: scheduled.call.toolCallId,
        toolName: scheduled.call.toolName,
        status,
        message,
        payload: { queue: scheduled.queue },
      })
      progress.push(event)
      this.lastProgressEvents.push(event)
    }

    const executionContext = {
      turnId: scheduled.call.turnId,
    }

    if (tool.validateInput) {
      const validation = tool.validateInput(
        normalizedInput,
        executionContext,
      )
      if (!validation.valid) {
        const error = createToolErrorIR({
          toolCallId: scheduled.call.toolCallId,
          toolName: scheduled.call.toolName,
          code: 'validation_failed',
          message: validation.reason || 'Tool input validation failed',
        })
        const result = createToolResultIR({
          toolCallId: scheduled.call.toolCallId,
          toolName: scheduled.call.toolName,
          success: false,
          message: error.message,
          error,
        })
        const state = {
          call: { ...scheduled.call, status: 'failed', finishedAt: result.timestamp },
          context: executionContext,
          progress,
          result,
        }
        this.lastPipelineState = state
        return { state, result }
      }
    }

    if (tool.checkPermissions) {
      const permission = await tool.checkPermissions(
        normalizedInput,
        executionContext,
      )
      if (!permission.allowed) {
        const error = createToolErrorIR({
          toolCallId: scheduled.call.toolCallId,
          toolName: scheduled.call.toolName,
          code: 'permission_denied',
          message: permission.reason || 'Tool permission denied',
        })
        const result = createToolResultIR({
          toolCallId: scheduled.call.toolCallId,
          toolName: scheduled.call.toolName,
          success: false,
          message: error.message,
          error,
        })
        const state = {
          call: { ...scheduled.call, status: 'failed', finishedAt: result.timestamp },
          context: executionContext,
          progress,
          result,
        }
        this.lastPipelineState = state
        return { state, result }
      }
    }

    pushProgress('queued', `Queued ${scheduled.call.toolName}`)
    pushProgress('running', `Executing ${scheduled.call.toolName}`)

    try {
      const output = await tool.handler?.(
        normalizedInput,
        executionContext,
      )
      const result = createToolResultIR({
        toolCallId: scheduled.call.toolCallId,
        toolName: scheduled.call.toolName,
        success: output?.success ?? false,
        message: output?.message || '',
        payload: output?.data,
      })
      const state = {
        call: {
          ...scheduled.call,
          status: result.success ? 'completed' : 'failed',
          finishedAt: result.timestamp,
        },
        context: executionContext,
        progress,
        result,
      }
      this.lastPipelineState = state
      return { state, result }
    } catch (error) {
      const toolError = createToolErrorIR({
        toolCallId: scheduled.call.toolCallId,
        toolName: scheduled.call.toolName,
        code: 'execution_failed',
        message: (error as Error).message || String(error),
      })
      const result = createToolResultIR({
        toolCallId: scheduled.call.toolCallId,
        toolName: scheduled.call.toolName,
        success: false,
        message: toolError.message,
        error: toolError,
      })
      const state = {
        call: { ...scheduled.call, status: 'failed', finishedAt: result.timestamp },
        context: executionContext,
        progress,
        result,
      }
      this.lastPipelineState = state
      return { state, result }
    }
  }

  async execute(
    id: string,
    args: Record<string, string>,
  ): Promise<AgentToolExecutionResult | null> {
    const resolved = this.registry.resolve(id)
    if (!resolved) return null

    const tool = resolved.definition as ExecutableAgentTool

    const call = createToolCallIR({
      toolName: resolved.id,
      input: args,
      source: 'synthetic',
      metadata:
        resolved.id === id
          ? undefined
          : {
              requestedToolName: id,
              aliasResolvedFrom: id,
            },
      domain: tool.domain,
      mutation: tool.mutation,
      concurrency: tool.concurrency,
    })

    const executed = await this.executeCall(call)
    if (!executed) return null

    const { result } = executed
    return {
      tool: result.toolName,
      success: result.success,
      message: result.message,
      data: result.payload,
    }
  }

  getLastProgressEvents(): ToolProgressIR[] {
    return [...this.lastProgressEvents]
  }

  getLastPipelineState(): ToolExecutionPipelineState | null {
    return this.lastPipelineState
  }

  snapshot() {
    return {
      count: this.registry.snapshot().count,
      ids: this.registry.ids(),
      progressEvents: this.lastProgressEvents.length,
    }
  }
}
