export type AgentToolDomain =
  | 'core'
  | 'word'
  | 'workspace'
  | 'ppt'
  | 'excel'
  | 'web'
  | 'memory'
  | 'knowledge'
  | 'legacy'

export type AgentToolMutation =
  | 'read'
  | 'write'
  | 'create'
  | 'delete'
  | 'transform'
  | 'external'

export type AgentToolConcurrencyPolicy = 'serial' | 'parallel_safe'

export interface AgentToolExecutionContext {
  turnId?: string
  documentId?: string
  workspacePath?: string
}

export type AgentToolHandler<
  TInput extends Record<string, unknown> = Record<string, unknown>,
  TOutput = unknown,
> = (input: TInput, context: AgentToolExecutionContext) => Promise<TOutput>

export type AgentToolInterruptBehavior = 'cancel' | 'block'

export interface AgentToolDefinition<
  TInput extends Record<string, unknown> = Record<string, unknown>,
  TOutput = unknown,
> {
  id: string
  displayName: string
  description: string
  domain: AgentToolDomain
  mutation: AgentToolMutation
  concurrency: AgentToolConcurrencyPolicy
  tags?: string[]
  inputKeys?: string[]
  inputSchema?: Record<string, unknown>
  outputKeys?: string[]
  legacyAliases?: string[]
  prompt?: string
  validateInput?: (
    input: TInput,
    context: AgentToolExecutionContext,
  ) => { valid: boolean; reason?: string }
  checkPermissions?: (
    input: TInput,
    context: AgentToolExecutionContext,
  ) => Promise<{ allowed: boolean; reason?: string }>
  isReadOnly?: boolean
  isDestructive?: boolean
  interruptBehavior?: AgentToolInterruptBehavior
  renderToolUse?: (input: TInput) => string
  renderToolResult?: (output: TOutput) => string
  toTelemetryPayload?: (
    input: TInput,
    output?: TOutput,
  ) => Record<string, unknown>
  handler?: AgentToolHandler<TInput, TOutput>
}

export function defineAgentTool<
  TInput extends Record<string, unknown> = Record<string, unknown>,
  TOutput = unknown,
>(
  definition: AgentToolDefinition<TInput, TOutput>,
): AgentToolDefinition<TInput, TOutput> {
  return definition
}

export function isMutationTool(definition: AgentToolDefinition): boolean {
  return definition.mutation !== 'read' && definition.mutation !== 'external'
}
