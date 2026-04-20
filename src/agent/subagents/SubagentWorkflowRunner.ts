import type { AgentRuntimeMessage } from '../core/messageTypes'
import type { ToolResult } from '../core/runtimeTypes'
import type { ToolExecutionPipelineState } from '../tools/ir'
import type { SubagentRecord } from './SubagentManager'
import type { SubagentProfileDefinition } from './SubagentProfiles'

export interface ParsedSubagentCommand {
  profileId: string
  mode?: 'sync' | 'background'
  request: string
}

export interface SubagentToolExecution {
  result: ToolResult
  pipelineState: ToolExecutionPipelineState | null
}

export interface RunSubagentWorkflowParams {
  profile: SubagentProfileDefinition
  request: string
  mode?: 'sync' | 'background'
  parentTurnId?: string | null
  documentContext?: string
  filesContext?: string
  workspaceProfile?: string
  images?: string[]
}

export interface RunSubagentWorkflowResult {
  subagent: SubagentRecord
  summary: string
  toolResults: ToolResult[]
}

export interface SubagentWorkflowRuntime {
  spawnSubagent: (params: {
    profileId: string
    label?: string
    mode?: 'sync' | 'background'
    parentTurnId?: string | null
  }) => SubagentRecord
  startSubagent: (subagentId: string) => SubagentRecord | null
  completeSubagent: (params: {
    subagentId: string
    summary?: string
    outputPath?: string
  }) => SubagentRecord | null
  failSubagent: (params: {
    subagentId: string
    error: string
  }) => SubagentRecord | null
  appendSubagentTranscript: (params: {
    subagentId: string
    messages?: AgentRuntimeMessage[]
    toolCalls?: ToolExecutionPipelineState['call'][]
    toolProgress?: ToolExecutionPipelineState['progress']
    toolResults?: NonNullable<ToolExecutionPipelineState['result']>[]
  }) => void
  invokeTool: (
    tool: string,
    args: Record<string, string>,
  ) => Promise<SubagentToolExecution>
  runSession?: (params: {
    content: string
    systemPrompt?: string
    documentContext?: string
    filesContext?: string
    callbacks?: {
      onToolCall?: (tool: string, args: Record<string, string>) => Promise<ToolResult>
      onToolCallStart?: (tool: string) => void
      onToolCallPreview?: (tool: string, args: Record<string, string>) => void
      onToolCallSkipped?: (
        tool: string,
        args: Record<string, string>,
        reason: string,
      ) => void
    }
    memoryContext?: {
      workspaceProfile?: string
      activeSkill?: Record<string, unknown>
      availableToolsDelta?: Record<string, unknown>
    }
    images?: string[]
  }) => Promise<{
    finalContent: string
    toolResults: ToolResult[]
    reasoning?: string
    iteration: number
  }>
}

function createTranscriptMessage(
  role: AgentRuntimeMessage['role'],
  content: string,
  metadata?: AgentRuntimeMessage['metadata'],
): AgentRuntimeMessage {
  return {
    id: `msg-${Date.now()}-${Math.random().toString(16).slice(2)}`,
    role,
    kind: 'text',
    content,
    createdAt: new Date(),
    metadata,
  }
}

export function parseSubagentCommand(input: string): ParsedSubagentCommand | null {
  const trimmed = input.trim()
  if (!trimmed.startsWith('/')) return null

  const aliasCommands: Record<string, ParsedSubagentCommand> = {
    '/workspace-explore': {
      profileId: 'workspace-explore',
      request: '',
    },
    '/doc-explore': {
      profileId: 'doc-explore',
      request: '',
    },
    '/verification': {
      profileId: 'verification',
      request: '',
    },
    '/verify': {
      profileId: 'verification',
      request: '',
    },
    '/doc-editor': {
      profileId: 'doc-editor',
      request: '',
    },
    '/ppt-builder': {
      profileId: 'ppt-builder',
      request: '',
    },
    '/excel-operator': {
      profileId: 'excel-operator',
      request: '',
    },
    '/bg-workspace-explore': {
      profileId: 'workspace-explore',
      mode: 'background',
      request: '',
    },
    '/bg-ppt-builder': {
      profileId: 'ppt-builder',
      mode: 'background',
      request: '',
    },
  }

  const firstSpace = trimmed.indexOf(' ')
  const command = (firstSpace >= 0 ? trimmed.slice(0, firstSpace) : trimmed).toLowerCase()
  const rest = firstSpace >= 0 ? trimmed.slice(firstSpace + 1).trim() : ''

  if (aliasCommands[command]) {
    return {
      ...aliasCommands[command],
      request: rest,
    }
  }

  if (command === '/subagent' || command === '/bg-subagent') {
    if (!rest) return null
    const profileSpace = rest.indexOf(' ')
    const profileId = (profileSpace >= 0 ? rest.slice(0, profileSpace) : rest).trim()
    const request = profileSpace >= 0 ? rest.slice(profileSpace + 1).trim() : ''
    if (!profileId) return null
    return {
      profileId,
      mode: command === '/bg-subagent' ? 'background' : 'sync',
      request,
    }
  }

  return null
}

function buildWorkflowSteps(
  profileId: string,
  request: string,
): Array<{ tool: string; args: Record<string, string> }> {
  if (profileId === 'workspace-explore') {
    return [
      {
        tool: 'workspace_profile',
        args: { refresh: 'true' },
      },
      {
        tool: 'workspace_list',
        args: { refresh: 'false' },
      },
    ]
  }

  if (profileId === 'doc-explore') {
    return [
      {
        tool: 'word.read',
        args: { target: 'outline' },
      },
      {
        tool: 'word.read',
        args: { target: 'selection' },
      },
    ]
  }

  if (profileId === 'verification') {
    return [
      {
        tool: 'workspace_profile',
        args: { refresh: 'false' },
      },
      {
        tool: 'word.read',
        args: { target: 'outline' },
      },
      {
        tool: 'word.read',
        args: { target: 'selection' },
      },
    ]
  }

  if (profileId === 'excel-operator') {
    return [
      {
        tool: 'excel_read',
        args: {
          sheet: 'Sheet1',
          range: 'A1:E10',
        },
      },
    ]
  }

  return [
    {
      tool: 'workspace_profile',
      args: { refresh: 'false' },
    },
  ]
}

function formatSubagentSummary(
  profile: SubagentProfileDefinition,
  request: string,
  toolResults: ToolResult[],
): string {
  const lines = toolResults.map((result) => {
    const status = result.success ? 'success' : 'failed'
    return `- ${result.tool}: ${status}\n  ${result.message}`
  })

  return [
    `Subagent ${profile.displayName} completed.`,
    request ? `Request: ${request}` : '',
    lines.length > 0 ? `Results:\n${lines.join('\n')}` : 'No tool results were produced.',
  ]
    .filter(Boolean)
    .join('\n\n')
}

export class SubagentWorkflowRunner {
  constructor(private readonly runtime: SubagentWorkflowRuntime) {}

  private shouldUseModelSession(profile: SubagentProfileDefinition): boolean {
    return true
  }

  async run(
    params: RunSubagentWorkflowParams,
  ): Promise<RunSubagentWorkflowResult> {
    const subagent = this.runtime.spawnSubagent({
      profileId: params.profile.id,
      label: params.profile.displayName,
      mode: params.mode || params.profile.defaultMode,
      parentTurnId: params.parentTurnId,
    })

    this.runtime.appendSubagentTranscript({
      subagentId: subagent.subagentId,
      messages: [
        createTranscriptMessage(
          'user',
          params.request || `Run subagent profile ${params.profile.id}`,
          {
            origin: 'runtime',
            turnId: params.parentTurnId || undefined,
          },
        ),
      ],
    })

    this.runtime.startSubagent(subagent.subagentId)

    const toolResults: ToolResult[] = []

    try {
      if (this.runtime.runSession && this.shouldUseModelSession(params.profile)) {
        const sessionResult = await this.runtime.runSession({
          content: [
            `[Subagent Profile] ${params.profile.displayName} (${params.profile.id})`,
            params.profile.prompt,
            params.request ? `User request:\n${params.request}` : '',
          ]
            .filter(Boolean)
            .join('\n\n'),
          systemPrompt: [
            'You are running as an isolated subagent.',
            `Profile: ${params.profile.displayName}`,
            `Safety: ${params.profile.safety}`,
            params.profile.allowedToolIds?.length
              ? `Allowed tools: ${params.profile.allowedToolIds.join(', ')}`
              : '',
            'Stay within your profile boundary and return a concise result.',
          ]
            .filter(Boolean)
            .join('\n\n'),
          documentContext: params.documentContext,
          filesContext: params.filesContext,
          callbacks: {
            onToolCall: async (tool, args) => {
              if (
                params.profile.allowedToolIds?.length &&
                !params.profile.allowedToolIds.includes(tool)
              ) {
                return {
                  tool,
                  success: false,
                  message: `Tool ${tool} is outside profile boundary.`,
                }
              }
              const execution = await this.runtime.invokeTool(tool, args)
              if (execution.pipelineState) {
                this.runtime.appendSubagentTranscript({
                  subagentId: subagent.subagentId,
                  toolCalls: [execution.pipelineState.call],
                  toolProgress: execution.pipelineState.progress,
                  toolResults: execution.pipelineState.result
                    ? [execution.pipelineState.result]
                    : [],
                  messages: [
                    createTranscriptMessage(
                      'assistant',
                      `${tool}: ${execution.result.message}`,
                      {
                        origin: 'tool',
                        toolId: tool,
                        turnId: execution.pipelineState.call.turnId,
                      },
                    ),
                  ],
                })
              }
              return execution.result
            },
          },
          memoryContext: {
            workspaceProfile: params.workspaceProfile,
            activeSkill: {
              id: params.profile.id,
              displayName: params.profile.displayName,
              safety: params.profile.safety,
            },
            availableToolsDelta: params.profile.allowedToolIds?.length
              ? {
                  source: 'subagent_profile',
                  profileId: params.profile.id,
                  allowedToolIds: params.profile.allowedToolIds,
                }
              : undefined,
          },
          images: params.images,
        })

        toolResults.push(...sessionResult.toolResults)
        const summary = sessionResult.finalContent || formatSubagentSummary(
          params.profile,
          params.request,
          toolResults,
        )
        const completed =
          this.runtime.completeSubagent({
            subagentId: subagent.subagentId,
            summary,
          }) || subagent

        this.runtime.appendSubagentTranscript({
          subagentId: subagent.subagentId,
          messages: [
            createTranscriptMessage('assistant', summary, {
              origin: 'runtime',
              turnId: params.parentTurnId || undefined,
            }),
          ],
        })

        return {
          subagent: completed,
          summary,
          toolResults,
        }
      }

      const steps = buildWorkflowSteps(params.profile.id, params.request)
      for (const step of steps) {
        if (
          params.profile.allowedToolIds?.length &&
          !params.profile.allowedToolIds.includes(step.tool)
        ) {
          toolResults.push({
            tool: step.tool,
            success: false,
            message: `Skipped because tool ${step.tool} is outside profile boundary.`,
          })
          continue
        }

        const execution = await this.runtime.invokeTool(step.tool, step.args)
        toolResults.push(execution.result)

        if (execution.pipelineState) {
          this.runtime.appendSubagentTranscript({
            subagentId: subagent.subagentId,
            toolCalls: [execution.pipelineState.call],
            toolProgress: execution.pipelineState.progress,
            toolResults: execution.pipelineState.result
              ? [execution.pipelineState.result]
              : [],
            messages: [
              createTranscriptMessage(
                'assistant',
                `${step.tool}: ${execution.result.message}`,
                {
                  origin: 'tool',
                  toolId: step.tool,
                  turnId: execution.pipelineState.call.turnId,
                },
              ),
            ],
          })
        }
      }

      const summary = formatSubagentSummary(params.profile, params.request, toolResults)
      const completed =
        this.runtime.completeSubagent({
          subagentId: subagent.subagentId,
          summary,
        }) || subagent

      this.runtime.appendSubagentTranscript({
        subagentId: subagent.subagentId,
        messages: [
          createTranscriptMessage('assistant', summary, {
            origin: 'runtime',
            turnId: params.parentTurnId || undefined,
          }),
        ],
      })

      return {
        subagent: completed,
        summary,
        toolResults,
      }
    } catch (error) {
      const message = (error as Error).message || String(error)
      const failed =
        this.runtime.failSubagent({
          subagentId: subagent.subagentId,
          error: message,
        }) || subagent

      this.runtime.appendSubagentTranscript({
        subagentId: subagent.subagentId,
        messages: [
          createTranscriptMessage('assistant', `Subagent failed: ${message}`, {
            origin: 'runtime',
            turnId: params.parentTurnId || undefined,
          }),
        ],
      })

      throw Object.assign(new Error(message), {
        subagent: failed,
      })
    }
  }
}
