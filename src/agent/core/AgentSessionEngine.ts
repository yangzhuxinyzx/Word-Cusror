import type { ChatMessage } from '../../types'
import { AnthropicToolUseAdapter } from '../adapters/providers/AnthropicToolUseAdapter'
import { LegacyTextToolAdapter } from '../adapters/providers/LegacyTextToolAdapter'
import { OpenAICompatibleToolCallAdapter } from '../adapters/providers/OpenAICompatibleToolCallAdapter'
import type { ProviderToolCallAdapter } from '../adapters/providers/types'
import { HookRegistry } from '../hooks/HookRegistry'
import { ReplayStore } from '../storage/ReplayStore'
import { ResumeLoader } from '../storage/ResumeLoader'
import { SessionTranscriptStore } from '../storage/SessionTranscriptStore'
import { ToolResultStore } from '../storage/ToolResultStore'
import { SubagentManager, type SubagentRecord } from '../subagents/SubagentManager'
import type { SubagentProfileDefinition } from '../subagents/SubagentProfiles'
import { TaskNotificationCenter } from '../tasks/TaskNotifications'
import { TaskRegistry } from '../tasks/TaskRegistry'
import type { AgentSkillDefinition } from '../skills/SkillRegistry'
import type { AgentToolDefinition } from '../tools/contracts'
import { createLegacyAgentRuntime, type LegacyAgentRuntimeSnapshot } from '../compat/LegacyRuntime'
import {
  createEmptyConversationState,
  type AgentConversationState,
  type AgentRuntimePhase,
} from './ConversationState'
import type { AgentRuntimeMessage } from './messageTypes'
import type {
  ToolCallIR,
  ToolExecutionPipelineState,
  ToolProgressIR,
  ToolResultIR,
} from '../tools/ir'

export interface AgentSessionSnapshot {
  runtime: LegacyAgentRuntimeSnapshot
  conversation: AgentConversationState
  providers: {
    adapters: ProviderToolCallAdapter[]
  }
  skills: {
    count: number
    ids: string[]
  }
  messages: {
    count: number
    lastMessageId: string | null
    lastRole: AgentRuntimeMessage['role'] | null
  }
  tasks: {
    count: number
    running: number
    background: number
    subagents: number
  }
  notifications: {
    count: number
    items: Array<{
      taskId: string
      message: string
      createdAt: string
    }>
  }
  subagents: {
    profiles: SubagentProfileDefinition[]
    agents: SubagentRecord[]
  }
  tools: {
    calls: ToolCallIR[]
    progress: ToolProgressIR[]
    results: ToolResultIR[]
  }
}

function toRuntimeMessage(message: ChatMessage): AgentRuntimeMessage {
  return {
    id: message.id,
    role: message.role === 'system' ? 'system' : message.role,
    kind: 'text',
    content: message.content,
    createdAt: new Date(message.timestamp),
    metadata: {
      origin: 'legacy_ai_context',
    },
  }
}

function createRuntimeSessionId(): string {
  return `session-${Date.now()}-${Math.random().toString(16).slice(2)}`
}

export class AgentSessionEngine {
  private readonly sessionId = createRuntimeSessionId()
  private readonly runtime = createLegacyAgentRuntime()
  private conversation = createEmptyConversationState()
  private toolCalls: ToolCallIR[] = []
  private toolProgress: ToolProgressIR[] = []
  private toolResults: ToolResultIR[] = []
  private readonly providerAdapters: ProviderToolCallAdapter[] = [
    OpenAICompatibleToolCallAdapter,
    AnthropicToolUseAdapter,
    LegacyTextToolAdapter,
  ]
  readonly transcriptStore = new SessionTranscriptStore()
  readonly replayStore = new ReplayStore()
  readonly resumeLoader = new ResumeLoader()
  readonly resultStore = new ToolResultStore()
  readonly taskRegistry = new TaskRegistry()
  readonly notificationCenter = new TaskNotificationCenter()
  readonly hookRegistry = new HookRegistry()
  readonly subagentManager = new SubagentManager(
    this.taskRegistry,
    this.transcriptStore,
    this.notificationCenter,
  )

  syncLegacyMessages(messages: ChatMessage[]): void {
    this.runtime.messageStore.clear()
    this.runtime.messageStore.appendMany(messages.map(toRuntimeMessage))
  }

  setPhase(phase: AgentRuntimePhase): void {
    this.conversation = {
      ...this.conversation,
      phase,
    }
  }

  setCurrentTurnId(turnId: string | null): void {
    this.conversation = {
      ...this.conversation,
      currentTurnId: turnId,
    }
  }

  setPendingAttachmentTypes(types: string[]): void {
    this.conversation = {
      ...this.conversation,
      pendingAttachmentTypes: Array.from(new Set(types)),
    }
  }

  getCurrentTurnId(): string | null {
    return this.conversation.currentTurnId
  }

  getPhase(): AgentRuntimePhase {
    return this.conversation.phase
  }

  listTools(): AgentToolDefinition[] {
    return this.runtime.toolRegistry.list()
  }

  syncToolDefinitions(definitions: readonly AgentToolDefinition[]): void {
    this.runtime.syncToolDefinitions(definitions)
  }

  findTool(id: string): AgentToolDefinition | undefined {
    return this.runtime.toolRegistry.get(id)
  }

  listSkills(): AgentSkillDefinition[] {
    return this.runtime.skillRegistry.list()
  }

  findSkill(id: string): AgentSkillDefinition | undefined {
    return this.runtime.skillRegistry.get(id)
  }

  syncWorkspaceSkills(skills: AgentSkillDefinition[]): void {
    this.runtime.skillRegistry.replaceBySource('workspace', skills)
  }

  listSubagentProfiles(): SubagentProfileDefinition[] {
    return this.subagentManager.listProfiles()
  }

  findSubagentProfile(id: string): SubagentProfileDefinition | undefined {
    return this.subagentManager.getProfile(id)
  }

  spawnSubagent(params: {
    profileId: string
    label?: string
    mode?: 'sync' | 'background'
    parentTurnId?: string | null
  }): SubagentRecord {
    return this.subagentManager.spawn({
      parentSessionId: this.sessionId,
      parentTurnId: params.parentTurnId || this.conversation.currentTurnId,
      profileId: params.profileId,
      label: params.label,
      mode: params.mode,
    })
  }

  startSubagent(subagentId: string): SubagentRecord | null {
    return this.subagentManager.start(subagentId)
  }

  completeSubagent(params: {
    subagentId: string
    summary?: string
    outputPath?: string
  }): SubagentRecord | null {
    return this.subagentManager.complete(params)
  }

  failSubagent(params: { subagentId: string; error: string }): SubagentRecord | null {
    return this.subagentManager.fail(params)
  }

  cancelSubagent(subagentId: string): SubagentRecord | null {
    return this.subagentManager.cancel(subagentId)
  }

  appendSubagentTranscript(params: {
    subagentId: string
    messages?: AgentRuntimeMessage[]
    toolCalls?: ToolCallIR[]
    toolProgress?: ToolProgressIR[]
    toolResults?: ToolResultIR[]
  }): void {
    this.subagentManager.appendTranscript(params)
  }

  loadSubagentTranscript(subagentId: string) {
    return this.subagentManager.getTranscript(subagentId)
  }

  getRuntimeSnapshot(): LegacyAgentRuntimeSnapshot {
    return this.runtime.snapshot()
  }

  appendToolCall(call: ToolCallIR): void {
    this.toolCalls = [...this.toolCalls, { ...call, input: { ...call.input } }]
  }

  appendToolProgress(progress: ToolProgressIR[]): void {
    if (progress.length === 0) return
    this.toolProgress = [
      ...this.toolProgress,
      ...progress.map((item) => ({
        ...item,
        payload: item.payload ? { ...item.payload } : undefined,
      })),
    ]
  }

  appendToolResult(result: ToolResultIR): void {
    this.toolResults = [
      ...this.toolResults,
      {
        ...result,
        payload: result.payload ? { ...result.payload } : undefined,
        error: result.error
          ? {
              ...result.error,
              details: result.error.details ? { ...result.error.details } : undefined,
            }
          : undefined,
      },
    ]
  }

  recordToolExecution(state: ToolExecutionPipelineState): void {
    this.appendToolCall(state.call)
    this.appendToolProgress(state.progress)
    if (state.result) {
      this.appendToolResult(state.result)
    }
  }

  clearToolEvents(): void {
    this.toolCalls = []
    this.toolProgress = []
    this.toolResults = []
  }

  snapshot(): AgentSessionSnapshot {
    return {
      runtime: this.runtime.snapshot(),
      conversation: { ...this.conversation },
      providers: {
        adapters: this.providerAdapters.map((adapter) => ({
          ...adapter,
          capabilities: { ...adapter.capabilities },
        })),
      },
      skills: this.runtime.skillRegistry.snapshot(),
      messages: this.runtime.messageStore.snapshot(),
      tasks: this.taskRegistry.snapshot(),
      notifications: {
        count: this.notificationCenter.list().length,
        items: this.notificationCenter.list(),
      },
      subagents: this.subagentManager.snapshot(),
      tools: {
        calls: this.toolCalls.map((call) => ({
          ...call,
          input: { ...call.input },
          metadata: call.metadata ? { ...call.metadata } : undefined,
        })),
        progress: this.toolProgress.map((item) => ({
          ...item,
          payload: item.payload ? { ...item.payload } : undefined,
        })),
        results: this.toolResults.map((item) => ({
          ...item,
          payload: item.payload ? { ...item.payload } : undefined,
          error: item.error
            ? {
                ...item.error,
                details: item.error.details ? { ...item.error.details } : undefined,
              }
            : undefined,
        })),
      },
    }
  }
}
