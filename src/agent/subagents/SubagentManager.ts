import type { AgentRuntimeMessage } from '../core/messageTypes'
import type { SessionTranscriptRecord } from '../storage/SessionTranscriptStore'
import { SessionTranscriptStore } from '../storage/SessionTranscriptStore'
import type { ToolCallIR, ToolProgressIR, ToolResultIR } from '../tools/ir'
import { TaskNotificationCenter } from '../tasks/TaskNotifications'
import {
  TaskRegistry,
  type RuntimeTask,
  type RuntimeTaskMode,
  type RuntimeTaskStatus,
} from '../tasks/TaskRegistry'
import {
  BUILTIN_SUBAGENT_PROFILES,
  type SubagentExecutionMode,
  type SubagentProfileDefinition,
  type SubagentSafety,
} from './SubagentProfiles'

export type SubagentStatus = RuntimeTaskStatus

export interface SubagentRecord {
  subagentId: string
  parentSessionId: string
  parentTurnId?: string | null
  profileId: string
  label: string
  status: SubagentStatus
  mode: SubagentExecutionMode
  safety: SubagentSafety
  transcriptSessionId: string
  taskId: string
  createdAt: string
  updatedAt: string
  startedAt?: string
  completedAt?: string
  summary?: string
  outputPath?: string
  error?: string
  allowedToolIds?: string[]
}

export interface SpawnSubagentParams {
  parentSessionId: string
  parentTurnId?: string | null
  profileId: string
  label?: string
  mode?: SubagentExecutionMode
}

export interface CompleteSubagentParams {
  subagentId: string
  summary?: string
  outputPath?: string
}

export interface FailSubagentParams {
  subagentId: string
  error: string
}

export interface SubagentSnapshot {
  profiles: SubagentProfileDefinition[]
  agents: SubagentRecord[]
}

function createId(prefix: string): string {
  return `${prefix}-${Date.now()}-${Math.random().toString(16).slice(2)}`
}

function toTaskMode(mode: SubagentExecutionMode): RuntimeTaskMode {
  return mode === 'background' ? 'background' : 'foreground'
}

export class SubagentManager {
  private readonly profiles = new Map<string, SubagentProfileDefinition>()
  private readonly agents = new Map<string, SubagentRecord>()

  constructor(
    private readonly taskRegistry = new TaskRegistry(),
    private readonly transcriptStore = new SessionTranscriptStore(),
    private readonly notifications = new TaskNotificationCenter(),
    profiles: readonly SubagentProfileDefinition[] = BUILTIN_SUBAGENT_PROFILES,
  ) {
    this.registerMany(profiles)
  }

  register(profile: SubagentProfileDefinition): void {
    this.profiles.set(profile.id, { ...profile })
  }

  registerMany(profiles: readonly SubagentProfileDefinition[]): void {
    profiles.forEach((profile) => this.register(profile))
  }

  listProfiles(): SubagentProfileDefinition[] {
    return Array.from(this.profiles.values()).sort((left, right) =>
      left.id.localeCompare(right.id),
    )
  }

  getProfile(id: string): SubagentProfileDefinition | undefined {
    return this.profiles.get(id)
  }

  spawn(params: SpawnSubagentParams): SubagentRecord {
    const profile = this.getProfile(params.profileId)
    if (!profile) {
      throw new Error(`Unknown subagent profile: ${params.profileId}`)
    }

    const mode = params.mode || profile.defaultMode
    if (mode === 'background' && profile.canRunInBackground === false) {
      throw new Error(`Subagent profile ${profile.id} cannot run in background mode`)
    }

    const createdAt = new Date().toISOString()
    const subagentId = createId('subagent')
    const taskId = createId('task')
    const transcriptSessionId = `${params.parentSessionId}:${subagentId}`

    const record: SubagentRecord = {
      subagentId,
      parentSessionId: params.parentSessionId,
      parentTurnId: params.parentTurnId || null,
      profileId: profile.id,
      label: params.label || profile.displayName,
      status: 'queued',
      mode,
      safety: profile.safety,
      transcriptSessionId,
      taskId,
      createdAt,
      updatedAt: createdAt,
      allowedToolIds: profile.allowedToolIds ? [...profile.allowedToolIds] : undefined,
    }

    this.agents.set(record.subagentId, { ...record })
    this.taskRegistry.upsert(this.toTaskRecord(record))
    this.transcriptStore.save({
      sessionId: transcriptSessionId,
      messages: [],
      toolCalls: [],
      toolProgress: [],
      toolResults: [],
      updatedAt: createdAt,
    })

    if (mode === 'background') {
      this.notifications.push({
        taskId,
        message: `Started background subagent ${record.label}`,
        createdAt,
      })
    }

    return this.cloneRecord(record)
  }

  start(subagentId: string): SubagentRecord | null {
    return this.update(subagentId, (current) => {
      const startedAt = new Date().toISOString()
      return {
        ...current,
        status: 'running',
        startedAt,
        updatedAt: startedAt,
      }
    })
  }

  complete(params: CompleteSubagentParams): SubagentRecord | null {
    return this.update(params.subagentId, (current) => {
      const completedAt = new Date().toISOString()
      const next: SubagentRecord = {
        ...current,
        status: 'completed',
        completedAt,
        updatedAt: completedAt,
        summary: params.summary || current.summary,
        outputPath: params.outputPath || current.outputPath,
      }
      if (next.mode === 'background') {
        this.notifications.push({
          taskId: next.taskId,
          message: `${next.label} completed`,
          createdAt: completedAt,
        })
      }
      return next
    })
  }

  fail(params: FailSubagentParams): SubagentRecord | null {
    return this.update(params.subagentId, (current) => {
      const failedAt = new Date().toISOString()
      const next: SubagentRecord = {
        ...current,
        status: 'failed',
        completedAt: failedAt,
        updatedAt: failedAt,
        error: params.error,
      }
      if (next.mode === 'background') {
        this.notifications.push({
          taskId: next.taskId,
          message: `${next.label} failed: ${params.error}`,
          createdAt: failedAt,
        })
      }
      return next
    })
  }

  cancel(subagentId: string): SubagentRecord | null {
    return this.update(subagentId, (current) => {
      const cancelledAt = new Date().toISOString()
      const next: SubagentRecord = {
        ...current,
        status: 'cancelled',
        completedAt: cancelledAt,
        updatedAt: cancelledAt,
      }
      if (next.mode === 'background') {
        this.notifications.push({
          taskId: next.taskId,
          message: `${next.label} cancelled`,
          createdAt: cancelledAt,
        })
      }
      return next
    })
  }

  appendTranscript(params: {
    subagentId: string
    messages?: AgentRuntimeMessage[]
    toolCalls?: ToolCallIR[]
    toolProgress?: ToolProgressIR[]
    toolResults?: ToolResultIR[]
  }): void {
    const record = this.get(params.subagentId)
    if (!record) return
    const existing = this.transcriptStore.load(record.transcriptSessionId)
    this.transcriptStore.save({
      sessionId: record.transcriptSessionId,
      messages: [
        ...(existing?.messages || []),
        ...((params.messages || []).map((message) => ({
          ...message,
          metadata: message.metadata ? { ...message.metadata } : undefined,
        })) as AgentRuntimeMessage[]),
      ],
      toolCalls: [
        ...(existing?.toolCalls || []),
        ...((params.toolCalls || []).map((call) => ({
          ...call,
          input: { ...call.input },
          metadata: call.metadata ? { ...call.metadata } : undefined,
        })) as ToolCallIR[]),
      ],
      toolProgress: [
        ...(existing?.toolProgress || []),
        ...((params.toolProgress || []).map((item) => ({
          ...item,
          payload: item.payload ? { ...item.payload } : undefined,
        })) as ToolProgressIR[]),
      ],
      toolResults: [
        ...(existing?.toolResults || []),
        ...((params.toolResults || []).map((item) => ({
          ...item,
          payload: item.payload ? { ...item.payload } : undefined,
          error: item.error
            ? {
                ...item.error,
                details: item.error.details ? { ...item.error.details } : undefined,
              }
            : undefined,
        })) as ToolResultIR[]),
      ],
      updatedAt: new Date().toISOString(),
    })
  }

  get(subagentId: string): SubagentRecord | null {
    const found = this.agents.get(subagentId)
    return found ? this.cloneRecord(found) : null
  }

  list(): SubagentRecord[] {
    return Array.from(this.agents.values()).map((record) => this.cloneRecord(record))
  }

  listByParentSession(sessionId: string): SubagentRecord[] {
    return this.list().filter((record) => record.parentSessionId === sessionId)
  }

  getTranscriptSessionId(subagentId: string): string | null {
    return this.agents.get(subagentId)?.transcriptSessionId || null
  }

  getTranscript(subagentId: string): SessionTranscriptRecord | null {
    const transcriptSessionId = this.getTranscriptSessionId(subagentId)
    if (!transcriptSessionId) return null
    return this.transcriptStore.load(transcriptSessionId)
  }

  snapshot(): SubagentSnapshot {
    return {
      profiles: this.listProfiles(),
      agents: this.list(),
    }
  }

  private update(
    subagentId: string,
    updater: (current: SubagentRecord) => SubagentRecord,
  ): SubagentRecord | null {
    const current = this.agents.get(subagentId)
    if (!current) return null
    const next = updater(this.cloneRecord(current))
    this.agents.set(subagentId, this.cloneRecord(next))
    this.taskRegistry.upsert(this.toTaskRecord(next))
    return this.get(subagentId)
  }

  private toTaskRecord(record: SubagentRecord): RuntimeTask {
    return {
      taskId: record.taskId,
      label: record.label,
      kind: 'subagent',
      mode: toTaskMode(record.mode),
      ownerId: record.subagentId,
      status: record.status,
      outputPath: record.outputPath,
      transcriptPath: record.transcriptSessionId,
      createdAt: record.createdAt,
      updatedAt: record.updatedAt,
      startedAt: record.startedAt,
      completedAt: record.completedAt,
      summary: record.summary,
      error: record.error,
    }
  }

  private cloneRecord(record: SubagentRecord): SubagentRecord {
    return {
      ...record,
      allowedToolIds: record.allowedToolIds ? [...record.allowedToolIds] : undefined,
    }
  }
}
