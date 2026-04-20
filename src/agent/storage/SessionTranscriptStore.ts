import type { AgentRuntimeMessage } from '../core/messageTypes'
import type { ToolCallIR, ToolProgressIR, ToolResultIR } from '../tools/ir'

export interface SessionTranscriptRecord {
  sessionId: string
  messages: AgentRuntimeMessage[]
  toolCalls: ToolCallIR[]
  toolProgress: ToolProgressIR[]
  toolResults: ToolResultIR[]
  updatedAt: string
}

export class SessionTranscriptStore {
  private records = new Map<string, SessionTranscriptRecord>()

  save(record: SessionTranscriptRecord): void {
    this.records.set(record.sessionId, {
      ...record,
      messages: record.messages.map((message) => ({
        ...message,
        createdAt: new Date(message.createdAt),
        metadata: message.metadata ? { ...message.metadata } : undefined,
      })),
      toolCalls: record.toolCalls.map((call) => ({
        ...call,
        input: { ...call.input },
        metadata: call.metadata ? { ...call.metadata } : undefined,
      })),
      toolProgress: record.toolProgress.map((item) => ({
        ...item,
        payload: item.payload ? { ...item.payload } : undefined,
      })),
      toolResults: record.toolResults.map((item) => ({
        ...item,
        payload: item.payload ? { ...item.payload } : undefined,
        error: item.error
          ? {
              ...item.error,
              details: item.error.details ? { ...item.error.details } : undefined,
            }
          : undefined,
      })),
    })
  }

  load(sessionId: string): SessionTranscriptRecord | null {
    return this.records.get(sessionId) || null
  }
}

