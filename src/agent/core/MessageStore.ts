import type { AgentRuntimeMessage } from './messageTypes'

export interface MessageStoreSnapshot {
  count: number
  lastMessageId: string | null
  lastRole: AgentRuntimeMessage['role'] | null
}

function cloneMessage(message: AgentRuntimeMessage): AgentRuntimeMessage {
  return {
    ...message,
    createdAt: new Date(message.createdAt),
    metadata: message.metadata ? { ...message.metadata } : undefined,
  }
}

export class MessageStore {
  private messages: AgentRuntimeMessage[]

  constructor(initialMessages: AgentRuntimeMessage[] = []) {
    this.messages = initialMessages.map(cloneMessage)
  }

  getAll(): AgentRuntimeMessage[] {
    return this.messages.map(cloneMessage)
  }

  append(message: AgentRuntimeMessage): AgentRuntimeMessage[] {
    this.messages = [...this.messages, cloneMessage(message)]
    return this.getAll()
  }

  appendMany(messages: AgentRuntimeMessage[]): AgentRuntimeMessage[] {
    this.messages = [...this.messages, ...messages.map(cloneMessage)]
    return this.getAll()
  }

  replaceLast(message: AgentRuntimeMessage): AgentRuntimeMessage[] {
    if (this.messages.length === 0) {
      return this.append(message)
    }

    const next = this.messages.slice()
    next[next.length - 1] = cloneMessage(message)
    this.messages = next
    return this.getAll()
  }

  updateLast(
    updater: (message: AgentRuntimeMessage | null) => AgentRuntimeMessage | null,
  ): AgentRuntimeMessage[] {
    const current = this.messages.length > 0 ? cloneMessage(this.messages[this.messages.length - 1]) : null
    const updated = updater(current)

    if (updated === null) {
      return this.getAll()
    }

    return this.replaceLast(updated)
  }

  clear(): void {
    this.messages = []
  }

  snapshot(): MessageStoreSnapshot {
    const lastMessage = this.messages[this.messages.length - 1]

    return {
      count: this.messages.length,
      lastMessageId: lastMessage?.id || null,
      lastRole: lastMessage?.role || null,
    }
  }
}

