import type { AgentToolDefinition } from './contracts'
import type { ToolCallIR } from './ir'

export interface ScheduledToolCall {
  call: ToolCallIR
  queue: 'serial' | 'parallel_safe'
}

export class AgentToolScheduler {
  schedule(
    call: ToolCallIR,
    tool: AgentToolDefinition,
  ): ScheduledToolCall {
    return {
      call: {
        ...call,
        status: 'scheduled',
        concurrency: tool.concurrency,
        domain: tool.domain,
        mutation: tool.mutation,
      },
      queue: tool.concurrency === 'parallel_safe' ? 'parallel_safe' : 'serial',
    }
  }
}

