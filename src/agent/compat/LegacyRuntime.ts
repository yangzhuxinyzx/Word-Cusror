import { MessageStore } from '../core/MessageStore'
import { BUILTIN_SKILLS } from '../skills/builtinSkills'
import type { AgentToolDomain } from '../tools/contracts'
import { AgentToolRegistry } from '../tools/registry'
import { SkillRegistry } from '../skills/SkillRegistry'
import { LEGACY_TOOL_DEFINITIONS } from './legacyTools'
import type { AgentToolDefinition } from '../tools/contracts'

export interface LegacyAgentRuntimeSnapshot {
  toolCount: number
  skillCount: number
  skillIds: string[]
  messageCount: number
  domains: AgentToolDomain[]
}

export class LegacyAgentRuntime {
  readonly toolRegistry = new AgentToolRegistry()
  readonly skillRegistry = new SkillRegistry()
  readonly messageStore = new MessageStore()

  constructor() {
    this.toolRegistry.registerMany(LEGACY_TOOL_DEFINITIONS)
    this.skillRegistry.registerMany(BUILTIN_SKILLS)
  }

  syncToolDefinitions(definitions: readonly AgentToolDefinition[]): void {
    this.toolRegistry.replaceMany(definitions)
  }

  snapshot(): LegacyAgentRuntimeSnapshot {
    const toolSnapshot = this.toolRegistry.snapshot()
    const skillSnapshot = this.skillRegistry.snapshot()
    const messageSnapshot = this.messageStore.snapshot()

    return {
      toolCount: toolSnapshot.count,
      skillCount: skillSnapshot.count,
      skillIds: skillSnapshot.ids,
      messageCount: messageSnapshot.count,
      domains: Array.from(new Set(this.toolRegistry.list().map((tool) => tool.domain))),
    }
  }
}

export function createLegacyAgentRuntime(): LegacyAgentRuntime {
  return new LegacyAgentRuntime()
}
