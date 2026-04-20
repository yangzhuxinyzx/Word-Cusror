export type AgentSkillSource = 'builtin' | 'workspace'

export type AgentSkillExecutionKind = 'prompt_transform' | 'workflow'

export type AgentSkillSafety =
  | 'read_only'
  | 'mutating'
  | 'planning'
  | 'verification'

export interface AgentSkillInvocation {
  slashCommands?: string[]
  aliases?: string[]
}

export interface AgentSkillDefinition {
  id: string
  displayName: string
  description: string
  source?: AgentSkillSource
  executionKind?: AgentSkillExecutionKind
  safety?: AgentSkillSafety
  prompt?: string
  toolIds?: string[]
  tags?: string[]
  invocation?: AgentSkillInvocation
  hidden?: boolean
}

function normalizeSkillKey(value: string): string {
  const normalized = value.trim().toLowerCase()
  return normalized.startsWith('/') ? normalized.slice(1) : normalized
}

export class SkillRegistry {
  private skills = new Map<string, AgentSkillDefinition>()
  private aliases = new Map<string, string>()

  register(skill: AgentSkillDefinition): void {
    this.skills.set(skill.id, skill)

    const aliasCandidates = new Set<string>([
      skill.id,
      ...(skill.invocation?.aliases || []),
      ...(skill.invocation?.slashCommands || []),
    ])

    aliasCandidates.forEach((alias) => {
      const normalized = normalizeSkillKey(alias)
      if (!normalized) return
      this.aliases.set(normalized, skill.id)
    })
  }

  registerMany(skills: readonly AgentSkillDefinition[]): void {
    skills.forEach((skill) => this.register(skill))
  }

  unregister(id: string): void {
    if (!this.skills.has(id)) return
    this.skills.delete(id)
    for (const [alias, resolvedId] of this.aliases.entries()) {
      if (resolvedId === id) {
        this.aliases.delete(alias)
      }
    }
  }

  clear(): void {
    this.skills.clear()
    this.aliases.clear()
  }

  removeBySource(source: AgentSkillSource): void {
    this.list()
      .filter((skill) => (skill.source || 'builtin') === source)
      .forEach((skill) => this.unregister(skill.id))
  }

  replaceBySource(
    source: AgentSkillSource,
    skills: readonly AgentSkillDefinition[],
  ): void {
    this.removeBySource(source)
    this.registerMany(
      skills.map((skill) => ({
        ...skill,
        source,
      })),
    )
  }

  resolve(idOrAlias: string):
    | { id: string; definition: AgentSkillDefinition }
    | undefined {
    const normalized = normalizeSkillKey(idOrAlias)
    if (!normalized) return undefined

    const resolvedId =
      this.skills.has(normalized)
        ? normalized
        : this.aliases.get(normalized)
    if (!resolvedId) return undefined

    const definition = this.skills.get(resolvedId)
    if (!definition) return undefined

    return {
      id: resolvedId,
      definition,
    }
  }

  get(id: string): AgentSkillDefinition | undefined {
    return this.resolve(id)?.definition
  }

  list(): AgentSkillDefinition[] {
    return Array.from(this.skills.values())
  }

  ids(): string[] {
    return this.list().map((skill) => skill.id)
  }

  commandIds(): string[] {
    return Array.from(this.aliases.keys()).sort((left, right) =>
      left.localeCompare(right),
    )
  }

  snapshot() {
    const list = this.list()
    return {
      count: list.length,
      ids: this.ids(),
      aliasCount: this.aliases.size,
      commands: this.commandIds(),
      builtinCount: list.filter((skill) => (skill.source || 'builtin') === 'builtin')
        .length,
      workspaceCount: list.filter((skill) => skill.source === 'workspace').length,
    }
  }
}
