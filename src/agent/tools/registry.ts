import type { AgentToolDefinition } from './contracts'

export class AgentToolRegistry {
  private tools = new Map<string, AgentToolDefinition>()
  private aliases = new Map<string, string>()

  clear(): void {
    this.tools.clear()
    this.aliases.clear()
  }

  register(definition: AgentToolDefinition): void {
    for (const [alias, canonicalId] of this.aliases.entries()) {
      if (canonicalId === definition.id) {
        this.aliases.delete(alias)
      }
    }
    this.tools.set(definition.id, definition)
    definition.legacyAliases?.forEach((alias) => {
      if (!alias || alias === definition.id) return
      this.aliases.set(alias, definition.id)
    })
  }

  registerMany(definitions: readonly AgentToolDefinition[]): void {
    definitions.forEach((definition) => this.register(definition))
  }

  replaceMany(definitions: readonly AgentToolDefinition[]): void {
    this.clear()
    this.registerMany(definitions)
  }

  resolveId(id: string): string | undefined {
    if (this.tools.has(id)) return id
    return this.aliases.get(id)
  }

  resolve(id: string):
    | { id: string; definition: AgentToolDefinition }
    | undefined {
    const resolvedId = this.resolveId(id)
    if (!resolvedId) return undefined
    const definition = this.tools.get(resolvedId)
    if (!definition) return undefined
    return {
      id: resolvedId,
      definition,
    }
  }

  get(id: string): AgentToolDefinition | undefined {
    return this.resolve(id)?.definition
  }

  list(): AgentToolDefinition[] {
    return Array.from(this.tools.values())
  }

  ids(): string[] {
    return this.list().map((definition) => definition.id)
  }

  has(id: string): boolean {
    return !!this.resolveId(id)
  }

  snapshot() {
    return {
      count: this.tools.size,
      ids: this.ids(),
      aliasCount: this.aliases.size,
      aliases: Array.from(this.aliases.entries()).map(([alias, canonicalId]) => ({
        alias,
        canonicalId,
      })),
    }
  }
}
