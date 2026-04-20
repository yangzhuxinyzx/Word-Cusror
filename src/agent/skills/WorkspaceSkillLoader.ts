import type { FileItem, FileReadResult, FolderReadResult } from '../../types/electron'
import type { AgentSkillDefinition } from './SkillRegistry'

const SKILL_ID_REGEX = /^[a-z0-9][a-z0-9-_]{1,63}$/i

export interface WorkspaceSkillLoaderOptions {
  readFile?: (filePath: string) => Promise<FileReadResult>
  readFolder?: (folderPath: string) => Promise<FolderReadResult>
}

export interface WorkspaceSkillLoadResult {
  skills: AgentSkillDefinition[]
  warnings: string[]
  sources: string[]
}

type RawSkillManifest =
  | AgentSkillDefinition
  | AgentSkillDefinition[]
  | { skills?: AgentSkillDefinition[] }

function joinWindowsPath(...segments: Array<string | null | undefined>): string {
  return segments
    .filter((segment): segment is string => !!segment)
    .map((segment, index) =>
      index === 0
        ? segment.replace(/[\\/]+$/, '')
        : segment.replace(/^[\\/]+|[\\/]+$/g, ''),
    )
    .join('\\')
}

function flattenFiles(items: FileItem[] | undefined): FileItem[] {
  const result: FileItem[] = []
  const walk = (nodes: FileItem[]) => {
    nodes.forEach((node) => {
      if (node.type === 'file') {
        result.push(node)
        return
      }
      if (node.children?.length) {
        walk(node.children)
      }
    })
  }
  if (items?.length) {
    walk(items)
  }
  return result
}

function toArray(manifest: RawSkillManifest): AgentSkillDefinition[] {
  if (Array.isArray(manifest)) return manifest
  if ('skills' in manifest && Array.isArray(manifest.skills)) {
    return manifest.skills
  }
  return [manifest]
}

function normalizeStringArray(value: unknown): string[] | undefined {
  if (!Array.isArray(value)) return undefined
  const normalized = value
    .filter((item): item is string => typeof item === 'string')
    .map((item) => item.trim())
    .filter(Boolean)
  return normalized.length > 0 ? normalized : undefined
}

function normalizeWorkspaceSkill(
  raw: AgentSkillDefinition,
  sourcePath: string,
): { skill: AgentSkillDefinition | null; warnings: string[] } {
  const warnings: string[] = []

  if (!raw || typeof raw !== 'object' || typeof raw.id !== 'string') {
    return { skill: null, warnings: [`Skipped invalid skill in ${sourcePath}`] }
  }

  const id = raw.id.trim()
  if (!id) {
    return { skill: null, warnings: [`Skipped unnamed skill in ${sourcePath}`] }
  }

  if (!SKILL_ID_REGEX.test(id)) {
    return {
      skill: null,
      warnings: [
        `Skipped workspace skill "${id}" in ${sourcePath}: id must match ${SKILL_ID_REGEX}`,
      ],
    }
  }

  const displayName =
    typeof raw.displayName === 'string' && raw.displayName.trim()
      ? raw.displayName.trim()
      : id

  const description =
    typeof raw.description === 'string' && raw.description.trim()
      ? raw.description.trim()
      : `Workspace skill loaded from ${sourcePath}`

  const executionKind =
    raw.executionKind === 'workflow' ? 'workflow' : 'prompt_transform'
  if (executionKind === 'workflow') {
    warnings.push(
      `Workspace skill "${id}" in ${sourcePath} requested workflow mode; runtime currently downgrades workspace skills to prompt_transform.`,
    )
  }

  const prompt = typeof raw.prompt === 'string' ? raw.prompt.trim() : undefined
  if (!prompt) {
    warnings.push(
      `Workspace skill "${id}" in ${sourcePath} has no prompt; it may not be useful until a prompt is provided.`,
    )
  }

  return {
    skill: {
    id,
    displayName,
    description,
    source: 'workspace',
    executionKind: 'prompt_transform',
    safety:
      raw.safety === 'read_only' ||
      raw.safety === 'planning' ||
      raw.safety === 'verification'
        ? raw.safety
        : 'mutating',
    prompt,
    toolIds: normalizeStringArray(raw.toolIds),
    tags: normalizeStringArray(raw.tags),
    hidden: raw.hidden === true,
    invocation: {
      slashCommands: normalizeStringArray(raw.invocation?.slashCommands),
      aliases: normalizeStringArray(raw.invocation?.aliases),
    },
    },
    warnings,
  }
}

export class WorkspaceSkillLoader {
  constructor(private readonly options: WorkspaceSkillLoaderOptions) {}

  private async readOptionalJsonFile(filePath: string): Promise<{
    content: string
    path: string
  } | null> {
    if (!this.options.readFile) return null
    try {
      const result = await this.options.readFile(filePath)
      if (!result?.success || !result.data) return null
      return {
        content: result.data,
        path: filePath,
      }
    } catch {
      return null
    }
  }

  async load(workspacePath: string): Promise<WorkspaceSkillLoadResult> {
    if (!workspacePath || !this.options.readFile) {
      return { skills: [], warnings: [], sources: [] }
    }

    const warnings: string[] = []
    const sources: string[] = []
    const manifests: Array<{ content: string; path: string }> = []

    const manifestPath = joinWindowsPath(workspacePath, '.word-cursor', 'skills.json')
    const manifest = await this.readOptionalJsonFile(manifestPath)
    if (manifest) {
      manifests.push(manifest)
      sources.push(manifest.path)
    }

    if (this.options.readFolder) {
      const skillDirPath = joinWindowsPath(workspacePath, '.word-cursor', 'skills')
      try {
        const folder = await this.options.readFolder(skillDirPath)
        if (folder?.success && folder.data) {
          const files = flattenFiles(folder.data).filter((file) =>
            file.name.toLowerCase().endsWith('.json'),
          )
          for (const file of files) {
            const loaded = await this.readOptionalJsonFile(file.path)
            if (loaded) {
              manifests.push(loaded)
              sources.push(loaded.path)
            }
          }
        }
      } catch {
        // ignore missing workspace skill directory
      }
    }

    const deduped = new Map<string, AgentSkillDefinition>()

    for (const manifestEntry of manifests) {
      try {
        const parsed = JSON.parse(manifestEntry.content) as RawSkillManifest
        for (const rawSkill of toArray(parsed)) {
          const normalized = normalizeWorkspaceSkill(rawSkill, manifestEntry.path)
          warnings.push(...normalized.warnings)
          if (!normalized.skill) {
            continue
          }
          deduped.set(normalized.skill.id, normalized.skill)
        }
      } catch (error) {
        warnings.push(
          `Failed to parse workspace skill manifest ${manifestEntry.path}: ${String(error)}`,
        )
      }
    }

    return {
      skills: Array.from(deduped.values()).sort((left, right) =>
        left.id.localeCompare(right.id),
      ),
      warnings,
      sources,
    }
  }
}
