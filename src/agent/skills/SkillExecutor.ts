import type { AgentToolExecutionResult } from '../tools/results'
import {
  SkillRegistry,
  type AgentSkillDefinition,
} from './SkillRegistry'

export interface SkillExecutionContext {
  invokeTool?: (
    toolId: string,
    args: Record<string, string>,
  ) => Promise<AgentToolExecutionResult | null>
  buildWorkspaceProfile?: (refresh?: boolean) => Promise<string>
  currentFileName?: string | null
  workspacePath?: string | null
  workspaceSkillWarnings?: string[]
  workspaceSkillSources?: string[]
}

export interface SkillExecutionResult {
  skill: AgentSkillDefinition
  mode: 'handled' | 'transform'
  assistantMessage?: string
  transformedInput?: string
  workspaceProfile?: string
  allowedToolIds?: string[]
}

interface SkillInvocationMatch {
  skill: AgentSkillDefinition
  command: string
  body: string
}

interface WorkspaceProfilePayload {
  folderPath?: string
  totalFiles?: number
  fileTypes?: Record<string, number>
  topFiles?: Array<{
    name: string
    path: string
    relativePath?: string
    extension?: string
  }>
  summary?: string
}

function normalizeSkillCommand(value: string): string {
  const normalized = value.trim().toLowerCase()
  return normalized.startsWith('/') ? normalized.slice(1) : normalized
}

function tryParseWorkspaceProfile(
  content: string,
): WorkspaceProfilePayload | null {
  if (!content.trim()) return null
  try {
    const parsed = JSON.parse(content) as WorkspaceProfilePayload
    if (!parsed || typeof parsed !== 'object') return null
    return parsed
  } catch {
    return null
  }
}

function formatWorkspaceProfileSummary(
  payload: WorkspaceProfilePayload | null,
  rawProfile: string,
): string {
  if (!payload) {
    return [
      '**/init Result**',
      '',
      'Workspace profile created, but it could not be parsed as structured JSON.',
      '',
      '```json',
      rawProfile,
      '```',
    ].join('\n')
  }

  const fileTypes = Object.entries(payload.fileTypes || {})
    .slice(0, 8)
    .map(([ext, count]) => `${ext}: ${count}`)

  const topFiles = (payload.topFiles || [])
    .slice(0, 6)
    .map((file) => `- ${file.relativePath || file.name}`)

  const suggestions: string[] = [
    'Use `workspace_read` to inspect the highest-value files first.',
  ]
  const hasDoc = Object.keys(payload.fileTypes || {}).some((ext) =>
    ['docx', 'doc', 'md', 'txt', 'pdf'].includes(ext),
  )
  const hasPpt = Object.keys(payload.fileTypes || {}).some((ext) =>
    ['pptx', 'ppt'].includes(ext),
  )
  const hasExcel = Object.keys(payload.fileTypes || {}).some((ext) =>
    ['xlsx', 'xls', 'csv'].includes(ext),
  )
  if (hasDoc) {
    suggestions.push('For document cleanup, use `/rewrite-formal` or `/document-proofread`.')
  }
  if (hasPpt) {
    suggestions.push('For deck generation or refinement, use `/ppt-from-outline`.')
  }
  if (hasExcel) {
    suggestions.push('For workbook cleanup, use `/excel-cleanup`.')
  }

  return [
    '**/init Result**',
    '',
    `Workspace: ${payload.folderPath || '(unknown)'}`,
    `Total files: ${payload.totalFiles || 0}`,
    fileTypes.length > 0 ? `File types: ${fileTypes.join(', ')}` : '',
    topFiles.length > 0 ? 'Priority files:\n' + topFiles.join('\n') : '',
    payload.summary ? `Auto summary:\n${payload.summary}` : '',
    suggestions.length > 0 ? 'Suggested next actions:\n- ' + suggestions.join('\n- ') : '',
    '',
    '```json',
    rawProfile,
    '```',
  ]
    .filter(Boolean)
    .join('\n\n')
}

export class SkillExecutor {
  constructor(private readonly registry: SkillRegistry) {}

  private matchInvocation(input: string): SkillInvocationMatch | null {
    const trimmed = input.trim()
    if (!trimmed) return null

    if (trimmed.startsWith('/')) {
      const withoutSlash = trimmed.slice(1).trim()
      if (!withoutSlash) return null
      const firstSpace = withoutSlash.indexOf(' ')
      const command =
        firstSpace >= 0 ? withoutSlash.slice(0, firstSpace) : withoutSlash
      const body =
        firstSpace >= 0 ? withoutSlash.slice(firstSpace + 1).trim() : ''
      const resolved = this.registry.resolve(command)
      if (!resolved) return null
      return {
        skill: resolved.definition,
        command: normalizeSkillCommand(command),
        body,
      }
    }

    const resolved = this.registry.resolve(trimmed)
    if (!resolved) return null
    return {
      skill: resolved.definition,
      command: normalizeSkillCommand(trimmed),
      body: '',
    }
  }

  async execute(
    input: string,
    context: SkillExecutionContext = {},
  ): Promise<SkillExecutionResult | null> {
    const match = this.matchInvocation(input)
    if (!match) return null

    if ((match.skill.executionKind || 'prompt_transform') === 'workflow') {
      return this.executeWorkflowSkill(match, context)
    }

    return {
      skill: match.skill,
      mode: 'transform',
      transformedInput: this.buildPromptTransform(match, context),
      allowedToolIds: match.skill.toolIds,
    }
  }

  private buildPromptTransform(
    match: SkillInvocationMatch,
    context: SkillExecutionContext,
  ): string {
    const body = match.body.trim()
    const target =
      body ||
      (context.currentFileName
        ? `Apply this skill to the current file "${context.currentFileName}".`
        : 'Apply this skill to the current active context.')
    const toolHint =
      match.skill.toolIds && match.skill.toolIds.length > 0
        ? `Preferred tools: ${match.skill.toolIds.join(', ')}`
        : ''
    const toolBoundary =
      match.skill.toolIds && match.skill.toolIds.length > 0
        ? `Do not use tools outside this skill boundary: ${match.skill.toolIds.join(', ')}`
        : ''

    return [
      `[Skill Activation] ${match.skill.displayName} (${match.skill.id})`,
      `Description: ${match.skill.description}`,
      match.skill.prompt ? `Instructions: ${match.skill.prompt}` : '',
      toolHint,
      toolBoundary,
      `User request:\n${target}`,
    ]
      .filter(Boolean)
      .join('\n\n')
  }

  private async executeWorkflowSkill(
    match: SkillInvocationMatch,
    context: SkillExecutionContext,
  ): Promise<SkillExecutionResult> {
    if (match.skill.id === 'init') {
      return this.executeInitSkill(match.skill, context)
    }

    if (match.skill.id === 'skills') {
      return {
        skill: match.skill,
        mode: 'handled',
        assistantMessage: this.formatSkillList(context),
      }
    }

    return {
      skill: match.skill,
      mode: 'handled',
      assistantMessage: `Skill workflow not implemented: ${match.skill.id}`,
    }
  }

  private async executeInitSkill(
    skill: AgentSkillDefinition,
    context: SkillExecutionContext,
  ): Promise<SkillExecutionResult> {
    let rawProfile = ''

    if (context.invokeTool) {
      const result = await context.invokeTool('workspace_profile', {
        refresh: 'true',
        ...(context.workspacePath ? { path: context.workspacePath } : {}),
      })
      if (!result) {
        return {
          skill,
          mode: 'handled',
          assistantMessage: 'Unable to run `workspace_profile` for /init.',
        }
      }
      if (!result.success) {
        return {
          skill,
          mode: 'handled',
          assistantMessage: result.message,
        }
      }
      rawProfile =
        typeof result.data?.profile === 'string'
          ? result.data.profile
          : result.message
    } else if (context.buildWorkspaceProfile) {
      rawProfile = await context.buildWorkspaceProfile(true)
    }

    if (!rawProfile.trim()) {
      return {
        skill,
        mode: 'handled',
        assistantMessage:
          'No workspace profile was produced. Open a workspace or file first, then run `/init` again.',
      }
    }

    const payload = tryParseWorkspaceProfile(rawProfile)
    return {
      skill,
      mode: 'handled',
      assistantMessage: formatWorkspaceProfileSummary(payload, rawProfile),
      workspaceProfile: rawProfile,
    }
  }

  private formatSkillList(context: SkillExecutionContext): string {
    const skills = this.registry
      .list()
      .filter((skill) => !skill.hidden)
      .sort((left, right) => left.id.localeCompare(right.id))

    const lines = skills.map((skill) => {
      const commands = skill.invocation?.slashCommands?.length
        ? skill.invocation.slashCommands.map((command) => `/${command}`).join(', ')
        : `/${skill.id}`
      const source = skill.source || 'builtin'
      const safety = skill.safety || 'mutating'
      const toolScope =
        skill.toolIds?.length
          ? `\n  tools: ${skill.toolIds.join(', ')}`
          : ''
      return `- ${commands}\n  ${skill.displayName} | ${source} | ${safety}\n  ${skill.description}${toolScope}`
    })

    return [
      '**Available Skills**',
      '',
      context.workspaceSkillSources?.length
        ? `Workspace manifests:\n- ${context.workspaceSkillSources.join('\n- ')}`
        : '',
      context.workspaceSkillWarnings?.length
        ? `Workspace warnings:\n- ${context.workspaceSkillWarnings.join('\n- ')}`
        : '',
      context.workspaceSkillSources?.length || context.workspaceSkillWarnings?.length
        ? ''
        : '',
      ...lines,
    ].join('\n')
  }
}
