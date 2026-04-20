import {
  CORE_ATTACHMENT_DEFINITIONS,
  type AgentAttachment,
  type AgentAttachmentDefinition,
  type AgentAttachmentKind,
  type AgentAttachmentScope,
} from './AttachmentTypes'

export interface CreateAttachmentOptions<TPayload = unknown> {
  type: string
  payload: TPayload
  scope?: AgentAttachmentScope
  kind?: AgentAttachmentKind
  id?: string
}

const LEGACY_ATTACHMENT_LABELS: Record<string, string> = {
  active_skill: '当前激活技能',
  skill_descriptions: '可用技能说明',
  current_document_summary: '当前文档内容',
  document_structure_delta: '文档结构增量',
  available_tools_delta: '可用工具变化',
  workspace_context_delta: '附加文件内容',
  workspace_profile: '工作区画像',
  relevant_memories: '记忆检索',
  ppt_edit_context: 'PPT 编辑上下文',
  excel_sheet_context: 'Excel 工作表上下文',
}

function createAttachmentId(type: string): string {
  return `attachment-${type}-${Date.now()}-${Math.random().toString(16).slice(2)}`
}

function normalizeAttachmentText(payload: unknown): string {
  if (payload === null || payload === undefined) return ''
  if (typeof payload === 'string') return payload.trim()
  try {
    return JSON.stringify(payload, null, 2)
  } catch {
    return String(payload)
  }
}

export class AttachmentManager {
  private definitions = new Map<string, AgentAttachmentDefinition>()

  constructor(
    definitions: readonly AgentAttachmentDefinition[] = CORE_ATTACHMENT_DEFINITIONS,
  ) {
    this.registerMany(definitions)
  }

  register(definition: AgentAttachmentDefinition): void {
    this.definitions.set(definition.type, definition)
  }

  registerMany(definitions: readonly AgentAttachmentDefinition[]): void {
    definitions.forEach((definition) => this.register(definition))
  }

  getDefinition(type: string): AgentAttachmentDefinition | undefined {
    return this.definitions.get(type)
  }

  createAttachment<TPayload>(
    options: CreateAttachmentOptions<TPayload>,
  ): AgentAttachment<TPayload> {
    const definition = this.getDefinition(options.type)
    return {
      id: options.id || createAttachmentId(options.type),
      type: options.type,
      scope: options.scope || definition?.scope || 'turn',
      kind: options.kind || definition?.kind || 'text',
      payload: options.payload,
    }
  }

  serializeForLegacyPrompt(attachment: AgentAttachment): string {
    const label =
      LEGACY_ATTACHMENT_LABELS[attachment.type] ||
      this.getDefinition(attachment.type)?.description ||
      attachment.type
    const body = normalizeAttachmentText(attachment.payload)
    if (!body) return ''
    return `[${label}]\n${body}`
  }

  serializeManyForLegacyPrompt(attachments: readonly AgentAttachment[]): string {
    return attachments
      .map((attachment) => this.serializeForLegacyPrompt(attachment))
      .filter(Boolean)
      .join('\n\n')
  }

  snapshot() {
    return {
      count: this.definitions.size,
      types: Array.from(this.definitions.keys()),
    }
  }
}
