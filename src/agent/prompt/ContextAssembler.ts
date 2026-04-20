import type { AgentAttachment } from '../attachments/AttachmentTypes'
import { AttachmentManager } from '../attachments/AttachmentManager'

export interface AssembleAgentContextOptions {
  userInput: string
  skillDescriptions?: string
  activeSkill?: Record<string, unknown> | null
  availableToolsDelta?: Record<string, unknown> | null
  documentContext?: string
  filesContext?: string
  relevantMemories?: string
  workspaceProfile?: string
  pptEditContext?: Record<string, unknown> | null
  excelSheetContext?: Record<string, unknown> | null
}

export interface AssembledAgentContext {
  attachments: AgentAttachment[]
  userContent: string
}

export class ContextAssembler {
  constructor(private readonly attachmentManager: AttachmentManager) {}

  buildAttachments(
    options: AssembleAgentContextOptions,
  ): AgentAttachment[] {
    const attachments: AgentAttachment[] = []

    if (options.skillDescriptions?.trim()) {
      attachments.push(
        this.attachmentManager.createAttachment({
          type: 'skill_descriptions',
          payload: options.skillDescriptions,
          scope: 'session',
          kind: 'text',
        }),
      )
    }

    if (options.activeSkill) {
      attachments.push(
        this.attachmentManager.createAttachment({
          type: 'active_skill',
          payload: options.activeSkill,
          scope: 'turn',
          kind: 'json',
        }),
      )
    }

    if (options.availableToolsDelta) {
      attachments.push(
        this.attachmentManager.createAttachment({
          type: 'available_tools_delta',
          payload: options.availableToolsDelta,
          scope: 'turn',
          kind: 'json',
        }),
      )
    }

    if (options.documentContext?.trim()) {
      attachments.push(
        this.attachmentManager.createAttachment({
          type: 'current_document_summary',
          payload: options.documentContext,
          scope: 'document',
          kind: 'text',
        }),
      )
    }

    if (options.filesContext?.trim()) {
      attachments.push(
        this.attachmentManager.createAttachment({
          type: 'workspace_context_delta',
          payload: options.filesContext,
          scope: 'workspace',
          kind: 'text',
        }),
      )
    }

    if (options.relevantMemories?.trim()) {
      attachments.push(
        this.attachmentManager.createAttachment({
          type: 'relevant_memories',
          payload: options.relevantMemories,
          scope: 'turn',
          kind: 'text',
        }),
      )
    }

    if (options.workspaceProfile?.trim()) {
      attachments.push(
        this.attachmentManager.createAttachment({
          type: 'workspace_profile',
          payload: options.workspaceProfile,
          scope: 'workspace',
          kind: 'text',
        }),
      )
    }

    if (options.pptEditContext) {
      attachments.push(
        this.attachmentManager.createAttachment({
          type: 'ppt_edit_context',
          payload: options.pptEditContext,
          scope: 'document',
          kind: 'json',
        }),
      )
    }

    if (options.excelSheetContext) {
      attachments.push(
        this.attachmentManager.createAttachment({
          type: 'excel_sheet_context',
          payload: options.excelSheetContext,
          scope: 'document',
          kind: 'json',
        }),
      )
    }

    return attachments
  }

  assembleAgentContext(
    options: AssembleAgentContextOptions,
  ): AssembledAgentContext {
    const attachments = this.buildAttachments(options)
    const attachmentText =
      this.attachmentManager.serializeManyForLegacyPrompt(attachments)

    return {
      attachments,
      userContent: attachmentText
        ? `${options.userInput}\n\n${attachmentText}`
        : options.userInput,
    }
  }
}
