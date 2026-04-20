import {
  createAttachmentUsageSection,
  createPromptBoundarySection,
  renderPromptSections,
  type PromptSection,
} from './PromptSections'

export interface ComposeSystemPromptOptions {
  basePrompt: string
  sections?: readonly PromptSection[]
  includeAttachmentProtocol?: boolean
}

export class SystemPromptComposer {
  compose(options: ComposeSystemPromptOptions): string {
    const sections: PromptSection[] = [...(options.sections || [])]
    if (options.includeAttachmentProtocol !== false) {
      sections.push(createAttachmentUsageSection(), createPromptBoundarySection())
    }

    const renderedSections = renderPromptSections(sections)
    if (!renderedSections.trim()) return options.basePrompt.trim()
    return `${options.basePrompt.trim()}\n\n${renderedSections}`
  }
}
