export interface PromptSection {
  id: string
  title: string
  content: string
  order?: number
  enabled?: boolean
}

export function createPromptSection(
  section: PromptSection,
): PromptSection {
  return section
}

export function sortPromptSections(
  sections: readonly PromptSection[],
): PromptSection[] {
  return [...sections].sort((left, right) => {
    const leftOrder = left.order ?? 0
    const rightOrder = right.order ?? 0
    return leftOrder - rightOrder
  })
}

export function renderPromptSections(
  sections: readonly PromptSection[],
): string {
  return sortPromptSections(sections)
    .filter((section) => section.enabled !== false && section.content.trim())
    .map((section) => `${section.title}\n${section.content.trim()}`)
    .join('\n\n')
}

export function createAttachmentUsageSection(): PromptSection {
  return createPromptSection({
    id: 'attachment-usage',
    title: '<attachment_protocol>',
    order: 100,
    content: [
      '附加上下文会以结构化 attachment block 形式注入。',
      '优先使用 attachment 中的内容，而不是猜测文档或工作区状态。',
      '不要逐字复述整块 attachment，除非用户明确要求。',
      '若 attachment 信息不足，再调用工具补充，而不是直接编造。',
      '</attachment_protocol>',
    ].join('\n'),
  })
}

export function createPromptBoundarySection(): PromptSection {
  return createPromptSection({
    id: 'prompt-boundary',
    title: '<runtime_contract>',
    order: 110,
    content: [
      '静态 system prompt、动态 prompt sections、attachments、tool descriptions 由 runtime 分层组装。',
      '当 attachment 与历史聊天冲突时，优先相信最新 attachment。',
      '工具结果会由 runtime 回填，不要假设旧的文本协议一定是唯一来源。',
      '</runtime_contract>',
    ].join('\n'),
  })
}
