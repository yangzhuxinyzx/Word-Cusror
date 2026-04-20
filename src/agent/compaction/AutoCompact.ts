export interface AutoCompactDecision {
  shouldCompact: boolean
  reason?: string
}

export function shouldAutoCompact(params: {
  messageCount: number
  attachmentChars: number
  thresholdMessages?: number
  thresholdAttachmentChars?: number
}): AutoCompactDecision {
  const thresholdMessages = params.thresholdMessages ?? 120
  const thresholdAttachmentChars = params.thresholdAttachmentChars ?? 40_000

  if (params.messageCount >= thresholdMessages) {
    return {
      shouldCompact: true,
      reason: `message_count:${params.messageCount}`,
    }
  }

  if (params.attachmentChars >= thresholdAttachmentChars) {
    return {
      shouldCompact: true,
      reason: `attachment_chars:${params.attachmentChars}`,
    }
  }

  return { shouldCompact: false }
}
