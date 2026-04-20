export function extractLegacyThinking(content: string): {
  thinking: string
  cleaned: string
} {
  const thinkMatch = content.match(/<think>([\s\S]*?)<\/think>/g)
  let thinking = ''
  if (thinkMatch) {
    thinking = thinkMatch.map((m) => m.replace(/<\/?think>/g, '')).join('\n')
  }
  const cleaned = content.replace(/<think>[\s\S]*?<\/think>/g, '')
  return { thinking, cleaned }
}

export function extractLegacyStreamText(value: unknown): string {
  if (!value) return ''
  if (typeof value === 'string') return value
  if (Array.isArray(value)) {
    return value
      .map((part) => {
        if (!part) return ''
        if (typeof part === 'string') return part
        if (typeof (part as { text?: unknown }).text === 'string') {
          return (part as { text: string }).text
        }
        if (typeof (part as { content?: unknown }).content === 'string') {
          return (part as { content: string }).content
        }
        return ''
      })
      .join('')
  }
  if (typeof value === 'object') {
    const obj = value as { text?: unknown; content?: unknown }
    if (typeof obj.text === 'string') return obj.text
    if (typeof obj.content === 'string') return obj.content
  }
  return ''
}

export function cleanLegacyModelOutput(content: string): string {
  let cleaned = content
  cleaned = cleaned.replace(/<think>[\s\S]*?<\/think>/g, '')
  cleaned = cleaned.replace(/<\|.*?\|>/g, '')
  cleaned = cleaned.replace(/\n{3,}/g, '\n\n').trim()
  return cleaned || content
}

