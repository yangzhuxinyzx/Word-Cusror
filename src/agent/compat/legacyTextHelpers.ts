import { stripLegacyToolBlocks } from './legacyProtocol'
import { cleanLegacyModelOutput } from './legacyModelOutput'

export function cleanLegacyMessageForSend(content: string): string {
  const withoutModelMarkers = (content || '')
    .replace(/<\|.*?\|>/g, '')
    .replace(/<think>[\s\S]*?<\/think>/g, '')

  return stripLegacyToolBlocks(withoutModelMarkers).trim()
}

export function truncateLegacyTextForMemory(
  text: string,
  maxLen: number,
): string {
  const normalized = (text || '').replace(/\s+/g, ' ').trim()
  if (normalized.length <= maxLen) return normalized
  return normalized.slice(0, maxLen) + '...'
}

export function buildLegacyMemoryFlushText(
  recentMessages: Array<{ role: string; content: string }>,
  currentUserContent: string,
  maxChars = 1500,
): string {
  const lines: string[] = []
  const tail = recentMessages.slice(-6)
  tail.forEach((message) => {
    const label =
      message.role === 'user'
        ? 'User'
        : message.role === 'assistant'
          ? 'Assistant'
          : message.role
    lines.push(
      `${label}: ${truncateLegacyTextForMemory(message.content, 200)}`,
    )
  })
  if (currentUserContent) {
    lines.push(
      `Current request: ${truncateLegacyTextForMemory(currentUserContent, 300)}`,
    )
  }
  const summary = lines.join('\n')
  if (summary.length <= maxChars) return summary
  return `${summary.slice(0, maxChars)}\n...(truncated)`
}

export function buildLegacyToolFailureHint(
  latestDoc: string,
  search: string,
): string {
  const normalizedSearch = (search || '').trim().replace(/^['"]|['"]$/g, '')
  if (!normalizedSearch) return ''

  const lines = latestDoc
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)

  const keywordCandidates = Array.from(
    new Set([
      normalizedSearch,
      normalizedSearch.replace(/\s+/g, ''),
      normalizedSearch.slice(0, Math.min(8, normalizedSearch.length)),
      normalizedSearch.slice(Math.max(0, normalizedSearch.length - 8)),
    ]),
  ).filter((item) => item.length >= 2)

  const matched = lines
    .filter((line) =>
      keywordCandidates.some((keyword) => line.includes(keyword)),
    )
    .slice(0, 3)
    .map((line, index) => `${index + 1}. ${line.slice(0, 120)}`)

  if (matched.length === 0) {
    return 'Supplement: text not found in latest document snapshot. Try a shorter and exact source snippet as search.'
  }

  return `Supplement: related lines from current document\n${matched.join('\n')}`
}

export function extractLegacyAssistantText(content: string): string {
  return cleanLegacyModelOutput(stripLegacyToolBlocks(content))
}
