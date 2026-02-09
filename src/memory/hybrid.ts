import type { MemorySearchResult } from './schema'

export const formatMemoryResults = (results: MemorySearchResult[], maxChars = 2000) => {
  if (!results.length) return ''
  const lines = results.map((item, index) => {
    const location = item.startLine && item.endLine
      ? `（${item.startLine}-${item.endLine} 行）`
      : ''
    return `#${index + 1} [${item.source}] ${item.path}${location}\n${item.snippet}`
  })
  const output = lines.join('\n\n')
  if (output.length <= maxChars) return output
  return output.slice(0, maxChars) + '\n\n... (记忆结果已截断)'
}
