const SAFE_TOOL_NAME_PATTERN = /^[a-zA-Z0-9_-]+$/
const ENCODED_PREFIX = 'wc__'
const ENCODED_CHAR_PATTERN = /__x([0-9a-f]+)__/gi

function encodeUnsafeChar(char: string): string {
  return `__x${char.codePointAt(0)?.toString(16) || '0'}__`
}

export function encodeOpenAIToolName(toolName: string): string {
  if (!toolName) return toolName
  if (SAFE_TOOL_NAME_PATTERN.test(toolName)) return toolName

  return `${ENCODED_PREFIX}${Array.from(toolName)
    .map((char) => (SAFE_TOOL_NAME_PATTERN.test(char) ? char : encodeUnsafeChar(char)))
    .join('')}`
}

export function decodeOpenAIToolName(toolName: string): string {
  if (!toolName || !toolName.startsWith(ENCODED_PREFIX)) return toolName

  const encoded = toolName.slice(ENCODED_PREFIX.length)
  return encoded.replace(ENCODED_CHAR_PATTERN, (_, hex: string) => {
    const codePoint = parseInt(hex, 16)
    if (!Number.isFinite(codePoint)) return _
    try {
      return String.fromCodePoint(codePoint)
    } catch {
      return _
    }
  })
}

export function usesEncodedOpenAIToolName(toolName: string): boolean {
  return encodeOpenAIToolName(toolName) !== toolName
}
