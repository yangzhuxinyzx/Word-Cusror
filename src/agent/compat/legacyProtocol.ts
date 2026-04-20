import { DOC_EDIT_END, DOC_EDIT_START, DOC_SUMMARY_END, DOC_SUMMARY_START } from '../../utils/aiMarkers'
import {
  LEGACY_TOOL_NAMES,
  createLegacyXmlToolBlockRegex,
  createLegacyXmlToolOpenTagRegex,
} from './legacyTools'
import { createToolCallIR, type ToolCallIR, type ToolCallSource } from '../tools/ir'

const TOOL_CALL_BLOCK_REGEX = /\[TOOL_CALL\][\s\S]*?\[\/TOOL_CALL\]/g
const TOOL_RESULT_BLOCK_REGEX = /\[TOOL_RESULT\][\s\S]*?\[\/TOOL_RESULT\]/g
const TOOL_USE_BLOCK_REGEX = /<tool_use>[\s\S]*?<\/tool_use>/gi
const TOOL_USE_START_WITH_NAME_REGEX = /<tool_use>[\s\S]*?<tool_name>\s*([a-zA-Z0-9_.-]+)\s*<\/tool_name>/gi
const TOOL_USE_NAME_REGEX = /<tool_name>\s*([\s\S]*?)\s*<\/tool_name>/i
const TOOL_USE_PARAM_PAIR_REGEX = /<parameter_name>\s*([\s\S]*?)\s*<\/parameter_name>\s*<parameter_value>\s*([\s\S]*?)\s*<\/parameter_value>/gi
const TOOL_USE_INPUT_JSON_REGEX = /<tool_input>\s*([\s\S]*?)\s*<\/tool_input>/i
const LEGACY_XML_TOOL_ARG_REGEX = /<([a-zA-Z][\w-]*)>([\s\S]*?)<\/\1>/g

const EDIT_START_REGEX = new RegExp(escapeRegExp(DOC_EDIT_START), 'g')
const EDIT_END_REGEX = new RegExp(escapeRegExp(DOC_EDIT_END), 'g')
const SUMMARY_BLOCK_REGEX = new RegExp(
  `${escapeRegExp(DOC_SUMMARY_START)}[\\s\\S]*?(?:${escapeRegExp(DOC_SUMMARY_END)})?`,
  'g',
)

const LEGACY_XML_TOOL_BLOCK_REGEX = createLegacyXmlToolBlockRegex()
const LEGACY_XML_TOOL_OPEN_TAG_REGEX = createLegacyXmlToolOpenTagRegex()

export interface LegacyParsedToolCall {
  tool: string
  args: Record<string, string>
  source?: ToolCallSource
  rawInput?: string
}

export interface LegacyToolCallStartState {
  startedSignatures: Set<string>
}

export interface LegacyToolCallPreviewState {
  emitted: Set<string>
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

function cleanXmlTagText(value: string): string {
  return (value || '').replace(/<[^>]+>/g, '').trim()
}

function decodeXmlEntities(value: string): string {
  return (value || '')
    .replace(/&quot;/g, '"')
    .replace(/&#34;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#39;/g, "'")
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')
}

function normalizeToolUseArgValue(value: unknown): string {
  if (value === null || value === undefined) return ''
  if (typeof value === 'string') return value.trim()
  if (typeof value === 'number' || typeof value === 'boolean') return String(value)
  if (Array.isArray(value) || typeof value === 'object') {
    try {
      return JSON.stringify(value)
    } catch {
      return ''
    }
  }
  return String(value)
}

function extractToolUseInputArgs(block: string): Record<string, string> {
  const args: Record<string, string> = {}
  const inputMatch = block.match(TOOL_USE_INPUT_JSON_REGEX)
  if (!inputMatch) return args

  const rawInput = decodeXmlEntities((inputMatch[1] || '').trim())
  if (!rawInput) return args

  const candidate = rawInput
    .replace(/^```json\s*/i, '')
    .replace(/^```\s*/i, '')
    .replace(/\s*```$/i, '')
    .trim()

  if (!candidate) return args

  try {
    const parsed = JSON.parse(candidate)
    if (Array.isArray(parsed)) {
      args.ops = JSON.stringify(parsed)
      return args
    }
    if (parsed && typeof parsed === 'object') {
      for (const [key, value] of Object.entries(parsed as Record<string, unknown>)) {
        const normalized = normalizeToolUseArgValue(value)
        if (!key || !normalized) continue
        args[key] = normalized
      }
    }
  } catch {
    // Ignore malformed JSON and fallback to other parameter extraction paths.
  }

  return args
}

export function extractLegacyJsonObjectAfterKey(
  argsText: string,
  key: string,
): string | null {
  const keyRegex = new RegExp(`^\\s*${key}\\s*[:=]\\s*`, 'm')
  const match = keyRegex.exec(argsText)
  if (!match) return null

  const startIndex = match.index + match[0].length
  const openIndex = argsText.indexOf('{', startIndex)
  if (openIndex === -1) return null

  let depth = 0
  let inString = false
  let escaped = false

  for (let i = openIndex; i < argsText.length; i++) {
    const ch = argsText[i]
    if (inString) {
      if (escaped) {
        escaped = false
      } else if (ch === '\\') {
        escaped = true
      } else if (ch === '"') {
        inString = false
      }
      continue
    }

    if (ch === '"') {
      inString = true
      continue
    }

    if (ch === '{') depth++
    if (ch === '}') {
      depth--
      if (depth === 0) {
        return argsText.slice(openIndex, i + 1).trim()
      }
    }
  }

  return null
}

export function normalizeLegacyToolUseTagCalls(content: string): string {
  if (!content || content.indexOf('<tool_use>') === -1) return content

  return content.replace(TOOL_USE_BLOCK_REGEX, (block) => {
    const toolNameMatch = block.match(TOOL_USE_NAME_REGEX)
    let toolName = cleanXmlTagText(toolNameMatch?.[1] || '')
    if (!toolName) return ''

    const argsMap: Record<string, string> = {
      ...extractToolUseInputArgs(block),
    }

    const pairRegex = new RegExp(TOOL_USE_PARAM_PAIR_REGEX.source, 'gi')
    let pairMatch: RegExpExecArray | null
    while ((pairMatch = pairRegex.exec(block)) !== null) {
      const key = cleanXmlTagText(pairMatch[1] || '')
      const value = cleanXmlTagText(pairMatch[2] || '')
      if (!key || !value) continue
      argsMap[key] = value
    }

    if (!argsMap.action && argsMap.operation) argsMap.action = argsMap.operation
    if (!argsMap.search && argsMap.search_text) argsMap.search = argsMap.search_text
    if (!argsMap.replace && argsMap.replace_text) argsMap.replace = argsMap.replace_text

    if (toolName === 'word_edit_ops') {
      const operation = (argsMap.operation || argsMap.action || '').toLowerCase()
      const searchText = argsMap.search || argsMap.search_text || ''
      const replaceText = argsMap.replace || argsMap.replace_text || ''
      if (searchText && replaceText && (operation === 'search_replace' || operation === 'replace' || operation === 'find_replace' || !operation)) {
        toolName = 'replace'
        argsMap.search = searchText
        argsMap.replace = replaceText
      }
    }

    const argLines = Object.entries(argsMap)
      .filter(([key, value]) => key && value)
      .map(([key, value]) => `${key}: ${value}`)

    if (argLines.length === 0) {
      return `[TOOL_CALL] ${toolName}\n[/TOOL_CALL]`
    }

    return `[TOOL_CALL] ${toolName}\n${argLines.join('\n')}\n[/TOOL_CALL]`
  })
}

export function normalizeLegacyXmlToolCalls(content: string): string {
  if (!content || content.indexOf('<') === -1) return content

  return content.replace(LEGACY_XML_TOOL_BLOCK_REGEX, (_full, rawTool: string, rawBody: string) => {
    let toolName = (rawTool || '').trim()
    const body = rawBody || ''
    const args: string[] = []
    const argRegex = new RegExp(LEGACY_XML_TOOL_ARG_REGEX.source, 'g')
    let argMatch: RegExpExecArray | null

    while ((argMatch = argRegex.exec(body)) !== null) {
      const key = (argMatch[1] || '').trim()
      const value = (argMatch[2] || '').trim()
      if (!key || !value) continue
      args.push(`${key}: ${value}`)
    }

    if (toolName === 'word_edit_ops') {
      const hasSearch = args.some((line) => line.startsWith('search:'))
      const hasReplace = args.some((line) => line.startsWith('replace:'))
      const hasAction = args.some((line) => line.startsWith('action:'))
      if (hasSearch && hasReplace && !hasAction) {
        toolName = 'replace'
      }
    }

    if (args.length === 0) {
      const compactBody = body.trim()
      if (!compactBody) return ''
      args.push(`content: ${compactBody}`)
    }

    return `[TOOL_CALL] ${toolName}\n${args.join('\n')}\n[/TOOL_CALL]`
  })
}

export function trimTrailingOpenLegacyToolBlock(displayText: string): string {
  let cleaned = displayText || ''

  const lastOpenIdx = cleaned.lastIndexOf('[TOOL_CALL]')
  if (lastOpenIdx !== -1 && cleaned.indexOf('[/TOOL_CALL]', lastOpenIdx) === -1) {
    cleaned = cleaned.substring(0, lastOpenIdx).trim()
  }

  const lowered = cleaned.toLowerCase()
  const toolUseOpenIdx = lowered.lastIndexOf('<tool_use>')
  if (toolUseOpenIdx !== -1 && lowered.indexOf('</tool_use>', toolUseOpenIdx) === -1) {
    cleaned = cleaned.substring(0, toolUseOpenIdx).trim()
  }

  for (const tool of LEGACY_TOOL_NAMES) {
    const openTag = `<${tool}>`
    const closeTag = `</${tool}>`
    const openIdx = cleaned.lastIndexOf(openTag)
    if (openIdx === -1) continue
    if (cleaned.indexOf(closeTag, openIdx + openTag.length) === -1) {
      cleaned = cleaned.substring(0, openIdx).trim()
    }
  }

  return cleaned
}

export function stripLegacyToolBlocks(content: string): string {
  let cleaned = content || ''
  cleaned = cleaned.replace(TOOL_CALL_BLOCK_REGEX, '')
  cleaned = cleaned.replace(TOOL_RESULT_BLOCK_REGEX, '')
  cleaned = cleaned.replace(LEGACY_XML_TOOL_BLOCK_REGEX, '')
  cleaned = cleaned.replace(TOOL_USE_BLOCK_REGEX, '')
  return cleaned
}

function extractSummaryBlock(content: string): string {
  const startIndex = content.indexOf(DOC_SUMMARY_START)
  if (startIndex === -1) return ''
  const start = startIndex + DOC_SUMMARY_START.length
  const endIndex = content.indexOf(DOC_SUMMARY_END, start)
  const summary = endIndex === -1 ? content.slice(start) : content.slice(start, endIndex)
  return summary.trim()
}

export function parseLegacyAssistantOutput(content: string): {
  displayText: string
  summary: string
  phase: 'idle' | 'editing' | 'done'
} {
  let phase: 'idle' | 'editing' | 'done' = 'idle'
  if (content.includes(DOC_EDIT_START)) phase = 'editing'
  if (content.includes(DOC_EDIT_END)) phase = 'done'

  const summary = extractSummaryBlock(content)
  let displayText = content
  displayText = displayText.replace(SUMMARY_BLOCK_REGEX, '')
  displayText = displayText.replace(EDIT_START_REGEX, '').replace(EDIT_END_REGEX, '')
  displayText = stripLegacyToolBlocks(displayText)
  displayText = trimTrailingOpenLegacyToolBlock(displayText)
  displayText = displayText.replace(/\n{3,}/g, '\n\n').trim()

  return { displayText, summary, phase }
}

export function extractLegacyTextContent(content: string): string {
  let text = stripLegacyToolBlocks(content || '')
  text = text.replace(/\n{3,}/g, '\n\n').trim()
  return text
}

export function sanitizeLegacyAssistantText(content: string): string {
  return parseLegacyAssistantOutput(content).displayText
}

export function buildLegacyToolCallSignature(
  tool: string,
  args: Record<string, string>,
): string {
  const normalizedArgs = Object.keys(args || {})
    .sort()
    .map((key) => {
      const value = String(args[key] ?? '')
        .replace(/\s+/g, ' ')
        .trim()
      return `${key}=${value}`
    })

  return `${tool}::${normalizedArgs.join('||')}`
}

export function parseLegacyToolCalls(content: string): LegacyParsedToolCall[] {
  const toolCalls: LegacyParsedToolCall[] = []
  const toolUseNormalized = normalizeLegacyToolUseTagCalls(content)
  const normalizedContent = normalizeLegacyXmlToolCalls(toolUseNormalized)

  const toolCallRegex = /\[TOOL_CALL\]\s*([a-zA-Z0-9_.-]+)\s*\n([\s\S]*?)\[\/TOOL_CALL\]/g
  let match: RegExpExecArray | null

  while ((match = toolCallRegex.exec(normalizedContent)) !== null) {
    const toolName = match[1]
    const argsText = match[2]
    const args: Record<string, string> = {}

    if (toolName === 'create' || toolName === 'word.create') {
      const titleMatch = argsText.match(/^\s*title\s*[:=]\s*(.+?)(?:\n|$)/m)
      if (titleMatch) {
        args.title = titleMatch[1].trim()
      }

       const newTitleMatch = argsText.match(/^\s*newTitle\s*[:=]\s*(.+?)(?:\n|$)/m)
       if (newTitleMatch) {
         args.newTitle = newTitleMatch[1].trim()
       }

      const modeMatch = argsText.match(/^\s*mode\s*[:=]\s*(.+?)(?:\n|$)/m)
      if (modeMatch) {
        args.mode = modeMatch[1].trim()
      }

      const dslJson = extractLegacyJsonObjectAfterKey(argsText, 'dsl')
      if (dslJson) {
        args.dsl = dslJson
      }

      const elementsMatch = argsText.match(
        /^\s*elements\s*[:=]\s*(\[[\s\S]*?\])(?:\n|$)/m,
      )
      if (elementsMatch) {
        args.elements = elementsMatch[1].trim()
      }

      const styleRefPathMatch = argsText.match(
        /^\s*styleRefPath\s*[:=]\s*(.+?)(?:\n|$)/m,
      )
      if (styleRefPathMatch) {
        args.styleRefPath = styleRefPathMatch[1].trim()
      }
      const styleRefFileNameMatch = argsText.match(
        /^\s*styleRefFileName\s*[:=]\s*(.+?)(?:\n|$)/m,
      )
      if (styleRefFileNameMatch) {
        args.styleRefFileName = styleRefFileNameMatch[1].trim()
      }

      const templatePathMatch = argsText.match(
        /^\s*templatePath\s*[:=]\s*(.+?)(?:\n|$)/m,
      )
      if (templatePathMatch) {
        args.templatePath = templatePathMatch[1].trim()
      }
      const templateFileNameMatch = argsText.match(
        /^\s*templateFileName\s*[:=]\s*(.+?)(?:\n|$)/m,
      )
      if (templateFileNameMatch) {
        args.templateFileName = templateFileNameMatch[1].trim()
      }

      const contentRefPathMatch = argsText.match(
        /^\s*contentRefPath\s*[:=]\s*(.+?)(?:\n|$)/m,
      )
      if (contentRefPathMatch) {
        args.contentRefPath = contentRefPathMatch[1].trim()
      }
      const contentRefFileNameMatch = argsText.match(
        /^\s*contentRefFileName\s*[:=]\s*(.+?)(?:\n|$)/m,
      )
      if (contentRefFileNameMatch) {
        args.contentRefFileName = contentRefFileNameMatch[1].trim()
      }

      const contentMatch = argsText.match(/^\s*content\s*[:=]\s*([\s\S]*)$/m)
      if (contentMatch && !args.elements && !args.dsl) {
        let contentValue = contentMatch[1]
        const titleIndex = contentValue.indexOf('\ntitle:')
        if (titleIndex > -1) {
          contentValue = contentValue.substring(0, titleIndex)
        }
        args.content = contentValue.trim()
      }

      const replacementsMatch = argsText.match(
        /^\s*replacements\s*[:=]\s*(\[[\s\S]*?\])(?:\n|$)/m,
      )
      if (replacementsMatch) {
        args.replacements = replacementsMatch[1].trim()
      }
    } else if (
      toolName === 'copy_template' ||
      toolName === 'create_from_template'
    ) {
      const titleMatch = argsText.match(
        /^\s*newTitle\s*[:=]\s*(.+?)(?:\n|$)/m,
      )
      if (titleMatch) {
        args.newTitle = titleMatch[1].trim()
      }

      const replacementsMatch = argsText.match(
        /^\s*replacements\s*[:=]\s*(\[[\s\S]*?\])(?:\n|$)/m,
      )
      if (replacementsMatch) {
        args.replacements = replacementsMatch[1].trim()
      }

      const templatePathMatch = argsText.match(
        /^\s*templatePath\s*[:=]\s*(.+?)(?:\n|$)/m,
      )
      if (templatePathMatch) {
        args.templatePath = templatePathMatch[1].trim()
      }
      const templateFileNameMatch = argsText.match(
        /^\s*templateFileName\s*[:=]\s*(.+?)(?:\n|$)/m,
      )
      if (templateFileNameMatch) {
        args.templateFileName = templateFileNameMatch[1].trim()
      }
    } else if (toolName === 'word_edit_ops' || toolName === 'word.format') {
      const dryRunMatch = argsText.match(
        /^\s*dryRun\s*[:=]\s*(true|false)\s*(?:\n|$)/mi,
      )
      if (dryRunMatch) {
        args.dryRun = dryRunMatch[1].toLowerCase()
      }

      const modeMatch = argsText.match(/^\s*mode\s*[:=]\s*(.+?)(?:\n|$)/m)
      if (modeMatch) {
        args.mode = modeMatch[1].trim()
      }

      const opsMatch = argsText.match(/^\s*ops\s*[:=]\s*(\[[\s\S]*?\])(?:\n|$)/m)
      if (opsMatch) {
        args.ops = opsMatch[1].trim()
      }
    } else if (toolName === 'word.edit') {
      const dslJson = extractLegacyJsonObjectAfterKey(argsText, 'dsl')
      if (dslJson) {
        args.dsl = dslJson
      }

      const argLines = argsText.split('\n')
      for (const line of argLines) {
        const colonMatch = line.match(/^\s*(\w+)\s*[:=]\s*(.+?)\s*$/)
        if (colonMatch) {
          args[colonMatch[1]] = colonMatch[2]
        }
      }
    } else {
      const argLines = argsText.split('\n')
      for (const line of argLines) {
        const colonMatch = line.match(/^\s*(\w+)\s*[:=]\s*(.+?)\s*$/)
        if (colonMatch) {
          args[colonMatch[1]] = colonMatch[2]
        }
      }
    }

    if (
      (toolName === 'word_edit_ops' || toolName === 'word.format') &&
      !args.action &&
      !args.mode &&
      args.search &&
      args.replace
    ) {
      toolCalls.push({
        tool: 'replace',
        args: {
          search: args.search,
          replace: args.replace,
        },
        source: detectLegacyToolCallSource(content),
        rawInput: argsText,
      })
      continue
    }

    toolCalls.push({
      tool: toolName,
      args,
      source: detectLegacyToolCallSource(content),
      rawInput: argsText,
    })
  }

  return toolCalls
}

function detectLegacyToolCallSource(content: string): ToolCallSource {
  if (content.includes('[TOOL_CALL]')) return 'legacy_bracket'
  if (content.includes('<tool_use>')) return 'legacy_tool_use'
  if (content.includes('<')) return 'legacy_xml'
  return 'synthetic'
}

export function parseLegacyToolCallsToIR(
  content: string,
  options?: {
    turnId?: string
    metadata?: Record<string, unknown>
  },
): ToolCallIR[] {
  return parseLegacyToolCalls(content).map((call) =>
    createToolCallIR({
      toolName: call.tool,
      input: call.args,
      source: call.source || 'synthetic',
      rawInput: call.rawInput,
      turnId: options?.turnId,
      metadata: options?.metadata,
    }),
  )
}

export function parseLegacyToolCallsForCallbacks(
  content: string,
): Array<{ tool: string; args: Record<string, string> }> {
  return parseLegacyToolCalls(content).map((call) => ({
    tool: call.tool,
    args: call.args,
  }))
}

export function emitLegacyToolCallStartFromRaw(
  rawContent: string,
  state: LegacyToolCallStartState,
  onStart?: (tool: string) => void,
  maxStarts = Number.POSITIVE_INFINITY,
): void {
  if (!onStart) return

  const startCandidates: Array<{
    index: number
    tool: string
    signature: string
  }> = []

  const bracketStartRegex = /\[TOOL_CALL\]\s*([a-zA-Z0-9_.-]+)?/g
  let bracketMatch: RegExpExecArray | null
  while ((bracketMatch = bracketStartRegex.exec(rawContent)) !== null) {
    const tool = (bracketMatch[1] || '').trim()
    if (!tool) continue
    startCandidates.push({
      index: bracketMatch.index,
      tool,
      signature: `bracket:${bracketMatch.index}:${tool}`,
    })
  }

  const xmlStartRegex = new RegExp(LEGACY_XML_TOOL_OPEN_TAG_REGEX.source, 'gi')
  let xmlMatch: RegExpExecArray | null
  while ((xmlMatch = xmlStartRegex.exec(rawContent)) !== null) {
    const tool = (xmlMatch[1] || '').trim()
    if (!tool) continue
    startCandidates.push({
      index: xmlMatch.index,
      tool,
      signature: `xml:${xmlMatch.index}:${tool}`,
    })
  }

  const toolUseStartRegex = new RegExp(TOOL_USE_START_WITH_NAME_REGEX.source, 'gi')
  let toolUseMatch: RegExpExecArray | null
  while ((toolUseMatch = toolUseStartRegex.exec(rawContent)) !== null) {
    const tool = (toolUseMatch[1] || '').trim()
    if (!tool) continue

    let normalizedTool = tool
    if (tool === 'word_edit_ops') {
      const toolUseSnippet = rawContent
        .slice(toolUseMatch.index, toolUseMatch.index + 600)
        .toLowerCase()
      const looksLikeSearchReplace =
        toolUseSnippet.includes('search_replace') ||
        toolUseSnippet.includes('<parameter_name>search_text</parameter_name>') ||
        toolUseSnippet.includes('<parameter_name>replace_text</parameter_name>') ||
        toolUseSnippet.includes('"search_text"') ||
        toolUseSnippet.includes('"replace_text"')
      if (looksLikeSearchReplace) {
        normalizedTool = 'replace'
      }
    }

    startCandidates.push({
      index: toolUseMatch.index,
      tool: normalizedTool,
      signature: `tool_use:${toolUseMatch.index}:${normalizedTool}`,
    })
  }

  if (startCandidates.length === 0) return

  startCandidates.sort((a, b) => a.index - b.index)

  for (const candidate of startCandidates) {
    if (state.startedSignatures.size >= maxStarts) break
    if (state.startedSignatures.has(candidate.signature)) continue

    state.startedSignatures.add(candidate.signature)
    onStart(candidate.tool)
  }
}

export function emitLegacyToolCallPreviewFromRaw(
  rawContent: string,
  contentForParsing: string,
  state: LegacyToolCallPreviewState,
  onPreview?: (tool: string, args: Record<string, string>) => void,
  maxPreviews = Number.POSITIVE_INFINITY,
): void {
  if (!onPreview) return

  const calls = parseLegacyToolCalls(contentForParsing || rawContent)
  if (calls.length === 0) return

  for (const call of calls) {
    if (state.emitted.size >= maxPreviews) break

    const signature = buildLegacyToolCallSignature(call.tool, call.args)
    if (state.emitted.has(signature)) continue
    state.emitted.add(signature)
    onPreview(call.tool, { ...call.args })
  }
}

export function hasLegacyToolCall(content: string): boolean {
  if (!content) return false
  if (content.includes('[TOOL_CALL]')) return true
  if (LEGACY_XML_TOOL_OPEN_TAG_REGEX.test(content)) {
    LEGACY_XML_TOOL_OPEN_TAG_REGEX.lastIndex = 0
    return true
  }
  LEGACY_XML_TOOL_OPEN_TAG_REGEX.lastIndex = 0
  return /<tool_use>[\s\S]*?<\/tool_use>/i.test(content)
}
