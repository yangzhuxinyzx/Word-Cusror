import { useState, useRef, useEffect, useCallback, useMemo, type ReactNode } from 'react'
import { 
  Send,
  FileText,
  X,
  Settings,
  Paperclip,
  CheckCircle,
  FileEdit,
  FilePlus,
  Eye,
  Loader2,
  CheckCircle2,
  Circle,
  Bot,
  Table,
  ImagePlus,
  Folder,
  Square
} from 'lucide-react'
import ReactMarkdown from 'react-markdown'
import remarkGfm from 'remark-gfm'
import { motion, AnimatePresence } from 'framer-motion'
import { useAI, ToolResult, type AgentDebugEvent } from '../context/AIContext'
import { useDocument } from '../context/useDocument'
import { FileItem, AgentStep, AgentFileChange, ChatMessage } from '../types'
import { runWebSearch, WebSearchResponse } from '../utils/webSearch'
import { parseDocxToHtmlForAgent } from '../utils/docxParser'
import { generateDocxAgentContextFromFilePath } from '../utils/docxAgentContext'
import { extractTypographyProfileFromArrayBuffer, formatTypographyProfileForAgent } from '../utils/docxTypography'
import { docxHtmlToElements, elementsToHtmlPreview } from '../utils/docxHtmlToElements'
import { validateDocDsl, dslToHtml } from '../utils/docDsl'
import { DOC_EDIT_START, DOC_EDIT_END, DOC_SUMMARY_START, DOC_SUMMARY_END } from '../utils/aiMarkers'
import type { DocDsl } from '../types/docDsl'
import type { ChartConfig, ChartSeries } from '../utils/chartParser'
import { toolCallLogger } from '../utils/toolCallLogger'
import { htmlToDsl } from '../utils/htmlToDsl'
import { serializeDslForAI } from '../utils/dslSerializer'
import { createLegacyXmlToolBlockRegex } from '../agent/compat/legacyTools'
import CinematicTyper from './CinematicTyper'
import { createPhaseOneToolExecutor } from '../agent/adapters/chatPanel/createPhaseOneToolExecutor'
import { createDocumentAgentAdapter } from '../agent/adapters/document/DocumentAgentAdapter'
import { ExcelDomainAdapter } from '../agent/adapters/excel/ExcelDomainAdapter'
import { PptDomainAdapter } from '../agent/adapters/ppt/PptDomainAdapter'
import { WorkspaceDomainAdapter } from '../agent/adapters/workspace/WorkspaceDomainAdapter'
import { WorkspaceSummaryAdapter } from '../agent/adapters/workspace/WorkspaceSummaryAdapter'
import { BUILTIN_SKILLS } from '../agent/skills/builtinSkills'
import { SkillExecutor } from '../agent/skills/SkillExecutor'
import type { AgentSkillDefinition } from '../agent/skills/SkillRegistry'
import { SkillRegistry } from '../agent/skills/SkillRegistry'
import { WorkspaceSkillLoader } from '../agent/skills/WorkspaceSkillLoader'
import {
  parseSubagentCommand,
  SubagentWorkflowRunner,
} from '../agent/subagents/SubagentWorkflowRunner'

type PptOutlineSlideDraft = {
  pageNumber: number
  pageType?: string
  headline: string
  subheadline?: string
  bullets?: string[]
  footerNote?: string
  layoutIntent?: string
}

const TOOL_CALL_BLOCK_REGEX = /\[TOOL_CALL\][\s\S]*?\[\/TOOL_CALL\]/g
const TOOL_RESULT_BLOCK_REGEX = /\[TOOL_RESULT\][\s\S]*?\[\/TOOL_RESULT\]/g
const LEGACY_XML_TOOL_BLOCK_REGEX = createLegacyXmlToolBlockRegex()
const TOOL_USE_BLOCK_REGEX = /<tool_use>[\s\S]*?<\/tool_use>/gi

const escapeRegExp = (value: string) => value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
const EDIT_START_REGEX = new RegExp(escapeRegExp(DOC_EDIT_START), 'g')
const EDIT_END_REGEX = new RegExp(escapeRegExp(DOC_EDIT_END), 'g')
const SUMMARY_BLOCK_REGEX = new RegExp(
  `${escapeRegExp(DOC_SUMMARY_START)}[\\s\\S]*?(?:${escapeRegExp(DOC_SUMMARY_END)})?`,
  'g'
)

const stripToolBlocks = (text: string) =>
  text
    .replace(TOOL_CALL_BLOCK_REGEX, '')
    .replace(TOOL_RESULT_BLOCK_REGEX, '')
    .replace(LEGACY_XML_TOOL_BLOCK_REGEX, '')
    .replace(TOOL_USE_BLOCK_REGEX, '')

const joinLocalPath = (dir: string, fileName: string) => {
  const cleanDir = String(dir || '').replace(/[\\/]+$/, '')
  const separator = cleanDir.includes('\\') && !cleanDir.includes('/') ? '\\' : '/'
  return `${cleanDir}${separator}${fileName}`
}

const stripMarkers = (text: string) =>
  text
    .replace(SUMMARY_BLOCK_REGEX, '')
    .replace(EDIT_START_REGEX, '')
    .replace(EDIT_END_REGEX, '')

const sanitizeAssistantText = (text: string) =>
  stripMarkers(stripToolBlocks(text)).replace(/\n{3,}/g, '\n\n').trim()

const unwrapOuterQuotes = (value: string): string => {
  const trimmed = (value || '').trim()
  const pairs: Array<[string, string]> = [
    ['"', '"'],
    ["'", "'"],
    ['“', '”'],
    ['‘', '’'],
    ['「', '」'],
    ['『', '』'],
    ['《', '》'],
  ]

  for (const [left, right] of pairs) {
    if (trimmed.length > left.length + right.length && trimmed.startsWith(left) && trimmed.endsWith(right)) {
      return trimmed.slice(left.length, trimmed.length - right.length).trim()
    }
  }

  return trimmed
}

const stripHtmlLikeSearch = (value: string): string => {
  return (value || '')
    .replace(/<br\s*\/?\s*>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .trim()
}

const normalizeSmartQuotesToAscii = (value: string): string => {
  return (value || '')
    .replace(/[“”]/g, '"')
    .replace(/[‘’]/g, "'")
}

const normalizeAsciiQuotesToCnPairs = (value: string): string => {
  let doubleOpen = true
  let singleOpen = true

  return (value || '').replace(/["']/g, (ch) => {
    if (ch === '"') {
      const token = doubleOpen ? '“' : '”'
      doubleOpen = !doubleOpen
      return token
    }

    const token = singleOpen ? '‘' : '’'
    singleOpen = !singleOpen
    return token
  })
}

const HALF_TO_FULL_PUNCT: Record<string, string> = {
  ',': '，',
  '.': '。',
  ':': '：',
  ';': '；',
  '!': '！',
  '?': '？',
  '(': '（',
  ')': '）',
  '[': '【',
  ']': '】',
}

const FULL_TO_HALF_PUNCT: Record<string, string> = {
  '，': ',',
  '。': '.',
  '：': ':',
  '；': ';',
  '！': '!',
  '？': '?',
  '（': '(',
  '）': ')',
  '【': '[',
  '】': ']',
}

const mapChars = (value: string, mapping: Record<string, string>): string => {
  if (!value) return ''
  return value
    .split('')
    .map((char) => mapping[char] || char)
    .join('')
}

const removeEllipsisMarkers = (value: string): string => {
  return (value || '').replace(/(?:\.\.\.|…)+/g, '').trim()
}

const buildReplaceSearchCandidates = (search: string): string[] => {
  const base = (search || '').trim()
  if (!base) return []

  const variants = [
    base,
    unwrapOuterQuotes(base),
    stripHtmlLikeSearch(base),
    stripHtmlLikeSearch(unwrapOuterQuotes(base)),
    normalizeSmartQuotesToAscii(base),
    normalizeSmartQuotesToAscii(stripHtmlLikeSearch(unwrapOuterQuotes(base))),
    normalizeAsciiQuotesToCnPairs(base),
    normalizeAsciiQuotesToCnPairs(stripHtmlLikeSearch(unwrapOuterQuotes(base))),
    mapChars(base, HALF_TO_FULL_PUNCT),
    mapChars(base, FULL_TO_HALF_PUNCT),
    mapChars(normalizeSmartQuotesToAscii(base), HALF_TO_FULL_PUNCT),
    mapChars(normalizeSmartQuotesToAscii(base), FULL_TO_HALF_PUNCT),
    removeEllipsisMarkers(base),
    removeEllipsisMarkers(stripHtmlLikeSearch(unwrapOuterQuotes(base))),
  ]

  const deduped: string[] = []
  const seen = new Set<string>()

  for (const candidate of variants) {
    const normalized = (candidate || '').trim()
    if (!normalized || normalized.length < 2) continue
    if (seen.has(normalized)) continue
    seen.add(normalized)
    deduped.push(normalized)
  }

  return deduped
}

// Keep full model/tool debug logs by default so trace files contain complete payloads.
const AGENT_DEBUG_MAX_CHARS: number | null = null

const truncateDebugText = (value: string, max: number | null = AGENT_DEBUG_MAX_CHARS): string => {
  const text = value || ''
  if (max == null || max <= 0) return text
  if (text.length <= max) return text
  const extra = text.length - max
  return `${text.slice(0, max)}\n... [truncated ${extra} chars]`
}

const stringifyDebugData = (value: unknown, max: number | null = AGENT_DEBUG_MAX_CHARS): string => {
  try {
    return truncateDebugText(JSON.stringify(value, null, 2), max)
  } catch (error) {
    return `Unserializable debug payload: ${String(error)}`
  }
}

const toDebugCodeBlock = (content: string, lang = 'text') => [
  `\`\`\`${lang}`,
  truncateDebugText(content),
  '```',
].join('\n')

const formatAgentDebugEventMarkdown = (event: AgentDebugEvent): string => {
  switch (event.type) {
    case 'turn_start':
      return [
        `## Turn ${event.turnId}`,
        `- time: ${event.timestamp}`,
        `- model: ${event.model}`,
        `- baseUrl: ${event.baseUrl}`,
        `- recentMessages: ${event.recentMessagesCount}`,
        `- hasDocumentContext: ${event.hasDocumentContext}`,
        `- hasFilesContext: ${event.hasFilesContext}`,
        `- imageCount: ${event.imageCount}`,
        '',
        '### User Input',
        toDebugCodeBlock(event.userInput || ''),
        '',
      ].join('\n')

    case 'api_response_raw': {
      const rawResponseValue = event.rawResponse ?? event.response ?? ''
      const rawResponse =
        typeof rawResponseValue === 'string'
          ? rawResponseValue
          : stringifyDebugData(rawResponseValue)
      const cleanedResponse = event.response || ''
      const isIdentical = rawResponse === cleanedResponse

      return [
        `### API Response [iter ${event.iteration}]`,
        `- time: ${event.timestamp}`,
        `- stage: ${event.stage}`,
        `- hasToolCall: ${event.hasToolCall}`,
        `- rawLength: ${rawResponse.length}`,
        `- cleanedLength: ${cleanedResponse.length}`,
        '',
        '#### Raw Response',
        toDebugCodeBlock(rawResponse),
        '',
        `#### Cleaned Response${isIdentical ? ' (same as raw)' : ''}`,
        toDebugCodeBlock(cleanedResponse),
        '',
      ].join('\n')
    }

    case 'tool_calls_parsed':
      return [
        `### Tool Calls Parsed [iter ${event.iteration}]`,
        `- time: ${event.timestamp}`,
        `- count: ${event.calls.length}`,
        toDebugCodeBlock(stringifyDebugData(event.calls), 'json'),
        '',
      ].join('\n')

    case 'tool_call_skipped':
      return [
        `### Tool Call Skipped [iter ${event.iteration}]`,
        `- time: ${event.timestamp}`,
        `- tool: ${event.tool}`,
        `- reason: ${event.reason}`,
        toDebugCodeBlock(stringifyDebugData({ args: event.args }), 'json'),
        '',
      ].join('\n')

    case 'tool_result':
      return [
        `### Tool Result [iter ${event.iteration}]`,
        `- time: ${event.timestamp}`,
        `- tool: ${event.tool}`,
        `- index: ${event.index}/${event.total}`,
        `- success: ${event.result.success}`,
        `- message: ${event.result.message}`,
        toDebugCodeBlock(stringifyDebugData({ args: event.args, result: event.result }), 'json'),
        '',
      ].join('\n')

    case 'final_summary':
      return [
        `### Final Summary [${event.source}]`,
        `- time: ${event.timestamp}`,
        `- iteration: ${event.iteration}`,
        toDebugCodeBlock(event.content || ''),
        '',
      ].join('\n')

    case 'turn_complete':
      return [
        `### Turn Complete`,
        `- time: ${event.timestamp}`,
        `- totalIterations: ${event.totalIterations}`,
        `- toolResults: ${event.toolResults.length}`,
        toDebugCodeBlock(stringifyDebugData(event.toolResults), 'json'),
        '',
        '---',
        '',
      ].join('\n')

    case 'turn_error':
      return [
        `### Turn Error`,
        `- time: ${event.timestamp}`,
        `- iteration: ${event.iteration}`,
        `- aborted: ${event.aborted}`,
        `- name: ${event.name || 'UnknownError'}`,
        `- message: ${event.message}`,
        toDebugCodeBlock(event.stack || ''),
        '',
        '---',
        '',
      ].join('\n')

    default:
      return ''
  }
}

type TableCardData = {
  headers: string[]
  rows: string[][]
}

const getMarkdownNodeText = (node: any): string => {
  if (!node) return ''
  if (node.type === 'text') return node.value || ''
  if (Array.isArray(node.children)) return node.children.map(getMarkdownNodeText).join('')
  return ''
}

const normalizeCellText = (text: string) => (text || '').replace(/\s+/g, ' ').trim()

const extractTableCardData = (node: any): TableCardData | null => {
  if (!node || node.type !== 'table' || !Array.isArray(node.children)) return null
  const rows: string[][] = node.children
    .filter((row: any) => row?.type === 'tableRow' && Array.isArray(row.children))
    .map((row: any) =>
      row.children.map((cell: any) => normalizeCellText(getMarkdownNodeText(cell)))
    )

  if (rows.length === 0) return null

  const bodyRows: string[][] = rows.slice(1)
  const maxCols = Math.max(0, ...rows.map((row) => row.length))
  if (maxCols === 0) return null

  let headers: string[] = rows[0] || []
  if (headers.every((h) => !h)) {
    headers = Array.from({ length: maxCols }, (_, i) => `字段${i + 1}`)
  } else if (headers.length < maxCols) {
    headers = headers.slice()
    for (let i = headers.length; i < maxCols; i += 1) {
      headers.push(`字段${i + 1}`)
    }
  }

  const normalizedRows: string[][] = bodyRows.map((row) => {
    const r = row.slice(0, maxCols)
    while (r.length < maxCols) r.push('')
    return r
  })

  return { headers, rows: normalizedRows }
}

const tryExtractDocDsl = (text: string): { dsl: DocDsl; sourceBlock: string } | null => {
  const candidates: Array<{ json: string; sourceBlock: string }> = []
  const trimmed = text.trim()
  if (trimmed.startsWith('{') && trimmed.endsWith('}')) {
    candidates.push({ json: trimmed, sourceBlock: trimmed })
  }

  const codeBlockRegex = /```(?:json)?\s*([\s\S]*?)```/g
  let match: RegExpExecArray | null
  while ((match = codeBlockRegex.exec(text)) !== null) {
    const json = match[1]?.trim()
    if (json) {
      candidates.push({ json, sourceBlock: match[0] })
    }
  }

  for (const candidate of candidates) {
    try {
      const dslObj = JSON.parse(candidate.json)
      const validation = validateDocDsl(dslObj)
      if (validation.valid) {
        return { dsl: dslObj, sourceBlock: candidate.sourceBlock }
      }
    } catch {
      // ignore invalid JSON
    }
  }
  return null
}

type PptOutlineDraft = {
  title?: string
  theme?: string
  styleHint?: string
  slides: PptOutlineSlideDraft[]
}

function stripPptOutlineJsonFromText(text: string): string {
  if (!text) return ''
  // remove fenced json first
  let out = text.replace(/```json\s*[\s\S]*?\s*```/gi, '').trim()
  // remove best-effort object containing slides/pages/outline/content array
  out = out.replace(/\{[\s\S]*?"(?:slides|pages|outline|content|page_title)"\s*:\s*[\[\{][\s\S]*?\}[\s\S]*?\}/gi, '').trim()
  // cleanup excessive blank lines
  out = out.replace(/\n{3,}/g, '\n\n').trim()
  return out
}

function tryParsePptOutlineDraft(text: string): { draft: PptOutlineDraft; rawJson: string } | null {
  if (!text) return null

  const tryCandidates: string[] = []
  const fenced = text.match(/```json\s*([\s\S]*?)\s*```/i)
  if (fenced?.[1]) tryCandidates.push(fenced[1].trim())

  // best-effort: extract a JSON object that contains slides/pages/outline array
  const idx = text.indexOf('{')
  const last = text.lastIndexOf('}')
  if (idx !== -1 && last !== -1 && last > idx) {
    const maybe = text.slice(idx, last + 1).trim()
    // 支持更多字段名：slides, pages, outline, content, 页面, 幻灯片 等
    if (/"(?:slides|pages|outline|content|页面|幻灯片|ppt_outline|ppt_pages)"\s*:\s*\[/i.test(maybe)) {
      tryCandidates.push(maybe)
    }
    // 也检测包含 page_title 的数组结构
    if (/"page_title"\s*:/i.test(maybe) && /\[\s*\{/.test(maybe)) {
      tryCandidates.push(maybe)
    }
  }

  // fallback regex to find any object containing slides/pages array
  const objMatch = text.match(/\{[\s\S]*?"(?:slides|pages|outline|content)"\s*:\s*\[[\s\S]*?\][\s\S]*?\}/i)
  if (objMatch?.[0]) tryCandidates.push(objMatch[0].trim())

  for (const cand of tryCandidates) {
    try {
      const parsedAny = JSON.parse(cand) as any
      if (!parsedAny || typeof parsedAny !== 'object') continue
      // support multiple field names for slides array
      const rawSlides = parsedAny.slides ?? parsedAny.pages ?? parsedAny.outline ?? parsedAny.content ?? parsedAny.页面 ?? parsedAny.幻灯片 ?? parsedAny.ppt_outline ?? parsedAny.ppt_pages
      if (!Array.isArray(rawSlides) || rawSlides.length === 0) continue

      const normalizedSlides: PptOutlineSlideDraft[] = rawSlides.map((s: any, idx: number) => {
        const pageNumberRaw =
          s?.pageNumber ?? s?.page ?? s?.pageIndex ?? s?.index ?? s?.no ?? s?.页码 ?? s?.页数 ?? idx + 1
        const pageNumber = typeof pageNumberRaw === 'number' ? pageNumberRaw : Number(pageNumberRaw) || idx + 1

        const headline =
          (s?.headline ?? s?.title ?? s?.heading ?? s?.标题 ?? s?.主标题 ?? s?.pageTitle ?? s?.page_title ?? s?.slidetitle ?? s?.slide_title ?? '').toString().trim()

        const subheadlineRaw = s?.subheadline ?? s?.subtitle ?? s?.副标题 ?? s?.subTitle ?? s?.subHeading ?? s?.sub_title
        const subheadline = subheadlineRaw ? subheadlineRaw.toString().trim() : undefined

        const bulletsRaw = s?.bullets ?? s?.points ?? s?.keyPoints ?? s?.mainPoints ?? s?.content_points ?? s?.contentPoints ?? s?.要点 ?? s?.内容 ?? s?.items ?? s?.key_points ?? s?.main_points
        const bullets = Array.isArray(bulletsRaw)
          ? bulletsRaw.map((b: any) => (b ?? '').toString().trim()).filter(Boolean)
          : undefined

        const footerRaw = s?.footerNote ?? s?.footer ?? s?.页脚 ?? s?.footnote
        const footerNote = footerRaw ? footerRaw.toString().trim() : undefined

        const layoutRaw = s?.layoutIntent ?? s?.layout ?? s?.布局 ?? s?.layoutHint
        const layoutIntent = layoutRaw ? layoutRaw.toString().trim() : undefined

        const pageTypeRaw = s?.pageType ?? s?.type ?? s?.页类型
        const pageType = pageTypeRaw ? pageTypeRaw.toString().trim() : undefined

        return {
          pageNumber,
          pageType,
          headline,
          subheadline,
          bullets,
          footerNote,
          layoutIntent,
        }
      })

      // must have at least one slide with headline
      if (!normalizedSlides.some((s) => s.headline)) continue

      const draft: PptOutlineDraft = {
        title: (parsedAny.title ?? parsedAny.标题 ?? parsedAny.pptTitle ?? parsedAny.topic ?? '').toString().trim() || undefined,
        theme: (parsedAny.theme ?? parsedAny.主题 ?? parsedAny.topic ?? '').toString().trim() || undefined,
        styleHint: (parsedAny.styleHint ?? parsedAny.style ?? parsedAny.风格 ?? parsedAny.visualStyle ?? '').toString().trim() || undefined,
        slides: normalizedSlides.map((s, i) => ({ ...s, pageNumber: s.pageNumber || i + 1 })),
      }

      const rawJson = JSON.stringify(parsedAny, null, 2)
      return { draft, rawJson }
    } catch {
      // continue
    }
  }
  return null
}

// Framer Motion 变体配置 - 使用正确的 Easing 类型
const messageVariants = {
  hidden: { opacity: 0, y: 8 },
  visible: { 
    opacity: 1, 
    y: 0,
    transition: { duration: 0.25, ease: [0.25, 0.46, 0.45, 0.94] as const } // easeOut
  },
  exit: { 
    opacity: 0, 
    y: -4,
    transition: { duration: 0.15, ease: [0.55, 0.06, 0.68, 0.19] as const } // easeIn
  }
}

const streamingVariants = {
  hidden: { opacity: 0, y: 4 },
  visible: { 
    opacity: 1, 
    y: 0,
    transition: { duration: 0.2, ease: [0.25, 0.46, 0.45, 0.94] as const }
  }
}

const controlBarVariants = {
  hidden: { opacity: 0, y: 4, scale: 0.95 },
  visible: { 
    opacity: 1, 
    y: 0, 
    scale: 1,
    transition: { duration: 0.2, ease: [0.25, 0.46, 0.45, 0.94] as const }
  },
  exit: { 
    opacity: 0, 
    y: -4, 
    scale: 0.95,
    transition: { duration: 0.15, ease: [0.55, 0.06, 0.68, 0.19] as const }
  }
}

type ToolActivityItem = {
  id: string
  tool: string
  label: string
  status: 'running' | 'success' | 'error' | 'skipped'
  detail?: string
  searchText?: string
  replaceText?: string
}

type StreamItem =
  | { type: 'text'; id: string; content: string }
  | { type: 'tool'; id: string; data: ToolActivityItem }

const truncateLabel = (text: string, limit = 32) => {
  if (!text) return ''
  return text.length > limit ? `${text.slice(0, limit)}…` : text
}

const buildToolCallSignature = (tool: string, args: Record<string, string>) => {
  const normalizedArgs = Object.keys(args || {})
    .sort()
    .map((key) => {
      const value = String(args[key] ?? '').replace(/\s+/g, ' ').trim()
      return `${key}=${value}`
    })
  return `${tool}::${normalizedArgs.join('||')}`
}

const formatSearchResults = (response: WebSearchResponse, query: string) => {
  const sections = response.sections
  const webResults = sections?.web ?? response.results ?? []
  const lines: string[] = []

  if (webResults.length > 0) {
    lines.push('【Brave Web】')
    lines.push(
      webResults
        .map((item, index) => {
          const snippet = item.snippet ? item.snippet.replace(/\s+/g, ' ').trim() : ''
          return `${index + 1}. ${item.title}\n${item.link}\n${snippet}`
        })
        .join('\n\n')
    )
  }

  if (sections?.faq?.length) {
    const faqBlock = sections.faq
      .slice(0, 3)
      .map((faq, idx) => `Q${idx + 1}: ${faq.question}\nA: ${faq.answer}`)
      .join('\n\n')
    lines.push('【FAQ】')
    lines.push(faqBlock)
  }

  if (sections?.news?.length) {
    const newsBlock = sections.news
      .slice(0, 3)
      .map((news) => `${news.title}${news.source ? ` - ${news.source}` : ''}\n${news.link}`)
      .join('\n\n')
    lines.push('【新闻】')
    lines.push(newsBlock)
  }

  if (sections?.videos?.length) {
    const videoBlock = sections.videos
      .slice(0, 2)
      .map(
        (video) =>
          `${video.title}${video.duration ? ` (${video.duration})` : ''}\n${video.link}`
      )
      .join('\n\n')
    lines.push('【视频】')
    lines.push(videoBlock)
  }

  if (sections?.discussions?.length) {
    const discussionBlock = sections.discussions
      .slice(0, 2)
      .map(
        (discussion) =>
          `${discussion.forumName ?? '讨论'}：${discussion.question ?? ''}\n${discussion.link}`
      )
      .join('\n\n')
    lines.push('【讨论】')
    lines.push(discussionBlock)
  }

  if (response.summarizerKey) {
    lines.push(`Summarizer key: ${response.summarizerKey}`)
  }

  return `【Brave 搜索】${query}\n\n${lines.join('\n\n')}`
}

export default function ChatPanel() {
  const {
    messages,
    isLoading,
    streamingContent,
    streamingReasoning,
    requestRunState,
    settings,
    addMessage,
    sendAgentMessage,
    clearMessages,
    stopGeneration,
    retryLastAgentRequest,
    recordRuntimeToolExecution,
    syncRuntimeTools,
    syncRuntimeSkills,
    getAgentRuntimeSnapshot,
    spawnRuntimeSubagent,
    startRuntimeSubagent,
    completeRuntimeSubagent,
    failRuntimeSubagent,
    cancelRuntimeSubagent,
    appendRuntimeSubagentTranscript,
    loadRuntimeSubagentTranscript,
    getLastKnowledgeHits,
    runRuntimeSubagentSession,
  } = useAI()

  const { 
    document, 
    createNewDocument,
    createDocumentFromDsl,
    isElectron, 
    currentFile, 
    replaceInDocument,
    insertInDocument,
    deleteInDocument,
    replaceViaDsl,
    insertViaDsl,
    deleteViaDsl,
    openFile, 
    files, 
    workspacePath,
    editorMode,
    setEditorMode,
    refreshFiles,
    silentSaveToFile,
    getTiptapDocumentStructure,
    replaceWithFormat,
    documentAgentApi,
    excelData,
    refreshExcelData,
    previewWordOps,
    applyWordOps,
    getLatestContent,
    scrollToDiffId,
  } = useDocument()
  const [input, setInput] = useState('')
  const [pendingImages, setPendingImages] = useState<string[]>([]) // 待发送的图片 base64 URL
  const [attachedFiles, setAttachedFiles] = useState<FileItem[]>([])
  const [attachedFolders, setAttachedFolders] = useState<FileItem[]>([])
  const [isDragOver, setIsDragOver] = useState(false)
  const messagesEndRef = useRef<HTMLDivElement>(null)
  const chatContainerRef = useRef<HTMLDivElement>(null)
  const userScrolledUpRef = useRef(false)
  const inputRef = useRef<HTMLTextAreaElement>(null)
  const [runtimeClock, setRuntimeClock] = useState(() => Date.now())
  const [outlineJsonOpen, setOutlineJsonOpen] = useState<Record<string, boolean>>({})
  const [pendingPptOutline, setPendingPptOutline] = useState<{
    draft: PptOutlineDraft
    rawJson: string
    sourceMessageId: string
  } | null>(null)
  const [pendingWordOps, setPendingWordOps] = useState<{
    ops: any[]
    previewMessage: string
    previewLines: string[]
  } | null>(null)
  const [wordOpsApplying, setWordOpsApplying] = useState(false)
  const [pptGenerating, setPptGenerating] = useState(false)

  useEffect(() => {
    if (!isLoading) return
    setRuntimeClock(Date.now())
    const timer = window.setInterval(() => {
      setRuntimeClock(Date.now())
    }, 1000)
    return () => window.clearInterval(timer)
  }, [isLoading])

  // 打开设置（由 App.tsx 监听 open-settings 事件来弹出 SettingsModal）
  const openSettings = useCallback(() => {
    window.dispatchEvent(new CustomEvent('open-settings'))
  }, [])
  
  // ========== PPT 编辑上下文（拖拽/框选嵌入） ==========
  const [pptEditContext, setPptEditContext] = useState<{
    pageNumber: number
    imageBase64: string
    regionRect?: { x: number; y: number; w: number; h: number }
    pptxPath?: string
    isRegion?: boolean // 是否是框选区域（vs 整页）
  } | null>(null)
  const [isPptDragOver, setIsPptDragOver] = useState(false)
  const pptDragCounterRef = useRef(0)
  
  // 跳转到编辑器中的修改位置
  const scrollToChange = useCallback((text: string) => {
    console.log('scrollToChange called with:', text)
    // 触发自定义事件，让 WordEditor 处理滚动和高亮
    const event = new CustomEvent('scroll-to-text', { 
      detail: { text },
      bubbles: true
    })
    console.log('Dispatching event:', event)
    window.dispatchEvent(event)
  }, [])

  const jumpToDiffChange = useCallback((diff: { diffId?: string; replaceText: string; searchText: string }) => {
    if (diff.diffId) {
      scrollToDiffId(diff.diffId)
      return
    }
    const fallbackText = diff.replaceText || diff.searchText
    if (fallbackText) {
      scrollToChange(fallbackText)
    }
  }, [scrollToChange, scrollToDiffId])
  
  // 打开创建的文档
  const openCreatedFile = useCallback(async (fileName: string) => {
    // 在文件列表中查找匹配的文件
    const findFile = (items: FileItem[]): FileItem | null => {
      for (const item of items) {
        if (item.type === 'file' && item.name === fileName) {
          return item
        }
        if (item.children) {
          const found = findFile(item.children)
          if (found) return found
        }
      }
      return null
    }
    
    let file = findFile(files)
    
    // 如果在列表中没找到，尝试直接构建路径
    if (!file && workspacePath) {
      const filePath = joinLocalPath(workspacePath, fileName)
      file = { name: fileName, path: filePath, type: 'file' }
    }
    
    if (file) {
      // 无论文件是否已打开，都重新加载它
      await openFile(file)
      
      // 滚动编辑器到顶部
      setTimeout(() => {
        const editorElement = window.document.querySelector('.word-editor-content')
        if (editorElement) {
          editorElement.scrollTo({ top: 0, behavior: 'smooth' })
        }
        // 也滚动父容器
        const wordPage = window.document.querySelector('.word-page')
        if (wordPage?.parentElement) {
          wordPage.parentElement.scrollTo({ top: 0, behavior: 'smooth' })
        }
      }, 100)
    }
  }, [files, openFile, workspacePath])
  
  // Agent 进度状态 - 直接在聊天中显示
  const [agentProgress, setAgentProgress] = useState<{
    isActive: boolean
    currentAction: string
    steps: AgentStep[]
    fileChanges: AgentFileChange[]
    startTime: number | null
    thinkingTime: number
  }>({
    isActive: false,
    currentAction: '',
    steps: [],
    fileChanges: [],
    startTime: null,
    thinkingTime: 0
  })
  const [toolActivity, setToolActivity] = useState<ToolActivityItem[]>([])
  const [streamItems, setStreamItems] = useState<StreamItem[]>([])
  const streamItemsRef = useRef<StreamItem[]>([])
  const previewToolActivityBySignatureRef = useRef<Map<string, string>>(new Map())
  const pendingPreviewToolIdsRef = useRef<string[]>([])
  const startedPreviewToolQueueRef = useRef<Map<string, string[]>>(new Map())

  // Keep ref in sync with state
  useEffect(() => {
    streamItemsRef.current = streamItems
  }, [streamItems])

  const resetToolActivity = useCallback(() => {
    setToolActivity([])
    setStreamItems([])
    streamItemsRef.current = []
    previewToolActivityBySignatureRef.current.clear()
    pendingPreviewToolIdsRef.current = []
    startedPreviewToolQueueRef.current.clear()
  }, [])

  const resetCurrentTurnToolActivity = useCallback(() => {
    // Keep previous stream cards in chat history; only reset current-turn matching state.
    previewToolActivityBySignatureRef.current.clear()
    pendingPreviewToolIdsRef.current = []
    startedPreviewToolQueueRef.current.clear()
  }, [])

  const clearTrailingStreamTextItems = useCallback(() => {
    setStreamItems((prev) => {
      let cut = prev.length
      while (cut > 0 && prev[cut - 1].type === 'text') {
        cut--
      }
      const next = prev.slice(0, cut)
      streamItemsRef.current = next
      return next
    })
  }, [])

  const removeLatestMatchingStreamTextItem = useCallback((targetText: string) => {
    const normalizedTarget = sanitizeAssistantText(stripToolBlocks(targetText || '')).trim()
    if (!normalizedTarget) return

    setStreamItems((prev) => {
      let targetIndex = -1
      for (let i = prev.length - 1; i >= 0; i--) {
        const item = prev[i]
        if (item.type !== 'text') continue
        const normalizedItem = sanitizeAssistantText(stripToolBlocks(item.content || '')).trim()
        if (normalizedItem === normalizedTarget) {
          targetIndex = i
          break
        }
      }

      if (targetIndex === -1) return prev
      const next = [...prev.slice(0, targetIndex), ...prev.slice(targetIndex + 1)]
      streamItemsRef.current = next
      return next
    })
  }, [])

  const flushUiFrame = useCallback(async () => {
    await new Promise<void>((resolve) => {
      let settled = false
      const finish = () => {
        if (settled) return
        settled = true
        resolve()
      }

      const timeoutId = window.setTimeout(finish, 80)

      if (typeof window !== 'undefined' && typeof window.requestAnimationFrame === 'function') {
        window.requestAnimationFrame(() => {
          window.clearTimeout(timeoutId)
          finish()
        })
        return
      }
      finish()
    })
  }, [])

  const handleNewConversation = useCallback(() => {
    clearMessages()
    resetToolActivity()

    setInput('')
    setAttachedFiles([])
    setAttachedFolders([])
    setPendingImages([])
    setPendingWordOps(null)
    setPendingPptOutline(null)
    setPptEditContext(null)
    setWordOpsApplying(false)

    setAgentProgress({
      isActive: false,
      currentAction: '',
      steps: [],
      fileChanges: [],
      startTime: null,
      thinkingTime: 0,
    })
  }, [clearMessages, resetToolActivity])

  const registerToolActivity = useCallback((tool: string, label: string, extra?: { searchText?: string; replaceText?: string }) => {
    const id = `${tool}-${Date.now()}-${Math.random().toString(16).slice(2)}`
    const item: ToolActivityItem = { id, tool, label, status: 'running', ...extra }
    setToolActivity(prev => [...prev, item])
    // Push tool card into streamItems for interleaved rendering
    setStreamItems(prev => [...prev, { type: 'tool', id, data: item }])
    return id
  }, [])

  const completeToolActivity = useCallback((id: string, status: 'success' | 'error' | 'skipped', detail?: string) => {
    setToolActivity(prev =>
      prev.map(item =>
        item.id === id ? { ...item, status, detail: detail ?? item.detail } : item
      )
    )
    // Keep streamItems tool card state in sync for interleaved rendering
    setStreamItems(prev => prev.map(item =>
      item.type === 'tool' && item.id === id
        ? { ...item, data: { ...item.data, status, detail: detail ?? item.data.detail } }
        : item
    ))
  }, [])

  const patchToolActivity = useCallback((id: string, patch: Partial<ToolActivityItem>) => {
    setToolActivity((prev) =>
      prev.map((item) => (item.id === id ? { ...item, ...patch } : item))
    )
    setStreamItems((prev) =>
      prev.map((item) =>
        item.type === 'tool' && item.id === id
          ? { ...item, data: { ...item.data, ...patch } }
          : item
      )
    )
  }, [])

  const enqueueStartedToolPreview = useCallback((tool: string, activityId: string) => {
    const queue = startedPreviewToolQueueRef.current.get(tool) || []
    if (!queue.includes(activityId)) {
      queue.push(activityId)
      startedPreviewToolQueueRef.current.set(tool, queue)
    }
  }, [])

  const shiftStartedToolPreview = useCallback((tool: string): string | null => {
    const queue = startedPreviewToolQueueRef.current.get(tool)
    if (!queue || queue.length === 0) return null
    const activityId = queue.shift() || null
    if (queue.length === 0) {
      startedPreviewToolQueueRef.current.delete(tool)
    } else {
      startedPreviewToolQueueRef.current.set(tool, queue)
    }
    return activityId
  }, [])

  const removeStartedToolPreview = useCallback((tool: string, activityId: string) => {
    const queue = startedPreviewToolQueueRef.current.get(tool)
    if (!queue || queue.length === 0) return
    const next = queue.filter((id) => id !== activityId)
    if (next.length === 0) {
      startedPreviewToolQueueRef.current.delete(tool)
    } else {
      startedPreviewToolQueueRef.current.set(tool, next)
    }
  }, [])

  const clearPreviewMappingsForActivityId = useCallback((activityId: string) => {
    for (const [signature, mappedId] of previewToolActivityBySignatureRef.current.entries()) {
      if (mappedId === activityId) {
        previewToolActivityBySignatureRef.current.delete(signature)
      }
    }
  }, [])

  const resolvePreviewToolMeta = useCallback((tool: string, args: Record<string, string>) => {
    if (tool === 'word.edit') {
      const operation = args.operation || args.mode || args.action || (
        args.reason || args.type
          ? 'review'
          : args.position || args.content || args.dsl
            ? 'insert'
            : args.target && !args.search
              ? 'delete'
              : 'replace'
      )
      if (operation === 'review') {
        const searchText = args.search || ''
        const reason = args.reason || searchText
        const type = args.type || 'review'
        return {
          label: `[${type}] ${truncateLabel(reason, 30)}`,
          searchText,
          replaceText: args.replace || '',
        }
      }

      if (operation === 'insert') {
        return {
          label: `Insert: ${args.position || 'after'}`,
          searchText: args.target || '',
          replaceText: args.content || '',
        }
      }

      if (operation === 'delete') {
        const target = args.target || ''
        return {
          label: `Delete: ${truncateLabel(target, 24)}`,
          searchText: target,
          replaceText: '',
        }
      }

      const searchText = args.search || ''
      return {
        label: `Replace: ${truncateLabel(searchText, 24)}`,
        searchText,
        replaceText: args.replace || '',
      }
    }

    if (tool === 'word.format') {
      let opsCount = 0
      if (args.ops) {
        try {
          const parsed = JSON.parse(args.ops)
          if (Array.isArray(parsed)) {
            opsCount = parsed.length
          }
        } catch {
          // ignore parse error for preview label
        }
      }
      const mode = (args.mode || args.action || '').toLowerCase()
      return {
        label: `WordFormat: ${mode === 'preview' ? 'preview' : mode === 'apply' ? 'apply' : opsCount > 0 ? `${opsCount} ops` : 'running'}`,
      }
    }

    if (tool === 'word.resolve_change') {
      const action = args.action || 'resolve'
      const target = args.changeId || args.id || (args.all === 'true' ? 'all' : '')
      return {
        label: `${action}: ${truncateLabel(target || 'pending changes', 24)}`,
        searchText: target,
        replaceText: '',
      }
    }

    if (tool === 'replace') {
      const searchText = args.search || ''
      return {
        label: `Replace: ${truncateLabel(searchText, 24)}`,
        searchText,
        replaceText: args.replace || '',
      }
    }

    if (tool === 'review') {
      const searchText = args.search || ''
      const reason = args.reason || searchText
      const type = args.type || 'review'
      return {
        label: `[${type}] ${truncateLabel(reason, 30)}`,
        searchText,
        replaceText: args.replace || '',
      }
    }

    if (tool === 'insert') {
      return {
        label: `Insert: ${args.position || 'after'}`,
        searchText: args.target || '',
        replaceText: args.content || '',
      }
    }

    if (tool === 'word_chart' || tool === 'word.chart') {
      return {
        label: `Chart: ${args.type || 'bar'}`,
        searchText: '',
        replaceText: args.title || args.type || 'chart',
      }
    }

    if (tool === 'delete') {
      const target = args.target || ''
      return {
        label: `Delete: ${truncateLabel(target, 24)}`,
        searchText: target,
        replaceText: '',
      }
    }

    if (tool === 'word_edit_ops') {
      let opsCount = 0
      if (args.ops) {
        try {
          const parsed = JSON.parse(args.ops)
          if (Array.isArray(parsed)) {
            opsCount = parsed.length
          }
        } catch {
          // ignore parse error for preview label
        }
      }
      return {
        label: `WordOps: ${opsCount > 0 ? `${opsCount} ops` : 'running'}`,
      }
    }

    const hint = args.search || args.target || args.title || ''
    return {
      label: `${tool}: ${truncateLabel(hint || 'running', 24)}`,
      searchText: args.search || args.target || '',
      replaceText: args.replace || args.content || '',
    }
  }, [])

  const registerToolStart = useCallback((tool: string) => {
    const trackedTools = new Set([
      'replace',
      'review',
      'insert',
      'delete',
      'word_edit_ops',
      'word_chart',
      'word.edit',
      'word.format',
      'word.resolve_change',
      'word.chart',
    ])
    if (!trackedTools.has(tool)) return

    const startLabel =
      tool === 'word_edit_ops'
        ? 'WordOps: running'
        : tool === 'word.format'
          ? 'WordFormat: running'
          : `${tool}: running`
    const activityId = registerToolActivity(tool, startLabel)
    enqueueStartedToolPreview(tool, activityId)

    if (!pendingPreviewToolIdsRef.current.includes(activityId)) {
      pendingPreviewToolIdsRef.current.push(activityId)
    }
  }, [enqueueStartedToolPreview, registerToolActivity])

  const registerToolPreview = useCallback((tool: string, args: Record<string, string>) => {
    const trackedTools = new Set([
      'replace',
      'review',
      'insert',
      'delete',
      'word_edit_ops',
      'word_chart',
      'word.edit',
      'word.format',
      'word.resolve_change',
      'word.chart',
    ])
    if (!trackedTools.has(tool)) return

    const signature = buildToolCallSignature(tool, args)
    if (previewToolActivityBySignatureRef.current.has(signature)) return

    const meta = resolvePreviewToolMeta(tool, args)
    const startedActivityId = shiftStartedToolPreview(tool)

    if (startedActivityId) {
      patchToolActivity(startedActivityId, {
        tool,
        label: meta.label,
        status: 'running',
        detail: undefined,
        searchText: meta.searchText,
        replaceText: meta.replaceText,
      })
      previewToolActivityBySignatureRef.current.set(signature, startedActivityId)
      return
    }

    const activityId = registerToolActivity(tool, meta.label, {
      searchText: meta.searchText,
      replaceText: meta.replaceText,
    })
    previewToolActivityBySignatureRef.current.set(signature, activityId)
    if (!pendingPreviewToolIdsRef.current.includes(activityId)) {
      pendingPreviewToolIdsRef.current.push(activityId)
    }
  }, [patchToolActivity, registerToolActivity, resolvePreviewToolMeta, shiftStartedToolPreview])

  const markToolPreviewSkipped = useCallback((tool: string, args: Record<string, string>, reason: string) => {
    const signature = buildToolCallSignature(tool, args)
    const mappedActivityId = previewToolActivityBySignatureRef.current.get(signature)
    const activityId = mappedActivityId || shiftStartedToolPreview(tool)
    if (!activityId) return

    clearPreviewMappingsForActivityId(activityId)
    removeStartedToolPreview(tool, activityId)
    const shortReason = reason.length > 30 ? `${reason.slice(0, 30)}...` : reason
    completeToolActivity(activityId, 'skipped', shortReason)
  }, [clearPreviewMappingsForActivityId, completeToolActivity, removeStartedToolPreview, shiftStartedToolPreview])

  const claimOrRegisterToolActivity = useCallback((
    tool: string,
    args: Record<string, string>,
    label: string,
    extra?: { searchText?: string; replaceText?: string }
  ) => {
    const signature = buildToolCallSignature(tool, args)
    const mappedActivityId = previewToolActivityBySignatureRef.current.get(signature)
    const previewActivityId = mappedActivityId || shiftStartedToolPreview(tool)

    if (previewActivityId) {
      clearPreviewMappingsForActivityId(previewActivityId)
      removeStartedToolPreview(tool, previewActivityId)
      patchToolActivity(previewActivityId, {
        tool,
        label,
        status: 'running',
        detail: undefined,
        searchText: extra?.searchText,
        replaceText: extra?.replaceText,
      })
      return previewActivityId
    }

    return registerToolActivity(tool, label, extra)
  }, [clearPreviewMappingsForActivityId, patchToolActivity, registerToolActivity, removeStartedToolPreview, shiftStartedToolPreview])

  useEffect(() => {
    let interval: NodeJS.Timeout
    if (agentProgress.startTime) {
      interval = setInterval(() => {
        setAgentProgress(prev => ({
          ...prev,
          thinkingTime: Math.floor((Date.now() - (prev.startTime || Date.now())) / 1000)
        }))
      }, 1000)
    }
    return () => clearInterval(interval)
  }, [agentProgress.startTime])

  // Agent 操作函数
  const startAgentProgress = useCallback((operation: 'create' | 'edit') => {
    const initialSteps: AgentStep[] = operation === 'edit' 
      ? [
          { id: '1', type: 'reading', description: '读取当前文档', status: 'running' },
          { id: '2', type: 'thinking', description: '分析修改需求', status: 'pending' },
          { id: '3', type: 'editing', description: '执行修改', status: 'pending' },
        ]
      : [
          { id: '1', type: 'thinking', description: '分析需求', status: 'running' },
          { id: '2', type: 'creating', description: '生成内容', status: 'pending' },
          { id: '3', type: 'editing', description: '写入文件', status: 'pending' },
        ]
    
    setAgentProgress({
      isActive: true,
      currentAction: operation === 'edit' ? '正在修改文档...' : '正在创建文档...',
      steps: initialSteps,
      fileChanges: [{ name: '当前文档', additions: 0, deletions: 0, status: 'pending', operations: [] }],
      startTime: Date.now(),
      thinkingTime: 0
    })
  }, [])

  const updateAgentAction = useCallback((action: string) => {
    setAgentProgress(prev => ({ ...prev, currentAction: action }))
  }, [])

  const completeAgentStep = useCallback(() => {
    setAgentProgress(prev => {
      const runningIndex = prev.steps.findIndex(s => s.status === 'running')
      if (runningIndex === -1) return prev
      
      const newSteps = [...prev.steps]
      newSteps[runningIndex] = { ...newSteps[runningIndex], status: 'completed', timestamp: new Date() }
      
      if (runningIndex + 1 < newSteps.length) {
        newSteps[runningIndex + 1] = { ...newSteps[runningIndex + 1], status: 'running' }
      }
      
      return { ...prev, steps: newSteps }
    })
  }, [])

  const updateAgentFile = useCallback((updates: Partial<AgentFileChange>) => {
    setAgentProgress(prev => ({
      ...prev,
      fileChanges: prev.fileChanges.map((f, i) => i === 0 ? { ...f, ...updates } : f)
    }))
  }, [])

  const addAgentFileOperation = useCallback((operation: string) => {
    setAgentProgress(prev => ({
      ...prev,
      fileChanges: prev.fileChanges.map((f, i) => 
        i === 0 ? { ...f, operations: [...(f.operations || []), operation] } : f
      )
    }))
  }, [])

  const finishAgentProgress = useCallback(() => {
    setAgentProgress(prev => ({
      ...prev,
      isActive: false,
      steps: prev.steps.map(s => ({ ...s, status: 'completed' as const, timestamp: s.timestamp || new Date() })),
      fileChanges: prev.fileChanges.map(f => ({ ...f, status: 'done' as const })),
      startTime: null
    }))
    // 不清空 toolActivity —— 卡片在完成后仍需保持可见，直到下次发送消息时才清除
  }, [])

  // ========== 直接执行 PPT 生成（确认按钮用） ==========
  const pptDomainAdapter = useMemo(() => {
    return new PptDomainAdapter({
      isElectron,
      currentFilePath: currentFile?.path || null,
      currentFileName: currentFile?.name || null,
      workspacePath,
      refreshFiles,
      openFile,
    })
  }, [
    isElectron,
    currentFile?.path,
    currentFile?.name,
    workspacePath,
    refreshFiles,
    openFile,
  ])

  const executePptCreate = useCallback(async (draft: PptOutlineDraft, rawJson: string) => {
    if (pptGenerating) return
    setPptGenerating(true)

    const title = (draft.title || '新建演示文稿').trim()
    const theme = (draft.theme || '').trim()
    const outline = rawJson

    // 添加用户确认消息
    addMessage({ role: 'user', content: `✅ 确认大纲，开始生成 PPT：${title}` })

    // 启动进度
    setAgentProgress({
      isActive: true,
      currentAction: '正在准备生成 PPT...',
      steps: [
        { id: '1', type: 'thinking', description: '分析大纲', status: 'completed', timestamp: new Date() },
        { id: '2', type: 'creating', description: 'Gemini 设计视觉', status: 'running' },
        { id: '3', type: 'editing', description: '生成图片', status: 'pending' },
        { id: '4', type: 'editing', description: '导出 PPTX', status: 'pending' },
      ],
      fileChanges: [{ name: `${title}.pptx`, additions: 0, deletions: 0, status: 'writing', operations: [] }],
      startTime: Date.now(),
      thinkingTime: 0
    })

    // 注意：这里必须用 try/finally 包住，避免任何早期异常导致 pptGenerating 卡住为 true
    let activityId: string | null = null
    try {
      console.log('[PPT] executePptCreate start:', { title, slideCount: draft.slides?.length || 0 })
      activityId = registerToolActivity('ppt_create', `PPT：${title.slice(0, 24)}`)

      if (!pptDomainAdapter.canUsePptTools()) {
        throw new Error('PPT 生成仅支持桌面版（Electron）')
      }

      // 输出路径
      const output = pptDomainAdapter.buildOutputPath(title)

      if (!output) {
        throw new Error('缺少工作区路径，请先打开一个文件夹')
      }

      const safeTitle = String(title).replace(/[<>:"/\\|?*]/g, '_').slice(0, 60) || '新建演示文稿'
      const { fileName: pptxName, outputPath } = output

      // 获取 API Keys
      const openRouterApiKey = settings?.openRouterApiKey || ''
      const mainApiKey = settings?.apiKey || ''
      const mainBaseUrl = settings?.baseUrl || ''
      if (!mainApiKey || !mainBaseUrl) {
        throw new Error('缺少主模型 API 配置。PPT 生图现在复用主模型 5.4 的 baseUrl 和 apiKey')
      }

      const estimatedSlideCount = draft.slides?.length || 3

      // ========== 阶段1：调用 Gemini 生成文生图提示词 ==========
      updateAgentAction(`正在让 Gemini 设计视觉风格...`)
      addAgentFileOperation(`PPT: 正在设计 ${estimatedSlideCount} 页视觉`)

      const geminiResult = await pptDomainAdapter.generatePromptSlides({
        title,
        outline,
        outputPath,
        slideCount: estimatedSlideCount,
        theme,
        style: draft.styleHint || '',
        pendingImages,
        settings,
        // 主模型回退参数（当没有 OpenRouter API Key 时使用）
        mainApiKey: settings?.apiKey || '',
        mainBaseUrl: settings?.baseUrl || '',
        mainModel: settings?.model || '',
      })

      if (!geminiResult.success || !geminiResult.slides) {
        throw new Error(`Gemini 生成提示词失败: ${geminiResult.error || '未知错误'}`)
      }

      const slides = geminiResult.slides.map((s) => ({
        prompt: s.prompt,
        negativePrompt: s.negativePrompt,
      }))

      // 更新进度
      completeAgentStep()
      updateAgentAction(`Gemini 设计完成，共 ${slides.length} 页，开始生成图片...`)
      addAgentFileOperation(`PPT: 生成 ${slides.length} 页图片`)

      // ========== 阶段2：调用主模型同源 images/generations 生成图片 ==========
      const negativeDefault =
        'watermark, logo, brand name text, badge, QR code, UI, screenshot, HUD, sci-fi interface, holographic UI, futuristic dashboard, neon cyberpunk, neon cyan, bright cyan, fluorescent cyan, neon teal, cheap turquoise, generic isometric city, isometric cityscape, circuit-board city, lowres, blurry, garbled Chinese, wrong characters, text distortion, misspelling, random letters, gibberish, extra text, english text, ugly typography, amateur layout, noisy background, oversaturated, cheap plastic, toy-like, glossy, harsh specular, overbloom, stock 3d icons, generic template, ai artifacts, uncanny'

      // 为每页 slide 添加大纲内容（用于后续编辑时恢复）
      const slidesWithContent = slides.map((s, idx) => {
        const draftSlide = draft.slides?.[idx]
        const chineseContent = draftSlide 
          ? [
              draftSlide.headline,
              draftSlide.subheadline,
              ...(draftSlide.bullets || []),
              draftSlide.footerNote
            ].filter(Boolean).join('\n')
          : ''
        return {
          prompt: s.prompt,
          negativePrompt: s.negativePrompt || negativeDefault,
          originalChineseContent: chineseContent,
        }
      })
      
      // 根据用户选择的模型决定分辨率（默认使用 gpt-image-2）
      const pptImageModel = settings?.pptImageModel || 'gpt-image-2'
      const imageSize = pptImageModel === 'gpt-image-2' ? '2048x1152' : '1024x1024'
      console.log(`[PPT] 使用生图模型: ${pptImageModel}`)

      const result = await pptDomainAdapter.generateDeck({
        outputPath,
        slides: slidesWithContent,
        settings,
        designConcept: geminiResult?.designConcept || '',
        colorPalette: geminiResult?.colorPalette || '',
        // 主模型 API（用于同源 images/generations 生图）
        mainApiKey,
        mainBaseUrl,
        dashscope: {
          apiKey: '',
          region: 'cn',
          size: imageSize,
          model: pptImageModel,
          promptExtend: false,
          watermark: false,
          negativePromptDefault: negativeDefault,
        },
        postprocess: { mode: 'letterbox' },
        repair: {
          enabled: !!openRouterApiKey, // 只有配置了 OpenRouter 才启用修复
          openRouterApiKey,
          model: 'google/gemini-3-pro-preview',
          maxAttempts: 2,
          deckContext: {
            designConcept: geminiResult?.designConcept || '',
            colorPalette: geminiResult?.colorPalette || '',
          },
        },
        outline: draft, // 传递完整大纲供后续编辑使用
      })

      if (!result.success || !result.path) {
        throw new Error(`PPT 生成失败: ${result.error || '未知错误'}`)
      }

      await pptDomainAdapter.refreshFiles()

      // 打开新生成的 PPT
      await pptDomainAdapter.openGeneratedDeck(pptxName, result.path)

      // 完成进度
      completeAgentStep()
      completeAgentStep()
      updateAgentFile({ additions: slides.length, status: 'done', name: pptxName })
      finishAgentProgress()
      completeToolActivity(activityId, 'success', `${slides.length} 页`)

      // 添加成功消息
      addMessage({
        role: 'assistant',
        content: `✅ PPT 生成完成！\n\n📄 \`${pptxName}\`\n\n共 ${slides.length} 页，已导出到工作区并自动打开。`
      })
    } catch (e: any) {
      console.error('PPT 生成失败:', e)
      if (activityId) completeToolActivity(activityId, 'error', '失败')
      finishAgentProgress()
      addMessage({
        role: 'assistant',
        content: `❌ PPT 生成失败：${e?.message || e}`
      })
    } finally {
      console.log('[PPT] executePptCreate end')
      setPptGenerating(false)
    }
  }, [pptGenerating, isElectron, currentFile, workspacePath, settings, addMessage, registerToolActivity, completeToolActivity, updateAgentAction, completeAgentStep, updateAgentFile, addAgentFileOperation, finishAgentProgress, refreshFiles, openFile, pendingImages, pptDomainAdapter])

  // ========== PPT 编辑：整页重做 / 局部编辑 ==========
  const [pptEditPending, setPptEditPending] = useState<{
    pptxPath: string
    pageNumbers: number[]
    mode: 'regenerate' | 'partial_edit'
  } | null>(null)
  const [pptEditFeedback, setPptEditFeedback] = useState('')

  const executePptEdit = useCallback(async (
    pptxPath: string,
    pageNumbers: number[],
    mode: 'regenerate' | 'partial_edit',
    feedback: string
  ) => {
    if (pptGenerating || !pptDomainAdapter.canUsePptTools()) return
    setPptGenerating(true)

    const modeLabel = mode === 'regenerate' ? '整页重做' : '局部编辑'
    const pagesLabel = pageNumbers.length === 1 ? `第 ${pageNumbers[0]} 页` : `${pageNumbers.length} 页`

    addMessage({
      role: 'user',
      content: `🎨 PPT ${modeLabel}：${pagesLabel}\n反馈：${feedback}`
    })

    // 立即添加一条 "正在处理" 的消息，让用户知道在工作
    addMessage({
      role: 'assistant',
      content: `⏳ 正在${modeLabel}中...\n\n🔄 Gemini 正在根据反馈重新设计第 ${pageNumbers.join('、')} 页...`,
    })

    const activityId = registerToolActivity('ppt_edit', `PPT ${modeLabel}：${pagesLabel}`)

    try {
      const openRouterApiKey = settings.openRouterApiKey || ''
      // 优先使用专门的 DashScope API Key
      const dashscopeApiKey = settings.dashscopeApiKey || settings.apiKey || ''

      if (!openRouterApiKey) {
        throw new Error('请先在 AI 设置中配置 OpenRouter API Key')
      }
      if (mode === 'partial_edit' && !dashscopeApiKey) {
        throw new Error('请先在 AI 设置中配置 DashScope API Key（局部编辑仍使用图像编辑接口）')
      }

      updateAgentAction(`正在${modeLabel}：${pagesLabel}...`)
      addAgentFileOperation(`PPT: ${modeLabel} ${pagesLabel}`)

      const result = await pptDomainAdapter.editSlides({
        pptxPath,
        pageNumbers,
        feedback,
        mode,
        settings,
        openRouterApiKey,
        dashscopeApiKey,
        mainApiKey: settings.apiKey || '',
        mainBaseUrl: settings.baseUrl || '',
        pptImageModel: settings.pptImageModel || 'gpt-image-2',
      })

      if (!result.success) {
        throw new Error(result.error || '编辑失败')
      }

      await pptDomainAdapter.refreshFiles()

      // 重新打开 PPT 以刷新预览，并跳转到被编辑的页面
      const pptxName = pptxPath.split(/[\\/]/).pop() || 'output.pptx'
      const firstEditedPage = (result.editedPages && result.editedPages.length > 0) ? result.editedPages[0] : pageNumbers[0]
      
      // 触发自定义事件通知 PptPreviewHtml 跳转到指定页
      window.dispatchEvent(new CustomEvent('ppt-jump-to-page', {
        detail: { pageNumber: firstEditedPage }
      }))
      
      await pptDomainAdapter.openGeneratedDeck(pptxName, result.path || pptxPath)

      completeToolActivity(activityId, 'success', `${result.editedPages?.length || pageNumbers.length} 页`)
      finishAgentProgress()

      addMessage({
        role: 'assistant',
        content: `✅ PPT ${modeLabel}完成！\n\n已更新：${(result.editedPages || pageNumbers).map(p => `第 ${p} 页`).join('、')}\n\n文件已自动刷新，已跳转到第 ${firstEditedPage} 页。`
      })
    } catch (e: any) {
      console.error('PPT 编辑失败:', e)
      completeToolActivity(activityId, 'error', '失败')
      finishAgentProgress()
      addMessage({
        role: 'assistant',
        content: `❌ PPT ${modeLabel}失败：${e?.message || e}`
      })
    } finally {
      setPptGenerating(false)
    }
  }, [pptGenerating, isElectron, settings, addMessage, registerToolActivity, completeToolActivity, updateAgentAction, addAgentFileOperation, finishAgentProgress, refreshFiles, openFile, pptDomainAdapter])

  // 监听 PPT 编辑请求事件
  useEffect(() => {
    const handlePptEditRequest = (event: CustomEvent<{
      pptxPath: string
      pageNumbers: number[]
      mode: 'regenerate' | 'partial_edit'
    }>) => {
      const { pptxPath, pageNumbers, mode } = event.detail
      setPptEditPending({ pptxPath, pageNumbers, mode })
      setPptEditFeedback('')
    }

    window.addEventListener('ppt-edit-request', handlePptEditRequest as EventListener)
    return () => {
      window.removeEventListener('ppt-edit-request', handlePptEditRequest as EventListener)
    }
  }, [])
  
  // 监听 PPT 框选区域事件（Ctrl+框选）
  useEffect(() => {
    const handleRegionSelected = (event: CustomEvent<{
      pageNumber: number
      regionBase64: string
      regionRect: { x: number; y: number; w: number; h: number }
      fullPageBase64: string
      pptxPath: string
    }>) => {
      const { pageNumber, regionBase64, regionRect, pptxPath } = event.detail
      setPptEditContext({
        pageNumber,
        imageBase64: regionBase64,
        regionRect,
        pptxPath,
        isRegion: true,
      })
      // 聚焦输入框
      inputRef.current?.focus()
    }
    
    window.addEventListener('ppt-region-selected', handleRegionSelected as EventListener)
    return () => {
      window.removeEventListener('ppt-region-selected', handleRegionSelected as EventListener)
    }
  }, [])

  useEffect(() => {
    if (!userScrolledUpRef.current) {
      messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' })
    }
  }, [messages, agentProgress, streamingContent, toolActivity]) // 更新依赖，使用 streamingContent

  // 自动识别"阶段1：PPT 大纲 JSON"
  useEffect(() => {
    // 如果正在生成 PPT，不要重新检测大纲（避免点击确认后提示条又弹出来）
    if (pptGenerating) return

    // 关键：向上回溯"最近一次包含大纲 JSON"的 assistant 消息
    for (let i = messages.length - 1; i >= 0; i--) {
      const m = messages[i]
      if (m?.role !== 'assistant') continue
      const parsed = tryParsePptOutlineDraft(m.content)
      if (!parsed) continue
      setPendingPptOutline((prev) => {
        if (prev?.sourceMessageId === m.id) return prev
        return { draft: parsed.draft, rawJson: parsed.rawJson, sourceMessageId: m.id }
      })
      break
    }
  }, [messages, pptGenerating])

  // 检测操作类型
  const detectOperation = (text: string): 'create' | 'edit' | 'analyze' | 'chat' => {
    // 创建类关键词 - 包含"总结文档"、"做一个总结"等需要创建新文件的操作
    const createKeywords = ['创建', '新建', '生成', '写一份', '帮我写', '起草', '撰写', '编写', '拟写', '拟定', '总结文档', '做一个总结', '做个总结', '写总结', '生成总结', '/会议纪要']
    // 编辑类关键词 - 包含快捷命令
    const editKeywords = [
      '修改', '编辑', '润色', '优化', '改成', '替换', '删除', '添加', '扩展', '精简', '翻译', '重写',
      '格式化', '统一格式', '编号', '标题编号', '公文格式', '转换为公文',
      '审查', '校对', '纠错', '检查文档', '审阅',
      '/润色', '/精简', '/翻译', '/格式化', '/编号', '/公文', '/总结', '/审查', '/校对'
    ]
    const analyzeKeywords = ['分析', '解释', '什么意思', '有哪些', '告诉我', '是什么', '检查', '论文检查']
    
    // 优先检测创建类（包括总结文档）
    if (createKeywords.some(k => text.includes(k))) return 'create'
    if (editKeywords.some(k => text.includes(k))) return 'edit'
    if (analyzeKeywords.some(k => text.includes(k))) return 'analyze'
    return 'chat'
  }

  const DOCX_AGENT_MAX_CHARS = 80_000
  const FILES_CONTEXT_MAX_CHARS = 160_000
  const WORKSPACE_CONTEXT_MAX_CHARS = 60_000
  const WORKSPACE_AUTO_SUMMARY_MAX_FILES = 3
  const WORKSPACE_SUMMARY_MAX_CHARS = 1600
  const WORKSPACE_READ_MAX_CHARS = 8000
  const WORKSPACE_PPTX_MAX_SLIDES = 6

  const docxAttachmentCacheRef = useRef<Map<string, { key: string; content: string }>>(new Map())

  const truncateWithNote = useCallback((text: string, maxLen: number, note: string) => {
    if (!text) return ''
    if (text.length <= maxLen) return text
    const suffix = `\n\n... (${note}，已截断，建议指定章节/标题再问)`
    const keep = Math.max(0, maxLen - suffix.length)
    return text.slice(0, keep) + suffix
  }, [])

  const fetchArrayBufferFromLocalFile = useCallback(async (filePath: string): Promise<ArrayBuffer> => {
    if (!window.electronAPI?.getLocalFileUrl) {
      throw new Error('getLocalFileUrl 不可用')
    }
    const url = await window.electronAPI.getLocalFileUrl(filePath)
    const resp = await fetch(url)
    if (!resp.ok) throw new Error(`读取文件失败（HTTP ${resp.status}）`)
    return await resp.arrayBuffer()
  }, [])

  const splitPath = (fullPath: string) => {
    const sep = fullPath.includes('\\') ? '\\' : '/'
    const idx = fullPath.lastIndexOf(sep)
    return {
      dir: idx >= 0 ? fullPath.slice(0, idx) : '',
      base: idx >= 0 ? fullPath.slice(idx + 1) : fullPath,
      sep,
    }
  }

  const createTemplateCopy = useCallback(async (outputName?: string) => {
    if (!isElectron || !window.electronAPI?.readFile || !window.electronAPI?.writeBinaryFile) {
      return { success: false, message: '当前环境不支持模板复制' }
    }
    if (!currentFile?.path) {
      return { success: false, message: '未找到当前文档路径' }
    }
    const ext = (currentFile.name.split('.').pop() || '').toLowerCase()
    if (ext !== 'docx') {
      return { success: false, message: '仅支持 .docx 模板自动填充' }
    }

    const { dir, base, sep } = splitPath(currentFile.path)
    const baseName = base.replace(/\.[^.]+$/, '')
    let fileName = (outputName || `${baseName}-已填充`).trim()
    if (!fileName.toLowerCase().endsWith('.docx')) {
      fileName = `${fileName}.docx`
    }

    let newPath = `${dir}${sep}${fileName}`
    if (window.electronAPI.getFileInfo) {
      try {
        const info = await window.electronAPI.getFileInfo(newPath)
        if (info?.success) {
          fileName = `${baseName}-已填充-${Date.now()}.docx`
          newPath = `${dir}${sep}${fileName}`
        }
      } catch {
        // ignore
      }
    }

    const readResult = await window.electronAPI.readFile(currentFile.path)
    if (!readResult?.success || !readResult.data) {
      return { success: false, message: '读取模板失败' }
    }

    const writeResult = await window.electronAPI.writeBinaryFile(newPath, readResult.data)
    if (!writeResult?.success) {
      return { success: false, message: '写入新文档失败' }
    }

    await refreshFiles()
    return { success: true, file: { name: fileName, path: newPath, type: 'file' as const } }
  }, [currentFile, isElectron, refreshFiles])

  const prepareTemplateFillOutput = useCallback(async (ops: any[]) => {
    const templateOp = ops.find((op) => op?.type === 'template_fill' && String(op.params?.output || '').toLowerCase() === 'new_doc')
    if (!templateOp) return { success: true }

    if (editorMode === 'onlyoffice') {
      setEditorMode('tiptap')
      await new Promise(resolve => setTimeout(resolve, 100))
    }

    const outputName = templateOp.params?.outputName || templateOp.params?.fileName || templateOp.params?.title
    const created = await createTemplateCopy(typeof outputName === 'string' ? outputName : undefined)
    if (!created.success || !created.file) {
      return { success: false, message: created.message || '创建新文档失败' }
    }

    await openFile(created.file as FileItem)
    await new Promise(resolve => setTimeout(resolve, 100))
    return { success: true, file: created.file }
  }, [createTemplateCopy, editorMode, openFile, setEditorMode])

  // 获取文件内容
  const getFileContent = useCallback(async (file: FileItem): Promise<string> => {
    if (isElectron && window.electronAPI) {
      const result = await window.electronAPI.readFile(file.path)
      if (result.success && result.data) {
        if (result.type === 'docx') {
          // 让 Agent 读取 DOCX 全文与格式（不内联图片，避免 base64 过大）
          const cacheKey = `${file.path}:${result.data.length}`
          const cached = docxAttachmentCacheRef.current.get(file.path)
          if (cached?.key === cacheKey) {
            return cached.content
          }

          try {
            const parsed = await parseDocxToHtmlForAgent(result.data)

            // Typography profile（主题字体/Normal/Heading 等摘要）
            let typographyText = ''
            try {
              const ab = await fetchArrayBufferFromLocalFile(file.path)
              const { profile, outline } = await extractTypographyProfileFromArrayBuffer(ab)
              typographyText = formatTypographyProfileForAgent(profile, outline)
            } catch (e) {
              // ignore
            }

            const ps = parsed.pageSettings
            const pageLines: string[] = []
            if (ps) {
              pageLines.push(
                `页面: ${ps.orientation || 'portrait'}, size(pt)=${ps.width}×${ps.height}, margin(pt)=T${ps.marginTop}/B${ps.marginBottom}/L${ps.marginLeft}/R${ps.marginRight}, header=${ps.headerHeight}, footer=${ps.footerHeight}`
              )
            }

            const images = parsed.images || []
            const imageLines = images.slice(0, 200).map((img, idx) => {
              const size = img.widthPx || img.heightPx ? `${img.widthPx || '?'}x${img.heightPx || '?'}` : '?'
              const target = img.target ? ` target=${img.target}` : ''
              const alt = img.alt ? ` alt=${img.alt}` : ''
              const floating = img.floating ? ' floating=1' : ''
              return `- #${idx + 1} rid=${img.rId}${target} size=${size}${floating}${alt}`
            })

            const sections: string[] = []
            sections.push(`【Word DOCX】${file.name}`)
            if (pageLines.length) {
              sections.push('')
              sections.push('【页面设置】')
              sections.push(pageLines.join('\n'))
            }
            if (typographyText) {
              sections.push('')
              sections.push('【字体/排版摘要】')
              sections.push(typographyText)
            }
            sections.push('')
            sections.push(`【图片】${images.length} 张（仅元信息，不含二进制）`)
            if (imageLines.length) sections.push(imageLines.join('\n'))
            sections.push('')
            sections.push('【全文 HTML】')
            sections.push(parsed.html || '<p></p>')

            const combined = sections.join('\n')
            const truncated = truncateWithNote(combined, DOCX_AGENT_MAX_CHARS, `Word 附件 ${file.name}`)
            docxAttachmentCacheRef.current.set(file.path, { key: cacheKey, content: truncated })
            return truncated
          } catch (e) {
            return `【Word DOCX】${file.name}\n\n⚠️ 解析失败：${String(e)}`
          }
        }

        return result.data
      }
    }
    return file.content || `[文件: ${file.name}]`
  }, [isElectron, fetchArrayBufferFromLocalFile, truncateWithNote])

  const normalizePath = (value: string) => value.replace(/\\/g, '/').toLowerCase()

  const getParentDir = (filePath: string) => {
    const normalized = filePath.replace(/\\/g, '/')
    const idx = normalized.lastIndexOf('/')
    if (idx <= 0) return filePath
    const dir = normalized.slice(0, idx)
    return filePath.includes('\\') ? dir.replace(/\//g, '\\') : dir
  }

  const debugLogQueueRef = useRef<Promise<void>>(Promise.resolve())

  const resolveAgentDebugLogPath = useCallback((): string | null => {
    const rootPath = workspacePath || (currentFile?.path ? getParentDir(currentFile.path) : '')
    if (!rootPath) return null

    const dateTag = new Date().toISOString().slice(0, 10).replace(/-/g, '')
    const fileName = `agent-trace-${dateTag}.md`
    const needsSeparator = !rootPath.endsWith('\\') && !rootPath.endsWith('/')
    const separator = rootPath.includes('\\') ? '\\' : '/'
    return `${rootPath}${needsSeparator ? separator : ''}${fileName}`
  }, [workspacePath, currentFile?.path])

  const appendAgentDebugLog = useCallback(async (content: string) => {
    if (!content || !isElectron || !window.electronAPI) return

    const logPath = resolveAgentDebugLogPath()
    if (!logPath) return

    debugLogQueueRef.current = debugLogQueueRef.current
      .then(async () => {
        if (window.electronAPI?.appendFile) {
          const appendResult = await window.electronAPI.appendFile(logPath, content)
          if (!appendResult.success) {
            console.warn('[AgentDebug] appendFile failed:', appendResult.error)
          }
          return
        }

        if (window.electronAPI?.readFile && window.electronAPI?.writeFile) {
          const readResult = await window.electronAPI.readFile(logPath)
          const previous = readResult.success ? (readResult.data || '') : ''
          const writeResult = await window.electronAPI.writeFile(logPath, previous + content)
          if (!writeResult.success) {
            console.warn('[AgentDebug] writeFile fallback failed:', writeResult.error)
          }
        }
      })
      .catch((error) => {
        console.warn('[AgentDebug] queue write failed:', error)
      })

    await debugLogQueueRef.current
  }, [isElectron, resolveAgentDebugLogPath])

  const workspaceSummaryAdapter = useMemo(() => {
    return new WorkspaceSummaryAdapter({
      isElectron,
      workspaceSummaryMaxChars: WORKSPACE_SUMMARY_MAX_CHARS,
      workspacePptxMaxSlides: WORKSPACE_PPTX_MAX_SLIDES,
      truncateWithNote,
      readFile: window.electronAPI?.readFile,
      getFileInfo: window.electronAPI?.getFileInfo,
      excelListSheets: window.electronAPI?.excelListSheets,
      excelReadCells: window.electronAPI?.excelReadCells,
    })
  }, [isElectron, truncateWithNote])

  const summarizeWorkspaceFile = useCallback(async (file: FileItem, options?: { maxChars?: number; maxSlides?: number; format?: string }) => {
    return workspaceSummaryAdapter.summarizeWorkspaceFile(file, options)
  }, [workspaceSummaryAdapter])

  const buildWorkspaceAutoSummaries = useCallback(async (flatFiles: FileItem[]) => {
    return workspaceSummaryAdapter.buildWorkspaceAutoSummaries(
      flatFiles,
      currentFile?.path,
      WORKSPACE_AUTO_SUMMARY_MAX_FILES,
    )
  }, [currentFile?.path, workspaceSummaryAdapter])

  const workspaceDomainAdapter = useMemo(() => {
    return new WorkspaceDomainAdapter({
      isElectron,
      workspacePath,
      currentFilePath: currentFile?.path || null,
      workspaceContextMaxChars: WORKSPACE_CONTEXT_MAX_CHARS,
      truncateWithNote,
      readFolder: window.electronAPI?.readFolder,
      buildWorkspaceAutoSummaries,
    })
  }, [buildWorkspaceAutoSummaries, currentFile?.path, isElectron, truncateWithNote, workspacePath])

  const buildWorkspaceContext = useCallback(async () => {
    return workspaceDomainAdapter.buildWorkspaceContext()
  }, [workspaceDomainAdapter])

  const buildWorkspaceProfile = useCallback(async () => {
    return workspaceDomainAdapter.buildWorkspaceProfile()
  }, [workspaceDomainAdapter])

  const excelDomainAdapter = useMemo(() => {
    return new ExcelDomainAdapter({
      currentFileName: currentFile?.name || null,
      currentFilePath: currentFile?.path || null,
      workspacePath,
      refreshExcelData,
      refreshFiles,
      openFile,
    })
  }, [
    currentFile?.name,
    currentFile?.path,
    workspacePath,
    refreshExcelData,
    refreshFiles,
    openFile,
  ])

  const documentAgentAdapter = useMemo(() => {
    return createDocumentAgentAdapter({
      attachedFiles,
      fileName: documentAgentApi.getFileName(),
      currentFile: documentAgentApi.getCurrentFile(),
      workspacePath: documentAgentApi.getWorkspacePath(),
      editorMode,
      setEditorMode,
      isElectron: documentAgentApi.isElectron(),
      currentFilePath: documentAgentApi.getCurrentFilePath(),
      silentSaveToFile,
      replaceViaDsl: documentAgentApi.replaceViaDsl,
      replaceWithFormat: documentAgentApi.replaceWithFormat,
      replaceInDocument: documentAgentApi.replaceInDocument,
      insertViaDsl: documentAgentApi.insertViaDsl,
      insertInDocument: documentAgentApi.insertInDocument,
      deleteViaDsl: documentAgentApi.deleteViaDsl,
      deleteInDocument: documentAgentApi.deleteInDocument,
      previewWordOps: documentAgentApi.previewOps,
      applyWordOps: documentAgentApi.applyOps,
      prepareTemplateFillOutput: documentAgentApi.prepareTemplateFillOutput,
      createNewDocument: documentAgentApi.createDocument,
      createDocumentFromDsl: documentAgentApi.createDocumentFromDsl,
      getTiptapDocumentStructure: documentAgentApi.readOutline,
      getSelectionSnapshot: documentAgentApi.readSelection,
      readDocumentDsl: documentAgentApi.readDocumentDsl,
      listPendingChanges: documentAgentApi.listPendingChanges,
      acceptChange: documentAgentApi.acceptChange,
      rejectChange: documentAgentApi.rejectChange,
      acceptAllChanges: documentAgentApi.acceptAllChanges,
      rejectAllChanges: documentAgentApi.rejectAllChanges,
    })
  }, [
    attachedFiles,
    documentAgentApi,
    editorMode,
    setEditorMode,
    silentSaveToFile,
  ])

  // 构建文件上下文
  const buildFilesContext = useCallback(async () => {
    if (attachedFiles.length === 0) return ''
    const contents: string[] = []
    let totalLen = 0
    for (const file of attachedFiles) {
      const content = await getFileContent(file)
      const block = `=== ${file.name} ===\n${content}`
      const remaining = FILES_CONTEXT_MAX_CHARS - totalLen
      if (remaining <= 0) {
        break
      }
      const out = block.length > remaining
        ? truncateWithNote(block, remaining, '附加文件上下文总长度')
        : block
      contents.push(out)
      totalLen += out.length
      if (block.length > remaining) break
    }
    return contents.join('\n\n')
  }, [attachedFiles, getFileContent, truncateWithNote])

  // 构建文件夹上下文（只提供路径，让 AI 用工具自行查看）
  const buildFoldersContext = useCallback(() => {
    if (attachedFolders.length === 0) return ''
    const lines = attachedFolders.map(f =>
      `📁 文件夹：${f.name}\n   路径：${f.path}\n   请使用 workspace_list 工具（folder 参数填此路径）查看文件列表，再用 workspace_read 读取需要的文件。`
    )
    return `=== 用户拖入的文件夹 ===\n${lines.join('\n\n')}\n\n⚠️ 不要一次性读取所有文件，请根据需要选择性查看。`
  }, [attachedFolders])

  const phaseOneToolExecutor = useMemo(() => {
    return createPhaseOneToolExecutor({
      word: {
        adapter: documentAgentAdapter,
        fileName: currentFile?.name || '褰撳墠鏂囨。',
        currentFile: currentFile || null,
        attachedFiles,
        workspacePath,
        editorMode,
        setEditorMode,
        flushUiFrame,
        isElectron,
        currentFilePath: currentFile?.path || null,
        silentSaveToFile,
        claimOrRegisterToolActivity,
        completeToolActivity,
        patchToolActivity,
        updateAgentAction,
        completeAgentStep,
        updateAgentFile,
        addAgentFileOperation,
        buildReplaceSearchCandidates,
        replaceViaDsl,
        replaceWithFormat,
        replaceInDocument,
        insertViaDsl,
        insertInDocument,
        deleteViaDsl,
        deleteInDocument,
        previewWordOps,
        applyWordOps,
        prepareTemplateFillOutput,
        setPendingWordOps,
        createNewDocument,
        createDocumentFromDsl,
        refreshFiles,
        openFile,
        resolveWorkspaceFile: (args) => workspaceDomainAdapter.resolveWorkspaceFile(args),
      },
      excel: {
        adapter: excelDomainAdapter,
        registerToolActivity,
        completeToolActivity,
        updateAgentAction,
        truncateLabel,
      },
      ppt: {
        adapter: pptDomainAdapter,
        settings,
        pendingImages,
        pptEditContext,
        registerToolActivity,
        completeToolActivity,
        updateAgentAction,
        completeAgentStep,
        updateAgentFile,
        addAgentFileOperation,
        finishAgentProgress,
        truncateLabel,
      },
      memory: {
        registerToolActivity,
        completeToolActivity,
        updateAgentAction,
        truncateLabel,
      },
      knowledge: {
        registerToolActivity,
        completeToolActivity,
        updateAgentAction,
        truncateLabel,
      },
      workspace: {
        workspaceReadMaxChars: WORKSPACE_READ_MAX_CHARS,
        workspaceSummaryMaxChars: WORKSPACE_SUMMARY_MAX_CHARS,
        workspacePptxMaxSlides: WORKSPACE_PPTX_MAX_SLIDES,
        getWorkspaceFolderPath: () => workspaceDomainAdapter.getWorkspaceFolderPath(),
        registerToolActivity,
        completeToolActivity,
        updateAgentAction,
        truncateLabel,
        buildWorkspaceIndex: (folderPath, refresh) =>
          workspaceDomainAdapter.buildWorkspaceIndex(folderPath, refresh),
        formatWorkspaceIndex: (flatFiles, folderPath) =>
          workspaceDomainAdapter.formatWorkspaceIndex(flatFiles, folderPath),
        resolveWorkspaceFile: (args) => workspaceDomainAdapter.resolveWorkspaceFile(args),
        summarizeWorkspaceFile,
        buildWorkspaceProfile: (folderPath, refresh) =>
          workspaceDomainAdapter.buildWorkspaceProfile(folderPath, refresh),
        openFile,
      },
      web: {
        registerToolActivity,
        completeToolActivity,
        updateAgentAction,
        truncateLabel,
      },
    })
  }, [
    attachedFiles,
    documentAgentAdapter,
    excelDomainAdapter,
    pptDomainAdapter,
    workspaceDomainAdapter,
    claimOrRegisterToolActivity,
    completeAgentStep,
    completeToolActivity,
    createDocumentFromDsl,
    createNewDocument,
    currentFile?.name,
    currentFile?.path,
    deleteInDocument,
    deleteViaDsl,
    editorMode,
    flushUiFrame,
    insertInDocument,
    insertViaDsl,
    pendingImages,
    patchToolActivity,
    pptEditContext,
    prepareTemplateFillOutput,
    previewWordOps,
    applyWordOps,
    refreshFiles,
    registerToolActivity,
    replaceInDocument,
    replaceViaDsl,
    replaceWithFormat,
    setEditorMode,
    settings,
    silentSaveToFile,
    summarizeWorkspaceFile,
    truncateLabel,
    updateAgentAction,
    updateAgentFile,
    addAgentFileOperation,
    finishAgentProgress,
    openFile,
  ])

  useEffect(() => {
    syncRuntimeTools(phaseOneToolExecutor.listDefinitions())
  }, [phaseOneToolExecutor, syncRuntimeTools])

  const [workspaceSkills, setWorkspaceSkills] = useState<AgentSkillDefinition[]>([])
  const [workspaceSkillWarnings, setWorkspaceSkillWarnings] = useState<string[]>([])
  const [workspaceSkillSources, setWorkspaceSkillSources] = useState<string[]>([])
  const workspaceSkillLoader = useMemo(() => {
    return new WorkspaceSkillLoader({
      readFile: window.electronAPI?.readFile,
      readFolder: window.electronAPI?.readFolder,
    })
  }, [isElectron])

  const workspaceSkillRoot = useMemo(() => {
    return workspacePath || workspaceDomainAdapter.getWorkspaceFolderPath()
  }, [workspaceDomainAdapter, workspacePath])

  useEffect(() => {
    let cancelled = false

    if (!isElectron || !workspaceSkillRoot) {
      setWorkspaceSkills([])
      setWorkspaceSkillWarnings([])
      setWorkspaceSkillSources([])
      syncRuntimeSkills([])
      return () => {
        cancelled = true
      }
    }

    workspaceSkillLoader
      .load(workspaceSkillRoot)
      .then((result) => {
        if (cancelled) return
        setWorkspaceSkills(result.skills)
        setWorkspaceSkillWarnings(result.warnings)
        setWorkspaceSkillSources(result.sources)
        syncRuntimeSkills(result.skills)
        result.warnings.forEach((warning) => {
          console.warn('[WorkspaceSkillLoader]', warning)
        })
      })
      .catch((error) => {
        if (cancelled) return
        console.warn('[WorkspaceSkillLoader] load failed:', error)
        setWorkspaceSkills([])
        setWorkspaceSkillWarnings([String(error)])
        setWorkspaceSkillSources([])
        syncRuntimeSkills([])
      })

    return () => {
      cancelled = true
    }
  }, [isElectron, workspaceSkillLoader, workspaceSkillRoot, syncRuntimeSkills])

  const skillRegistry = useMemo(() => {
    const registry = new SkillRegistry()
    registry.registerMany(BUILTIN_SKILLS)
    registry.registerMany(workspaceSkills)
    return registry
  }, [workspaceSkills])

  const skillExecutor = useMemo(() => {
    return new SkillExecutor(skillRegistry)
  }, [skillRegistry])

  const availableSkillDescriptions = useMemo(() => {
    return skillRegistry
      .list()
      .filter((skill) => !skill.hidden)
      .sort((left, right) => left.id.localeCompare(right.id))
      .map((skill) => {
        const commands = skill.invocation?.slashCommands?.length
          ? skill.invocation.slashCommands.map((command) => `/${command}`).join(', ')
          : `/${skill.id}`
        const source = skill.source || 'builtin'
        const safety = skill.safety || 'mutating'
        return `${commands} | ${source} | ${safety}\n${skill.description}`
      })
      .join('\n\n')
  }, [skillRegistry])

  const executeRuntimeTool = useCallback(async (
    tool: string,
    args: Record<string, string>,
  ): Promise<ToolResult> => {
    const phaseOneDelegatedResult = await phaseOneToolExecutor.execute(tool, args)
    if (phaseOneDelegatedResult) {
      const pipelineState = phaseOneToolExecutor.getLastPipelineState()
      if (pipelineState) {
        recordRuntimeToolExecution(pipelineState)
      }
      return phaseOneDelegatedResult
    }

    return {
      tool,
      success: false,
      message: `Unsupported tool in phase1 runtime: ${tool}`,
    }
  }, [phaseOneToolExecutor, recordRuntimeToolExecution])

  const executeSubagentTool = useCallback(async (
    tool: string,
    args: Record<string, string>,
  ) => {
    const result = await phaseOneToolExecutor.execute(tool, args)
    const pipelineState = phaseOneToolExecutor.getLastPipelineState()

    return {
      result: result || {
        tool,
        success: false,
        message: `Unsupported subagent tool: ${tool}`,
      },
      pipelineState,
    }
  }, [phaseOneToolExecutor])

  const subagentWorkflowRunner = useMemo(() => {
    return new SubagentWorkflowRunner({
      spawnSubagent: spawnRuntimeSubagent,
      startSubagent: startRuntimeSubagent,
      completeSubagent: completeRuntimeSubagent,
      failSubagent: failRuntimeSubagent,
      appendSubagentTranscript: appendRuntimeSubagentTranscript,
      invokeTool: executeSubagentTool,
      runSession: runRuntimeSubagentSession,
    })
  }, [
    appendRuntimeSubagentTranscript,
    completeRuntimeSubagent,
    executeSubagentTool,
    failRuntimeSubagent,
    runRuntimeSubagentSession,
    spawnRuntimeSubagent,
    startRuntimeSubagent,
  ])

  const [runtimeSnapshotState, setRuntimeSnapshotState] = useState(() =>
    getAgentRuntimeSnapshot(),
  )
  const [selectedSubagentId, setSelectedSubagentId] = useState<string | null>(null)
  const [selectedSubagentTranscript, setSelectedSubagentTranscript] = useState<ReturnType<typeof loadRuntimeSubagentTranscript>>(null)

  useEffect(() => {
    const updateSnapshot = () => {
      setRuntimeSnapshotState(getAgentRuntimeSnapshot())
    }

    updateSnapshot()
    const timer = window.setInterval(updateSnapshot, 1000)
    return () => {
      window.clearInterval(timer)
    }
  }, [getAgentRuntimeSnapshot])

  useEffect(() => {
    if (!selectedSubagentId) {
      setSelectedSubagentTranscript(null)
      return
    }
    setSelectedSubagentTranscript(loadRuntimeSubagentTranscript(selectedSubagentId))
  }, [loadRuntimeSubagentTranscript, runtimeSnapshotState, selectedSubagentId])

  const handleCancelSubagent = useCallback((subagentId: string) => {
    const cancelled = cancelRuntimeSubagent(subagentId)
    if (!cancelled) return
    addMessage({
      role: 'assistant',
      content: `Subagent ${cancelled.label} cancelled.`,
    })
  }, [addMessage, cancelRuntimeSubagent])

  const buildSubagentDocumentContext = useCallback(() => {
    if (!currentFile) return ''

    const lowerName = currentFile.name.toLowerCase()
    const isExcelFile = lowerName.endsWith('.xlsx') || lowerName.endsWith('.xls')

    if (isExcelFile && excelData?.sheets?.length) {
      const sheetNames = excelData.sheets.map((sheet) => sheet.name).join(', ')
      const firstSheet = excelData.sheets[0]
      const previewLines: string[] = []

      if (firstSheet?.cells?.length) {
        const maxRows = 5
        const maxCols = 6
        const cellMap = new Map<string, string>()
        firstSheet.cells.forEach((cell) => {
          if (cell.r < maxRows && cell.c < maxCols) {
            cellMap.set(
              `${cell.r}-${cell.c}`,
              cell.display || cell.w || String(cell.v || ''),
            )
          }
        })

        for (let row = 0; row < maxRows; row += 1) {
          const cols: string[] = []
          for (let col = 0; col < maxCols; col += 1) {
            cols.push(cellMap.get(`${row}-${col}`) || '')
          }
          if (cols.some(Boolean)) {
            previewLines.push(cols.join('\t'))
          }
        }
      }

      return [
        `Active workbook: ${currentFile.name}`,
        `Sheets: ${sheetNames}`,
        previewLines.length > 0
          ? `Preview of ${firstSheet?.name || 'Sheet1'}:\n${previewLines.join('\n')}`
          : '',
      ]
        .filter(Boolean)
        .join('\n\n')
    }

    return getTiptapDocumentStructure()
  }, [currentFile, excelData, getTiptapDocumentStructure])

  const buildSubagentFilesContext = useCallback(async (profileId: string) => {
    const blocks: string[] = []
    const attachedContext = await buildFilesContext()
    if (attachedContext) {
      blocks.push(attachedContext)
    }

    if (
      profileId === 'workspace-explore' ||
      profileId === 'ppt-builder' ||
      profileId === 'verification'
    ) {
      const workspaceContext = await buildWorkspaceContext()
      if (workspaceContext) {
        blocks.push(workspaceContext)
      }
    }

    return blocks.join('\n\n')
  }, [buildFilesContext, buildWorkspaceContext])

  const handleSend = useCallback(async () => {
    if ((!input.trim() && pendingImages.length === 0) || isLoading) return

    const userMessage = input.trim() || 'Please analyze the attached image(s).'
    setInput('')
    resetToolActivity()
    userScrolledUpRef.current = false

    // 保存 PPT 编辑上下文（如果有）并清除状态
    const currentPptEditContext = pptEditContext
    if (pptEditContext) {
      setPptEditContext(null)
    }
    
    let runtimeUserMessage = userMessage
    let activeSkill: AgentSkillDefinition | null = null
    let activeSkillToolIds: string[] | undefined
    const fileNames = attachedFiles.map(f => f.name).join(', ')
    const allImages = [
      ...(currentPptEditContext?.imageBase64 ? [currentPptEditContext.imageBase64] : []),
      ...pendingImages
    ]
    
    // 构建用户消息内容（包含 PPT 编辑上下文标记）
    let displayMessage = userMessage
    if (currentPptEditContext) {
      displayMessage = `🖼️ [第 ${currentPptEditContext.pageNumber} 页${currentPptEditContext.isRegion ? '（框选区域）' : ''}] ${userMessage}`
    } else if (attachedFiles.length > 0 || attachedFolders.length > 0) {
      const parts: string[] = []
      if (attachedFiles.length > 0) parts.push(`📎 ${fileNames}`)
      if (attachedFolders.length > 0) parts.push(`📁 ${attachedFolders.map(f => f.name).join(', ')}`)
      displayMessage = `${userMessage}\n${parts.join('  ')}`
    }
    
    // 添加用户消息
    addMessage({ 
      role: 'user', 
      content: displayMessage
    })

    const parsedSubagentCommand = parseSubagentCommand(userMessage)
    if (parsedSubagentCommand) {
      const runtimeSnapshot = getAgentRuntimeSnapshot()
      const profile = runtimeSnapshot.subagents.profiles.find(
        (item) => item.id === parsedSubagentCommand.profileId,
      )

      if (!profile) {
        addMessage({
          role: 'assistant',
          content: `Unknown subagent profile: ${parsedSubagentCommand.profileId}`,
        })
        if (pendingImages.length > 0) {
          setPendingImages([])
        }
        return
      }

      const requestedMode = parsedSubagentCommand.mode || profile.defaultMode
      const needsDocumentContext = [
        'doc-explore',
        'doc-editor',
        'verification',
        'excel-operator',
      ].includes(profile.id)
      const needsFilesContext =
        attachedFiles.length > 0 ||
        [
          'workspace-explore',
          'ppt-builder',
          'verification',
        ].includes(profile.id)
      const needsWorkspaceProfile = [
        'workspace-explore',
        'ppt-builder',
        'verification',
      ].includes(profile.id)

      const subagentDocumentContext = needsDocumentContext
        ? buildSubagentDocumentContext()
        : ''
      const subagentFilesContext = needsFilesContext
        ? await buildSubagentFilesContext(profile.id)
        : ''
      const subagentWorkspaceProfile = needsWorkspaceProfile
        ? await buildWorkspaceProfile()
        : ''
      const runPromise = subagentWorkflowRunner.run({
        profile,
        request: parsedSubagentCommand.request,
        mode: requestedMode,
        parentTurnId: runtimeSnapshot.conversation.currentTurnId,
        documentContext: subagentDocumentContext || undefined,
        filesContext: subagentFilesContext || undefined,
        workspaceProfile: subagentWorkspaceProfile || undefined,
        images: profile.id === 'ppt-builder' ? allImages : undefined,
      })

      const latestSubagent =
        getAgentRuntimeSnapshot().subagents.agents
          .filter((agent) => agent.profileId === profile.id)
          .sort((left, right) => right.createdAt.localeCompare(left.createdAt))[0] ||
        null

      if (requestedMode === 'background') {
        addMessage({
          role: 'assistant',
          content: latestSubagent
            ? `Started background subagent ${profile.displayName}.\n\nTask: ${latestSubagent.taskId}\nSubagent: ${latestSubagent.subagentId}`
            : `Started background subagent ${profile.displayName}.`,
        })
        if (pendingImages.length > 0) {
          setPendingImages([])
        }
        void runPromise
          .then((result) => {
            addMessage({
              role: 'assistant',
              content: `Background subagent ${result.subagent.label} finished.\n\n${result.summary}`,
            })
          })
          .catch((error) => {
            addMessage({
              role: 'assistant',
              content: `Background subagent ${profile.displayName} failed: ${error instanceof Error ? error.message : String(error)}`,
            })
          })
        return
      }

      try {
        const result = await runPromise
        addMessage({
          role: 'assistant',
          content: `Subagent ${result.subagent.label} finished.\n\n${result.summary}`,
        })
      } catch (error) {
        addMessage({
          role: 'assistant',
          content: `Subagent ${profile.displayName} failed: ${error instanceof Error ? error.message : String(error)}`,
        })
      }
      if (pendingImages.length > 0) {
        setPendingImages([])
      }
      return
    }

    const skillExecution = await skillExecutor.execute(userMessage, {
      invokeTool: executeRuntimeTool,
      buildWorkspaceProfile: (refresh) =>
        workspaceDomainAdapter.buildWorkspaceProfile(undefined, !!refresh),
      currentFileName: currentFile?.name || null,
      workspacePath: workspaceDomainAdapter.getWorkspaceFolderPath(),
      workspaceSkillWarnings,
      workspaceSkillSources,
    })

    if (skillExecution?.mode === 'handled') {
      addMessage({
        role: 'assistant',
        content:
          skillExecution.assistantMessage ||
          `Skill completed: ${skillExecution.skill.displayName}`,
      })
      return
    }

    if (skillExecution?.mode === 'transform' && skillExecution.transformedInput) {
      activeSkill = skillExecution.skill
      activeSkillToolIds = skillExecution.allowedToolIds
      runtimeUserMessage = skillExecution.transformedInput
    }

    const operation = detectOperation(runtimeUserMessage)

    if (pendingImages.length > 0) {
      setPendingImages([])
    }

    // 启动 Agent 进度（在聊天中显示）
    if (operation === 'create' || operation === 'edit') {
      startAgentProgress(operation)
    }

    // 构建附加文件上下文 - 不再自动清除附加文件，由用户手动取消
    const attachedContext = await buildFilesContext()

    const fileName = currentFile?.name || '当前文档'
    let totalReplacements = 0
    
    // 构建完整的文档上下文
    // 1. 当前编辑器中的文档内容（默认始终包含）
    // 2. 用户拖拽的附加文件内容
    // 3. 用户拖拽的文件夹路径（AI 用工具自行查看）
    const foldersContext = buildFoldersContext()
    let fullContext = [attachedContext, foldersContext].filter(Boolean).join('\n\n')
    // 给 Agent 的“当前文档内容”（由 AIContext 附加到 user 消息）；需要严格控大小，避免图片 base64 撑爆上下文
    let documentContextForAI: string | undefined
    
    // 检查是否是 Excel 文件
    const isExcelFile = currentFile?.name?.toLowerCase().endsWith('.xlsx') || currentFile?.name?.toLowerCase().endsWith('.xls')
    
    // 如果当前文件不在附加文件列表中，也把它的内容加进去
    const currentFileInAttached = attachedFiles.some(f => f.path === currentFile?.path)
    if (currentFile && !currentFileInAttached) {
      // 如果是 Excel 文件，提供 Excel 特定的上下文
      if (isExcelFile && excelData?.sheets) {
        const sheetNames = excelData.sheets.map(s => s.name).join(', ')
        const firstSheet = excelData.sheets[0]
        let preview = ''
        if (firstSheet?.cells) {
          // 构建简单的数据预览（前几行）
          const maxRows = 5
          const cellMap: Record<string, string> = {}
          firstSheet.cells.forEach(cell => {
            if (cell.r < maxRows) {
              const key = `${cell.r}-${cell.c}`
              cellMap[key] = cell.display || cell.w || String(cell.v || '')
            }
          })
          const rows: string[] = []
          for (let r = 0; r < maxRows; r++) {
            const cols: string[] = []
            for (let c = 0; c < 10; c++) {
              cols.push(cellMap[`${r}-${c}`] || '')
            }
            if (cols.some(c => c)) {
              rows.push(cols.join('\t'))
            }
          }
          if (rows.length > 0) {
            preview = '\n\n数据预览（前几行）：\n' + rows.join('\n')
          }
        }
        
        const excelContext = `=== ${currentFile.name} (Excel 表格) ===
【文件类型】Excel 电子表格 (.${currentFile.name.split('.').pop()})
【工作表】${sheetNames}
【当前工作表】${firstSheet?.name || 'Sheet1'}${preview}

⚠️ 重要提示：这是 Excel 文件！请使用 Excel 专用工具：
- 删除行：excel_delete_rows（参数：sheet, startRow, count）
- 插入行：excel_insert_rows（参数：sheet, startRow, count, data）
- 删除列：excel_delete_columns（参数：sheet, startCol, count）
- 插入列：excel_insert_columns（参数：sheet, startCol, count）
- 修改单元格：excel_write（参数：sheet, updates）
- 合并单元格：excel_merge（参数：sheet, range）
- 新建工作表：excel_add_sheet（参数：name）
- 删除工作表：excel_delete_sheet（参数：name）
- ⭐生成图表：excel_chart（参数：sheet, type, dataRange, title, position）
  - 用于数据可视化：饼图(pie)、柱状图(column)、折线图(line)等
  - sheet 必须填当前工作表名称：${firstSheet?.name || 'Sheet1'}

❌ 不要使用 word.read / word.create / word.edit / word.format / word.resolve_change / word.chart 这些 Word 工具！`
        
        fullContext = fullContext ? `${excelContext}\n\n${fullContext}` : excelContext
      } else {
        // Word/文本文档处理
        let docContent = document.content
        let docStructure = ''
        let onlyOfficeExtra = ''

        // 给 AI 的文档内容：移除图片 base64（否则会把 prompt 撑爆导致超时/失败）
        const sanitizeHtmlForAI = (html: string) => {
          const input = html || ''
          if (!input) return input

          // 1) 先把 data:image base64 替换掉（避免后续 <img> 标签匹配时字符串过大）
          let out = input.replace(
            /data:image\/[^;]+;base64,[A-Za-z0-9+/=]+/gi,
            '[omitted:data-image]'
          )

          // 2) 把 img 变成可读占位，保留少量元信息（rid/尺寸/alt）
          out = out.replace(/<img\b[^>]*>/gi, (tag) => {
            const getAttr = (name: string) => {
              const m =
                tag.match(new RegExp(`${name}=\"([^\"]*)\"`, 'i')) ||
                tag.match(new RegExp(`${name}='([^']*)'`, 'i'))
              return m?.[1] || ''
            }
            const rid = getAttr('data-rid')
            const w = getAttr('data-w')
            const h = getAttr('data-h')
            const alt = getAttr('alt')
            const floating = getAttr('data-floating')
            const parts: string[] = ['[图片]']
            if (rid) parts.push(`rid=${rid}`)
            if (w || h) parts.push(`size=${w || '?'}x${h || '?'}`)
            if (floating === '1') parts.push('floating=1')
            if (alt) parts.push(`alt=${alt}`)
            return parts.join(' ')
          })

          return out
        }
        
        // AI 始终使用内置编辑器（Tiptap）的内容和结构
        // 这样可以保证 AI 编辑功能的稳定性
        try {
          const structure = getTiptapDocumentStructure()
          if (structure) {
            docStructure = '\n\n' + structure
          }
        } catch (e) {
          console.log('获取文档结构失败')
        }

        // ONLYOFFICE 预览模式下：补充主题/分节/样式 JSON（默认不包含全文 content，避免超大）
        if (editorMode === 'onlyoffice' && window.onlyOfficeConnector?.getDocumentJson) {
          try {
            const json = await window.onlyOfficeConnector.getDocumentJson({
              writeDefaultTextPr: true,
              writeDefaultParaPr: true,
              writeTheme: true,
              writeSectionPr: true,
              writeNumberings: false,
              writeStyles: true,
              includeContent: false
            })
            if (json && json.trim()) {
              const clipped = truncateWithNote(json, 50_000, 'ONLYOFFICE 文档 JSON')
              onlyOfficeExtra = `\n\n【ONLYOFFICE 文档 JSON（主题/分节/样式）】\n${clipped}`
            }
          } catch (e) {
            // ignore
          }
        }
        
        if (docContent) {
          // 给 AI：将 HTML 转为 DSL JSON 发给模型（结构化、token 更少、模型理解更好）
          try {
            const docDsl = htmlToDsl(docContent, { stripDiffMarkers: true })
            documentContextForAI = truncateWithNote(
              serializeDslForAI(docDsl),
              120_000,
              '当前文档内容（DSL 格式）'
            )
          } catch (e) {
            // DSL 转换失败时回退到 HTML
            console.warn('[ChatPanel] htmlToDsl failed, falling back to sanitizeHtmlForAI:', e)
            documentContextForAI = truncateWithNote(
              sanitizeHtmlForAI(docContent),
              120_000,
              '当前文档 HTML（已移除图片 base64）'
            )
          }

          const formatNote = '\n\n[提示：文档内容以 DSL JSON 格式提供，每个块有 _i 索引。编辑时可使用 blockIndex 精确定位。' +
            (editorMode === 'onlyoffice' ? ' 当前预览模式为 ONLYOFFICE。]' : ']')
          const currentFileContext = `=== ${currentFile.name} (当前编辑) ===\n${docStructure}${onlyOfficeExtra}${formatNote}`
          fullContext = fullContext ? `${currentFileContext}\n\n${fullContext}` : currentFileContext
        }
      }
    }
    
    const workspaceContext = await buildWorkspaceContext()
    const workspaceProfile = await buildWorkspaceProfile()
    if (workspaceContext) {
      fullContext = fullContext ? `${fullContext}\n\n${workspaceContext}` : workspaceContext
    }

    // 如果有 PPT 编辑上下文，添加到 fullContext 中
    if (currentPptEditContext) {
      const pptEditInfo = `
=== PPT 编辑请求 ===
【页码】第 ${currentPptEditContext.pageNumber} 页
【编辑类型】${currentPptEditContext.isRegion ? '框选区域编辑' : '整页编辑'}
【PPTX 路径】${currentPptEditContext.pptxPath || '（未知）'}
${currentPptEditContext.regionRect ? `【框选区域】x=${currentPptEditContext.regionRect.x}, y=${currentPptEditContext.regionRect.y}, w=${currentPptEditContext.regionRect.w}, h=${currentPptEditContext.regionRect.h}` : ''}

⚠️ 重要：用户拖拽/框选了 PPT 页面并发送了修改要求。**此请求与 Word 文档无关**，必须使用 **ppt_edit** 工具来处理。
🚫 禁止：word.read / word.create / word.edit / word.format / word.resolve_change / word.chart（这些是 Word 工具，会导致错误操作）
根据用户的描述判断：
- 如果用户对整体不满意（太丑、换风格、重做等），使用 mode="regenerate"
- 如果用户只想修改局部细节（改颜色、换文字、调整位置等），使用 mode="partial_edit"
`
      fullContext = fullContext ? `${pptEditInfo}\n\n${fullContext}` : pptEditInfo
    }

    const memoryWorkspaceKey = currentFile?.path
      ? getParentDir(currentFile.path)
      : (workspacePath || '')
    const memoryWorkspaceSummary = workspaceContext
      ? truncateWithNote(workspaceContext, 1200, '工作夹摘要')
      : ''

    // ─── 工具调用日志：记录发给模型的文档上下文 ───
    toolCallLogger.setWorkDir(workspacePath || (currentFile?.path ? getParentDir(currentFile.path) : ''))
    toolCallLogger.log({
      type: 'request_context',
      data: {
        userMessage: runtimeUserMessage.slice(0, 500),
        documentContextLength: documentContextForAI?.length || 0,
        fullContextLength: fullContext?.length || 0,
        hasPptEditContext: !!currentPptEditContext,
        currentFileName: currentFile?.name || null,
      },
    })

    // 使用 Agent 模式发送消息
    await sendAgentMessage(
      runtimeUserMessage,
      documentContextForAI,
      fullContext || undefined,
      {
        onDebugEvent: async (event) => {
          const markdown = formatAgentDebugEventMarkdown(event)
          if (!markdown) return
          await appendAgentDebugLog(markdown)
        },

        onTextChunk: (text) => {
          const cleaned = sanitizeAssistantText(stripToolBlocks(text || '')).trim()
          if (!cleaned) {
            pendingPreviewToolIdsRef.current = []
            return
          }

          const id = `text-${Date.now()}-${Math.random().toString(16).slice(2)}`
          const pendingIds = new Set(pendingPreviewToolIdsRef.current)

          setStreamItems((prev) => {
            const last = prev[prev.length - 1]
            if (last?.type === 'text' && last.content === cleaned) {
              return prev
            }

            if (pendingIds.size > 0) {
              const firstPendingToolIndex = prev.findIndex(
                (item) => item.type === 'tool' && pendingIds.has(item.id)
              )
              if (firstPendingToolIndex >= 0) {
                return [
                  ...prev.slice(0, firstPendingToolIndex),
                  { type: 'text', id, content: cleaned },
                  ...prev.slice(firstPendingToolIndex),
                ]
              }
            }

            return [...prev, { type: 'text', id, content: cleaned }]
          })

          pendingPreviewToolIdsRef.current = []
        },

        // 工具调用处理
        onToolCallStart: (tool) => {
          registerToolStart(tool)
        },

        onToolCallPreview: (tool, args) => {
          registerToolPreview(tool, args)
        },

        onToolCallSkipped: (tool, args, reason) => {
          markToolPreviewSkipped(tool, args, reason)
        },

        onToolCall: async (tool, args): Promise<ToolResult> => {
          if (activeSkillToolIds?.length && !activeSkillToolIds.includes(tool)) {
            const reason = `blocked by skill ${activeSkill?.id || 'unknown'}`
            markToolPreviewSkipped(tool, args, reason)
            return {
              tool,
              success: false,
              message: `Tool "${tool}" is not allowed while skill "${activeSkill?.displayName || activeSkill?.id || 'unknown'}" is active.`,
            }
          }
          return executeRuntimeTool(tool, args)
        },

        onComplete: (content, toolResults) => {
          // 完成 Agent 进度
          finishAgentProgress()
          const knowledgeHits = getLastKnowledgeHits()

          // 快照当前 streamItems 中的工具卡片，附加到消息上用于历史内联展示
          const toolCards = streamItemsRef.current
            .filter((si): si is { type: 'tool'; id: string; data: ToolActivityItem } => si.type === 'tool')
            .map(si => ({ ...si.data }))

          // 构建完整的交替快照（文字 + 工具卡片按原始顺序），用于历史消息交替渲染
          const streamSnapshot = streamItemsRef.current.map(si => {
            if (si.type === 'text') {
              return { type: 'text' as const, id: si.id, content: si.content }
            }
            return { type: 'tool' as const, id: si.id, toolCard: { ...si.data } }
          })

          console.log('[onComplete] content:', content?.substring(0, 200))
          console.log('[onComplete] toolResults:', toolResults.length)

          // streamItems 由交替卡片实时展示，消息只放简短总结
          // 如果有工具调用结果，显示统计
          if (toolResults.length > 0) {
            // Remove duplicated final summary from stream cards (it may be inserted before tool cards).
            removeLatestMatchingStreamTextItem(content || '')
            // Also remove any trailing text-only stream cards.
            clearTrailingStreamTextItems()
            const successCount = toolResults.filter(r => r.success).length
            const replaceResults = toolResults.filter((r) =>
              r.success && (
                r.tool === 'replace' ||
                r.tool === 'word.replace_via_dsl' ||
                (r.tool === 'word.edit' && r.data?.operation === 'replace')
              ),
            )
            const reviewResults = toolResults.filter((r) =>
              r.success && (
                r.tool === 'review' ||
                (r.tool === 'word.edit' && r.data?.operation === 'review')
              ),
            )
            const createResults = toolResults.filter(r => r.tool === 'create' && r.success)
            const excelCreateResults = toolResults.filter(r => r.tool === 'excel_create' && r.success)
            
            // 构建状态标签
            let statusBadge = ''
            let resultFileName = fileName
            
            if (createResults.length > 0) {
              const created = createResults[0]
              statusBadge = `\n\n---\n✅ **已创建文档** 📄 \`${created.data?.fileName}\` (+${created.data?.lines || 0} 行)`
              resultFileName = created.data?.fileName as string
            } else if (excelCreateResults.length > 0) {
              const created = excelCreateResults[0]
              statusBadge = `\n\n---\n✅ **已创建表格** 📊 \`${created.data?.fileName}\``
              resultFileName = created.data?.fileName as string
            } else if (replaceResults.length > 0 || reviewResults.length > 0) {
              const diffSource = [...replaceResults, ...reviewResults]
              const diffChanges = diffSource.map(r => ({
                diffId: r.data?.diffId as string | undefined,
                searchText: r.data?.searchText as string || '',
                replaceText: r.data?.replaceText as string || '',
                count: (r.data?.count as number) || 0
              }))
              const totalCount = diffChanges.reduce((sum, d) => sum + d.count, 0)
              statusBadge = `\n\n---\n✅ **已更新文档** 📄 \`${fileName}\` (~${totalCount} 处修改)`
              
              // 消息只放简短总结（步骤内容由 toolCards 在历史消息中内联展示）
              addMessage({
                role: 'assistant',
                content: content?.trim() || 'Edit completed',
                diffChanges,
                fileName,
                knowledgeHits,
                toolCards,
                streamSnapshot
              })
              resetToolActivity()
              return
            } else {
              // PPT 编辑：补齐状态徽章（避免只有“已更新”卡片/无总结）
              const pptEditResults = toolResults.filter(r => r.tool === 'ppt_edit' && r.success)
              if (pptEditResults.length > 0) {
                const pages = pptEditResults
                  .map(r => Number((r.data as any)?.pageNumber))
                  .filter(n => Number.isFinite(n) && n > 0)
                const uniquePages = Array.from(new Set(pages)).sort((a, b) => a - b)
                const pptNameFromResult =
                  (pptEditResults[0].data as any)?.fileName ||
                  (pptEditResults[0].data as any)?.pptxName ||
                  ''
                const pptDisplayName = String(pptNameFromResult || currentFile?.name || '演示文稿.pptx')
                const pageStats = uniquePages.length > 0 ? `第 ${uniquePages.join('、')} 页` : '已更新页面'
                
                statusBadge = `\n\n---\n✅ **已更新 PPT** 📄 \`${pptDisplayName}\` ${pageStats}`
                resultFileName = pptDisplayName
              }
            }
            
            if (successCount === 0 && toolResults.length > 0) {
              addMessage({
                role: 'assistant',
                content: content || '操作未能完成，请检查文档内容是否匹配',
                knowledgeHits,
                toolCards,
                streamSnapshot
              })
            } else {
              const summaryText = content?.trim()
                ? content + statusBadge
                : (statusBadge ? `任务已完成！${statusBadge}` : '任务已完成')
              addMessage({
                role: 'assistant',
                content: summaryText,
                fileName: resultFileName,
                knowledgeHits,
                toolCards,
                streamSnapshot
              })
            }
            resetToolActivity()
          } else {
            // No tool calls in this turn: if model claims doc updates, show explicit warning.
            const rawContent = content || 'Done'
            const claimsDocChanged = /\b(completed|updated|modified|replaced|formatted|created)\b/i.test(rawContent) || /已(创建|修改|替换|生成|完成|更新|删除|插入|添加|写好|写完)/.test(rawContent)
            const needsToolButMissing = (operation === 'edit' || operation === 'create') && claimsDocChanged

            addMessage({
              role: 'assistant',
              content: needsToolButMissing
                ? `⚠️ 未检测到工具调用，文档实际未被创建或修改。请重新发送指令，确保模型使用 create 工具。\n\n${rawContent}`
                : rawContent
              ,
              knowledgeHits,
            })
            resetToolActivity()
          }
        },
        
        // 获取最新文档内容（结构化纯文本+格式标注，不含 HTML）
        getLatestDocument: () => {
          return getTiptapDocumentStructure()
        }
      },
      {
        workspaceKey: memoryWorkspaceKey,
        workspaceSummary: memoryWorkspaceSummary,
        workspaceProfile: workspaceProfile || undefined,
        skillDescriptions: availableSkillDescriptions || undefined,
        activeSkill: activeSkill
          ? {
              id: activeSkill.id,
              displayName: activeSkill.displayName,
              description: activeSkill.description,
              safety: activeSkill.safety || 'mutating',
              toolIds: activeSkillToolIds,
            }
          : undefined,
        availableToolsDelta: activeSkillToolIds?.length
          ? {
              source: 'active_skill',
              skillId: activeSkill?.id,
              allowedToolIds: activeSkillToolIds,
            }
          : undefined,
      },
      allImages.length > 0 ? allImages : undefined
    )
  }, [
    input,
    isLoading,
    pptEditContext,
    attachedFiles,
    addMessage,
    sendAgentMessage,
    appendAgentDebugLog,
    document.content,
    availableSkillDescriptions,
    buildSubagentDocumentContext,
    buildSubagentFilesContext,
    buildFilesContext,
    buildWorkspaceContext,
    buildWorkspaceProfile,
    createNewDocument,
    createDocumentFromDsl,
    currentFile?.name,
    replaceInDocument,
    startAgentProgress,
    updateAgentAction,
    completeAgentStep,
    updateAgentFile,
    addAgentFileOperation,
    finishAgentProgress,
    flushUiFrame,
    clearTrailingStreamTextItems,
    removeLatestMatchingStreamTextItem,
    insertInDocument,
    deleteInDocument,
    currentFile?.path,
    handleCancelSubagent,
    getAgentRuntimeSnapshot,
    resetCurrentTurnToolActivity,
    registerToolActivity,
    claimOrRegisterToolActivity,
    registerToolStart,
    registerToolPreview,
    markToolPreviewSkipped,
    completeToolActivity,
    excelData,
    refreshExcelData,
    settings,
    refreshFiles,
    openFile,
    executeRuntimeTool,
    subagentWorkflowRunner,
    workspacePath,
    workspaceDomainAdapter,
    prepareTemplateFillOutput,
    getLatestContent,
    pendingImages,
    skillExecutor,
    workspaceSkillSources,
    workspaceSkillWarnings,
  ])

  const handleKeyDown = useCallback((e: React.KeyboardEvent) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault()
      handleSend()
    }
  }, [handleSend])

  // 拖拽处理
  // image upload/paste helpers
  const imageInputRef = useRef<HTMLInputElement>(null)

  const handlePaste = useCallback((e: React.ClipboardEvent) => {
    const items = e.clipboardData?.items
    if (!items) return
    for (let i = 0; i < items.length; i++) {
      const item = items[i]
      if (!item.type.startsWith('image/')) continue
      e.preventDefault()
      const file = item.getAsFile()
      if (!file) continue
      const reader = new FileReader()
      reader.onload = () => {
        const base64Url = reader.result as string
        setPendingImages(prev => [...prev, base64Url])
      }
      reader.readAsDataURL(file)
    }
  }, [])

  const handleImageUpload = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files
    if (!files) return
    Array.from(files).forEach((file) => {
      if (!file.type.startsWith('image/')) return
      const reader = new FileReader()
      reader.onload = () => {
        const base64Url = reader.result as string
        setPendingImages(prev => [...prev, base64Url])
      }
      reader.readAsDataURL(file)
    })
    e.target.value = ''
  }, [])

  const removePendingImage = useCallback((index: number) => {
    setPendingImages(prev => prev.filter((_, i) => i !== index))
  }, [])

  const handleDragOver = (e: React.DragEvent) => {
    // PPT 页面拖拽：交给输入框区域处理，避免整面板闪烁遮挡
    if (e.dataTransfer.types.includes('application/ppt-page')) return
    e.preventDefault()
    setIsDragOver(true)
  }

  const handleDragLeave = (e: React.DragEvent) => {
    if (e.dataTransfer.types.includes('application/ppt-page')) return
    e.preventDefault()
    setIsDragOver(false)
  }

  const handleDrop = (e: React.DragEvent) => {
    // PPT 页面拖拽：交给输入框区域处理
    if (e.dataTransfer.getData('application/ppt-page')) return
    e.preventDefault()
    setIsDragOver(false)
    try {
      const data = e.dataTransfer.getData('application/json')
      if (data) {
        const file = JSON.parse(data) as FileItem
        if (file && file.type === 'file' && !attachedFiles.find(f => f.path === file.path)) {
          setAttachedFiles(prev => [...prev, file])
        } else if (file && file.type === 'folder' && !attachedFolders.find(f => f.path === file.path)) {
          setAttachedFolders(prev => [...prev, file])
        }
      }
    } catch (error) {
      console.error('Drop error:', error)
    }
  }

  const removeAttachedFile = (path: string) => {
    setAttachedFiles(prev => prev.filter(f => f.path !== path))
  }

  const removeAttachedFolder = (path: string) => {
    setAttachedFolders(prev => prev.filter(f => f.path !== path))
  }

  // 快捷命令
  const quickCommands = [
    { icon: <Folder className="w-3 h-3" />, label: '初始化', command: '/init' },
    { icon: <Bot className="w-3 h-3" />, label: '技能', command: '/skills' },
    { icon: <Folder className="w-3 h-3" />, label: '探索工作区', command: '/workspace-explore' },
    { icon: <Eye className="w-3 h-3" />, label: '验证', command: '/verification' },
    { icon: <FilePlus className="w-3 h-3" />, label: '创建', command: '帮我创建一份' },
    { icon: <FileEdit className="w-3 h-3" />, label: '润色', command: '润色当前文档' },
    { icon: <Eye className="w-3 h-3" />, label: '总结', command: '总结要点' },
  ]

  // Sidebar 触发：新建 PPT（由 Agent 自动调用 ppt_create）
  useEffect(() => {
    const handler = (event: Event) => {
      const detail = (event as CustomEvent<{ topic: string; slideCount: number }>).detail
      if (!detail?.topic) return
      const slideCount = detail.slideCount || 12
      const userMessage =
        `我们要做“海报式 image-only PPTX”（每页是一张完整成片，**文字与排版也必须在图里**）。\n` +
        `主题/需求：${detail.topic}\n` +
        `页数：${slideCount}\n\n` +
        `请严格按两阶段执行（功能优先）：\n` +
        `**阶段1：只输出 PPT 大纲（不要调用任何工具）**\n` +
        `- 只输出一个 JSON（不要 Markdown、不要多余解释），字段如下：\n` +
        `  {\n` +
        `    "title": "...",\n` +
        `    "theme": "...",\n` +
        `    "styleHint": "...(可空)",\n` +
        `    "slides": [\n` +
        `      {\n` +
        `        "pageNumber": 1,\n` +
        `        "pageType": "cover|section|content|diagram|ending",\n` +
        `        "headline": "该页主标题（中文，必须可直接上屏）",\n` +
        `        "subheadline": "副标题（可空）",\n` +
        `        "bullets": ["要点1","要点2","要点3"],\n` +
        `        "footerNote": "页脚/注释（可空）",\n` +
        `        "layoutIntent": "排版意图（例如：左文右图/居中标题+下方三要点/时间轴等）"\n` +
        `      }\n` +
        `    ]\n` +
        `  }\n` +
        `- slides 数组长度必须等于页数；每页文案要完整且专业，便于后续直接用于排版。\n\n` +
        `用户确认后我会回复“开始生成”。\n` +
        `**阶段2：收到“开始生成”后，再调用 ppt_create 工具一次性导出 PPTX**（不要让我手动复制提示词）。\n` +
        `硬性要求：\n` +
        `1) slides 数组长度必须等于页数；\n` +
        `2) 每页 prompt 必须包含该页所有中文文案 + 明确排版（层级/对齐/留白/网格）；\n` +
        `3) 禁止水印/徽章/二维码/乱码/错别字；中文必须清晰准确。\n`

      setInput(userMessage)
      setTimeout(() => {
        handleSend()
      }, 50)
    }

    window.addEventListener('ppt-create-request', handler as EventListener)
    return () => window.removeEventListener('ppt-create-request', handler as EventListener)
  }, [handleSend])

  const displayMessages = messages.filter(m => m.content.trim() !== '')

  const requestElapsedSec = requestRunState.requestStartedAt
    ? Math.max(0, Math.floor((runtimeClock - requestRunState.requestStartedAt) / 1000))
    : 0
  const lastActivityAgoSec = requestRunState.lastActivityAt
    ? Math.max(0, Math.floor((runtimeClock - requestRunState.lastActivityAt) / 1000))
    : 0
  const canRetryCurrentRequest =
    isLoading &&
    requestRunState.stallLevel === 'stalled' &&
    lastActivityAgoSec >= 45

  const waitingPrimaryLabel = (() => {
    switch (requestRunState.phase) {
      case 'awaiting_model':
        return '等待模型首个响应'
      case 'streaming_text':
        return '模型正在实时输出'
      case 'streaming_reasoning':
        return '模型正在输出思维链'
      case 'executing_tools':
        return requestRunState.activeToolLabel
          ? `正在调用工具：${requestRunState.activeToolLabel}`
          : '正在调用工具'
      case 'applying_document':
        return requestRunState.activeToolLabel
          ? `正在写回文档：${requestRunState.activeToolLabel}`
          : '正在写回文档'
      case 'stalled': {
        if (requestRunState.stallReason === 'waiting_first_token') {
          return '疑似卡住：模型迟迟没有首个响应'
        }
        if (requestRunState.stallReason === 'waiting_next_token') {
          return '疑似卡住：模型输出中断'
        }
        if (requestRunState.stallReason === 'tool_no_progress') {
          return '疑似卡住：工具执行长时间无进展'
        }
        if (requestRunState.stallReason === 'writeback_no_progress') {
          return '疑似卡住：文档写回长时间无进展'
        }
        return '疑似卡住：长时间无新活动'
      }
      case 'error':
        return '请求失败'
      case 'aborted':
        return '请求已停止'
      case 'completed':
        return '处理完成'
      default:
        return streamingContent ? '正在处理' : '等待响应'
    }
  })()

  const waitingSecondaryLabel = (() => {
    if (requestRunState.stallLevel === 'slow') {
      return `响应较慢，最近一次活动在 ${lastActivityAgoSec}s 前`
    }
    if (requestRunState.stallLevel === 'stalled') {
      return `最近 ${lastActivityAgoSec}s 无新活动`
    }
    if (requestRunState.activeToolLabel && requestRunState.phase !== 'executing_tools' && requestRunState.phase !== 'applying_document') {
      return `当前步骤：${requestRunState.activeToolLabel}`
    }
    if (lastActivityAgoSec > 0) {
      return `最近一次活动 ${lastActivityAgoSec}s 前`
    }
    return '请求已发出，等待新的流式信号'
  })()

  const markdownComponents: any = {
    h1: ({ children }: { children?: ReactNode }) => (
      <h1 className="text-[15px] font-semibold text-text mt-3 mb-2 pb-1 border-b border-border">{children}</h1>
    ),
    h2: ({ children }: { children?: ReactNode }) => (
      <h2 className="text-[14px] font-semibold text-text mt-3 mb-1.5 flex items-center gap-1.5">{children}</h2>
    ),
    h3: ({ children }: { children?: ReactNode }) => (
      <h3 className="text-[13px] font-medium text-text-secondary mt-2 mb-1">{children}</h3>
    ),
    p: ({ children }: { children?: ReactNode }) => <p className="mb-2 last:mb-0">{children}</p>,
    ul: ({ children }: { children?: ReactNode }) => <ul className="list-none ml-0 mb-2 space-y-1">{children}</ul>,
    ol: ({ children }: { children?: ReactNode }) => <ol className="list-decimal ml-4 mb-2 space-y-1">{children}</ol>,
    li: ({ children }: { children?: ReactNode }) => (
      <li className="text-[13px] leading-relaxed flex items-start gap-1.5">
        <span className="text-text-muted mt-0.5">•</span>
        <span className="flex-1">{children}</span>
      </li>
    ),
    strong: ({ children }: { children?: ReactNode }) => <strong className="font-semibold text-text">{children}</strong>,
    em: ({ children }: { children?: ReactNode }) => <em className="italic text-text-secondary">{children}</em>,
    code: ({ children, className }: { children?: ReactNode; className?: string }) => {
      const isBlock = className?.includes('language-')
      if (isBlock) {
        return (
          <code className="block bg-black/10 dark:bg-black/35 text-text p-2 rounded text-[12px] font-mono overflow-x-auto my-2 border border-border">
            {children}
          </code>
        )
      }
      return (
        <code className="bg-black/10 dark:bg-black/35 text-text px-1 py-0.5 rounded text-[12px] font-mono border border-border">
          {children}
        </code>
      )
    },
    pre: ({ children }: { children?: ReactNode }) => (
      <pre className="bg-black/10 dark:bg-black/35 rounded-md overflow-hidden my-2 border border-border">{children}</pre>
    ),
    a: ({ href, children }: { href?: string; children?: ReactNode }) => (
      <a href={href} className="text-accent hover:underline" target="_blank" rel="noopener noreferrer">
        {children}
      </a>
    ),
    blockquote: ({ children }: { children?: ReactNode }) => (
      <blockquote className="border-l-2 border-accent pl-3 my-2 text-text-muted italic">{children}</blockquote>
    ),
    hr: () => <hr className="border-border my-3" />,
    table: ({ node, children, ...props }: any) => {
      const data = extractTableCardData(node)
      if (!data || data.rows.length === 0) {
        return (
          <table className="ai-markdown-table" {...props}>
            {children}
          </table>
        )
      }
      const cards = data.rows
        .map((row, rowIndex) => {
          const fields: JSX.Element[] = []
          data.headers.forEach((header, colIndex) => {
            const value = row[colIndex]
            if (!value) return
            const label = header || `字段${colIndex + 1}`
            fields.push(
              <div className="ai-table-field" key={`field-${rowIndex}-${colIndex}`}>
                <div className="ai-table-field-label">{label}</div>
                <div className="ai-table-field-value">{value}</div>
              </div>
            )
          })
          if (fields.length === 0) return null
          return (
            <div className="ai-table-card" key={`row-${rowIndex}`}>
              {fields}
            </div>
          )
        })
        .filter((card): card is JSX.Element => card !== null)

      if (cards.length === 0) {
        return (
          <table className="ai-markdown-table" {...props}>
            {children}
          </table>
        )
      }

      return <div className="ai-table-cards">{cards}</div>
    },
    thead: ({ children }: { children?: ReactNode }) => <thead className="ai-markdown-thead">{children}</thead>,
    tbody: ({ children }: { children?: ReactNode }) => <tbody className="ai-markdown-tbody">{children}</tbody>,
    tr: ({ children }: { children?: ReactNode }) => <tr className="ai-markdown-tr">{children}</tr>,
    th: ({ children }: { children?: ReactNode }) => <th className="ai-markdown-th">{children}</th>,
    td: ({ children }: { children?: ReactNode }) => <td className="ai-markdown-td">{children}</td>,
  }

  // 渲染历史消息中内联的工具卡片（从 onComplete 快照）
  const renderInlineToolCards = (cards: ChatMessage['toolCards']) => {
    if (!cards || cards.length === 0) return null
    return (
      <div className="space-y-1.5 mb-2">
        {cards.map(card => {
          const jumpTarget = card.replaceText || card.searchText || ''
          const canJump = !!jumpTarget
          const cardClass = `glass-card-soft rounded-xl border border-border px-3 py-2 ${canJump ? 'hover:border-accent/40 hover:bg-accent/5 cursor-pointer' : ''}`
          const cardBody = (
            <>
              <div className="flex items-center gap-2 text-[12px]">
                {card.status === 'success' ? (
                  <CheckCircle2 className="w-3.5 h-3.5 text-success flex-shrink-0" />
                ) : card.status === 'skipped' ? (
                  <Circle className="w-3.5 h-3.5 text-text-muted flex-shrink-0" />
                ) : card.status === 'error' ? (
                  <X className="w-3.5 h-3.5 text-error flex-shrink-0" />
                ) : (
                  <CheckCircle2 className="w-3.5 h-3.5 text-text-muted flex-shrink-0" />
                )}
                <span className="px-1.5 py-0.5 rounded bg-accent/10 text-accent text-[11px] font-medium">
                  {card.tool}
                </span>
                <span className="text-text truncate flex-1" title={card.label}>
                  {card.label}
                </span>
                {card.detail && (
                  <span className="text-[10px] text-text-muted flex-shrink-0">{card.detail}</span>
                )}
              </div>
              {(card.searchText || card.replaceText) && (
                <div className="mt-1.5 text-[11px] text-text-secondary flex items-center gap-2">
                  <span className="text-error truncate max-w-[40%]" title={card.searchText}>
                    {card.searchText || '-'}
                  </span>
                  <span className="text-text-muted">&rarr;</span>
                  <span className="text-success truncate max-w-[40%]" title={card.replaceText}>
                    {card.replaceText || '-'}
                  </span>
                </div>
              )}
            </>
          )
          if (canJump) {
            return (
              <button key={card.id} type="button" onClick={() => scrollToChange(jumpTarget)} className={`${cardClass} w-full text-left`}>
                {cardBody}
              </button>
            )
          }
          return <div key={card.id} className={cardClass}>{cardBody}</div>
        })}
      </div>
    )
  }

  const renderKnowledgeHits = (hits: ChatMessage['knowledgeHits']) => {
    if (!hits || hits.length === 0) return null

    const scopeLabelMap: Record<string, string> = {
      workspace: '工作区',
      global: '长期库',
      profile: '画像',
      daily: '日记忆',
      sessions: '会话',
    }

    const scopeClassMap: Record<string, string> = {
      workspace: 'bg-accent/10 text-accent',
      global: 'bg-amber-500/10 text-amber-400',
      profile: 'bg-emerald-500/10 text-emerald-400',
      daily: 'bg-sky-500/10 text-sky-400',
      sessions: 'bg-violet-500/10 text-violet-400',
    }

    return (
      <div className="space-y-1.5 mb-2">
        <div className="glass-card-soft rounded-xl border border-border px-3 py-2">
          <div className="flex items-center gap-2 text-[12px] text-text mb-2">
            <Eye className="w-3.5 h-3.5 text-accent flex-shrink-0" />
            <span className="font-medium">知识命中来源</span>
            <span className="text-[10px] text-text-muted">{hits.length} 条</span>
          </div>
          <div className="space-y-2">
            {hits.slice(0, 4).map((hit, index) => {
              const scopeKey = String(hit.sourceScope || '')
              const scopeLabel = scopeLabelMap[scopeKey] || scopeKey || '未知'
              const scopeClass = scopeClassMap[scopeKey] || 'bg-black/5 dark:bg-white/5 text-text-secondary'
              const pathText = hit.relativePath || hit.sourcePath || hit.title || ''
              const summaryText = hit.statement || hit.snippet

              return (
                <div
                  key={`${scopeKey}-${hit.sourcePath || hit.relativePath || hit.title || 'hit'}-${index}`}
                  className="rounded-lg border border-black/10 dark:border-white/10 bg-black/5 dark:bg-white/5 px-2.5 py-2"
                >
                  <div className="flex items-center gap-2 mb-1.5">
                    <span className={`px-1.5 py-0.5 rounded text-[10px] font-medium ${scopeClass}`}>
                      {scopeLabel}
                    </span>
                    <span className="text-[11px] text-text truncate flex-1" title={pathText}>
                      {pathText}
                    </span>
                    <span className="text-[10px] text-text-muted flex-shrink-0">
                      {typeof hit.score === 'number' ? hit.score.toFixed(2) : '-'}
                    </span>
                  </div>
                  <div className="text-[11px] text-text-secondary leading-relaxed line-clamp-3" title={summaryText}>
                    {summaryText}
                  </div>
                </div>
              )
            })}
          </div>
        </div>
      </div>
    )
  }

  // 渲染交替快照（文字 + 工具卡片按原始执行顺序）
  const renderStreamSnapshot = (snapshot: NonNullable<ChatMessage['streamSnapshot']>) => {
    return (
      <div className="space-y-2">
        {snapshot.map(item => {
          if (item.type === 'text') {
            return (
              <div key={item.id} className="glass-card-soft rounded-2xl rounded-tl-sm px-4 py-3 border border-border">
                <div className="text-[13px] leading-relaxed text-text prose prose-sm max-w-none">
                  <ReactMarkdown remarkPlugins={[remarkGfm]} components={markdownComponents}>
                    {item.content}
                  </ReactMarkdown>
                </div>
              </div>
            )
          }
          const card = item.toolCard
          const jumpTarget = card.replaceText || card.searchText || ''
          const canJump = !!jumpTarget
          const cardClass = `glass-card-soft rounded-xl border border-border px-3 py-2 ${canJump ? 'hover:border-accent/40 hover:bg-accent/5 cursor-pointer' : ''}`
          const cardBody = (
            <>
              <div className="flex items-center gap-2 text-[12px]">
                {card.status === 'success' ? (
                  <CheckCircle2 className="w-3.5 h-3.5 text-success flex-shrink-0" />
                ) : card.status === 'skipped' ? (
                  <Circle className="w-3.5 h-3.5 text-text-muted flex-shrink-0" />
                ) : card.status === 'error' ? (
                  <X className="w-3.5 h-3.5 text-error flex-shrink-0" />
                ) : (
                  <CheckCircle2 className="w-3.5 h-3.5 text-text-muted flex-shrink-0" />
                )}
                <span className="px-1.5 py-0.5 rounded bg-accent/10 text-accent text-[11px] font-medium">
                  {card.tool}
                </span>
                <span className="text-text truncate flex-1" title={card.label}>
                  {card.label}
                </span>
                {card.detail && (
                  <span className="text-[10px] text-text-muted flex-shrink-0">{card.detail}</span>
                )}
              </div>
              {(card.searchText || card.replaceText) && (
                <div className="mt-1.5 text-[11px] text-text-secondary flex items-center gap-2">
                  <span className="text-error truncate max-w-[40%]" title={card.searchText}>
                    {card.searchText || '-'}
                  </span>
                  <span className="text-text-muted">&rarr;</span>
                  <span className="text-success truncate max-w-[40%]" title={card.replaceText}>
                    {card.replaceText || '-'}
                  </span>
                </div>
              )}
            </>
          )
          if (canJump) {
            return (
              <button key={card.id} type="button" onClick={() => scrollToChange(jumpTarget)} className={`${cardClass} w-full text-left`}>
                {cardBody}
              </button>
            )
          }
          return <div key={card.id} className={cardClass}>{cardBody}</div>
        })}
      </div>
    )
  }

  const renderSubagentMonitor = () => {
    if (
      runtimeSnapshotState.subagents.agents.length === 0 &&
      runtimeSnapshotState.notifications.count === 0
    ) {
      return null
    }

    const sortedAgents = runtimeSnapshotState.subagents.agents
      .slice()
      .sort((left, right) => right.createdAt.localeCompare(left.createdAt))

    return (
      <div className="glass-card-soft border border-border rounded-xl px-3 py-2 mt-3">
        <div className="text-[10px] text-text-dim uppercase tracking-wider mb-2">
          Subagent Runtime
        </div>
        <div className="space-y-1 max-h-[180px] overflow-y-auto">
          {sortedAgents.slice(0, 6).map((agent) => {
            const isSelected = selectedSubagentId === agent.subagentId
            return (
              <button
                key={agent.subagentId}
                type="button"
                onClick={() =>
                  setSelectedSubagentId((current) =>
                    current === agent.subagentId ? null : agent.subagentId,
                  )
                }
                className={`w-full text-left rounded px-1.5 py-1 transition-colors border ${
                  isSelected
                    ? 'border-accent/40 bg-accent/5'
                    : 'border-black/10 dark:border-white/10 bg-black/5 dark:bg-white/5 hover:bg-white/8'
                }`}
              >
                <div className="flex items-center gap-1.5 text-[11px] text-text-secondary">
                  {agent.status === 'running' || agent.status === 'queued' ? (
                    <Loader2 className="w-3 h-3 text-accent animate-spin flex-shrink-0" />
                  ) : agent.status === 'completed' ? (
                    <CheckCircle2 className="w-3 h-3 text-success flex-shrink-0" />
                  ) : agent.status === 'cancelled' ? (
                    <Circle className="w-3 h-3 text-text-muted flex-shrink-0" />
                  ) : (
                    <X className="w-3 h-3 text-error flex-shrink-0" />
                  )}
                  <span className="truncate flex-1">{agent.label}</span>
                  <span className="text-[10px] text-text-muted flex-shrink-0">{agent.mode}</span>
                  {(agent.status === 'running' || agent.status === 'queued') && (
                    <span
                      role="button"
                      tabIndex={0}
                      onClick={(event) => {
                        event.stopPropagation()
                        handleCancelSubagent(agent.subagentId)
                      }}
                      onKeyDown={(event) => {
                        if (event.key === 'Enter' || event.key === ' ') {
                          event.preventDefault()
                          event.stopPropagation()
                          handleCancelSubagent(agent.subagentId)
                        }
                      }}
                      className="text-[10px] text-error flex-shrink-0 hover:underline"
                    >
                      cancel
                    </span>
                  )}
                </div>
                <div className="mt-0.5 text-[10px] text-text-dim">
                  {agent.profileId} | {agent.status}
                  {agent.summary ? ` | ${agent.summary.slice(0, 96)}` : ''}
                </div>
              </button>
            )
          })}
        </div>

        {selectedSubagentTranscript && (
          <div className="mt-2 border-t border-border pt-2">
            <div className="text-[10px] text-text-dim uppercase tracking-wider mb-1">
              Transcript Preview
            </div>
            <div className="text-[10px] text-text-dim mb-2">
              messages: {selectedSubagentTranscript.messages.length} | tool calls: {selectedSubagentTranscript.toolCalls.length} | tool results: {selectedSubagentTranscript.toolResults.length}
            </div>
            <div className="space-y-1 max-h-[160px] overflow-y-auto">
              {selectedSubagentTranscript.messages.slice(-6).map((message) => (
                <div
                  key={message.id}
                  className="rounded px-1.5 py-1 bg-black/5 dark:bg-white/5 border border-black/10 dark:border-white/10"
                >
                  <div className="text-[10px] text-text-muted uppercase tracking-wider mb-0.5">
                    {message.role}
                  </div>
                  <div className="text-[11px] text-text-secondary whitespace-pre-wrap">
                    {message.content}
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}

        {runtimeSnapshotState.notifications.count > 0 && (
          <div className="mt-2 border-t border-border pt-2">
            <div className="text-[10px] text-text-dim uppercase tracking-wider mb-1">
              Notifications ({runtimeSnapshotState.notifications.count})
            </div>
            <div className="space-y-1 max-h-[100px] overflow-y-auto">
              {runtimeSnapshotState.notifications.items
                .slice()
                .sort((left, right) => right.createdAt.localeCompare(left.createdAt))
                .slice(0, 4)
                .map((item) => (
                  <div
                    key={`${item.taskId}-${item.createdAt}`}
                    className="rounded px-1.5 py-1 bg-black/5 dark:bg-white/5 border border-black/10 dark:border-white/10 text-[10px] text-text-dim"
                  >
                    {item.message}
                  </div>
                ))}
            </div>
          </div>
        )}
      </div>
    )
  }

  return (
    <div 
      className={`flex flex-col h-full bg-transparent ${isDragOver ? 'ring-2 ring-accent/40 ring-inset' : ''}`}
      onDragOver={handleDragOver}
      onDragLeave={handleDragLeave}
      onDrop={handleDrop}
    >
      {/* 头部 - 柔和玻璃态风格 */}
      <div className="flex items-center justify-between px-4 py-3 border-b border-border">
        <div className="flex items-center gap-2.5">
          <div className="w-7 h-7 rounded-xl bg-gradient-to-br from-accent to-accent-hover flex items-center justify-center shadow-md shadow-accent/20">
            <Bot className="w-4 h-4 text-white" />
          </div>
          <span className="text-sm font-semibold text-text tracking-wide">AI CHAT</span>
        </div>
        <div className="flex items-center gap-1.5">
          <button
            onClick={openSettings}
            className="p-2 rounded-xl text-text-muted hover:text-text hover:bg-black/5 dark:hover:bg-white/5 transition-all"
            title="设置"
          >
            <Settings className="w-4 h-4" />
          </button>
          <button
            onClick={handleNewConversation}
            className="p-2 rounded-xl text-text-muted hover:text-text hover:bg-black/5 dark:hover:bg-white/5 transition-all"
            title="New chat"
          >
            <FilePlus className="w-4 h-4" />
          </button>
        </div>
      </div>

      {/* 快捷命令 - 柔和标签风格 */}
      <div className="px-4 py-3 border-b border-border flex gap-2 overflow-x-auto scrollbar-none">
        {quickCommands.map((cmd, i) => (
          <button
            key={i}
            onClick={() => setInput(cmd.command)}
            className="flex items-center gap-1.5 px-3 py-1.5 bg-black/5 dark:bg-white/5 hover:bg-black/10 dark:hover:bg-white/10 text-[11px] text-text-muted hover:text-text rounded-xl transition-all whitespace-nowrap border border-black/10 dark:border-white/10 hover:border-accent/20"
          >
            {cmd.icon}
            <span>{cmd.label}</span>
          </button>
        ))}
      </div>

      {/* 拖拽提示 */}
      {isDragOver && (
        <div className="absolute inset-0 z-50 flex items-center justify-center bg-black/20 dark:bg-black/45 backdrop-blur-md">
          <div className="flex flex-col items-center gap-3 p-8 glass-card">
            <Paperclip className="w-10 h-10 text-accent" />
            <p className="text-sm text-text font-medium">释放以添加文件</p>
          </div>
        </div>
      )}

      {/* 消息列表 - 更清爽：减少硬编码深色块 */}
      <div 
        ref={chatContainerRef}
        className="flex-1 overflow-y-auto px-4 py-4 space-y-4 chat-scrollbar"
        onScroll={(e) => {
          const el = e.currentTarget
          const isNearBottom = el.scrollHeight - el.scrollTop - el.clientHeight < 120
          userScrolledUpRef.current = !isNearBottom
        }}
      >
        <AnimatePresence mode="popLayout">
        {displayMessages.map((message) => {
          const displayContent =
            message.role === 'assistant' ? sanitizeAssistantText(message.content) : stripToolBlocks(message.content)

          return (
            <motion.div
              key={message.id}
              layout
              variants={messageVariants}
              initial="hidden"
              animate="visible"
              exit="exit"
              className={`group ${message.role === 'user' ? 'flex flex-col items-end' : ''}`}
            >
              {/* 用户消息 */}
              {message.role === 'user' ? (
                <div className="max-w-[90%]">
                  <div className="bg-gradient-to-br from-accent/92 to-accent-hover/92 text-white rounded-2xl rounded-tr-sm px-4 py-2.5 shadow-sm shadow-black/10 dark:shadow-black/20 border border-white/12">
                    <p className="text-[13px] leading-relaxed whitespace-pre-wrap">{displayContent}</p>
                  </div>
                  <span className="text-[10px] text-text-dim mt-1.5 block text-right pr-1">
                    {message.timestamp.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
                  </span>
                </div>
              ) : displayContent.includes('\n---\n✅') ? (
              /* 操作完成消息 - 显示 AI 总结 + 状态卡片 */
              <div className="w-full space-y-3">
                {renderKnowledgeHits(message.knowledgeHits)}
                {message.streamSnapshot?.length ? renderStreamSnapshot(message.streamSnapshot) : renderInlineToolCards(message.toolCards)}
                {/* AI 总结内容 — streamSnapshot 已包含文字，只渲染状态卡片 */}
                {(() => {
                  const parts = displayContent.split('\n---\n')
                  const summaryContent = parts[0]
                  const statusContent = parts.slice(1).join('\n---\n')
                  const hasSnapshot = !!message.streamSnapshot?.length
                  return (
                    <>
                      {!hasSnapshot && summaryContent && (
                        <div className="text-[13px] leading-relaxed text-text prose prose-sm max-w-none">
                          <ReactMarkdown
                            remarkPlugins={[remarkGfm]}
                            components={markdownComponents}
                          >
                            {summaryContent}
                          </ReactMarkdown>
                        </div>
                      )}
                      {/* 状态卡片 */}
                      {statusContent && (
                        <div className="bg-black/5 dark:bg-white/5 border border-border rounded-lg overflow-hidden">
                          <div className="flex items-center gap-2 px-3 py-2 bg-success/10 border-b border-success/20">
                            <CheckCircle className="w-3.5 h-3.5 text-success" />
                            <span className="text-[12px] font-medium text-success">
                              {statusContent.includes('表格') ? '表格已创建' : statusContent.includes('创建') ? '文档已创建' : '文档已更新'}
                            </span>
                          </div>
                          <div className="px-3 py-2">
                            {statusContent.split('\n').map((line, i) => {
                              if (line.startsWith('📄') || line.startsWith('📊')) {
                                const emoji = line.startsWith('📊') ? '📊' : '📄'
                                const parts = line.replace(/^(📄|📊)\s*/, '').split(/\s+/)
                                const fileNamePart = parts[0]?.replace(/`/g, '')
                                const stats = parts.slice(1).join(' ')
                                return (
                                  <button
                                    key={i}
                                    onClick={() => fileNamePart && openCreatedFile(fileNamePart)}
                                    className="w-full flex items-center justify-between gap-2 py-1 hover:bg-black/10 dark:hover:bg-white/10 cursor-pointer rounded"
                                  >
                                    <div className="flex items-center gap-2 min-w-0">
                                      {emoji === '📊' ? (
                                        <Table className="w-3.5 h-3.5 text-success flex-shrink-0" />
                                      ) : (
                                        <FileText className="w-3.5 h-3.5 text-accent flex-shrink-0" />
                                      )}
                                      <span className="text-[12px] text-text font-mono truncate">{fileNamePart}</span>
                                    </div>
                                    {stats && (
                                      <span className="text-[10px] font-mono text-text-muted">{stats}</span>
                                    )}
                                  </button>
                                )
                              }
                              return null
                            })}
                          </div>
                        </div>
                      )}
                    </>
                  )
                })()}
                <span className="text-[10px] text-text-dim mt-1 block">
                  {message.timestamp.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
                </span>
              </div>
            ) : displayContent.startsWith('✅') ? (
              /* 简单操作完成消息 - Cursor 风格卡片 */
              <div className="w-full">
                {renderKnowledgeHits(message.knowledgeHits)}
                {message.streamSnapshot?.length ? renderStreamSnapshot(message.streamSnapshot) : renderInlineToolCards(message.toolCards)}
                <div className="glass-card-soft rounded-2xl overflow-hidden border border-border">
                  {/* 成功标题栏 */}
                  <div className="flex items-center gap-2 px-3 py-2 bg-white/5 border-b border-border">
                    <CheckCircle className="w-3.5 h-3.5 text-success" />
                    <span className="text-[12px] font-medium text-text">
                      {displayContent.includes('表格') ? '表格已创建' : displayContent.includes('创建') ? '文档已创建' : '文档已更新'}
                    </span>
                  </div>
                  {/* 文件信息 */}
                  <div className="px-3 py-2">
                    {displayContent.split('\n').slice(1).map((line, i) => {
                      if (line.startsWith('📄') || line.startsWith('📊')) {
                        const emoji = line.startsWith('📊') ? '📊' : '📄'
                        const parts = line.replace(/^(📄|📊)\s*/, '').split(/\s+/)
                        const fileNamePart = parts[0]?.replace(/`/g, '')
                        const stats = parts.slice(1).join(' ')
                        const isCreateMessage = displayContent.includes('创建')
                        return (
                          <button
                            key={i}
                            onClick={() => {
                              if (isCreateMessage && fileNamePart) {
                                openCreatedFile(fileNamePart)
                              }
                            }}
                            className={`w-full flex items-center justify-between gap-2 py-1 ${isCreateMessage ? 'hover:bg-white/5 cursor-pointer rounded-lg' : ''}`}
                          >
                            <div className="flex items-center gap-2 min-w-0">
                              {emoji === '📊' ? (
                                <Table className="w-3.5 h-3.5 text-success flex-shrink-0" />
                              ) : (
                                <FileText className="w-3.5 h-3.5 text-accent flex-shrink-0" />
                              )}
                              <span className="text-[12px] text-text font-mono truncate">{fileNamePart}</span>
                            </div>
                            <div className="flex items-center gap-1 flex-shrink-0">
                              {stats.includes('+') && (
                                <span className="text-[10px] font-mono text-success">
                                  {stats.match(/\+\d+/)?.[0]}
                                </span>
                              )}
                              {stats.includes('-') && (
                                <span className="text-[10px] font-mono text-rose-400">
                                  {stats.match(/-\d+/)?.[0]}
                                </span>
                              )}
                              {stats.includes('~') && (
                                <span className="text-[10px] font-mono text-warning">
                                  {stats.match(/~\d+/)?.[0]}
                                </span>
                              )}
                            </div>
                          </button>
                        )
                      }
                      return null
                    })}
                  </div>
                  
                  {/* Diff 详情 */}
                  {message.diffChanges && message.diffChanges.length > 0 && (
                    <div className="border-t border-border px-3 py-2">
                      <div className="text-[10px] text-text-muted mb-2">修改详情</div>
                      <div className="space-y-1">
                        {message.diffChanges.slice(0, 5).map((diff, i) => (
                          <button
                            key={i}
                            onClick={() => jumpToDiffChange(diff)}
                            className="w-full text-left px-2 py-1.5 rounded bg-black/5 dark:bg-white/5 hover:bg-black/10 dark:hover:bg-white/10 transition-colors border border-black/10 dark:border-white/10"
                          >
                            <div className="flex items-center gap-2 text-[11px]">
                              <span className="text-error line-through truncate flex-1" title={diff.searchText}>
                                {diff.searchText.slice(0, 25)}{diff.searchText.length > 25 ? '...' : ''}
                              </span>
                              <span className="text-text-dim">→</span>
                              <span className="text-success truncate flex-1" title={diff.replaceText}>
                                {diff.replaceText.slice(0, 25)}{diff.replaceText.length > 25 ? '...' : ''}
                              </span>
                            </div>
                          </button>
                        ))}
                        {message.diffChanges.length > 5 && (
                          <div className="text-[10px] text-text-muted text-center py-1">
                            还有 {message.diffChanges.length - 5} 处修改...
                          </div>
                        )}
                      </div>
                    </div>
                  )}
                </div>
                <span className="text-[10px] text-text-dim mt-1 block pl-1">
                  {message.timestamp.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
                </span>
              </div>
            ) : (
              /* AI 普通消息 - 使用 Markdown 渲染 */
              <div className="w-full">
                {renderKnowledgeHits(message.knowledgeHits)}
                {message.streamSnapshot?.length ? renderStreamSnapshot(message.streamSnapshot) : renderInlineToolCards(message.toolCards)}
                {!message.streamSnapshot?.length && (
                <div className="glass-card-soft border border-border rounded-2xl rounded-tl-sm px-3 py-2">
                  <div className="ai-markdown text-[13px] text-text leading-relaxed">
                    {(() => {
                      const baseText = displayContent
                      const parsed = tryParsePptOutlineDraft(baseText)
                      const cleanedText = parsed ? stripPptOutlineJsonFromText(baseText) : baseText
                      const jsonOpen = !!outlineJsonOpen[message.id]
                      const dslMatch = tryExtractDocDsl(cleanedText)
                      const markdownText = dslMatch
                        ? cleanedText.replace(dslMatch.sourceBlock, '').trim()
                        : cleanedText

                      return (
                        <>
                          {parsed && (
                            <div className="mb-3 bg-black/5 dark:bg-white/5 border border-border rounded-lg overflow-hidden">
                              <div className="flex items-center justify-between px-3 py-2 bg-black/5 dark:bg-white/5 border-b border-border">
                                <div className="min-w-0">
                                  <div className="text-[12px] text-text truncate">
                                    PPT 大纲：{parsed.draft.title || '未命名'}（{parsed.draft.slides.length} 页）
                                  </div>
                                  <div className="text-[10px] text-text-muted truncate">
                                    {parsed.draft.theme ? `主题：${parsed.draft.theme}  ` : ''}{parsed.draft.styleHint ? `风格：${parsed.draft.styleHint}` : ''}
                                  </div>
                                </div>
                                <button
                                  onClick={() =>
                                    setOutlineJsonOpen((prev) => ({ ...prev, [message.id]: !prev[message.id] }))
                                  }
                                  className="px-2 py-1 text-[10px] rounded bg-black/5 dark:bg-white/5 hover:bg-black/10 dark:hover:bg-white/10 text-text-secondary transition-colors flex-shrink-0 border border-black/10 dark:border-white/10"
                                  title={jsonOpen ? '收起 JSON' : '展开 JSON'}
                                >
                                  {jsonOpen ? '收起 JSON' : '展开 JSON'}
                                </button>
                              </div>

                              <div className="px-3 py-2 space-y-2">
                                {parsed.draft.slides.map((s, idx) => (
                                  <div key={`${s.pageNumber}-${idx}`} className="border border-border rounded-md bg-black/5 dark:bg-white/5">
                                    <div className="px-2.5 py-2 border-b border-border flex items-center justify-between gap-2">
                                      <div className="min-w-0">
                                        <div className="text-[12px] text-text truncate">
                                          第{s.pageNumber || idx + 1}页：{s.headline || '（未填写标题）'}
                                        </div>
                                        {s.subheadline && (
                                          <div className="text-[10px] text-text-secondary truncate">{s.subheadline}</div>
                                        )}
                                      </div>
                                      {s.layoutIntent && (
                                        <div className="text-[10px] text-text-muted flex-shrink-0 truncate max-w-[45%]" title={s.layoutIntent}>
                                          {s.layoutIntent}
                                        </div>
                                      )}
                                    </div>
                                    {(s.bullets?.length || s.footerNote) && (
                                      <div className="px-2.5 py-2">
                                        {s.bullets?.length ? (
                                          <ul className="space-y-1">
                                            {s.bullets.slice(0, 8).map((b, bi) => (
                                              <li key={bi} className="text-[12px] text-text-secondary leading-relaxed flex items-start gap-1.5">
                                                <span className="text-text-muted mt-0.5">•</span>
                                                <span className="flex-1">{b}</span>
                                              </li>
                                            ))}
                                          </ul>
                                        ) : null}
                                        {s.footerNote && (
                                          <div className="mt-2 text-[10px] text-text-muted border-t border-border pt-2">
                                            页脚：{s.footerNote}
                                          </div>
                                        )}
                                      </div>
                                    )}
                                  </div>
                                ))}

                                {jsonOpen && (
                                  <pre className="mt-2 bg-black/10 dark:bg-black/35 border border-border rounded-md p-2 text-[11px] text-text-secondary overflow-x-auto">
                                    {parsed.rawJson}
                                  </pre>
                                )}
                              </div>
                            </div>
                          )}

                          {dslMatch && (
                            <div className="dsl-preview mb-3">
                              <div className="dsl-preview-header">结构化文档预览</div>
                              <div
                                className="dsl-preview-body"
                                dangerouslySetInnerHTML={{ __html: dslToHtml(dslMatch.dsl) }}
                              />
                            </div>
                          )}
                          {markdownText && (
                            <ReactMarkdown
                              remarkPlugins={[remarkGfm]}
                              components={markdownComponents}
                            >
                              {markdownText}
                            </ReactMarkdown>
                          )}
                        </>
                      )
                    })()}
                  </div>
                </div>
                )}
                <span className="text-[10px] text-text-dim mt-1 block pl-1">
                  {message.timestamp.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
                </span>
              </div>
            )}
            </motion.div>
          )
        })}
        {streamItems.map((item) => (
          <motion.div
            key={`stream-${item.id}`}
            layout
            variants={messageVariants}
            initial="hidden"
            animate="visible"
            exit="exit"
            className="group"
          >
            {item.type === 'text' ? (
              <div className="w-full">
                <div className="glass-card-soft rounded-2xl rounded-tl-sm px-4 py-3 border border-border">
                  <div className="text-[13px] leading-relaxed text-text prose prose-sm max-w-none">
                    <ReactMarkdown remarkPlugins={[remarkGfm]} components={markdownComponents}>
                      {item.content}
                    </ReactMarkdown>
                  </div>
                </div>
              </div>
            ) : (
              <div className="w-full">
                {(() => {
                  const jumpTarget = item.data.replaceText || item.data.searchText || ''
                  const canJump = !!jumpTarget
                  const cardClass = `glass-card-soft rounded-xl border border-border px-3 py-2.5 ${canJump ? 'hover:border-accent/40 hover:bg-accent/5 cursor-pointer' : ''}`

                  const cardBody = (
                    <>
                      <div className="flex items-center gap-2 text-[12px]">
                        {item.data.status === 'running' ? (
                          <Loader2 className="w-3.5 h-3.5 text-accent animate-spin flex-shrink-0" />
                        ) : item.data.status === 'success' ? (
                          <CheckCircle2 className="w-3.5 h-3.5 text-success flex-shrink-0" />
                        ) : item.data.status === 'skipped' ? (
                          <Circle className="w-3.5 h-3.5 text-text-muted flex-shrink-0" />
                        ) : (
                          <X className="w-3.5 h-3.5 text-error flex-shrink-0" />
                        )}
                        <span className="px-1.5 py-0.5 rounded bg-accent/10 text-accent text-[11px] font-medium">
                          {item.data.tool}
                        </span>
                        <span className="text-text truncate flex-1" title={item.data.label}>
                          {item.data.label}
                        </span>
                        {item.data.detail && (
                          <span className="text-[10px] text-text-muted flex-shrink-0">{item.data.detail}</span>
                        )}
                      </div>
                      {(item.data.searchText || item.data.replaceText) && (
                        <div className="mt-2 text-[11px] text-text-secondary flex items-center gap-2">
                          <span className="text-error truncate max-w-[40%]" title={item.data.searchText}>
                            {item.data.searchText || '-'}
                          </span>
                          <span className="text-text-muted">-&gt;</span>
                          <span className="text-success truncate max-w-[40%]" title={item.data.replaceText}>
                            {item.data.replaceText || '-'}
                          </span>
                        </div>
                      )}
                      {canJump && (
                        <div className="mt-1 text-[10px] text-accent">Click to jump to modified location</div>
                      )}
                    </>
                  )

                  if (canJump) {
                    return (
                      <button
                        type="button"
                        onClick={() => scrollToChange(jumpTarget)}
                        className={`${cardClass} w-full text-left`}
                      >
                        {cardBody}
                      </button>
                    )
                  }

                  return <div className={cardClass}>{cardBody}</div>
                })()}
              </div>
            )}
          </motion.div>
        ))}
        {/* 实时流式文本 — 跟在 streamItems 后面交替显示 */}
        {isLoading && streamItems.length > 0 && streamingContent && (
          <motion.div
            key="stream-live-text"
            layout
            variants={messageVariants}
            initial="hidden"
            animate="visible"
            exit="exit"
            className="group"
          >
            <div className="w-full">
              <div className="glass-card-soft rounded-2xl rounded-tl-sm px-4 py-3 border border-border">
                <div className="text-[13px] leading-relaxed text-text prose prose-sm max-w-none">
                  <ReactMarkdown remarkPlugins={[remarkGfm]} components={markdownComponents}>
                    {streamingContent}
                  </ReactMarkdown>
                </div>
              </div>
            </div>
          </motion.div>
        )}
        </AnimatePresence>

        {/* 流式输出 - 实时显示 AI 响应（去掉 layout 防抖动） */}
        <AnimatePresence mode="wait">
          {isLoading && (
            <motion.div 
              className="w-full"
              variants={streamingVariants}
              initial="hidden"
              animate="visible"
              exit="exit"
            >
              <div className="glass-card-soft rounded-2xl rounded-tl-sm px-4 py-3 border border-border">
                <div className="flex items-center gap-2">
                  <span className="thinking-indicator">
                    <span className="thinking-pulse"></span>
                  </span>
                  <span className="text-[13px] text-text font-medium flex-1">
                    {waitingPrimaryLabel}
                  </span>
                  <span className="text-[11px] text-text-dim flex-shrink-0">
                    {requestElapsedSec}s
                  </span>
                </div>
                <div className="mt-2 text-[12px] text-text-secondary">
                  {waitingSecondaryLabel}
                </div>
                {requestRunState.activeToolLabel && (
                  <div className="mt-2 text-[11px] text-text-dim">
                    当前目标：{requestRunState.activeToolLabel}
                  </div>
                )}
                <div className="mt-3 flex items-center gap-2">
                  <button
                    type="button"
                    onClick={stopGeneration}
                    className="px-3 py-1.5 rounded-lg bg-black/10 dark:bg-white/10 border border-black/10 dark:border-white/10 text-[11px] text-text hover:bg-black/15 dark:hover:bg-white/15 transition-colors"
                  >
                    停止本次生成
                  </button>
                  {canRetryCurrentRequest && (
                    <button
                      type="button"
                      onClick={() => {
                        void retryLastAgentRequest()
                      }}
                      className="px-3 py-1.5 rounded-lg bg-warning/12 border border-warning/25 text-[11px] text-warning hover:bg-warning/18 transition-colors"
                    >
                      重新发送上一条请求
                    </button>
                  )}
                </div>
              </div>

              {requestRunState.hasRealReasoning && streamingReasoning.trim() && (
                <details
                  className="thinking-section mt-2"
                  open={!streamingContent || streamingContent.length < 10}
                >
                  <summary className="thinking-summary">
                    <span className="thinking-indicator">
                      <span className="thinking-pulse"></span>
                    </span>
                    <span>模型思维链（原始输出）</span>
                    <svg className="thinking-arrow" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <path d="M9 18l6-6-6-6" />
                    </svg>
                  </summary>
                  <div className="thinking-body">
                    <CinematicTyper text={streamingReasoning} isStreaming={isLoading} baseSpeed={2} maxSpeed={8} />
                  </div>
                </details>
              )}

              {streamingContent && streamItems.length === 0 && (
                <div className="glass-card-soft rounded-2xl rounded-tl-sm px-4 py-3 border border-border mt-2">
                  <div className="text-[11px] text-text-dim mb-1">Streaming output</div>
                  <div className="text-[13px] leading-relaxed text-text prose prose-sm max-w-none">
                    <ReactMarkdown remarkPlugins={[remarkGfm]} components={markdownComponents}>
                      {streamingContent}
                    </ReactMarkdown>
                  </div>
                </div>
              )}
            </motion.div>
          )}
        </AnimatePresence>

        {/* Agent 进度 - Cursor 风格 + 动画 */}
        <AnimatePresence>
          {agentProgress.isActive && (
            <motion.div 
              className="w-full"
              layout
              variants={controlBarVariants}
              initial="hidden"
              animate="visible"
              exit="exit"
            >
              <div className="glass-card-soft border border-border rounded-xl px-3 py-2">
                <div className="flex items-center gap-2">
                  <Loader2 className="w-3.5 h-3.5 text-accent animate-spin flex-shrink-0" />
                  <span className="text-[12px] text-text flex-1 truncate">
                    {agentProgress.currentAction}
                  </span>
                  {agentProgress.thinkingTime > 0 && (
                    <span className="text-[10px] text-text-muted flex-shrink-0">
                      {agentProgress.thinkingTime}s
                    </span>
                  )}
                </div>
                {toolActivity.length > 0 && (
                  <div className="mt-2 border-t border-border pt-2">
                    <div className="text-[10px] text-text-dim uppercase tracking-wider mb-1">工具调用 ({toolActivity.length})</div>
                    <div className="space-y-1 max-h-[240px] overflow-y-auto">
                      {toolActivity.map(activity => {
                        const jumpTarget = activity.replaceText || activity.searchText || ''
                        const canJump = !!jumpTarget

                        return (
                          <button
                            key={activity.id}
                            type="button"
                            onClick={() => {
                              if (canJump) scrollToChange(jumpTarget)
                            }}
                            className={`w-full text-left rounded px-1.5 py-1 transition-colors ${canJump ? 'hover:bg-white/8 cursor-pointer' : 'cursor-default'}`}
                          >
                            <div className="flex items-center gap-1.5 text-[11px] text-text-secondary">
                              {activity.status === 'running' ? (
                                <Loader2 className="w-3 h-3 text-accent animate-spin flex-shrink-0" />
                              ) : activity.status === 'success' ? (
                                <CheckCircle2 className="w-3 h-3 text-success flex-shrink-0" />
                              ) : (
                                <X className="w-3 h-3 text-error flex-shrink-0" />
                              )}
                              <span className="truncate flex-1">{activity.label}</span>
                              {activity.detail && (
                                <span className="text-[10px] text-text-muted flex-shrink-0">{activity.detail}</span>
                              )}
                            </div>
                            {(activity.searchText || activity.replaceText) && (
                              <div className="mt-0.5 text-[10px] text-text-dim truncate">
                                {activity.searchText || '-'} -&gt; {activity.replaceText || '-'}
                              </div>
                            )}
                          </button>
                        )
                      })}
                    </div>
                  </div>
                )}
              </div>
            </motion.div>
          )}
        </AnimatePresence>

        {renderSubagentMonitor()}

        <div ref={messagesEndRef} />
      </div>

      {/* 上下文文件显示 - Cursor 风格 */}
      <div className="px-3 py-2 border-t border-border bg-black/5 dark:bg-white/5">
        <div className="flex items-center gap-1.5 flex-wrap">
          <span className="text-[10px] text-text-muted">上下文:</span>
          
          {/* 当前编辑的文档 */}
          {currentFile && (
            <div className="flex items-center gap-1 px-1.5 py-0.5 bg-success/10 text-success text-[10px] rounded border border-success/20">
              <FileText className="w-2.5 h-2.5" />
              <span className="max-w-[80px] truncate">{currentFile.name}</span>
            </div>
          )}
          
          {/* 用户拖拽的附加文件 */}
          {attachedFiles.map((file) => (
            <div
              key={file.path}
              className="flex items-center gap-1 px-1.5 py-0.5 bg-accent/10 text-accent text-[10px] rounded border border-accent/20"
            >
              <FileText className="w-2.5 h-2.5" />
              <span className="max-w-[60px] truncate">{file.name}</span>
              <button onClick={() => removeAttachedFile(file.path)} className="hover:bg-accent/15 rounded p-0.5 -mr-0.5">
                <X className="w-2.5 h-2.5" />
              </button>
            </div>
          ))}

          {/* 用户拖拽的文件夹 */}
          {attachedFolders.map((folder) => (
            <div
              key={folder.path}
              className="flex items-center gap-1 px-1.5 py-0.5 bg-amber-500/10 text-amber-400 text-[10px] rounded border border-amber-500/20"
            >
              <Folder className="w-2.5 h-2.5" />
              <span className="max-w-[60px] truncate">{folder.name}</span>
              <button onClick={() => removeAttachedFolder(folder.path)} className="hover:bg-amber-500/15 rounded p-0.5 -mr-0.5">
                <X className="w-2.5 h-2.5" />
              </button>
            </div>
          ))}

          {!currentFile && attachedFiles.length === 0 && attachedFolders.length === 0 && (
            <span className="text-[10px] text-text-dim">拖拽文件或文件夹添加上下文</span>
          )}
        </div>
      </div>

      {/* 快捷命令提示 - Cursor 风格 */}
      {input.startsWith('/') && !isLoading && (
        <div className="px-3 py-2 border-t border-border bg-black/5 dark:bg-white/5">
          <div className="space-y-0.5">
            {[
              { cmd: '/审查', desc: '审查文档，找出问题并建议修改' },
              { cmd: '/校对', desc: '检查语法、用词、逻辑问题' },
              { cmd: '/润色', desc: '优化文字表达' },
              { cmd: '/精简', desc: '删除冗余内容' },
              { cmd: '/翻译', desc: '翻译成英文/中文' },
              { cmd: '/格式化', desc: '统一文档格式' },
              { cmd: '/编号', desc: '自动添加标题编号' },
              { cmd: '/公文', desc: '转换为公文格式' },
              { cmd: '/会议纪要', desc: '整理为会议纪要' },
              { cmd: '/总结', desc: '生成文档摘要' },
            ].filter(item => item.cmd.includes(input) || input === '/').map((item) => (
              <button
                key={item.cmd}
                onClick={() => setInput(item.cmd + ' ')}
                className="w-full flex items-center justify-between px-2 py-1.5 hover:bg-black/10 dark:hover:bg-white/10 rounded text-left"
              >
                <span className="text-[12px] text-accent">{item.cmd}</span>
                <span className="text-[10px] text-text-muted">{item.desc}</span>
              </button>
            ))}
          </div>
        </div>
      )}

      {/* Word 格式操作确认条（dryRun → apply） */}
      {pendingWordOps && !isLoading && (
        <div className="px-3 py-2 border-t border-border bg-black/5 dark:bg-white/5">
          <div className="flex items-center gap-2">
            <div className="flex-1 min-w-0">
              <div className="text-[12px] text-text truncate">
                {pendingWordOps.previewMessage || '已生成格式修改预览'}
              </div>
              <div className="text-[10px] text-text-muted truncate">
                {pendingWordOps.previewLines?.length
                  ? pendingWordOps.previewLines.join(' · ')
                  : '点击应用后将以“修订”方式写入，可逐条接受/拒绝'}
              </div>
            </div>
            <button
              disabled={wordOpsApplying}
              onClick={async () => {
                if (!pendingWordOps) return
                setWordOpsApplying(true)
                try {
                  const prep = await prepareTemplateFillOutput(pendingWordOps.ops as any)
                  if (!prep.success) {
                    addMessage({
                      role: 'assistant',
                      content: `应用失败：${prep.message || '模板填充准备失败'}`,
                    })
                    return
                  }
                  const result = applyWordOps(pendingWordOps.ops as any)
                  setPendingWordOps(null)
                  addMessage({
                    role: 'assistant',
                    content: result.success
                      ? `已应用格式修订：${result.message}`
                      : `应用失败：${result.message}`,
                  })
                } finally {
                  setWordOpsApplying(false)
                }
              }}
              className="flex items-center gap-1.5 px-2.5 py-1.5 bg-accent/12 border border-accent/25 hover:bg-accent/18 text-accent text-[11px] rounded-xl transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
              title="应用修订"
            >
              <CheckCircle2 className="w-3.5 h-3.5" />
              应用修订
            </button>
            <button
              disabled={wordOpsApplying}
              onClick={() => setPendingWordOps(null)}
              className="p-1.5 rounded-md text-text-muted hover:text-text hover:bg-black/10 dark:hover:bg-white/10 transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
              title="取消"
            >
              <X className="w-4 h-4" />
            </button>
          </div>
        </div>
      )}

      {/* PPT 大纲确认条（阶段1 → 阶段2） */}
      {pendingPptOutline && !pptGenerating && (
        <div className="px-3 py-2 border-t border-border bg-black/5 dark:bg-white/5">
          <div className="flex items-center gap-2">
            <div className="flex-1 min-w-0">
              <div className="text-[12px] text-text truncate">
                已检测到 PPT 大纲：{pendingPptOutline.draft.title || '未命名'}（{pendingPptOutline.draft.slides?.length || 0} 页）
              </div>
              <div className="text-[10px] text-text-muted truncate">
                点击确认后将直接开始生成（Gemini 设计视觉 → DashScope 生图 → 导出 PPTX）
              </div>
            </div>
            <button
              disabled={isLoading || pptGenerating}
              onClick={() => {
                const { draft, rawJson } = pendingPptOutline
                setPendingPptOutline(null)
                executePptCreate(draft, rawJson)
              }}
              className="flex items-center gap-1.5 px-2.5 py-1.5 bg-accent/12 border border-accent/25 hover:bg-accent/18 text-accent text-[11px] rounded-xl transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
              title="确认大纲并开始生成 PPT"
            >
              <CheckCircle2 className="w-3.5 h-3.5" />
              确认并开始生成
            </button>
            <button
              disabled={isLoading || pptGenerating}
              onClick={() => setPendingPptOutline(null)}
              className="p-1.5 rounded-md text-text-muted hover:text-text hover:bg-black/10 dark:hover:bg-white/10 transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
              title="关闭提示"
            >
              <X className="w-4 h-4" />
            </button>
          </div>
        </div>
      )}

      {/* PPT 编辑反馈输入区域 */}
      {pptEditPending && !pptGenerating && (
        <div className="px-3 py-2 border-t border-border bg-black/5 dark:bg-white/5">
          <div className="flex flex-col gap-2">
            <div className="flex items-center gap-2">
              <div className="flex-1 min-w-0">
                <div className="text-[12px] text-text">
                  {pptEditPending.mode === 'regenerate' ? '🔄 整页重做' : '🎨 局部编辑'}：
                  {pptEditPending.pageNumbers.length === 1 
                    ? `第 ${pptEditPending.pageNumbers[0]} 页`
                    : `${pptEditPending.pageNumbers.length} 页（${pptEditPending.pageNumbers.join(', ')}）`
                  }
                </div>
                <div className="text-[10px] text-text-muted">
                  {pptEditPending.mode === 'regenerate' 
                    ? '请描述你对这些页面不满意的地方，AI 将根据反馈重新生成'
                    : '请描述你想要修改的部分（如：换背景颜色、改文字大小等）'
                  }
                </div>
              </div>
              <button
                onClick={() => {
                  setPptEditPending(null)
                  setPptEditFeedback('')
                }}
                className="p-1.5 rounded-md text-text-muted hover:text-text hover:bg-black/10 dark:hover:bg-white/10 transition-colors"
                title="取消"
              >
                <X className="w-4 h-4" />
              </button>
            </div>
            <div className="flex gap-2">
              <input
                type="text"
                value={pptEditFeedback}
                onChange={(e) => setPptEditFeedback(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter' && !e.shiftKey && pptEditFeedback.trim()) {
                    e.preventDefault()
                    const { pptxPath, pageNumbers, mode } = pptEditPending
                    setPptEditPending(null)
                    executePptEdit(pptxPath, pageNumbers, mode, pptEditFeedback.trim())
                    setPptEditFeedback('')
                  }
                }}
                placeholder={pptEditPending.mode === 'regenerate' ? '例如：背景太暗，配色不协调，标题太小...' : '例如：背景换成蓝色渐变，标题放大一点...'}
                className="flex-1 glass-input rounded-xl px-3 py-1.5 text-[12px] text-text placeholder-text-dim focus:outline-none focus:border-accent/40 focus:ring-2 focus:ring-accent/10"
                autoFocus
              />
              <button
                disabled={!pptEditFeedback.trim()}
                onClick={() => {
                  const { pptxPath, pageNumbers, mode } = pptEditPending
                  setPptEditPending(null)
                  executePptEdit(pptxPath, pageNumbers, mode, pptEditFeedback.trim())
                  setPptEditFeedback('')
                }}
                className="flex items-center gap-1.5 px-3 py-1.5 bg-accent/12 border border-accent/25 hover:bg-accent/18 text-accent text-[11px] rounded-xl transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
              >
                <Send className="w-3.5 h-3.5" />
                开始{pptEditPending.mode === 'regenerate' ? '重做' : '编辑'}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* 输入区域 - 柔和玻璃态风格 */}
      <div 
        className={`p-4 border-t transition-colors ${
          isPptDragOver ? 'border-accent/50 bg-accent/5' : 'border-border'
        }`}
        onDragEnter={(e) => {
          if (!e.dataTransfer.types.includes('application/ppt-page')) return
          e.preventDefault()
          e.stopPropagation()
          pptDragCounterRef.current += 1
          setIsPptDragOver(true)
        }}
        onDragOver={(e) => {
          // 检查是否是 PPT 页面拖拽
          if (e.dataTransfer.types.includes('application/ppt-page')) {
            e.preventDefault()
            e.stopPropagation()
            e.dataTransfer.dropEffect = 'copy'
            // 不要在 onDragOver 里反复 setState，避免闪烁
          }
        }}
        onDragLeave={(e) => {
          if (!isPptDragOver) return
          e.preventDefault()
          e.stopPropagation()
          pptDragCounterRef.current = Math.max(0, pptDragCounterRef.current - 1)
          if (pptDragCounterRef.current === 0) {
            setIsPptDragOver(false)
          }
        }}
        onDrop={(e) => {
          const pptData = e.dataTransfer.getData('application/ppt-page')
          if (!pptData) return // 非 PPT 拖拽：交给外层文件拖拽逻辑

          e.preventDefault()
          e.stopPropagation()
          pptDragCounterRef.current = 0
          setIsPptDragOver(false)

          try {
            const { pageNumber, imageBase64, pptxPath } = JSON.parse(pptData)
            setPptEditContext({
              pageNumber,
              imageBase64,
              pptxPath,
              isRegion: false,
            })
            inputRef.current?.focus()
          } catch (err) {
            console.error('解析拖拽数据失败:', err)
          }
        }}
      >
        {/* PPT 编辑上下文预览 */}
        {pptEditContext && (
          <div className="mb-2 p-2 bg-black/5 dark:bg-white/5 rounded-xl border border-border flex items-start gap-3">
            <div className="relative flex-shrink-0">
              <img
                src={`data:image/png;base64,${pptEditContext.imageBase64}`}
                alt={`第 ${pptEditContext.pageNumber} 页${pptEditContext.isRegion ? '（框选区域）' : ''}`}
                className="w-[100px] h-[62px] object-contain rounded-lg border border-border bg-black/20"
              />
              <div className="absolute -top-1 -left-1 bg-accent text-[9px] text-white px-1.5 py-0.5 rounded-md shadow-sm shadow-accent/20">
                {pptEditContext.isRegion ? '框选' : `第 ${pptEditContext.pageNumber} 页`}
              </div>
            </div>
            <div className="flex-1 min-w-0">
              <div className="text-[11px] text-text mb-1">
                {pptEditContext.isRegion ? (
                  <>已框选第 <span className="text-accent font-medium">{pptEditContext.pageNumber}</span> 页的区域</>
                ) : (
                  <>已选择第 <span className="text-accent font-medium">{pptEditContext.pageNumber}</span> 页</>
                )}
              </div>
              <div className="text-[10px] text-text-muted">
                输入修改要求，AI 将自动判断是整页重做还是局部调整
              </div>
            </div>
            <button
              onClick={() => setPptEditContext(null)}
              className="p-1 text-text-muted hover:text-text hover:bg-black/10 dark:hover:bg-white/10 rounded-lg transition-colors"
              title="移除"
            >
              <X className="w-3.5 h-3.5" />
            </button>
          </div>
        )}
        
        {/* 拖拽提示 */}
        {isPptDragOver && (
          <div className="mb-2 p-3 border-2 border-dashed border-accent/50 rounded-xl bg-accent/6 text-center">
            <div className="text-[12px] text-accent">松开鼠标，将 PPT 页面添加到对话</div>
          </div>
        )}
        
        {/* 粘贴/上传的图片预览 */}
        {pendingImages.length > 0 && (
          <div className="flex gap-2 mb-2 flex-wrap">
            {pendingImages.map((img, idx) => (
              <div key={idx} className="relative group w-16 h-16 rounded-lg overflow-hidden border border-border/40 bg-bg-secondary">
                <img src={img} alt={`图片 ${idx + 1}`} className="w-full h-full object-cover" />
                <button
                  onClick={() => removePendingImage(idx)}
                  className="absolute top-0 right-0 p-0.5 bg-black/60 text-white rounded-bl-md opacity-0 group-hover:opacity-100 transition-opacity"
                >
                  <X className="w-3 h-3" />
                </button>
              </div>
            ))}
          </div>
        )}

        <div className="relative">
          <textarea
            ref={inputRef}
            value={input}
            onChange={(e) => setInput(e.target.value)}
            onKeyDown={handleKeyDown}
            onPaste={handlePaste}
            placeholder={
              pptEditContext 
                ? `描述如何修改第 ${pptEditContext.pageNumber} 页...` 
                : isLoading 
                  ? "AI 正在处理中..." 
                  : "输入问题或 / 查看命令...（可粘贴图片）"
            }
            className={`w-full glass-input rounded-2xl pl-10 pr-12 py-3 text-[13px] text-text placeholder-text-dim focus:outline-none transition-all resize-none scrollbar-none ${
              isLoading ? 'border-accent/20' : pptEditContext ? 'border-accent/35 focus:border-accent/50' : ''
            }`}
            rows={2}
            disabled={isLoading}
          />
          {/* 图片上传按钮 */}
          <button
            onClick={() => imageInputRef.current?.click()}
            disabled={isLoading}
            className="absolute left-3 bottom-3 p-1.5 rounded-lg text-text-muted hover:text-text hover:bg-black/10 dark:hover:bg-white/10 transition-all disabled:opacity-30 disabled:cursor-not-allowed"
            title="上传图片"
          >
            <ImagePlus className="w-4 h-4" />
          </button>
          <input
            ref={imageInputRef}
            type="file"
            accept="image/*"
            multiple
            className="hidden"
            onChange={handleImageUpload}
          />
          <button
            onClick={isLoading ? stopGeneration : handleSend}
            disabled={!isLoading && (!input.trim() && pendingImages.length === 0)}
            className={`absolute right-3 bottom-3 p-2 rounded-xl transition-all disabled:cursor-not-allowed ${
              isLoading
                ? 'text-red-400 bg-red-500/10 hover:bg-red-500/20'
                : 'text-text-muted hover:text-text hover:bg-black/10 dark:hover:bg-white/10 disabled:opacity-30'
            }`}
            title={isLoading ? '停止生成' : '发送'}
          >
            {isLoading ? (
              <Square className="w-4 h-4 fill-current" />
            ) : (
              <Send className="w-4 h-4" />
            )}
          </button>
        </div>
        
        <p className="text-[10px] text-text-dim text-center mt-2">
          {pptEditContext ? (
            <span className="text-accent">输入修改要求后按 Enter 发送</span>
          ) : (
            <>AI can make mistakes. Review generated code.</>
          )}
        </p>
      </div>
    </div>
  )
}
