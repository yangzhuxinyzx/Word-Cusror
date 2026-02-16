import { createContext, useContext, useState, useCallback, ReactNode, useRef, useEffect } from 'react'
import { DocumentContent, DocumentStyles, FileItem, PageSetup, HeaderFooterSetup, CustomStyle } from '../types'
import type { ExcelOpenResponse } from '../types/electron'
import { saveAs } from 'file-saver'
import { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell, WidthType, BorderStyle, UnderlineType } from 'docx'
import { extractTypographyProfileFromArrayBuffer } from '../utils/docxTypography'
import type { DocxTypographyProfile } from '../utils/docxTypography'
import { toChineseDefaultFallbackStack } from '../fonts/fontManifest'
import { useComments } from './CommentContext'
import { postProcessDocxWithAnnotations } from '../utils/docxExportWithTracking'
// DSL 支持
import type { DocDsl, DslBlock, DslRun, DslInline, DslBlockMeta, DslRunMeta } from '../types/docDsl'
import { validateDocDsl, dslToHtml, normalizeContent, extractPlainText } from '../utils/docDsl'
import { dslToDocxBlob } from '../utils/docDslToDocx'
import { htmlToDsl } from '../utils/htmlToDsl'
import { useMcpBridge } from '../hooks/useMcpBridge'

// 将 ArrayBuffer 转换为 Base64（分块处理，避免大文件导致栈溢出）
function arrayBufferToBase64(buffer: ArrayBuffer): string {
  const bytes = new Uint8Array(buffer)
  const chunkSize = 8192 // 每次处理 8KB
  let binary = ''
  
  for (let i = 0; i < bytes.length; i += chunkSize) {
    const chunk = bytes.subarray(i, Math.min(i + chunkSize, bytes.length))
    binary += String.fromCharCode.apply(null, Array.from(chunk))
  }
  
  const base64 = btoa(binary)
  return base64
}

function b64uEncodeUtf8(input: string): string {
  try {
    const b64 = btoa(unescape(encodeURIComponent(input)))
    return b64.replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/g, '')
  } catch {
    return ''
  }
}

function stripQuotes(s: string): string {
  return s.replace(/['"]/g, '').trim()
}

interface ReplaceResult {
  success: boolean
  count: number
  message: string
  searchText?: string
  replaceText?: string
  positions?: number[]  // match positions in plain-text projection
  debug?: {
    reasonCode: 'not_found' | 'empty_search' | 'already_applied'
    effectiveSearch?: string
    hints?: string[]
  }
}

// 单个替换记录
interface SingleReplacement {
  id: string  // 唯一标识
  searchText: string
  replaceText: string
  count: number
  timestamp: number
  // 审查专属字段（review 工具使用）
  reviewReason?: string
  reviewType?: string
}

// 操作类型
type WordEditOpType = 
  | 'replace_text' 
  | 'format_text' 
  | 'format_paragraph' 
  | 'apply_style' 
  | 'clear_format' 
  | 'copy_format' 
  | 'list_edit' 
  | 'insert_page_break' 
  | 'structure_edit' 
  | 'table_edit' 
  | 'image_edit'
  | 'page_setup'      // 页面设置
  | 'header_footer'   // 页眉页脚
  | 'section_break'   // 分节符
  | 'columns'         // 分栏
  | 'watermark'       // 水印
  | 'toc'             // 目录
  | 'define_style'    // 定义样式
  | 'modify_style'    // 修改样式
  | 'outline_summary' // 大纲/摘要生成
  | 'template_fill'   // 模板智能填充
  | 'citation_footnote' // 引用/脚注管理

// 统一的待审阅变更（M1：先由替换记录映射；M3 起会扩展为格式/样式变更）
export interface PendingChange {
  id: string
  kind: WordEditOpType
  scope: 'selection' | 'document' | 'rule'
  summary: string
  beforePreview?: string
  afterPreview?: string
  stats?: { matches: number }
  timestamp: number
  meta?: Record<string, unknown>
  // 审查专属字段（review 工具使用）
  reviewReason?: string   // 修改原因（如 "语病修正"、"用词不当"）
  reviewType?: string     // 问题类型（grammar/logic/style/typo/format）
}

export type WordEditOp = {
  opId?: string
  type: WordEditOpType
  target: {
    scope: 'selection' | 'anchor_text' | 'document' | 'rule'
    text?: string
    filter?: Record<string, unknown>
  }
  params?: Record<string, unknown>
  dryRun?: boolean
}

type TemplateFieldCandidate = {
  id: string
  label: string
  kind: 'colon' | 'blank' | 'table'
  context: string
  path?: string
  currentValue?: string
  fieldType?: string
  groupKey?: string
  meta?: Record<string, unknown>
}

// 最近的替换记录（支持多个）
interface ReplacementRecord {
  searchText: string
  replaceText: string
  count: number
  timestamp: number
  pending: boolean  // 是否待确认
}

// 待确认的替换列表
interface PendingReplacements {
  items: SingleReplacement[]
  total: number
}

// 编辑器模式
type EditorMode = 'tiptap' | 'onlyoffice'

interface DocumentContextType {
  document: DocumentContent
  files: FileItem[]
  currentFile: FileItem | null
  workspacePath: string | null
  isElectron: boolean
  hasUnsavedChanges: boolean
  docxData: string | null
  excelData: ExcelOpenResponse | null
  pptData: { pptxBase64: string } | null
  /** 当前打开 docx 的排版/字体配置（从 docx 解析得到；用于显示与导出保持一致） */
  typographyProfile: DocxTypographyProfile | null
  refreshExcelData: () => Promise<boolean>  // 刷新 Excel 数据
  lastReplacement: ReplacementRecord | null  // 最近的替换记录
  pendingChanges: PendingChange[] // 待审阅修改（逐条）
  pendingChangesTotal: number // 待审阅修改命中总数（用于 UI 显示）
  editorMode: EditorMode  // 当前编辑器模式
  setEditorMode: (mode: EditorMode) => void
  setDocument: React.Dispatch<React.SetStateAction<DocumentContent>>
  updateDocument: (updates: Partial<DocumentContent>) => void
  updateContent: (content: string) => void
  updateStyles: (styles: Partial<DocumentStyles>) => void
  setCurrentFile: (file: FileItem | null) => void
  addFile: (file: FileItem) => void
  createNewDocument: (title: string, content: string, elements?: FormattedElement[], styleRefPath?: string) => void
  /** DSL 方式创建文档 */
  createDocumentFromDsl: (title: string, dsl: DocDsl) => Promise<{ success: boolean; message: string; filePath?: string }>
  uploadDocxFile: (file: File) => Promise<void>
  saveDocument: () => Promise<void>
  applyAIEdit: (newContent: string) => void
  replaceInDocument: (search: string, replace: string, reviewMeta?: { reason?: string; type?: string }) => ReplaceResult
  insertInDocument: (position: string, content: string) => { success: boolean; message: string }
  deleteInDocument: (target: string) => { success: boolean; count: number; message: string }
  replaceViaDsl: (search: string, replace: string, options?: { blockIndex?: number; format?: Partial<import('../types/docDsl').DslRun> }) => { success: boolean; count: number; message: string }
  insertViaDsl: (position: string, dslBlocks: import('../types/docDsl').DslBlock[]) => { success: boolean; message: string }
  deleteViaDsl: (target: string, options?: { blockIndex?: number }) => { success: boolean; count: number; message: string }
  scrollToText: (text: string) => void  // 滚动到指定文本
  confirmReplacement: () => void  // 确认替换
  rejectReplacement: () => void   // 拒绝替换
  acceptChange: (id: string) => void // 逐条接受
  rejectChange: (id: string) => void // 逐条拒绝
  acceptAllChanges: () => void // 全部接受
  rejectAllChanges: () => void // 全部拒绝
  openFolder: () => Promise<void>
  openFile: (file: FileItem) => Promise<void>
  saveCurrentFile: () => Promise<void>
  /** Agent 静默保存：将编辑器内容写回当前 .docx 文件，无弹窗确认 */
  silentSaveToFile: () => Promise<{ success: boolean; error?: string }>
  refreshFiles: () => Promise<void>
  // ONLYOFFICE 专用操作
  onlyOfficeReplace: (search: string, replace: string) => Promise<ReplaceResult>
  onlyOfficeInsert: (text: string) => Promise<{ success: boolean; message: string }>
  onlyOfficeGetText: () => Promise<string>
  // ONLYOFFICE 格式化操作
  onlyOfficeAddParagraph: (text: string, options?: {
    fontSize?: number
    fontFamily?: string
    bold?: boolean
    italic?: boolean
    color?: string
    alignment?: 'left' | 'center' | 'right' | 'justify'
  }) => Promise<{ success: boolean; message: string }>
  onlyOfficeAddHeading: (text: string, level: 1 | 2 | 3 | 4 | 5 | 6) => Promise<{ success: boolean; message: string }>
  onlyOfficeAddTable: (rows: number, cols: number, data?: string[][]) => Promise<{ success: boolean; message: string }>
  // Tiptap 文档结构获取
  getTiptapDocumentStructure: () => string
  // 定位到指定 diffId（用于 RevisionPanel）
  scrollToDiffId: (diffId: string) => void
  // 仅登记一条待审阅修改（不改动文档内容；用于选区 AI 修订等场景）
  addPendingReplacementItem: (item: SingleReplacement) => void
  // word_edit_ops：预览 & 应用（用于样式/段落/字符格式）
  previewWordOps: (ops: WordEditOp[]) => { success: boolean; message: string; data?: Record<string, unknown> }
  applyWordOps: (ops: WordEditOp[]) => { success: boolean; message: string; data?: Record<string, unknown> }
  // 格式化替换
  replaceWithFormat: (search: string, replace: string, format?: {
    bold?: boolean
    italic?: boolean
    underline?: boolean
    color?: string
    backgroundColor?: string
    fontSize?: string
    fontFamily?: string
  }, reviewMeta?: { reason?: string; type?: string }) => ReplaceResult
  // 动画控制
  docEntryAnimationKey: number
  triggerDocEntryAnimation: () => void
  // 获取最新文档内容（使用 ref，避免闭包问题）
  getLatestContent: () => string
  // 页面设置
  pageSetup: PageSetup
  setPageSetup: (setup: Partial<PageSetup>) => void
  // 页眉页脚设置
  headerFooterSetup: HeaderFooterSetup
  setHeaderFooterSetup: (setup: Partial<HeaderFooterSetup>) => void
  // 自定义样式
  customStyles: Record<string, CustomStyle>
  defineStyle: (style: CustomStyle) => void
  modifyStyle: (name: string, updates: Partial<CustomStyle>) => void
  deleteStyle: (name: string) => void
  getStyleCSS: (styleName: string) => string
}

const defaultStyles: DocumentStyles = {
  fontSize: 14,
  fontFamily: '等线',
  lineHeight: 1.15,
  textAlign: 'left',
}

const defaultPageSetup: PageSetup = {
  paperSize: 'A4',
  orientation: 'portrait',
  margins: {
    top: '2.54cm',
    bottom: '2.54cm',
    left: '3.17cm',
    right: '3.17cm',
  },
}

const defaultHeaderFooterSetup: HeaderFooterSetup = {}

// 默认内置样式
const defaultCustomStyles: Record<string, CustomStyle> = {
  'Normal': {
    name: 'Normal',
    fontFamily: '等线',
    fontSize: '10.5pt',
    lineHeight: '1.15',
    textIndent: '0',
  },
  'Heading1': {
    name: 'Heading1',
    fontFamily: '黑体',
    fontSize: '20pt',
    bold: true,
    alignment: 'left',
    spaceBefore: '18pt',
    spaceAfter: '4pt',
  },
  'Heading2': {
    name: 'Heading2',
    fontFamily: '黑体',
    fontSize: '16pt',
    bold: true,
    spaceBefore: '8pt',
    spaceAfter: '4pt',
  },
  'Heading3': {
    name: 'Heading3',
    fontFamily: '黑体',
    fontSize: '14pt',
    bold: true,
    spaceBefore: '8pt',
    spaceAfter: '4pt',
  },
  'Quote': {
    name: 'Quote',
    fontFamily: '楷体',
    fontSize: '10.5pt',
    italic: true,
    color: '#666666',
    marginLeft: '2em',
    marginRight: '2em',
    border: '1px solid var(--word-rule)',
    backgroundColor: '#f9f9f9',
  },
}

const defaultDocument: DocumentContent = {
  title: '新建文档',
  content: '',
  styles: defaultStyles,
  lastModified: new Date(),
}

// Vite HMR: keep a stable context identity across hot updates.
// Otherwise, provider/consumer can mismatch and throw "useDocument must be used within a DocumentProvider".
const DocumentContext: React.Context<DocumentContextType | undefined> = (() => {
  try {
    const hot = (import.meta as any)?.hot
    const existing = hot?.data?.DocumentContext as React.Context<DocumentContextType | undefined> | undefined
    if (existing) return existing
    const ctx = createContext<DocumentContextType | undefined>(undefined)
    if (hot) hot.data.DocumentContext = ctx
    return ctx
  } catch {
    return createContext<DocumentContextType | undefined>(undefined)
  }
})()

// 检测是否在 Electron 环境
const isElectron = typeof window !== 'undefined' && !!window.electronAPI

async function fetchArrayBufferFromLocalFile(filePath: string): Promise<ArrayBuffer> {
  if (!window.electronAPI?.getLocalFileUrl) {
    throw new Error('getLocalFileUrl 不可用')
  }
  const url = await window.electronAPI.getLocalFileUrl(filePath)
  const resp = await fetch(url)
  if (!resp.ok) throw new Error(`读取文件失败（HTTP ${resp.status}）`)
  return await resp.arrayBuffer()
}

function cssLengthToTwips(value: string | undefined, baseFontPt = 10.5): number | undefined {
  const v = (value || '').trim().toLowerCase()
  if (!v) return undefined
  // pt
  const mPt = v.match(/^(\d+(?:\.\d+)?)\s*pt$/)
  if (mPt) return Math.round(Number(mPt[1]) * 20)
  // px (approx 1px = 0.75pt)
  const mPx = v.match(/^(\d+(?:\.\d+)?)\s*px$/)
  if (mPx) return Math.round(Number(mPx[1]) * 0.75 * 20)
  // cm (1in=2.54cm, 1in=1440 twips)
  const mCm = v.match(/^(\d+(?:\.\d+)?)\s*cm$/)
  if (mCm) return Math.round((Number(mCm[1]) / 2.54) * 1440)
  // em (relative to base font)
  const mEm = v.match(/^(\d+(?:\.\d+)?)\s*em$/)
  if (mEm) return Math.round(Number(mEm[1]) * baseFontPt * 20)
  return undefined
}

// Markdown 转换为 docx 段落
function markdownToDocxParagraphs(content: string, profile?: DocxTypographyProfile): Paragraph[] {
  const paragraphs: Paragraph[] = []
  const lines = content.split('\n')
  let inList = false
  let listType = ''

  const flushList = () => {
    inList = false
    listType = ''
  }

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i]
    const trimmedLine = line.trim()

    // 跳过空行
    if (!trimmedLine) {
      flushList()
      continue
    }

    // 处理分隔线
    if (/^(-{3,}|\*{3,}|_{3,})$/.test(trimmedLine)) {
      flushList()
      paragraphs.push(new Paragraph({
        children: [],
        border: { bottom: { style: 'single' as any, size: 6, space: 1, color: '999999' } },
        spacing: { before: 200, after: 200 },
      }))
      continue
    }

    // 处理标题
    if (trimmedLine.startsWith('### ')) {
      flushList()
      paragraphs.push(new Paragraph({
        children: parseInlineFormatting(trimmedLine.slice(4), true, profile),
        heading: HeadingLevel.HEADING_3,
        spacing: { before: 200, after: 100 },
      }))
      continue
    }
    if (trimmedLine.startsWith('## ')) {
      flushList()
      paragraphs.push(new Paragraph({
        children: parseInlineFormatting(trimmedLine.slice(3), true, profile),
        heading: HeadingLevel.HEADING_2,
        spacing: { before: 260, after: 130 },
      }))
      continue
    }
    if (trimmedLine.startsWith('# ')) {
      flushList()
      paragraphs.push(new Paragraph({
        children: parseInlineFormatting(trimmedLine.slice(2), true, profile),
        heading: HeadingLevel.HEADING_1,
        alignment: AlignmentType.CENTER,
        spacing: { before: 300, after: 200 },
      }))
      continue
    }

    // 处理无序列表
    if (/^[-*] /.test(trimmedLine)) {
      inList = true
      listType = 'bullet'
      paragraphs.push(new Paragraph({
        children: parseInlineFormatting(trimmedLine.slice(2), false, profile),
        bullet: { level: 0 },
        spacing: { after: 60 },
      }))
      continue
    }

    // 处理有序列表
    if (/^\d+\. /.test(trimmedLine)) {
      inList = true
      listType = 'number'
      const text = trimmedLine.replace(/^\d+\. /, '')
      paragraphs.push(new Paragraph({
        children: parseInlineFormatting(text, false, profile),
        numbering: { reference: 'default-numbering', level: 0 },
        spacing: { after: 60 },
      }))
      continue
    }

    // 处理引用
    if (trimmedLine.startsWith('> ')) {
      flushList()
      paragraphs.push(new Paragraph({
        children: parseInlineFormatting(trimmedLine.slice(2), false, profile),
        indent: { left: 720 },
        border: { left: { style: 'single' as any, size: 12, space: 10, color: 'CCCCCC' } },
        spacing: { after: 100 },
      }))
      continue
    }

    // 处理普通段落
    flushList()
    const firstLine = profile?.normal?.indent?.firstLineTwips ?? 480
    const after = profile?.normal?.spacing?.afterTwips ?? 120
    const lineTwips = profile?.normal?.spacing?.lineTwips ?? 360
    const align =
      profile?.normal?.alignment === 'justify'
        ? AlignmentType.JUSTIFIED
        : profile?.normal?.alignment === 'center'
          ? AlignmentType.CENTER
          : profile?.normal?.alignment === 'right'
            ? AlignmentType.RIGHT
            : AlignmentType.LEFT
    paragraphs.push(new Paragraph({
      children: parseInlineFormatting(trimmedLine, false, profile),
      indent: { firstLine }, // 默认首行缩进（可由参考样式覆盖）
      spacing: { after, line: lineTwips }, // 默认段后/行距（可由参考样式覆盖）
      alignment: align,
    }))
  }

  return paragraphs.length > 0 ? paragraphs : [new Paragraph({ children: [] })]
}

// 解析行内格式（粗体、斜体等）
function parseInlineFormatting(text: string, isHeading: boolean = false, profile?: DocxTypographyProfile): TextRun[] {
  const runs: TextRun[] = []
  const defaultSize = profile?.normal?.fontSizeHalfPoints ?? 21 // 10.5pt
  const fontSize = isHeading ? 28 : defaultSize
  const normalAscii = profile?.normal?.fontAscii || profile?.normal?.fontHAnsi
  const normalEastAsia = profile?.normal?.fontEastAsia
  const fallbackFont = normalEastAsia || normalAscii || '等线'
  const headingFont = '黑体'
  const font =
    isHeading
      ? { ascii: headingFont, hAnsi: headingFont, eastAsia: headingFont }
      : { ascii: normalAscii || fallbackFont, hAnsi: profile?.normal?.fontHAnsi || normalAscii || fallbackFont, eastAsia: normalEastAsia || fallbackFont }
  
  // 简化处理：用正则分割文本
  const regex = /(\*\*\*.+?\*\*\*|\*\*.+?\*\*|\*.+?\*|__.+?__|_.+?_)/g
  let lastIndex = 0
  let match

  while ((match = regex.exec(text)) !== null) {
    // 添加匹配前的普通文本
    if (match.index > lastIndex) {
      runs.push(new TextRun({
        text: text.slice(lastIndex, match.index),
        font,
        size: fontSize,
      }))
    }

    const matchedText = match[0]
    let content = matchedText
    let bold = false
    let italic = false

    // 粗斜体
    if (matchedText.startsWith('***') && matchedText.endsWith('***')) {
      content = matchedText.slice(3, -3)
      bold = true
      italic = true
    }
    // 粗体
    else if ((matchedText.startsWith('**') && matchedText.endsWith('**')) ||
             (matchedText.startsWith('__') && matchedText.endsWith('__'))) {
      content = matchedText.slice(2, -2)
      bold = true
    }
    // 斜体
    else if ((matchedText.startsWith('*') && matchedText.endsWith('*')) ||
             (matchedText.startsWith('_') && matchedText.endsWith('_'))) {
      content = matchedText.slice(1, -1)
      italic = true
    }

    runs.push(new TextRun({
      text: content,
      font,
      size: fontSize,
      bold,
      italics: italic,
    }))

    lastIndex = regex.lastIndex
  }

  // 添加剩余的普通文本
  if (lastIndex < text.length) {
    runs.push(new TextRun({
      text: text.slice(lastIndex),
      font,
      size: fontSize,
    }))
  }

  return runs.length > 0 ? runs : [new TextRun({ text, font, size: fontSize })]
}

// 将 HTML 转换为保留结构的格式化文本
function htmlToStructuredText(html: string): string {
  const parser = new DOMParser()
  const doc = parser.parseFromString(html, 'text/html')
  
  let listCounter = 0
  
  const processNode = (node: Node): string => {
    if (node.nodeType === Node.TEXT_NODE) {
      return node.textContent || ''
    }
    if (node.nodeType !== Node.ELEMENT_NODE) return ''
    
    const el = node as HTMLElement
    const tag = el.tagName.toLowerCase()
    
    const getChildren = (): string => {
      let result = ''
      for (const child of Array.from(el.childNodes)) {
        result += processNode(child)
      }
      return result
    }
    
    switch (tag) {
      case 'h1': return `# ${getChildren().trim()}\n\n`
      case 'h2': return `## ${getChildren().trim()}\n\n`
      case 'h3': return `### ${getChildren().trim()}\n\n`
      case 'h4': case 'h5': case 'h6': return `**${getChildren().trim()}**\n\n`
      case 'p': case 'div': {
        const text = getChildren().trim()
        return text ? `${text}\n\n` : ''
      }
      case 'br': return '\n'
      case 'strong': case 'b': return `**${getChildren()}**`
      case 'em': case 'i': return `*${getChildren()}*`
      case 'ul': {
        listCounter = 0
        let result = ''
        for (const li of Array.from(el.children)) {
          if (li.tagName.toLowerCase() === 'li') {
            result += `- ${processNode(li).trim()}\n`
          }
        }
        return result + '\n'
      }
      case 'ol': {
        listCounter = 0
        let result = ''
        for (const li of Array.from(el.children)) {
          if (li.tagName.toLowerCase() === 'li') {
            listCounter++
            result += `${listCounter}. ${processNode(li).trim()}\n`
          }
        }
        return result + '\n'
      }
      case 'li': return getChildren()
      case 'table': {
        let result = ''
        for (const row of Array.from(el.querySelectorAll('tr'))) {
          const cells = Array.from(row.querySelectorAll('td, th'))
          result += cells.map(c => c.textContent?.trim() || '').join('\t') + '\n'
        }
        return result + '\n'
      }
      default: return getChildren()
    }
  }
  
  let result = ''
  for (const child of Array.from(doc.body.childNodes)) {
    result += processNode(child)
  }
  return result.replace(/\n{3,}/g, '\n\n').trim()
}

// 创建 docx 文档
async function createDocxBlob(content: string, title: string, typographyProfile?: DocxTypographyProfile | null): Promise<Blob> {
  // 判断是 HTML 还是 Markdown
  const isHtml = content.trim().startsWith('<')

  const parseCssColorToHex = (value: string): string | undefined => {
    const v = (value || '').trim()
    if (!v) return undefined
    if (v.startsWith('#')) {
      const hex = v.slice(1).trim()
      if (/^[0-9a-fA-F]{6}$/.test(hex)) return hex.toUpperCase()
      if (/^[0-9a-fA-F]{3}$/.test(hex)) {
        return hex.split('').map(c => (c + c)).join('').toUpperCase()
      }
      return undefined
    }
    const rgb = v.match(/rgb\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)/i)
    if (rgb) {
      const r = Math.max(0, Math.min(255, Number(rgb[1] || 0)))
      const g = Math.max(0, Math.min(255, Number(rgb[2] || 0)))
      const b = Math.max(0, Math.min(255, Number(rgb[3] || 0)))
      const toHex = (n: number) => n.toString(16).padStart(2, '0').toUpperCase()
      return `${toHex(r)}${toHex(g)}${toHex(b)}`
    }
    return undefined
  }

  const parseCssFontSizePt = (value: string): number | undefined => {
    const v = (value || '').trim().toLowerCase()
    if (!v) return undefined
    const pt = v.match(/^(\d+(?:\.\d+)?)\s*pt$/)
    if (pt) return Number(pt[1])
    const px = v.match(/^(\d+(?:\.\d+)?)\s*px$/)
    if (px) return Number(px[1]) / 1.333
    return undefined
  }

  // 字体槽位分离：CJK 字体放 eastAsia，Latin 字体放 ascii/hAnsi
  const CJK_FONT_SET = new Set([
    '宋体', '黑体', '仿宋', '楷体', '微软雅黑', '华文中宋', '华文仿宋',
    '华文楷体', '华文宋体', '华文细黑', '方正小标宋简体', '方正仿宋简体',
    'SimSun', 'SimHei', 'FangSong', 'KaiTi', 'Microsoft YaHei',
    'STZhongsong', 'STFangsong', 'STKaiti', 'STSong', 'STXihei',
  ])
  const LATIN_FONT_SET = new Set([
    'Times New Roman', 'Arial', 'Calibri', 'Cambria', 'Georgia',
    'Verdana', 'Tahoma', 'Trebuchet MS', 'Garamond', 'Palatino Linotype',
    'Book Antiqua', 'Century', 'Consolas', 'Courier New', 'Segoe UI',
  ])
  const resolveFontSlots = (name: string): { ascii: string; hAnsi: string; eastAsia: string } => {
    const n = name.trim()
    const isCjk = CJK_FONT_SET.has(n) || /[\u4e00-\u9fff]/.test(n) || /^(ST|MS |Noto Sans (SC|TC|JP|KR)|Source Han)/i.test(n)
    if (isCjk) {
      return { ascii: 'Times New Roman', hAnsi: 'Times New Roman', eastAsia: n }
    }
    const latin = LATIN_FONT_SET.has(n) ? n : 'Times New Roman'
    return { ascii: latin, hAnsi: latin, eastAsia: '仿宋' }
  }

  const parseStyle = (style: string) => {
    const s = style || ''
    const get = (name: string) => {
      const m = s.match(new RegExp(`${name}\\s*:\\s*([^;]+)`, 'i'))
      return m?.[1]?.trim()
    }
    return {
      textAlign: get('text-align'),
      fontSize: get('font-size'),
      fontFamily: get('font-family'),
      color: get('color'),
      backgroundColor: get('background-color'),
    }
  }

  const toAlignment = (v?: string) => {
    const a = (v || '').toLowerCase()
    if (a === 'center') return AlignmentType.CENTER
    if (a === 'right') return AlignmentType.RIGHT
    if (a === 'justify') return AlignmentType.JUSTIFIED
    return AlignmentType.LEFT
  }

  const htmlToDocxChildren = (html: string, profile?: DocxTypographyProfile): (Paragraph | Table)[] => {
    const parser = new DOMParser()
    const doc = parser.parseFromString(html, 'text/html')

    const walkInline = (
      node: Node,
      inherited: {
        bold?: boolean
        italics?: boolean
        underline?: boolean
        color?: string
        font?: { ascii?: string; hAnsi?: string; eastAsia?: string }
        size?: number
      }
    ): TextRun[] => {
      // 忽略“被删旧内容”
      if (node.nodeType === Node.ELEMENT_NODE) {
        const el = node as HTMLElement
        const classList = Array.from(el.classList || [])
        if (classList.includes('diff-old')) return []
        if (el.getAttribute('data-diff-role') === 'old') return []
      }

      if (node.nodeType === Node.TEXT_NODE) {
        const text = (node.nodeValue || '').replace(/\u00A0/g, ' ')
        if (!text) return []
        return [
          new TextRun({
            text,
            bold: inherited.bold,
            italics: inherited.italics,
            underline: inherited.underline ? { type: UnderlineType.SINGLE } : undefined,
            color: inherited.color,
            font: inherited.font,
            size: inherited.size,
          }),
        ]
      }

      if (node.nodeType !== Node.ELEMENT_NODE) return []
      const el = node as HTMLElement
      const tag = el.tagName.toLowerCase()

      // === Word 原生修订标记（导出时转为 w:ins/w:del） ===
      if (tag === 'span' && (el.classList.contains('docx-track') || !!el.getAttribute('data-track-type'))) {
        const tRaw = (el.getAttribute('data-track-type') || 'insert').toLowerCase()
        const t = tRaw === 'delete' ? 'del' : 'ins'
        const id = stripQuotes(el.getAttribute('data-track-id') || '') || '0'
        const a = b64uEncodeUtf8(el.getAttribute('data-track-author') || '')
        const d = b64uEncodeUtf8(el.getAttribute('data-track-date') || '')
        const start = `[[[WC_TC_START|t=${t}|id=${id}|a=${a}|d=${d}]]]`
        const end = `[[[WC_TC_END]]]`
        const inner: TextRun[] = []
        el.childNodes.forEach((c) => inner.push(...walkInline(c, inherited)))
        return [new TextRun({ text: start }), ...inner, new TextRun({ text: end })]
      }

      // === Word 批注范围（导出时转为 commentRangeStart/End + comments.xml） ===
      if (tag === 'span' && (el.classList.contains('docx-comment') || !!el.getAttribute('data-comment-ids') || !!el.getAttribute('data-comment-id'))) {
        const raw = el.getAttribute('data-comment-ids') || el.getAttribute('data-comment-id') || ''
        const ids = raw.split(',').map(s => s.trim()).filter(Boolean)
        if (ids.length === 0) {
          const inner: TextRun[] = []
          el.childNodes.forEach((c) => inner.push(...walkInline(c, inherited)))
          return inner
        }
        const runs: TextRun[] = []
        for (const cid of ids) runs.push(new TextRun({ text: `[[[WC_CM_START|id=${stripQuotes(cid)}]]]` }))
        el.childNodes.forEach((c) => runs.push(...walkInline(c, inherited)))
        for (const cid of [...ids].reverse()) runs.push(new TextRun({ text: `[[[WC_CM_END|id=${stripQuotes(cid)}]]]` }))
        return runs
      }

      // diff-new：直接解析其子内容（相当于接受）
      const classList = Array.from(el.classList || [])
      const isDiffNew = classList.includes('diff-new') || el.getAttribute('data-diff-role') === 'new'

      const next = { ...inherited }

      // 基础标签
      if (tag === 'strong' || tag === 'b') next.bold = true
      if (tag === 'em' || tag === 'i') next.italics = true
      if (tag === 'u') next.underline = true

      // span style
      const style = el.getAttribute('style') || ''
      if (style) {
        const parsed = parseStyle(style)
        const color = parseCssColorToHex(parsed.color || '')
        if (color) next.color = color
        const fontFamily = parsed.fontFamily
          ? parsed.fontFamily.split(',')[0].replace(/['"]/g, '').trim()
          : ''
        if (fontFamily) next.font = resolveFontSlots(fontFamily)
        const pt = parseCssFontSizePt(parsed.fontSize || '')
        if (pt) next.size = Math.round(pt * 2)
      }

      if (tag === 'br') {
        return [new TextRun({ text: '', break: 1 })]
      }

      // 对 diff-new span，本质和普通 span 一样：解析 children
      const childRuns: TextRun[] = []
      el.childNodes.forEach((c) => childRuns.push(...walkInline(c, next)))
      return childRuns
    }

    const children: (Paragraph | Table)[] = []

    const parseBlockStyle = (style: string) => {
      const s = style || ''
      const get = (name: string) => {
        const m = s.match(new RegExp(`${name}\\s*:\\s*([^;]+)`, 'i'))
        return m?.[1]?.trim()
      }
      return {
        textAlign: get('text-align'),
        textIndent: get('text-indent'),
        lineHeight: get('line-height'),
        marginTop: get('margin-top'),
        marginBottom: get('margin-bottom'),
        marginLeft: get('margin-left'),
      }
    }

    const toSpacing = (style: ReturnType<typeof parseBlockStyle>) => {
      const basePt = (profile?.normal?.fontSizeHalfPoints ? profile.normal.fontSizeHalfPoints / 2 : 10.5)
      const before = cssLengthToTwips(style.marginTop, basePt)
      const after = cssLengthToTwips(style.marginBottom, basePt)

      // line-height: number(倍数) | px/pt
      let lineTwips: number | undefined
      const lh = (style.lineHeight || '').trim().toLowerCase()
      if (lh) {
        const mult = lh.match(/^(\d+(?:\.\d+)?)$/)
        if (mult) {
          const m = Number(mult[1])
          const baseLine = profile?.normal?.spacing?.lineTwips || Math.round(basePt * 20 * 1.15)
          lineTwips = Math.round(baseLine * m)
        } else {
          lineTwips = cssLengthToTwips(lh, basePt)
        }
      }

      return {
        before: before ?? profile?.normal?.spacing?.beforeTwips,
        after: after ?? profile?.normal?.spacing?.afterTwips,
        line: lineTwips ?? profile?.normal?.spacing?.lineTwips,
      }
    }

    const toIndent = (style: ReturnType<typeof parseBlockStyle>) => {
      const basePt = (profile?.normal?.fontSizeHalfPoints ? profile.normal.fontSizeHalfPoints / 2 : 10.5)
      const firstLine = cssLengthToTwips(style.textIndent, basePt)
      const left = cssLengthToTwips(style.marginLeft, basePt)
      return {
        firstLine: firstLine ?? profile?.normal?.indent?.firstLineTwips,
        left: left ?? profile?.normal?.indent?.leftTwips,
      }
    }

    const processBlock = (el: HTMLElement) => {
      const tag = el.tagName.toLowerCase()

      // 忽略 old 块（导出默认“接受”）
      if (el.getAttribute('data-diff-role') === 'old') return
      if (tag === 'span' && el.classList.contains('diff-old')) return

      const styleStr = el.getAttribute('style') || ''
      const { textAlign } = parseStyle(styleStr)
      const block = parseBlockStyle(styleStr)

      if (tag === 'h1' || tag === 'h2' || tag === 'h3') {
        const level =
          tag === 'h1' ? HeadingLevel.HEADING_1 : tag === 'h2' ? HeadingLevel.HEADING_2 : HeadingLevel.HEADING_3
        const headingProfile = tag === 'h1' ? profile?.heading1 : tag === 'h2' ? profile?.heading2 : profile?.heading3
        const headingAscii = headingProfile?.fontAscii || headingProfile?.fontHAnsi || profile?.normal?.fontAscii || profile?.normal?.fontHAnsi
        const headingEast = headingProfile?.fontEastAsia || profile?.normal?.fontEastAsia
        const headingFont = headingEast || headingAscii || '黑体'
        const headingRunFont = { ascii: headingAscii || headingFont, hAnsi: headingProfile?.fontHAnsi || headingAscii || headingFont, eastAsia: headingEast || headingFont }
        children.push(new Paragraph({
          heading: level,
          alignment: toAlignment(textAlign),
          children: walkInline(el, {
            font: headingRunFont,
            size: headingProfile?.fontSizeHalfPoints || profile?.normal?.fontSizeHalfPoints,
          }),
        }))
        return
      }

      if (tag === 'p') {
        const spacing = toSpacing(block)
        const indent = toIndent(block)
        const normalAscii = profile?.normal?.fontAscii || profile?.normal?.fontHAnsi
        const normalEast = profile?.normal?.fontEastAsia
        const normalFont = normalEast || normalAscii || '等线'
        const normalRunFont = { ascii: normalAscii || normalFont, hAnsi: profile?.normal?.fontHAnsi || normalAscii || normalFont, eastAsia: normalEast || normalFont }
        children.push(new Paragraph({
          alignment: toAlignment(textAlign || profile?.normal?.alignment),
          indent: {
            firstLine: indent.firstLine,
            left: indent.left,
          },
          spacing: {
            before: spacing.before,
            after: spacing.after,
            line: spacing.line,
          },
          children: walkInline(el, {
            font: normalRunFont,
            size: profile?.normal?.fontSizeHalfPoints,
          }),
        }))
        return
      }

      if (tag === 'ul') {
        const items = Array.from(el.querySelectorAll(':scope > li'))
        items.forEach((li) => {
          children.push(new Paragraph({
            bullet: { level: 0 },
            children: walkInline(li as any, {}),
          }))
        })
        return
      }

      if (tag === 'ol') {
        // 简化：先用纯文本编号，后续可升级 docx numbering
        const items = Array.from(el.querySelectorAll(':scope > li'))
        items.forEach((li, idx) => {
          const runs = walkInline(li as any, {})
          children.push(new Paragraph({
            children: [new TextRun({ text: `${idx + 1}. ` }), ...runs],
          }))
        })
        return
      }

      if (tag === 'table') {
        const rows = Array.from(el.querySelectorAll('tr'))
        const tableRows: TableRow[] = rows.map((tr) => {
          const cells = Array.from(tr.querySelectorAll('th,td'))
          return new TableRow({
            children: cells.map((cell) => {
              const isHeader = cell.tagName.toLowerCase() === 'th'
              const cellRuns = walkInline(cell as any, {})
              const cellParagraph = new Paragraph({
                children: isHeader ? [new TextRun({ text: (cell.textContent || '').trim(), bold: true })] : cellRuns,
              })
              return new TableCell({ children: [cellParagraph] })
            }),
          })
        })
        children.push(new Table({ rows: tableRows, width: { size: 100, type: WidthType.PERCENTAGE } }))
        children.push(new Paragraph({ text: '' }))
        return
      }

      // fallback：把任意块转成段落
      const text = (el.textContent || '').trim()
      if (text) {
        children.push(new Paragraph({ children: [new TextRun({ text })] }))
      }
    }

    Array.from(doc.body.children).forEach((child) => processBlock(child as HTMLElement))

    if (children.length === 0) {
      children.push(new Paragraph({ text: '' }))
    }
    return children
  }

  const paragraphsOrTables = isHtml ? htmlToDocxChildren(content, typographyProfile || undefined) : markdownToDocxParagraphs(content, typographyProfile || undefined)
  
  const margin = typographyProfile?.page?.margin
  const normalAscii = typographyProfile?.normal?.fontAscii || typographyProfile?.normal?.fontHAnsi
  const normalEastAsia = typographyProfile?.normal?.fontEastAsia
  // 默认字体：仿宋（中文正文）、Times New Roman（西文），更专业的公文/报告风格
  const normalFallbackEA = normalEastAsia || '仿宋'
  const normalFallbackLatin = normalAscii || 'Times New Roman'
  const normalRunFont = { ascii: normalFallbackLatin, hAnsi: typographyProfile?.normal?.fontHAnsi || normalFallbackLatin, eastAsia: normalFallbackEA }
  const normalRunSize = typographyProfile?.normal?.fontSizeHalfPoints ?? 24 // 默认小四(12pt=24半点)
  const normalParaIndent = typographyProfile?.normal?.indent
  const normalParaSpacing = typographyProfile?.normal?.spacing

  const h1 = typographyProfile?.heading1
  const h2 = typographyProfile?.heading2
  const h3 = typographyProfile?.heading3
  // 标题默认字体：黑体（中文）、Arial（西文）
  const headingFallbackEA = '黑体'
  const headingFallbackLatin = 'Arial'
  const headingRunFont = (h?: DocxTypographyProfile['normal']) => {
    const a = h?.fontAscii || h?.fontHAnsi || normalAscii || headingFallbackLatin
    const e = h?.fontEastAsia || normalEastAsia || headingFallbackEA
    return { ascii: a, hAnsi: h?.fontHAnsi || a, eastAsia: e }
  }

  const doc = new Document({
    creator: 'Word-Cursor',
    title: title,
    description: 'Created by Word-Cursor',
    styles: {
      default: {
        document: {
          run: {
            font: normalRunFont,
            size: normalRunSize,
          },
          paragraph: {
            indent: normalParaIndent
              ? { firstLine: normalParaIndent.firstLineTwips, left: normalParaIndent.leftTwips, right: normalParaIndent.rightTwips }
              : { firstLine: 480 }, // 默认首行缩进2字符（480twips ≈ 24pt ≈ 2em at 12pt）
            spacing: normalParaSpacing
              ? { before: normalParaSpacing.beforeTwips, after: normalParaSpacing.afterTwips, line: normalParaSpacing.lineTwips }
              : { before: 0, after: 0, line: 360 }, // 默认行距 18pt（固定值）
          },
        },
        heading1: {
          run: {
            font: headingRunFont(h1),
            size: h1?.fontSizeHalfPoints ?? 32, // 默认小二(16pt=32半点)
            bold: true,
            color: '000000', // 黑色，避免蓝色主题色
          },
          paragraph: {
            spacing: h1?.spacing
              ? { before: h1.spacing.beforeTwips, after: h1.spacing.afterTwips, line: h1.spacing.lineTwips }
              : { before: 240, after: 120, line: 360 }, // 段前12pt 段后6pt 行距18pt
          },
        },
        heading2: {
          run: {
            font: headingRunFont(h2),
            size: h2?.fontSizeHalfPoints ?? 28, // 默认三号(14pt=28半点)
            bold: true,
            color: '000000',
          },
          paragraph: {
            spacing: h2?.spacing
              ? { before: h2.spacing.beforeTwips, after: h2.spacing.afterTwips, line: h2.spacing.lineTwips }
              : { before: 200, after: 100, line: 360 },
          },
        },
        heading3: {
          run: {
            font: headingRunFont(h3),
            size: h3?.fontSizeHalfPoints ?? 26, // 默认小三(13pt=26半点)
            bold: true,
            color: '000000',
          },
          paragraph: {
            spacing: h3?.spacing
              ? { before: h3.spacing.beforeTwips, after: h3.spacing.afterTwips, line: h3.spacing.lineTwips }
              : { before: 160, after: 80, line: 340 },
          },
        },
      },
    },
    sections: [{
      properties: {
        page: {
          margin: {
            top: margin?.topTwips ?? 1440,
            right: margin?.rightTwips ?? 1440,
            bottom: margin?.bottomTwips ?? 1440,
            left: margin?.leftTwips ?? 1440,
          },
        },
      },
      children: paragraphsOrTables,
    }],
  })
  
  const blob = await Packer.toBlob(doc)
  return blob
}

// 格式化元素类型
interface FormattedElement {
  type: 'heading' | 'paragraph' | 'table'
  content?: string
  level?: number
  bold?: boolean
  fontSize?: number
  fontFamily?: string
  alignment?: 'left' | 'center' | 'right' | 'justify'
  rows?: number
  cols?: number
  data?: string[][]
}

// 创建带格式的 docx 文档
async function createFormattedDocxBlob(elements: FormattedElement[], title: string, typographyProfile?: DocxTypographyProfile | null): Promise<Blob> {
  const children: (Paragraph | Table)[] = []
  const normalAscii = typographyProfile?.normal?.fontAscii || typographyProfile?.normal?.fontHAnsi
  const normalEastAsia = typographyProfile?.normal?.fontEastAsia
  const defaultFont = normalEastAsia || normalAscii || '宋体'
  const defaultRunFont = { ascii: normalAscii || defaultFont, hAnsi: typographyProfile?.normal?.fontHAnsi || normalAscii || defaultFont, eastAsia: normalEastAsia || defaultFont }
  const defaultSize = typographyProfile?.normal?.fontSizeHalfPoints ?? 24 // default 12pt
  const normalIndent = typographyProfile?.normal?.indent
  const defaultIndent = normalIndent?.firstLineTwips
  const defaultIndentLeft = normalIndent?.leftTwips
  const defaultIndentRight = normalIndent?.rightTwips
  const defaultSpacing = typographyProfile?.normal?.spacing
  
  for (const elem of elements) {
    if (elem.type === 'heading' && elem.content) {
      // 标题
      const level = elem.level || 1
      const headingLevelMap: Record<number, typeof HeadingLevel[keyof typeof HeadingLevel]> = {
        1: HeadingLevel.HEADING_1,
        2: HeadingLevel.HEADING_2,
        3: HeadingLevel.HEADING_3,
        4: HeadingLevel.HEADING_4,
        5: HeadingLevel.HEADING_5,
        6: HeadingLevel.HEADING_6,
      }
      
      const alignmentMap: Record<string, typeof AlignmentType[keyof typeof AlignmentType]> = {
        'left': AlignmentType.LEFT,
        'center': AlignmentType.CENTER,
        'right': AlignmentType.RIGHT,
        'justify': AlignmentType.JUSTIFIED,
      }

      const headingStyle =
        (level === 1 ? typographyProfile?.heading1 : undefined) ||
        (level === 2 ? typographyProfile?.heading2 : undefined) ||
        (level === 3 ? typographyProfile?.heading3 : undefined) ||
        typographyProfile?.normal
      const headingAscii = headingStyle?.fontAscii || headingStyle?.fontHAnsi || normalAscii || defaultFont
      const headingEastAsia = headingStyle?.fontEastAsia || normalEastAsia || defaultFont
      const headingFont = { ascii: headingAscii, hAnsi: headingStyle?.fontHAnsi || headingAscii, eastAsia: headingEastAsia }
      const headingSize = headingStyle?.fontSizeHalfPoints ?? defaultSize
      const headingSpacing = headingStyle?.spacing
      const headingIndent = headingStyle?.indent

      children.push(new Paragraph({
        heading: headingLevelMap[level] || HeadingLevel.HEADING_1,
        children: [
          new TextRun({
            text: elem.content,
            size: headingSize,
            font: headingFont,
          }),
        ],
        alignment: elem.alignment ? alignmentMap[elem.alignment] : (headingStyle?.alignment ? alignmentMap[headingStyle.alignment] : AlignmentType.LEFT),
        // 标题通常不首行缩进；但如果模板设置了左右缩进，则继承
        indent: (headingIndent?.leftTwips || headingIndent?.rightTwips)
          ? { left: headingIndent.leftTwips, right: headingIndent.rightTwips, firstLine: 0 }
          : undefined,
        spacing: headingSpacing ? { before: headingSpacing.beforeTwips, after: headingSpacing.afterTwips, line: headingSpacing.lineTwips } : undefined,
      }))
    } else if (elem.type === 'paragraph' && elem.content) {
      // 段落
      const alignmentMap: Record<string, typeof AlignmentType[keyof typeof AlignmentType]> = {
        'left': AlignmentType.LEFT,
        'center': AlignmentType.CENTER,
        'right': AlignmentType.RIGHT,
        'justify': AlignmentType.JUSTIFIED,
      }
      
      children.push(new Paragraph({
        children: [
          new TextRun({
            text: elem.content,
            bold: elem.bold || false,
            size: elem.fontSize ? elem.fontSize * 2 : defaultSize, // docx 使用半点
            font: elem.fontFamily ? { ascii: elem.fontFamily, hAnsi: elem.fontFamily, eastAsia: elem.fontFamily } : defaultRunFont,
          }),
        ],
        alignment: elem.alignment
          ? alignmentMap[elem.alignment]
          : (typographyProfile?.normal?.alignment ? alignmentMap[typographyProfile.normal.alignment] : AlignmentType.LEFT),
        indent: (defaultIndent || defaultIndentLeft || defaultIndentRight)
          ? { firstLine: defaultIndent, left: defaultIndentLeft, right: defaultIndentRight }
          : undefined,
        spacing: defaultSpacing ? { before: defaultSpacing.beforeTwips, after: defaultSpacing.afterTwips, line: defaultSpacing.lineTwips } : undefined,
      }))
    } else if (elem.type === 'table' && elem.rows && elem.cols) {
      // 表格
      const tableRows: TableRow[] = []
      const data = elem.data || []
      
      for (let r = 0; r < elem.rows; r++) {
        const cells: TableCell[] = []
        for (let c = 0; c < elem.cols; c++) {
          const cellText = data[r]?.[c] || ''
          cells.push(new TableCell({
            children: [new Paragraph({
              children: [new TextRun({
                text: cellText,
                bold: r === 0, // 第一行加粗（表头）
                size: defaultSize,
                font: defaultRunFont,
              })],
            })],
            width: { size: 100 / elem.cols, type: WidthType.PERCENTAGE },
          }))
        }
        tableRows.push(new TableRow({ children: cells }))
      }
      
      children.push(new Table({
        rows: tableRows,
        width: { size: 100, type: WidthType.PERCENTAGE },
      }))
      
      // 表格后添加空行
      children.push(new Paragraph({ text: '' }))
    }
  }
  
  // 如果没有元素，添加一个空段落
  if (children.length === 0) {
    children.push(new Paragraph({ text: '' }))
  }
  
  const margin = typographyProfile?.page?.margin
  const h1 = typographyProfile?.heading1
  const h2 = typographyProfile?.heading2
  const h3 = typographyProfile?.heading3
  const headingRunFont = (h?: DocxTypographyProfile['normal']) => {
    const a = h?.fontAscii || h?.fontHAnsi || normalAscii
    const e = h?.fontEastAsia || normalEastAsia
    const fb = e || a || defaultFont
    return { ascii: a || fb, hAnsi: h?.fontHAnsi || a || fb, eastAsia: e || fb }
  }

  const doc = new Document({
    creator: 'Word-Cursor',
    title,
    description: 'Created by Word-Cursor',
    // 关键：设置 Document 级默认样式，保证 Word 打开时字体/行距/缩进继承稳定
    styles: {
      default: {
        document: {
          run: {
            font: defaultRunFont,
            size: typographyProfile?.normal?.fontSizeHalfPoints ?? defaultSize,
          },
          paragraph: {
            indent: typographyProfile?.normal?.indent
              ? {
                  firstLine: typographyProfile.normal.indent.firstLineTwips,
                  left: typographyProfile.normal.indent.leftTwips,
                  right: typographyProfile.normal.indent.rightTwips,
                }
              : undefined,
            spacing: typographyProfile?.normal?.spacing
              ? {
                  before: typographyProfile.normal.spacing.beforeTwips,
                  after: typographyProfile.normal.spacing.afterTwips,
                  line: typographyProfile.normal.spacing.lineTwips,
                }
              : undefined,
          },
        },
        heading1: {
          run: { font: headingRunFont(h1), size: h1?.fontSizeHalfPoints },
          paragraph: { spacing: h1?.spacing ? { before: h1.spacing.beforeTwips, after: h1.spacing.afterTwips, line: h1.spacing.lineTwips } : undefined },
        },
        heading2: {
          run: { font: headingRunFont(h2), size: h2?.fontSizeHalfPoints },
          paragraph: { spacing: h2?.spacing ? { before: h2.spacing.beforeTwips, after: h2.spacing.afterTwips, line: h2.spacing.lineTwips } : undefined },
        },
        heading3: {
          run: { font: headingRunFont(h3), size: h3?.fontSizeHalfPoints },
          paragraph: { spacing: h3?.spacing ? { before: h3.spacing.beforeTwips, after: h3.spacing.afterTwips, line: h3.spacing.lineTwips } : undefined },
        },
      },
    },
    sections: [{
      properties: {
        page: {
          margin: {
            top: margin?.topTwips ?? 1440,
            right: margin?.rightTwips ?? 1440,
            bottom: margin?.bottomTwips ?? 1440,
            left: margin?.leftTwips ?? 1440,
          },
        },
      },
      children,
    }],
  })
  
  return await Packer.toBlob(doc)
}

export function DocumentProvider({ children }: { children: ReactNode }) {
  const { comments: commentsForExport } = useComments()
  const [document, setDocument] = useState<DocumentContent>(defaultDocument)
  const [hasUnsavedChanges, setHasUnsavedChanges] = useState(false)
  const [workspacePath, setWorkspacePath] = useState<string | null>(null)
  const [files, setFiles] = useState<FileItem[]>([])
  const [docEntryAnimationKey, setDocEntryAnimationKey] = useState(0)
  
  // 使用 ref 跟踪最新的文档内容，解决连续替换时闭包问题
  const documentContentRef = useRef(document.content)
  
  // 同步更新 ref
  useEffect(() => {
    documentContentRef.current = document.content
  }, [document.content])
  const [currentFile, setCurrentFileState] = useState<FileItem | null>(null)
  const [docxData, setDocxData] = useState<string | null>(null)
  const [excelData, setExcelData] = useState<ExcelOpenResponse | null>(null)
  const [pptData, setPptData] = useState<{ pptxBase64: string } | null>(null)
  const [lastReplacement, setLastReplacement] = useState<ReplacementRecord | null>(null)
  const [pendingReplacements, setPendingReplacements] = useState<PendingReplacements>({ items: [], total: 0 })
  const [extraPendingChanges, setExtraPendingChanges] = useState<PendingChange[]>([])
  const [scrollTarget, setScrollTarget] = useState<string | null>(null)
  const [pageSetup, setPageSetupState] = useState<PageSetup>(defaultPageSetup)
  const [headerFooterSetup, setHeaderFooterSetupState] = useState<HeaderFooterSetup>(defaultHeaderFooterSetup)
  const [customStyles, setCustomStyles] = useState<Record<string, CustomStyle>>(defaultCustomStyles)
  const [typographyProfile, setTypographyProfile] = useState<DocxTypographyProfile | null>(null)

  // 将 docx 提取的默认字体同步到 CSS 变量（让预览/编辑默认字体尽量贴近原文）
  useEffect(() => {
    if (typeof window === 'undefined') return
    const root = window.document?.documentElement
    if (!root) return

    const normal = typographyProfile?.normal
    const eastAsia =
      normal?.fontEastAsia ||
      normal?.fontAscii ||
      normal?.fontHAnsi ||
      ''

    if (eastAsia) {
      root.style.setProperty('--word-font-family-cn', toChineseDefaultFallbackStack(eastAsia))
    } else {
      // 清空覆盖，让 index.css 的默认值接管
      root.style.removeProperty('--word-font-family-cn')
    }

    const fontSizeHalfPoints = normal?.fontSizeHalfPoints
    if (typeof fontSizeHalfPoints === 'number' && Number.isFinite(fontSizeHalfPoints) && fontSizeHalfPoints > 0) {
      const px = (fontSizeHalfPoints * 2) / 3
      root.style.setProperty('--word-font-size', `${px.toFixed(2).replace(/\.?0+$/, '')}px`)
    } else {
      root.style.removeProperty('--word-font-size')
    }

    const spacing = normal?.spacing
    const lineTwips = spacing?.lineTwips
    if (typeof lineTwips === 'number' && Number.isFinite(lineTwips) && lineTwips > 0) {
      if (spacing?.lineRule === 'exact' || spacing?.lineRule === 'atLeast') {
        const pt = lineTwips / 20
        root.style.setProperty('--word-line-height', `${pt.toFixed(2).replace(/\.?0+$/, '')}pt`)
      } else {
        const multiplier = lineTwips / 240
        root.style.setProperty('--word-line-height', multiplier.toFixed(2).replace(/\.?0+$/, ''))
      }
    } else {
      root.style.removeProperty('--word-line-height')
    }
  }, [typographyProfile])

  const triggerDocEntryAnimation = useCallback(() => {
    setDocEntryAnimationKey(Date.now())
  }, [])

  // 刷新 Excel 数据（重新读取文件）
  const refreshExcelData = useCallback(async () => {
    if (!currentFile || !isElectron || !window.electronAPI) return false
    
    const ext = (currentFile.name.split('.').pop() || '').toLowerCase()
    if (ext !== 'xlsx' && ext !== 'xls') return false
    
    try {
      // 先关闭缓存
      await window.electronAPI.excelClose?.(currentFile.path)
      // 重新读取
      const result = await window.electronAPI.excelOpen(currentFile.path)
      if (result.success && result.sheets) {
        setExcelData(result)
        return true
      }
    } catch (error) {
      console.error('刷新 Excel 数据失败:', error)
    }
    return false
  }, [currentFile, isElectron])

  const updateDocument = useCallback((updates: Partial<DocumentContent>) => {
    setDocument(prev => ({
      ...prev,
      ...updates,
      lastModified: new Date(),
    }))
    setHasUnsavedChanges(true)
  }, [])

  const updateContent = useCallback((content: string) => {
    setDocument(prev => ({
      ...prev,
      content,
      lastModified: new Date(),
    }))
    setHasUnsavedChanges(true)
  }, [])

  const updateStyles = useCallback((styles: Partial<DocumentStyles>) => {
    setDocument(prev => ({
      ...prev,
      styles: { ...prev.styles, ...styles },
      lastModified: new Date(),
    }))
  }, [])

  const addFile = useCallback((file: FileItem) => {
    setFiles(prev => [...prev, file])
  }, [])

  const setCurrentFile = useCallback((file: FileItem | null) => {
    setCurrentFileState(file)
  }, [])

  const createNewDocument = useCallback(async (title: string, content: string, elements?: FormattedElement[], styleRefPath?: string) => {
    console.log('createNewDocument 被调用:', { title, contentLength: content.length, elementsCount: elements?.length })
    setExcelData(null)
    
    // 清理文件名中的非法字符，并移除已有的 .docx 后缀（避免双重后缀）
    let safeTitle = title.replace(/[<>:"/\\|?*]/g, '_').slice(0, 50)
    if (safeTitle.toLowerCase().endsWith('.docx')) {
      safeTitle = safeTitle.slice(0, -5)
    }
    
    // 如果在 Electron 环境且有工作区路径，创建真实文件
    if (isElectron && window.electronAPI && workspacePath) {
      try {
        const fileName = `${safeTitle}.docx`
        const filePath = `${workspacePath}\\${fileName}`
        console.log('准备创建文件:', filePath)
        
        let success = false
        let nextProfile: DocxTypographyProfile | null = null

        // 如果指定了“样式参考 docx”，抽取 TypographyProfile（用于导出时继承字体/缩进/行距/页边距）
        if (styleRefPath) {
          try {
            const ab = await fetchArrayBufferFromLocalFile(styleRefPath)
            const { profile } = await extractTypographyProfileFromArrayBuffer(ab)
            nextProfile = profile
            setTypographyProfile(profile)
          } catch (e) {
            console.warn('样式参考解析失败，继续使用默认样式:', e)
            setTypographyProfile(null)
          }
        } else {
          setTypographyProfile(null)
        }
        
        // 如果有 elements，优先尝试使用 ONLYOFFICE Document Builder API（但在指定样式参考时，优先走 docx 库以应用 profile）
        if (elements && elements.length > 0) {
          console.log('尝试使用 ONLYOFFICE Document Builder API 创建文档，元素:', elements)
          
          if (!styleRefPath) {
            try {
              const builderResult = await window.electronAPI.createFormattedDocument({
                filePath,
                elements,
                title: safeTitle
              })
              
              if (builderResult.success) {
                console.log('ONLYOFFICE Document Builder 创建成功', builderResult.fallback ? '(使用回退方案)' : '')
                success = true
              } else {
                console.log('ONLYOFFICE Document Builder 失败，回退到 docx 库:', builderResult.error)
              }
            } catch (e) {
              console.log('ONLYOFFICE Document Builder 调用失败，回退到 docx 库:', e)
            }
          } else {
            console.log('指定了样式参考，跳过 Document Builder，直接使用 docx 库以应用 TypographyProfile')
          }
          
          // 如果 Document Builder 失败，回退到 docx 库
          if (!success) {
            console.log('使用 docx 库创建格式化文档')
            const blob = await createFormattedDocxBlob(elements, safeTitle, nextProfile)
            const arrayBuffer = await blob.arrayBuffer()
            // 使用分块方式将 ArrayBuffer 转换为 base64，避免大文件导致的栈溢出
            const base64 = arrayBufferToBase64(arrayBuffer)
            const result = await window.electronAPI.writeBinaryFile(filePath, base64)
            success = result.success
          }
        } else {
          // 纯文本文档，使用 docx 库
          console.log('使用纯文本方式创建文档')
          const blob = await createDocxBlob(content, safeTitle, nextProfile)
          const arrayBuffer = await blob.arrayBuffer()
          // 使用分块方式将 ArrayBuffer 转换为 base64，避免大文件导致的栈溢出
          const base64 = arrayBufferToBase64(arrayBuffer)
          const result = await window.electronAPI.writeBinaryFile(filePath, base64)
          success = result.success
        }
        
        if (success) {
          console.log('文件已创建:', filePath)
          
          // 刷新文件列表以显示新文件
          const folderResult = await window.electronAPI.readFolder(workspacePath)
          if (folderResult.success && folderResult.data) {
            const convertFiles = (items: any[]): FileItem[] => {
              return items.map(item => ({
                name: item.name,
                path: item.path,
                type: item.type,
                children: item.children ? convertFiles(item.children) : undefined,
              }))
            }
            setFiles(convertFiles(folderResult.data))
          }
          
          // 创建文件项并设置为当前文件
          const newFile: FileItem = {
            name: fileName,
            path: filePath,
            type: 'file',
          }

          // 同步更新 ref（关键！保证下一次工具调用拿到最新内容）
          documentContentRef.current = content
          dslCacheRef.current = null

          setCurrentFileState(newFile)
          setDocument({
            title: safeTitle,
            content,
            styles: defaultStyles,
            lastModified: new Date(),
          })
          triggerDocEntryAnimation()
          setDocxData(null)
          setHasUnsavedChanges(false) // 已保存
        } else {
          console.error('创建文件失败')
        }
      } catch (error) {
        console.error('创建文档失败:', error)
      }
    } else {
      // Web 模式或没有工作区，只在内存中创建
      const newFile: FileItem = {
        name: `${safeTitle}.docx`,
        path: `/${safeTitle}.docx`,
        type: 'file',
        content,
      }

      // 同步更新 ref（关键！保证下一次工具调用拿到最新内容）
      documentContentRef.current = content
      dslCacheRef.current = null

      setFiles(prev => [...prev, newFile])
      setCurrentFileState(newFile)
      setDocument({
        title: safeTitle,
        content,
        styles: defaultStyles,
        lastModified: new Date(),
      })
      triggerDocEntryAnimation()
      setDocxData(null)
      setHasUnsavedChanges(true)
    }
  }, [workspacePath, triggerDocEntryAnimation])

  // DSL 方式创建文档
  const createDocumentFromDsl = useCallback(async (title: string, dsl: DocDsl): Promise<{ success: boolean; message: string; filePath?: string }> => {
    // 校验 DSL
    const validation = validateDocDsl(dsl)
    if (!validation.valid) {
      const errorMessages = validation.errors.map(e => `${e.path}: ${e.message}`).join('\n')
      return { success: false, message: `DSL 校验失败:\n${errorMessages}` }
    }

    // 清理文件名
    let safeTitle = title.replace(/[<>:"/\\|?*]/g, '_').slice(0, 50)
    if (safeTitle.toLowerCase().endsWith('.docx')) {
      safeTitle = safeTitle.slice(0, -5)
    }

    // 在 Electron 环境创建真实文件
    if (isElectron && window.electronAPI && workspacePath) {
      try {
        const fileName = `${safeTitle}.docx`
        const filePath = `${workspacePath}\\${fileName}`
        console.log('DSL 创建文件:', filePath)

        // 使用 DSL 渲染器生成 DOCX
        const blob = await dslToDocxBlob(dsl)
        const arrayBuffer = await blob.arrayBuffer()
        const base64 = arrayBufferToBase64(arrayBuffer)
        
        const result = await window.electronAPI.writeBinaryFile(filePath, base64)
        
        if (result.success) {
          console.log('DSL 文件已创建:', filePath)
          
          // 刷新文件列表
          const folderResult = await window.electronAPI.readFolder(workspacePath)
          if (folderResult.success && folderResult.data) {
            const convertFiles = (items: any[]): FileItem[] => {
              return items.map(item => ({
                name: item.name,
                path: item.path,
                type: item.type,
                children: item.children ? convertFiles(item.children) : undefined,
              }))
            }
            setFiles(convertFiles(folderResult.data))
          }

          // 给所有块加 diff 标记（绿色高亮），让用户可以审查后确认
          const diffId = `diff-create-${Date.now()}`
          for (const block of dsl.blocks) {
            if ('_meta' in block || block.type === 'heading' || block.type === 'paragraph') {
              (block as any)._meta = { diffRole: 'new' as const, diffId }
            }
          }

          // 生成 HTML 预览用于编辑器
          const htmlContent = dslToHtml(dsl)

          // 同步更新 ref（关键！保证下一次工具调用拿到最新内容）
          documentContentRef.current = htmlContent
          dslCacheRef.current = null

          // 创建文件项并设置为当前文件
          const newFile: FileItem = {
            name: fileName,
            path: filePath,
            type: 'file',
          }

          setCurrentFileState(newFile)
          setDocument({
            title: safeTitle,
            content: htmlContent,
            styles: defaultStyles,
            lastModified: new Date(),
          })
          triggerDocEntryAnimation()
          setDocxData(null)
          setExcelData(null)
          setHasUnsavedChanges(false)

          // 注册为待确认修改，让 accept/reject 按钮出现
          const blockCount = dsl.blocks.length
          setPendingReplacements(prev => ({
            items: [...prev.items, {
              id: diffId,
              searchText: '',
              replaceText: `[创建文档: ${fileName}]`,
              count: blockCount,
              timestamp: Date.now(),
            }],
            total: prev.total + blockCount,
          }))

          return { success: true, message: `已创建文档: ${fileName}`, filePath }
        } else {
          return { success: false, message: '文件写入失败' }
        }
      } catch (error) {
        console.error('DSL 创建文档失败:', error)
        return { success: false, message: `创建失败: ${(error as Error).message}` }
      }
    } else {
      // Web 模式：只在内存中创建
      const htmlContent = dslToHtml(dsl)

      // 同步更新 ref（关键！保证下一次工具调用拿到最新内容）
      documentContentRef.current = htmlContent
      dslCacheRef.current = null

      const newFile: FileItem = {
        name: `${safeTitle}.docx`,
        path: `/${safeTitle}.docx`,
        type: 'file',
        content: htmlContent,
      }
      setFiles(prev => [...prev, newFile])
      setCurrentFileState(newFile)
      setDocument({
        title: safeTitle,
        content: htmlContent,
        styles: defaultStyles,
        lastModified: new Date(),
      })
      triggerDocEntryAnimation()
      setDocxData(null)
      setHasUnsavedChanges(true)

      return { success: true, message: `已创建文档: ${safeTitle}.docx` }
    }
  }, [workspacePath, isElectron, triggerDocEntryAnimation])

  // 打开本地文件夹 (Electron)
  const openFolder = useCallback(async () => {
    if (!isElectron || !window.electronAPI) {
      alert('此功能需要在桌面应用中使用')
      return
    }

    const folderPath = await window.electronAPI.selectFolder()
    if (!folderPath) return

    setWorkspacePath(folderPath)
    
    const result = await window.electronAPI.readFolder(folderPath)
    if (result.success && result.data) {
      const convertFiles = (items: any[]): FileItem[] => {
        return items.map(item => ({
          name: item.name,
          path: item.path,
          type: item.type,
          children: item.children ? convertFiles(item.children) : undefined,
        }))
      }
      setFiles(convertFiles(result.data))
    }
  }, [])

  // 将 xls 转换为 xlsx
  const convertXlsToXlsx = useCallback(async (xlsPath: string) => {
    if (!window.electronAPI?.excelConvertXlsToXlsx) {
      alert('转换功能不可用')
      return
    }
    
    try {
      const result = await window.electronAPI.excelConvertXlsToXlsx(xlsPath)
      if (result.success) {
        alert(result.message || '转换成功！')
        // 刷新文件列表
        if (workspacePath) {
          const folderResult = await window.electronAPI.readFolder(workspacePath)
          if (folderResult.success && folderResult.data) {
            const convertFilesLocal = (items: any[]): FileItem[] => {
              return items.map(item => ({
                name: item.name,
                path: item.path,
                type: item.type,
                children: item.children ? convertFilesLocal(item.children) : undefined,
              }))
            }
            setFiles(convertFilesLocal(folderResult.data))
          }
        }
      } else {
        alert('转换失败：' + (result.error || '未知错误'))
      }
    } catch (error) {
      alert('转换失败：' + (error as Error).message)
    }
  }, [workspacePath])

  // 文件内容缓存 - Cursor 风格：切换文件时自动缓存，保存时才写入磁盘
  const fileContentCacheRef = useRef<Map<string, { content: string; title: string; hasChanges: boolean }>>(new Map())

  // 打开文件 (Electron)
  const openFile = useCallback(async (file: FileItem) => {
    if (file.type !== 'file') return

    // 如果当前文件有未保存的更改，先缓存到内存（不弹出确认框）
    if (currentFile && hasUnsavedChanges) {
      fileContentCacheRef.current.set(currentFile.path, {
        content: documentContentRef.current,
        title: document.title,
        hasChanges: true
      })
      console.log(`[Cache] 缓存文件修改: ${currentFile.name}`)
    }

    // 检查目标文件是否有缓存的修改
    const cached = fileContentCacheRef.current.get(file.path)

    setCurrentFileState(file)

    // 如果有缓存的修改，优先使用缓存内容（Cursor 风格）
    if (cached && cached.hasChanges) {
      console.log(`[Cache] 恢复缓存内容: ${file.name}`)
      setExcelData(null)
      setDocxData(null)
      setPptData(null)
      documentContentRef.current = cached.content
      setDocument({
        title: cached.title,
        content: cached.content,
        styles: defaultStyles,
        lastModified: new Date(),
      })
      setHasUnsavedChanges(true)
      return
    }

    if (isElectron && window.electronAPI) {
      const ext = (file.name.split('.').pop() || '').toLowerCase()

      // Excel 预览
      if (ext === 'xlsx' || ext === 'xls') {
        const result = await window.electronAPI.excelOpen(file.path)
        if (result.success && result.sheets) {
          setExcelData(result)
          setDocxData(null)
          setTypographyProfile(null)
          setDocument({
            title: file.name.replace(/\.[^.]+$/, ''),
            content: '',
            styles: defaultStyles,
            lastModified: new Date(),
          })
          setHasUnsavedChanges(false)
          
          // 对于 xls 文件，显示警告并提供转换选项
          if (result.isXls && result.warning) {
            setTimeout(async () => {
              // 检查 LibreOffice 安装状态
              let libreOfficeInfo = null
              if (window.electronAPI?.checkLibreOffice) {
                libreOfficeInfo = await window.electronAPI.checkLibreOffice()
              }
              
              let message = result.warning + '\n\n是否现在将此文件转换为 xlsx 格式？\n\n'
              
              if (libreOfficeInfo?.installed) {
                message += '✅ 已检测到 LibreOffice，将使用它进行无损转换。'
              } else {
                message += '⚠️ 未检测到 LibreOffice，将尝试以下方式：\n' +
                  '1. Microsoft Excel（如果已安装）\n' +
                  '2. 基础转换（仅保留数据）\n\n' +
                  '💡 推荐安装免费的 LibreOffice 以获得完美转换：\n' +
                  'https://www.libreoffice.org/download/'
              }
              
              const shouldConvert = window.confirm(message)
              if (shouldConvert && result.originalPath) {
                convertXlsToXlsx(result.originalPath)
              }
            }, 500)
          }
        } else {
          alert(result.error || '读取 Excel 失败')
        }
        return
      }

      // 其它文件走原有逻辑
      const result = await window.electronAPI.readFile(file.path)
      
      if (result.success && result.data) {
        if (result.type === 'pptx') {
          // .pptx 文件 - 使用纯 JS 预览
          setExcelData(null)
          setDocxData(null)
          setTypographyProfile(null)
          setPptData({ pptxBase64: result.data })
          setDocument({
            title: file.name.replace(/\.[^.]+$/, ''),
            content: '',
            styles: defaultStyles,
            lastModified: new Date(),
          })
        } else if (result.type === 'docx') {
          // .docx 文件 - 使用前端解析器（保留样式）
          setExcelData(null)
          setDocxData(result.data)
          setTypographyProfile(null)
          setPptData(null)
          setDocument({
            title: file.name.replace(/\.[^.]+$/, ''),
            content: '',
            styles: defaultStyles,
            lastModified: new Date(),
          })
          // 尽量从 docx 提取默认字体/样式（用于 UI 字体显示与导出一致性）
          // 不阻塞打开流程：失败则忽略，保留默认值
          void (async () => {
            try {
              const ab = await fetchArrayBufferFromLocalFile(file.path)
              const { profile } = await extractTypographyProfileFromArrayBuffer(ab)
              setTypographyProfile(profile)
            } catch (e) {
              console.warn('[DOCX Typography] 提取失败，使用默认字体回退:', e)
              setTypographyProfile(null)
            }
          })()
        } else if (result.type === 'doc-html') {
          // .doc 文件 - 已经转换为 HTML
          setExcelData(null)
          setDocxData(null)
          setTypographyProfile(null)
          setPptData(null)
          setDocument({
            title: file.name.replace(/\.[^.]+$/, ''),
            content: result.data,
            styles: defaultStyles,
            lastModified: new Date(),
          })
        } else {
          // 文本文件
          setExcelData(null)
          setDocxData(null)
          setTypographyProfile(null)
          setPptData(null)
          setDocument({
            title: file.name.replace(/\.[^.]+$/, ''),
            content: result.data,
            styles: defaultStyles,
            lastModified: new Date(),
          })
        }
        setHasUnsavedChanges(false)
      }
    } else if (file.content) {
      setDocxData(null)
      setPptData(null)
      setDocument({
        title: file.name.replace('.docx', ''),
        content: file.content,
        styles: defaultStyles,
        lastModified: new Date(),
      })
      setHasUnsavedChanges(false)
    }
  }, [hasUnsavedChanges, currentFile, document.title])

  // 保存当前文件
  const saveCurrentFile = useCallback(async () => {
    if (!currentFile) return

    const pendingTotal =
      pendingReplacements.total +
      extraPendingChanges.reduce((sum, c) => sum + (c.stats?.matches ?? 1), 0)

    if (pendingTotal > 0) {
      const choice = window.prompt(
        `检测到未确认修订（共 ${pendingTotal} 处/块）。\n` +
          `输入 1=全部接受并保存，2=全部拒绝并保存，0=取消`,
        '1'
      )
      if (choice === null || choice.trim() === '' || choice.trim() === '0') return
      // 注意：这里不能在 useCallback deps 中引用 confirmReplacement/rejectReplacement（TDZ）。
      // 直接在此处 resolve，并同步清空待确认队列。
      const mode = choice.trim() === '2' ? 'reject' : 'accept'
      const resolved = resolveDiffContent(mode)
      documentContentRef.current = resolved
      setDocument(prev => ({
        ...prev,
        content: resolved,
        lastModified: new Date(),
      }))
      setPendingReplacements({ items: [], total: 0 })
      setExtraPendingChanges([])
      setLastReplacement(null)
      setHasUnsavedChanges(true)
    }

    if (isElectron && window.electronAPI) {
      const ext = currentFile.name.split('.').pop()?.toLowerCase()
      
      if (ext === 'docx') {
        const sourceHtmlRaw = documentContentRef.current || document.content
        const sourceHtml = stripDiffMarkupForExport(sourceHtmlRaw)
        const blob = await createDocxBlob(sourceHtml, document.title, typographyProfile)
        let arrayBuffer = await blob.arrayBuffer()

        const hasTrackOrComments =
          typeof sourceHtml === 'string' &&
          (sourceHtml.includes('data-track-type=') ||
            sourceHtml.includes('class="docx-track"') ||
            sourceHtml.includes('data-comment-ids=') ||
            sourceHtml.includes('class="docx-comment"'))

        if (hasTrackOrComments) {
          try {
            arrayBuffer = await postProcessDocxWithAnnotations(arrayBuffer, {
              comments: (commentsForExport || []).map((c) => ({
                id: c.id,
                author: c.author,
                date: c.date,
                text: c.text,
              })),
            })
          } catch (e) {
            console.warn('[docxExportWithTracking] postProcess failed, fallback to plain export:', e)
          }
        }

        const base64 = arrayBufferToBase64(arrayBuffer)
        await window.electronAPI.writeBinaryFile(currentFile.path, base64)
      } else {
        await window.electronAPI.writeFile(currentFile.path, document.content)
      }
      
      // 保存成功后清除该文件的缓存
      fileContentCacheRef.current.delete(currentFile.path)
      setHasUnsavedChanges(false)
    } else {
      const sourceHtmlRaw = documentContentRef.current || document.content
      const sourceHtml = stripDiffMarkupForExport(sourceHtmlRaw)
      const blob = await createDocxBlob(sourceHtml, document.title, typographyProfile)
      let arrayBuffer = await blob.arrayBuffer()

      const hasTrackOrComments =
        typeof sourceHtml === 'string' &&
        (sourceHtml.includes('data-track-type=') ||
          sourceHtml.includes('class="docx-track"') ||
          sourceHtml.includes('data-comment-ids=') ||
          sourceHtml.includes('class="docx-comment"'))

      if (hasTrackOrComments) {
        try {
          arrayBuffer = await postProcessDocxWithAnnotations(arrayBuffer, {
            comments: (commentsForExport || []).map((c) => ({
              id: c.id,
              author: c.author,
              date: c.date,
              text: c.text,
            })),
          })
        } catch (e) {
          console.warn('[docxExportWithTracking] postProcess failed, fallback to plain export:', e)
        }
      }

      saveAs(new Blob([arrayBuffer], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }), `${document.title}.docx`)
      // 保存成功后清除该文件的缓存
      if (currentFile) {
        fileContentCacheRef.current.delete(currentFile.path)
      }
      setHasUnsavedChanges(false)
    }
  }, [currentFile, document, pendingReplacements.total, extraPendingChanges, typographyProfile, commentsForExport])

  // 刷新文件列表
  const refreshFiles = useCallback(async () => {
    if (!workspacePath || !isElectron || !window.electronAPI) return

    const result = await window.electronAPI.readFolder(workspacePath)
    if (result.success && result.data) {
      const convertFiles = (items: any[]): FileItem[] => {
        return items.map(item => ({
          name: item.name,
          path: item.path,
          type: item.type,
          children: item.children ? convertFiles(item.children) : undefined,
        }))
      }
      setFiles(convertFiles(result.data))
    }
  }, [workspacePath])

  // 上传 docx 文件 (Web 模式)
  const uploadDocxFile = useCallback(async (file: File) => {
    try {
      const arrayBuffer = await file.arrayBuffer()
      const base64 = btoa(String.fromCharCode(...new Uint8Array(arrayBuffer)))
      
      const title = file.name.replace(/\.docx?$/i, '')
      
      const newFile: FileItem = {
        name: file.name,
        path: `/${file.name}`,
        type: 'file',
      }
      
      setFiles(prev => [...prev, newFile])
      setCurrentFileState(newFile)
      setDocxData(base64)
      setDocument({
        title,
        content: '',
        styles: defaultStyles,
        lastModified: new Date(),
      })
      setHasUnsavedChanges(false)
    } catch (error) {
      console.error('Failed to upload docx file:', error)
      throw error
    }
  }, [])

  // 保存文档
  const saveDocument = useCallback(async () => {
    if (currentFile && isElectron) {
      await saveCurrentFile()
    } else {
      const pendingTotal =
        pendingReplacements.total +
        extraPendingChanges.reduce((sum, c) => sum + (c.stats?.matches ?? 1), 0)
      if (pendingTotal > 0) {
        const choice = window.prompt(
          `检测到未确认修订（共 ${pendingTotal} 处/块）。\n` +
            `输入 1=全部接受并导出，2=全部拒绝并导出，0=取消`,
          '1'
        )
        if (choice === null || choice.trim() === '' || choice.trim() === '0') return
        const mode = choice.trim() === '2' ? 'reject' : 'accept'
        const resolved = resolveDiffContent(mode)
        documentContentRef.current = resolved
        setDocument(prev => ({
          ...prev,
          content: resolved,
          lastModified: new Date(),
        }))
        setPendingReplacements({ items: [], total: 0 })
        setExtraPendingChanges([])
        setLastReplacement(null)
        setHasUnsavedChanges(true)
      }
      const sourceHtmlRaw = documentContentRef.current || document.content
      const sourceHtml = stripDiffMarkupForExport(sourceHtmlRaw)
      const blob = await createDocxBlob(sourceHtml, document.title, typographyProfile)
      let arrayBuffer = await blob.arrayBuffer()

      const hasTrackOrComments =
        typeof sourceHtml === 'string' &&
        (sourceHtml.includes('data-track-type=') ||
          sourceHtml.includes('class="docx-track"') ||
          sourceHtml.includes('data-comment-ids=') ||
          sourceHtml.includes('class="docx-comment"'))

      if (hasTrackOrComments) {
        try {
          arrayBuffer = await postProcessDocxWithAnnotations(arrayBuffer, {
            comments: (commentsForExport || []).map((c) => ({
              id: c.id,
              author: c.author,
              date: c.date,
              text: c.text,
            })),
          })
        } catch (e) {
          console.warn('[docxExportWithTracking] postProcess failed, fallback to plain export:', e)
        }
      }

      saveAs(new Blob([arrayBuffer], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }), `${document.title}.docx`)
      setHasUnsavedChanges(false)
    }
  }, [currentFile, document, saveCurrentFile, pendingReplacements.total, extraPendingChanges, typographyProfile, commentsForExport])

  // Agent 静默保存：将当前编辑器内容写回 .docx 文件（无弹窗）
  // Strip unresolved diff markup before writing or exporting docx.
  const stripDiffMarkupForExport = useCallback((html: string): string => {
    if (!html) return html
    if (!html.includes('diff-old') && !html.includes('diff-new') && !html.includes('data-diff-role') && !html.includes('data-diff-id')) {
      return html
    }

    const unwrapNode = (node: Element) => {
      const parent = node.parentNode
      if (!parent) return
      while (node.firstChild) {
        parent.insertBefore(node.firstChild, node)
      }
      parent.removeChild(node)
    }

    const stripDiffStyles = (el: HTMLElement) => {
      const style = el.getAttribute('style')
      if (!style || !/(fecaca|bbf7d0|c8e6c9|b91c1c|15803d|line-through)/i.test(style)) {
        return
      }

      const cleanedDeclarations = style
        .split(';')
        .map((part) => part.trim())
        .filter(Boolean)
        .filter((declaration) => {
          const [rawName, rawValue = ''] = declaration.split(':')
          const name = rawName.trim().toLowerCase()
          const value = rawValue.trim().toLowerCase()

          if (name === 'text-decoration' && value.includes('line-through')) return false
          if (name === 'background-color' && /(fecaca|bbf7d0|c8e6c9|rgb\(254\s*,\s*202\s*,\s*202\)|rgb\(187\s*,\s*247\s*,\s*208\)|rgb\(200\s*,\s*230\s*,\s*201\))/.test(value)) return false
          if (name === 'color' && /(b91c1c|15803d|rgb\(185\s*,\s*28\s*,\s*28\)|rgb\(21\s*,\s*128\s*,\s*61\))/.test(value)) return false
          if (name === 'padding' && /(1px\s+2px|0\s+2px)/.test(value)) return false
          if (name === 'border-radius' && value === '2px') return false
          return true
        })

      if (cleanedDeclarations.length > 0) {
        el.setAttribute('style', cleanedDeclarations.join('; '))
      } else {
        el.removeAttribute('style')
      }
    }

    try {
      const parser = new DOMParser()
      const doc = parser.parseFromString(html, 'text/html')

      doc.querySelectorAll<HTMLElement>('.diff-old').forEach((el) => el.remove())

      doc.querySelectorAll<HTMLElement>('.diff-new').forEach((el) => {
        if (el.tagName.toLowerCase() === 'span') {
          unwrapNode(el)
          return
        }
        el.classList.remove('diff-new')
        stripDiffStyles(el)
      })

      doc.querySelectorAll<HTMLElement>('[data-diff-role="old"]').forEach((el) => el.remove())

      doc.querySelectorAll<HTMLElement>('[data-diff-role="new"]').forEach((el) => {
        const tag = el.tagName.toLowerCase()
        el.removeAttribute('data-diff-id')
        el.removeAttribute('data-diff-role')
        el.removeAttribute('data-diff-kind')
        el.classList.remove('diff-old', 'diff-new')
        stripDiffStyles(el)

        if (tag === 'span') {
          unwrapNode(el)
        }
      })

      doc.querySelectorAll<HTMLElement>('[data-diff-id]').forEach((el) => {
        el.removeAttribute('data-diff-id')
        el.removeAttribute('data-diff-role')
        el.removeAttribute('data-diff-kind')
        el.classList.remove('diff-old', 'diff-new')
        stripDiffStyles(el)
      })

      return doc.body.innerHTML
    } catch (error) {
      console.warn('[silentSaveToFile] diff cleanup failed, fallback to regex:', error)
      return html
        .replace(/<span class="diff-old"[^>]*>[\s\S]*?<\/span>/g, '')
        .replace(/<span class="diff-new"[^>]*>([\s\S]*?)<\/span>/g, '$1')
        .replace(/\sdata-diff-id="[^"]*"/g, '')
        .replace(/\sdata-diff-role="[^"]*"/g, '')
        .replace(/\sdata-diff-kind="[^"]*"/g, '')
    }
  }, [])

  const silentSaveToFile = useCallback(async (): Promise<{ success: boolean; error?: string }> => {
    if (!currentFile || !isElectron || !window.electronAPI?.writeBinaryFile) {
      return { success: false, error: '无当前文件或非 Electron 环境' }
    }
    const ext = currentFile.name.split('.').pop()?.toLowerCase()
    if (ext !== 'docx') {
      return { success: false, error: `不支持的文件类型: ${ext}` }
    }
    try {
      const sourceHtmlRaw = documentContentRef.current || document.content
      const sourceHtml = stripDiffMarkupForExport(sourceHtmlRaw)
      const blob = await createDocxBlob(sourceHtml, document.title, typographyProfile)
      const arrayBuffer = await blob.arrayBuffer()
      const base64 = arrayBufferToBase64(arrayBuffer)
      const result = await window.electronAPI.writeBinaryFile(currentFile.path, base64)
      if (result.success) {
        setHasUnsavedChanges(false)
        return { success: true }
      }
      return { success: false, error: result.error || '写入失败' }
    } catch (e) {
      return { success: false, error: (e as Error).message }
    }
  }, [currentFile, document, isElectron, typographyProfile, stripDiffMarkupForExport])

  // AI 编辑应用
  const applyAIEdit = useCallback((newContent: string) => {
    setDocument(prev => ({
      ...prev,
      content: newContent,
      lastModified: new Date(),
    }))
    setDocxData(null)
    setHasUnsavedChanges(true)
  }, [])

  // 获取 Tiptap 文档的格式化结构信息（供 AI 参考）
  // 输出纯文本 + 格式标注，不含 HTML 标签，确保 AI 用纯文本搜索
  const getTiptapDocumentStructure = useCallback((): string => {
    const content = document.content
    if (!content) return ''
    
    const parser = new DOMParser()
    const doc = parser.parseFromString(content, 'text/html')
    
    const elements: string[] = []
    elements.push('【文档内容与格式】')
    elements.push('说明：引号内为精确文字（用于 search），方括号内为格式信息。\n')
    
    // 从元素提取简化字体名（去掉 fallback 列表，只取第一个）
    const simplifyFont = (raw: string): string => {
      if (!raw) return ''
      // "仿宋, STFANGSO, FanSong, \"Fangsong SC\", serif" → "仿宋"
      const first = raw.split(',')[0].trim().replace(/["']/g, '')
      return first
    }
    
    // 从一个元素收集格式信息
    const collectFormat = (el: HTMLElement): string[] => {
      const info: string[] = []
      const style = el.getAttribute('style') || ''
      const tag = el.tagName.toLowerCase()
      
      // 字体
      const fontFamily = style.match(/font-family:\s*([^;]+)/)?.[1] || ''
      if (fontFamily) info.push(simplifyFont(fontFamily))
      
      // 字号
      const fontSize = style.match(/font-size:\s*([^;]+)/)?.[1]?.trim() || ''
      if (fontSize) info.push(fontSize)
      
      // 粗体
      if (tag === 'strong' || tag === 'b' || style.includes('font-weight: bold') || style.includes('font-weight:bold')) info.push('粗体')
      // 斜体
      if (tag === 'em' || tag === 'i' || style.includes('font-style: italic')) info.push('斜体')
      // 下划线
      if (tag === 'u' || (style.includes('text-decoration') && style.includes('underline'))) info.push('下划线')
      // 删除线
      if (tag === 's' || tag === 'del' || (style.includes('text-decoration') && style.includes('line-through'))) info.push('删除线')
      
      // 颜色（排除 inherit/transparent）
      const colorMatch = style.match(/(?:^|[^-])color:\s*([^;]+)/)?.[1]?.trim() || ''
      if (colorMatch && colorMatch !== 'inherit' && colorMatch !== 'transparent' && !colorMatch.startsWith('var(')) info.push(`颜色:${colorMatch}`)
      
      // 背景色
      const bgColor = style.match(/background-color:\s*([^;]+)/)?.[1]?.trim() || ''
      if (bgColor && bgColor !== 'transparent' && !bgColor.startsWith('var(')) info.push(`背景:${bgColor}`)
      
      // 对齐
      const alignment = style.match(/text-align:\s*(\w+)/)?.[1] || ''
      if (alignment && alignment !== 'left') info.push(alignment === 'center' ? '居中' : alignment === 'right' ? '右对齐' : alignment === 'justify' ? '两端对齐' : alignment)
      
      // 缩进
      const textIndent = style.match(/text-indent:\s*([^;]+)/)?.[1]?.trim() || ''
      if (textIndent && textIndent !== '0' && textIndent !== '0px') info.push(`缩进:${textIndent}`)
      
      // 行距
      const lineHeight = style.match(/line-height:\s*([^;]+)/)?.[1]?.trim() || ''
      if (lineHeight && lineHeight !== 'normal') info.push(`行距:${lineHeight}`)
      
      return info
    }
    
    // 提取段内混合格式：当 <p> 内有多个不同格式的 <span> 时拆分标注
    const extractInlineRuns = (el: HTMLElement): string => {
      const spans = el.querySelectorAll(':scope > span, :scope > strong, :scope > b, :scope > em, :scope > i, :scope > u')
      if (spans.length <= 1) {
        // 只有一个或没有 span，整段统一格式
        return ''
      }
      
      // 检查是否有不同格式
      const runs: { format: string[]; text: string }[] = []
      let hasVariation = false
      let prevFormatKey = ''
      
      for (const span of Array.from(spans)) {
        const text = span.textContent?.trim() || ''
        if (!text) continue
        const fmt = collectFormat(span as HTMLElement)
        const fmtKey = fmt.join(',')
        if (prevFormatKey && fmtKey !== prevFormatKey) hasVariation = true
        prevFormatKey = fmtKey
        runs.push({ format: fmt, text })
      }
      
      if (!hasVariation || runs.length <= 1) return ''
      
      // 有格式变化：输出分段标注
      return runs.map(r => {
        const fmtStr = r.format.length > 0 ? `[${r.format.join(',')}]` : ''
        return `${fmtStr}"${r.text}"`
      }).join(' + ')
    }
    
    // 处理表格
    const processTable = (table: HTMLTableElement, tableIndex: number) => {
      const rows = table.querySelectorAll('tr')
      const colCount = rows[0]?.querySelectorAll('td, th').length || 0
      elements.push(`\n📊 表格${tableIndex} (${rows.length}行×${colCount}列):`)
      
      rows.forEach((row, rowIdx) => {
        const cells = row.querySelectorAll('td, th')
        cells.forEach((cell, colIdx) => {
          const cellText = cell.textContent?.trim() || ''
          if (cellText) {
            const cellEl = cell as HTMLElement
            const fmt = collectFormat(cellEl)
            // 也检查单元格内的 span/strong
            const innerBold = cell.querySelector('strong, b')
            if (innerBold && !fmt.includes('粗体')) fmt.push('粗体')
            const innerSpan = cell.querySelector('span')
            if (innerSpan) {
              const spanFmt = collectFormat(innerSpan as HTMLElement)
              for (const f of spanFmt) {
                if (!fmt.includes(f)) fmt.push(f)
              }
            }
            const formatStr = fmt.length > 0 ? ` [${fmt.join(',')}]` : ''
            elements.push(`   [${rowIdx+1},${colIdx+1}]${formatStr}: "${cellText}"`)
          }
        })
      })
    }
    
    // 遍历所有顶级元素
    let tableIndex = 1
    const processedTables = new Set<HTMLTableElement>()
    
    const walkNodes = (node: Node) => {
      if (node.nodeType === Node.ELEMENT_NODE) {
        const el = node as HTMLElement
        const tag = el.tagName.toLowerCase()
        
        // 跳过已处理的表格内部元素
        if (el.closest('table') && processedTables.has(el.closest('table') as HTMLTableElement)) {
          return
        }
        
        // 标题：提取实际格式（字体/字号/对齐等）
        if (tag === 'h1' || tag === 'h2' || tag === 'h3' || tag === 'h4') {
          const text = el.textContent?.trim() || ''
          if (text) {
            const level = tag.charAt(1)
            const fmt = collectFormat(el)
            // 也检查标题内部的 span 格式
            const innerSpan = el.querySelector('span')
            if (innerSpan) {
              const spanFmt = collectFormat(innerSpan as HTMLElement)
              for (const f of spanFmt) {
                if (!fmt.includes(f)) fmt.push(f)
              }
            }
            const formatStr = fmt.length > 0 ? ` [${fmt.join(',')}]` : ''
            elements.push(`📌 标题${level}${formatStr}: "${text}"`)
          }
        } else if (tag === 'p') {
          const text = el.textContent?.trim() || ''
          if (!text) return // 空段落跳过
          
          // 检查段内混合格式
          const inlineRuns = extractInlineRuns(el)
          if (inlineRuns) {
            // 段落级格式（对齐/缩进/行距）
            const paraFmt: string[] = []
            const style = el.getAttribute('style') || ''
            const alignment = style.match(/text-align:\s*(\w+)/)?.[1] || ''
            if (alignment && alignment !== 'left') paraFmt.push(alignment === 'center' ? '居中' : alignment === 'right' ? '右对齐' : alignment)
            const textIndent = style.match(/text-indent:\s*([^;]+)/)?.[1]?.trim() || ''
            if (textIndent && textIndent !== '0' && textIndent !== '0px') paraFmt.push(`缩进:${textIndent}`)
            const paraFmtStr = paraFmt.length > 0 ? ` [${paraFmt.join(',')}]` : ''
            elements.push(`📝 段落${paraFmtStr}: ${inlineRuns}`)
          } else {
            // 统一格式段落
            const fmt = collectFormat(el)
            // 如果段落本身没字体信息，检查内部第一个 span
            const innerSpan = el.querySelector('span')
            if (innerSpan) {
              const spanFmt = collectFormat(innerSpan as HTMLElement)
              for (const f of spanFmt) {
                if (!fmt.includes(f)) fmt.push(f)
              }
            }
            const formatStr = fmt.length > 0 ? ` [${fmt.join(',')}]` : ''
            elements.push(`📝 段落${formatStr}: "${text}"`)
          }
        } else if (tag === 'table') {
          processTable(el as HTMLTableElement, tableIndex++)
          processedTables.add(el as HTMLTableElement)
          return
        } else if (tag === 'ul') {
          const items = el.querySelectorAll(':scope > li')
          if (items.length > 0) {
            elements.push(`📋 无序列表 (${items.length}项):`)
            items.forEach((_item, _i) => {
              const text = _item.textContent?.trim() || ''
              if (text) {
                const fmt = collectFormat(_item as HTMLElement)
                const innerSpan = _item.querySelector('span')
                if (innerSpan) {
                  const spanFmt = collectFormat(innerSpan as HTMLElement)
                  for (const f of spanFmt) { if (!fmt.includes(f)) fmt.push(f) }
                }
                const fmtStr = fmt.length > 0 ? ` [${fmt.join(',')}]` : ''
                elements.push(`   •${fmtStr} "${text}"`)
              }
            })
          }
          return
        } else if (tag === 'ol') {
          const items = el.querySelectorAll(':scope > li')
          if (items.length > 0) {
            elements.push(`📋 有序列表 (${items.length}项):`)
            items.forEach((_item, _i) => {
              const text = _item.textContent?.trim() || ''
              if (text) {
                const fmt = collectFormat(_item as HTMLElement)
                const innerSpan = _item.querySelector('span')
                if (innerSpan) {
                  const spanFmt = collectFormat(innerSpan as HTMLElement)
                  for (const f of spanFmt) { if (!fmt.includes(f)) fmt.push(f) }
                }
                const fmtStr = fmt.length > 0 ? ` [${fmt.join(',')}]` : ''
                elements.push(`   ${_i+1}.${fmtStr} "${text}"`)
              }
            })
          }
          return
        } else if (tag === 'img') {
          // 图片占位
          const alt = el.getAttribute('alt') || ''
          const width = el.getAttribute('width') || el.style.width || ''
          const height = el.getAttribute('height') || el.style.height || ''
          const sizeInfo = width || height ? ` ${width}×${height}` : ''
          elements.push(`🖼️ 图片${sizeInfo}${alt ? ` (${alt})` : ''}`)
          return
        } else if (tag === 'hr') {
          elements.push('--- 分割线 ---')
          return
        } else if (tag === 'br') {
          return // 换行不需要额外标注
        }
      }
      
      // 递归处理子节点
      node.childNodes.forEach(child => walkNodes(child))
    }
    
    walkNodes(doc.body)
    
    elements.push('\n【工具使用提示】')
    elements.push('- search 参数必须使用引号内的精确纯文本，不要包含 HTML 标签或格式代码')
    elements.push('- 方括号内的格式信息仅供参考，修改格式请用工具的格式参数（bold/fontSize/fontFamily/color 等）')
    elements.push('- 如需修改格式但不改文字，可用 word_edit_ops 工具的 format_text / format_paragraph')
    
    return elements.join('\n')
  }, [document.content])

  // 生成唯一 ID
  const generateDiffId = () => `diff-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`

  // 精准替换文档内容（支持格式保留，支持多个修改共存）
  // 使用 ref 来获取最新内容，解决连续调用时的闭包问题
  const replaceInDocument = useCallback((search: string, replace: string, reviewMeta?: { reason?: string; type?: string }): ReplaceResult => {
    if (!search) {
      return { success: false, count: 0, message: '搜索内容不能为空' }
    }

    // 保护：search 和 replace 纯文本相同时跳过（避免无意义的 diff 标记）
    const stripHtml = (s: string) => s.replace(/<[^>]+>/g, '').replace(/\s+/g, ' ').trim()
    if (stripHtml(search) === stripHtml(replace)) {
      console.log(`[replaceInDocument] 跳过：search 与 replace 内容相同`)
      return { success: true, count: 0, message: '内容相同，无需替换' }
    }

    // 使用 ref 获取最新的文档内容（解决连续替换时闭包问题）
    const content = documentContentRef.current

    console.log(`[replaceInDocument] 搜索: "${search.slice(0, 30)}..." 替换为: "${replace.slice(0, 30)}..."`)
    console.log(`[replaceInDocument] 当前内容长度: ${content.length}`)
    
    // 转义正则特殊字符
    const escapeRegex = (str: string) => str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
    
    // 创建智能匹配正则 - 忽略 HTML 标签内部的匹配
    // 先尝试精确匹配，但要排除已有 diff 标记内的文字
    // 移除 diff 标记后的内容用于匹配（使用非贪婪匹配 [^<]* 代替 .*? 避免灾难性回溯）
    const unwrapWrappedQuotes = (value: string): string => {
      const trimmed = (value || '').trim()
      if (
        (trimmed.startsWith('"') && trimmed.endsWith('"')) ||
        (trimmed.startsWith("'") && trimmed.endsWith("'"))
      ) {
        return trimmed.slice(1, -1).trim()
      }
      return value
    }

    // Common failure mode: model wraps search text with quotes.
    const effectiveSearch = unwrapWrappedQuotes(search)

    // Match against what user currently sees: remove old diff, keep new diff.
    const contentForMatch = content
      .replace(/<span class="diff-old"[^>]*>[\s\S]*?<\/span>/g, '')
      .replace(/<span class="diff-new"[^>]*>([\s\S]*?)<\/span>/g, '$1')

    let positions: number[] = []
    let match

    const textContent = contentForMatch.replace(/<[^>]+>/g, '')
    const textRegex = new RegExp(escapeRegex(effectiveSearch), 'g')
    while ((match = textRegex.exec(textContent)) !== null) {
      positions.push(match.index)
    }

    let count = positions.length
    let useFuzzy = false

    // Fallback: ignore whitespace differences.
    if (count === 0) {
      const fuzzySearch = effectiveSearch.replace(/\s+/g, '\\s*')
      positions = []

      const fuzzyTextRegex = new RegExp(fuzzySearch, 'g')
      while ((match = fuzzyTextRegex.exec(textContent)) !== null) {
        positions.push(match.index)
      }
      count = positions.length
      useFuzzy = count > 0
    }

    if (count === 0) {
      const preview = effectiveSearch.length > 30 ? effectiveSearch.substring(0, 30) + '...' : effectiveSearch
      const hints: string[] = []

      if (effectiveSearch !== search) {
        hints.push('检测到 search 外层引号，已自动去除后重试。')
      }

      const condensedSearch = effectiveSearch.replace(/\s+/g, '')
      const condensedText = textContent.replace(/\s+/g, '')
      if (condensedSearch && condensedText.includes(condensedSearch)) {
        hints.push('忽略空白后可匹配，可能存在换行或空格差异。')
      }

      const probe = condensedSearch.slice(0, Math.min(condensedSearch.length, 8))
      if (probe.length >= 2) {
        const probePos = condensedText.indexOf(probe)
        if (probePos >= 0) {
          const excerptStart = Math.max(0, probePos - 12)
          const excerptEnd = Math.min(condensedText.length, probePos + probe.length + 20)
          const excerpt = condensedText.slice(excerptStart, excerptEnd)
          hints.push(`文档中有近似片段：${excerpt}`)
        }
      }

      const hintMessage = hints.length
        ? `；诊断：${hints.join('；')}`
        : ''

      console.log(`[replaceInDocument] not found: "${preview}" (text length:${textContent.length})`)
      return {
        success: false,
        count: 0,
        message: `未找到「${preview}」，请检查文字是否完全一致（包括标点和空格）${hintMessage}`,
        debug: {
          reasonCode: 'not_found',
          effectiveSearch,
          hints,
        },
      }
    }

    console.log(`[replaceInDocument] matched ${count} occurrence(s)`)

    // Generate a unique diff id for this operation.
    const diffId = generateDiffId()

    const extractFormatTags = (htmlFragment: string): { openTags: string[]; closeTags: string[] } => {
      const openTags: string[] = []
      const closeTags: string[] = []
      
      // 匹配格式化标签（保留顺序）
      const formatTagRegex = /<(strong|em|u|s|b|i|sub|sup|span[^>]*|mark[^>]*)>/gi
      const closeTagRegex = /<\/(strong|em|u|s|b|i|sub|sup|span|mark)>/gi
      
      let match
      while ((match = formatTagRegex.exec(htmlFragment)) !== null) {
        openTags.push(match[0])
      }
      while ((match = closeTagRegex.exec(htmlFragment)) !== null) {
        closeTags.unshift(match[0]) // 反向添加以保持正确的嵌套顺序
      }
      
      return { openTags, closeTags }
    }
    
    // 创建带唯一 ID 的 Diff 标记（细粒度：只标记实际变化的字词）
    const createDiffHtml = (oldText: string, newText: string, originalHtml: string) => {
      const formatText = (text: string) => text.replace(/\n/g, '<br>')

      // 提取原有格式标签
      const { openTags, closeTags } = extractFormatTags(originalHtml)
      const openTagsStr = openTags.join('')
      const closeTagsStr = closeTags.join('')

      // 细粒度 diff：找出公共前缀和后缀，只标记中间变化的部分
      let prefixLen = 0
      const minLen = Math.min(oldText.length, newText.length)
      while (prefixLen < minLen && oldText[prefixLen] === newText[prefixLen]) {
        prefixLen++
      }

      let suffixLen = 0
      while (
        suffixLen < (minLen - prefixLen) &&
        oldText[oldText.length - 1 - suffixLen] === newText[newText.length - 1 - suffixLen]
      ) {
        suffixLen++
      }

      const commonPrefix = oldText.slice(0, prefixLen)
      const commonSuffix = suffixLen > 0 ? oldText.slice(oldText.length - suffixLen) : ''
      const oldDiff = oldText.slice(prefixLen, oldText.length - suffixLen)
      const newDiff = newText.slice(prefixLen, newText.length - suffixLen)

      // 如果变化部分为空（完全相同），不生成 diff 标记
      if (!oldDiff && !newDiff) {
        return originalHtml
      }

      // 构建细粒度 diff HTML
      const parts: string[] = []
      if (commonPrefix) parts.push(formatText(commonPrefix))
      if (oldDiff) {
        parts.push(`<span class="diff-old" data-diff-id="${diffId}" style="background-color: #fecaca; color: #b91c1c; text-decoration: line-through; padding: 1px 2px; border-radius: 2px;">${formatText(oldDiff)}</span>`)
      }
      if (newDiff) {
        parts.push(`<span class="diff-new" data-diff-id="${diffId}" style="background-color: #bbf7d0; color: #15803d; padding: 1px 2px; border-radius: 2px;">${formatText(newDiff)}</span>`)
      }
      if (commonSuffix) parts.push(formatText(commonSuffix))

      return openTagsStr + parts.join('') + closeTagsStr
    }
    
    // 分段替换策略：将内容按照已有的 diff 标记分割，只在非 diff 区域进行替换
    // 这样可以保留之前的修改标注（使用 [^<]* 代替 .*? 避免灾难性回溯）
    const diffPattern = /<span class="diff-(old|new)" data-diff-id="[^"]*"[^>]*>[\s\S]*?<\/span>/g
    
    // 找出所有已有的 diff 标记的位置
    const diffMatches: { start: number; end: number; content: string }[] = []
    let diffMatch
    while ((diffMatch = diffPattern.exec(content)) !== null) {
      diffMatches.push({
        start: diffMatch.index,
        end: diffMatch.index + diffMatch[0].length,
        content: diffMatch[0]
      })
    }
    
    // 智能替换逻辑 - 支持跨 HTML 标签的文本匹配
    let newContent = content
    let appliedCount = 0
    
    // 核心函数：在 HTML 中查找并替换文本（忽略标签，但保留格式）
    const replaceTextInHtml = (html: string, searchText: string, createReplacement: (matchedText: string, originalHtml: string) => string): { html: string; replacedCount: number } => {
      // 将 HTML 分解为文本节点和标签
      const parts: { type: 'text' | 'tag'; content: string; index: number }[] = []
      let lastIndex = 0
      const tagRegex = /<[^>]+>/g
      let tagMatch
      
      while ((tagMatch = tagRegex.exec(html)) !== null) {
        if (tagMatch.index > lastIndex) {
          parts.push({ type: 'text', content: html.slice(lastIndex, tagMatch.index), index: lastIndex })
        }
        parts.push({ type: 'tag', content: tagMatch[0], index: tagMatch.index })
        lastIndex = tagMatch.index + tagMatch[0].length
      }
      if (lastIndex < html.length) {
        parts.push({ type: 'text', content: html.slice(lastIndex), index: lastIndex })
      }
      
      // 提取纯文本并记录每个字符在原 HTML 中的位置
      let pureText = ''
      const charToHtmlIndex: number[] = [] // pureText 中每个字符对应的 html 索引
      
      for (const part of parts) {
        if (part.type === 'text') {
          for (let i = 0; i < part.content.length; i++) {
            charToHtmlIndex.push(part.index + i)
            pureText += part.content[i]
          }
        }
      }
      
      // 在纯文本中查找所有匹配
      const searchRegex = useFuzzy 
        ? new RegExp(searchText.replace(/\s+/g, '\\s*'), 'g')
        : new RegExp(escapeRegex(searchText), 'g')
      
      const matches: { start: number; end: number; text: string }[] = []
      let m
      while ((m = searchRegex.exec(pureText)) !== null) {
        matches.push({ start: m.index, end: m.index + m[0].length, text: m[0] })
      }
      
      if (matches.length === 0) {
        return { html, replacedCount: 0 }
      }
      
      // 从后向前替换（避免索引偏移问题）
      let result = html
      for (let i = matches.length - 1; i >= 0; i--) {
        const match = matches[i]
        const htmlStart = charToHtmlIndex[match.start]
        const htmlEnd = charToHtmlIndex[match.end - 1] + 1
        
        // 提取原始 HTML 片段（包含格式标签）
        const originalHtmlFragment = result.slice(htmlStart, htmlEnd)
        // 提取纯文本
        const originalText = originalHtmlFragment.replace(/<[^>]+>/g, '')
        
        // 创建替换内容（传递原始 HTML 以保留格式）
        const replacement = createReplacement(originalText, originalHtmlFragment)
        
        // 替换
        result = result.slice(0, htmlStart) + replacement + result.slice(htmlEnd)
      }
      
      return { html: result, replacedCount: matches.length }
    }
    
    if (diffMatches.length === 0) {
      // 没有已有标记，直接替换（保留原有格式）
      const replaced = replaceTextInHtml(content, effectiveSearch, (matchedText, originalHtml) => createDiffHtml(matchedText, replace, originalHtml))
      newContent = replaced.html
      appliedCount = replaced.replacedCount
    } else {
      // 有已有标记，分段处理
      const segments: { type: 'normal' | 'diff'; content: string }[] = []
      let lastEnd = 0
      
      for (const dm of diffMatches) {
        if (dm.start > lastEnd) {
          segments.push({ type: 'normal', content: content.slice(lastEnd, dm.start) })
        }
        segments.push({ type: 'diff', content: dm.content })
        lastEnd = dm.end
      }
      if (lastEnd < content.length) {
        segments.push({ type: 'normal', content: content.slice(lastEnd) })
      }
      
      newContent = segments.map(seg => {
        if (seg.type === 'diff') {
          return seg.content
        } else {
          const replaced = replaceTextInHtml(seg.content, effectiveSearch, (matchedText, originalHtml) => createDiffHtml(matchedText, replace, originalHtml))
          appliedCount += replaced.replacedCount
          return replaced.html
        }
      }).join('')
    }
    
    if (appliedCount <= 0 || newContent === content) {
      return {
        success: false,
        count: 0,
        message: `No editable match found for "${effectiveSearch}"; this text may have already been modified in an earlier step.`,
        debug: {
          reasonCode: 'already_applied',
          effectiveSearch,
          hints: ['matching text only appears inside existing diff ranges'],
        },
      }
    }

    // 同步更新 ref（关键！这样下一次调用就能获取最新内容）
    documentContentRef.current = newContent
    
    setDocument(prev => ({
      ...prev,
      content: newContent,
      lastModified: new Date(),
    }))
    setDocxData(null)
    setHasUnsavedChanges(true)
    
    // 添加到待确认列表（保留之前的记录）
    const newReplacement: SingleReplacement = {
      id: diffId,
      searchText: effectiveSearch,
      replaceText: replace,
      count: appliedCount,
      timestamp: Date.now(),
      ...(reviewMeta?.reason ? { reviewReason: reviewMeta.reason } : {}),
      ...(reviewMeta?.type ? { reviewType: reviewMeta.type } : {}),
    }
    
    setPendingReplacements(prev => ({
      items: [...prev.items, newReplacement],
      total: prev.total + appliedCount
    }))
    
    // 同时更新 lastReplacement 以保持向后兼容
    setLastReplacement({
      searchText: effectiveSearch,
      replaceText: replace,
      count: appliedCount,
      timestamp: Date.now(),
      pending: true
    })

    return { 
      success: true, 
      count: appliedCount, 
      message: `成功替换 ${appliedCount} 处`,
      searchText: effectiveSearch,
      replaceText: replace,
      positions
    }
  }, []) // 使用 ref 后不需要依赖 document.content

  const addPendingReplacementItem = useCallback((item: SingleReplacement) => {
    if (!item?.id) return
    const count = Number(item.count || 0) || 0
    setPendingReplacements(prev => ({
      items: [...prev.items, item],
      total: prev.total + count,
    }))
    setLastReplacement({
      searchText: item.searchText,
      replaceText: item.replaceText,
      count: item.count,
      timestamp: item.timestamp || Date.now(),
      pending: true,
    })
  }, [])

  // 格式化替换 - 替换文字并应用格式
  const replaceWithFormat = useCallback((
    search: string, 
    replace: string,
    format?: {
      bold?: boolean
      italic?: boolean
      underline?: boolean
      color?: string
      backgroundColor?: string
      fontSize?: string
      fontFamily?: string
    },
    reviewMeta?: { reason?: string; type?: string }
  ): ReplaceResult => {
    if (!search) {
      return { success: false, count: 0, message: '搜索内容不能为空' }
    }

    // 构建带格式的替换文本
    let formattedReplace = replace
    const styles: string[] = []
    
    if (format?.bold) formattedReplace = `<strong>${formattedReplace}</strong>`
    if (format?.italic) formattedReplace = `<em>${formattedReplace}</em>`
    if (format?.underline) formattedReplace = `<u>${formattedReplace}</u>`
    
    if (format?.color) styles.push(`color: ${format.color}`)
    if (format?.backgroundColor) styles.push(`background-color: ${format.backgroundColor}`)
    if (format?.fontSize) styles.push(`font-size: ${format.fontSize}`)
    if (format?.fontFamily) styles.push(`font-family: ${format.fontFamily}`)
    
    if (styles.length > 0) {
      formattedReplace = `<span style="${styles.join('; ')}">${formattedReplace}</span>`
    }

    // 统一走带 diff 的替换逻辑：避免“内容变了但没有修订标记”
    // formattedReplace 允许包含少量 HTML（如 <strong>/<span style>）
    const result = replaceInDocument(search, formattedReplace, reviewMeta)
    return {
      ...result,
      // 给 UI/日志展示时，replaceText 用纯文本版本更友好
      replaceText: replace,
      message: result.success ? `成功格式化替换 ${result.count} 处（已生成修订）` : result.message,
    }
  }, [replaceInDocument])
  
  
  const cleanDiffStyles = (html: string) => {
    let content = html
    content = content.replace(/color:\s*rgb\(21,\s*128,\s*61\);?/gi, '')
    content = content.replace(/color:\s*#15803d;?/gi, '')
    content = content.replace(/background-color:\s*rgb\(187,\s*247,\s*208\);?/gi, '')
    content = content.replace(/background-color:\s*#bbf7d0;?/gi, '')
    content = content.replace(/color:\s*rgb\(185,\s*28,\s*28\);?/gi, '')
    content = content.replace(/color:\s*#b91c1c;?/gi, '')
    content = content.replace(/background-color:\s*rgb\(254,\s*202,\s*202\);?/gi, '')
    content = content.replace(/background-color:\s*#fecaca;?/gi, '')
    content = content.replace(/text-decoration:\s*line-through;?/gi, '')
    content = content.replace(/\s*style="\s*"/g, '')
    return content
  }

  const unwrapDiffSpan = (span: Element) => {
    const parent = span.parentNode
    if (!parent) return
    while (span.firstChild) {
      parent.insertBefore(span.firstChild, span)
    }
    parent.removeChild(span)
  }

  const resolveDiffContent = (mode: 'accept' | 'reject', onlyDiffId?: string) => {
    const currentContent = documentContentRef.current
    if (!currentContent) return ''

    const parser = new DOMParser()
    const doc = parser.parseFromString(currentContent, 'text/html')
    const spans = Array.from(doc.querySelectorAll('span'))

    spans.forEach(span => {
      const classList = Array.from(span.classList || [])
      const isOld = classList.includes('diff-old')
      const isNew = classList.includes('diff-new')
      if (!isOld && !isNew) return
      if (onlyDiffId) {
        const diffId = span.getAttribute('data-diff-id') || ''
        if (diffId !== onlyDiffId) return
      }

      if (mode === 'accept') {
        if (isOld) {
          span.remove()
          return
        }
        if (isNew) {
          unwrapDiffSpan(span)
        }
      } else {
        if (isNew) {
          span.remove()
          return
        }
        if (isOld) {
          unwrapDiffSpan(span)
        }
      }
    })

    // 块级 diff（paragraph/heading old/new）
    const blocks = Array.from(doc.querySelectorAll<HTMLElement>('[data-diff-id][data-diff-role]'))
    blocks.forEach((el) => {
      const diffId = el.getAttribute('data-diff-id') || ''
      const role = el.getAttribute('data-diff-role') || ''
      if (!diffId || !role) return
      if (onlyDiffId && diffId !== onlyDiffId) return

      const isOld = role === 'old'
      const isNew = role === 'new'

      if (mode === 'accept') {
        if (isOld) {
          el.remove()
          return
        }
        if (isNew) {
          el.removeAttribute('data-diff-id')
          el.removeAttribute('data-diff-role')
          el.removeAttribute('data-diff-kind')
        }
      } else {
        if (isNew) {
          el.remove()
          return
        }
        if (isOld) {
          el.removeAttribute('data-diff-id')
          el.removeAttribute('data-diff-role')
          el.removeAttribute('data-diff-kind')
        }
      }
    })

    return cleanDiffStyles(doc.body.innerHTML)
  }

  // 确认替换 - 移除红色部分，保留绿色部分（处理所有待确认的修改）
  const confirmReplacement = useCallback(() => {
    if (pendingReplacements.items.length === 0 && extraPendingChanges.length === 0 && !lastReplacement) return

    const content = resolveDiffContent('accept')
    if (content === undefined) return

    documentContentRef.current = content
    
    setDocument(prev => ({
      ...prev,
      content,
      lastModified: new Date(),
    }))
    setHasUnsavedChanges(true)
    
    setPendingReplacements({ items: [], total: 0 })
    setExtraPendingChanges([])
    setLastReplacement(null)
    
    // 接受修改后自动保存到磁盘
    setTimeout(() => { silentSaveToFile().catch(() => {}) }, 200)
  }, [pendingReplacements, lastReplacement, extraPendingChanges, silentSaveToFile])
  
  
  // 拒绝替换 - 移除绿色部分，恢复红色部分（处理所有待确认的修改）
  const rejectReplacement = useCallback(() => {
    if (pendingReplacements.items.length === 0 && extraPendingChanges.length === 0 && !lastReplacement) return
    
    const content = resolveDiffContent('reject')
    if (content === undefined) return

    documentContentRef.current = content
    
    setDocument(prev => ({
      ...prev,
      content,
      lastModified: new Date(),
    }))
    setHasUnsavedChanges(true)
    
    // 清空所有待确认记录
    setPendingReplacements({ items: [], total: 0 })
    setExtraPendingChanges([])
    setLastReplacement(null)
    
    // 拒绝修改后也自动保存到磁盘
    setTimeout(() => { silentSaveToFile().catch(() => {}) }, 200)
  }, [pendingReplacements, lastReplacement, extraPendingChanges, silentSaveToFile])

  const acceptChange = useCallback((id: string) => {
    if (!id) return
    const exists = pendingReplacements.items.find(i => i.id === id)
    const existsExtra = extraPendingChanges.find(c => c.id === id)
    if (!exists && !existsExtra) return

    const content = resolveDiffContent('accept', id)
    if (content === undefined) return

    documentContentRef.current = content

    setDocument(prev => ({
      ...prev,
      content,
      lastModified: new Date(),
    }))
    setHasUnsavedChanges(true)

    if (exists) {
      const remainingItems = pendingReplacements.items.filter(i => i.id !== id)
      const remainingTotal = Math.max(0, pendingReplacements.total - (exists.count || 0))
      setPendingReplacements({ items: remainingItems, total: remainingTotal })

      if (remainingItems.length > 0) {
        const last = remainingItems[remainingItems.length - 1]
        setLastReplacement({
          searchText: last.searchText,
          replaceText: last.replaceText,
          count: last.count,
          timestamp: last.timestamp,
          pending: true,
        })
      } else {
        setLastReplacement(null)
      }
    }

    if (existsExtra) {
      setExtraPendingChanges(prev => prev.filter(c => c.id !== id))
    }
    
    // 接受单条修改后自动保存到磁盘
    setTimeout(() => { silentSaveToFile().catch(() => {}) }, 200)
  }, [pendingReplacements, extraPendingChanges, silentSaveToFile])

  const rejectChange = useCallback((id: string) => {
    if (!id) return
    const exists = pendingReplacements.items.find(i => i.id === id)
    const existsExtra = extraPendingChanges.find(c => c.id === id)
    if (!exists && !existsExtra) return

    const content = resolveDiffContent('reject', id)
    if (content === undefined) return

    documentContentRef.current = content

    setDocument(prev => ({
      ...prev,
      content,
      lastModified: new Date(),
    }))
    setHasUnsavedChanges(true)

    if (exists) {
      const remainingItems = pendingReplacements.items.filter(i => i.id !== id)
      const remainingTotal = Math.max(0, pendingReplacements.total - (exists.count || 0))
      setPendingReplacements({ items: remainingItems, total: remainingTotal })

      if (remainingItems.length > 0) {
        const last = remainingItems[remainingItems.length - 1]
        setLastReplacement({
          searchText: last.searchText,
          replaceText: last.replaceText,
          count: last.count,
          timestamp: last.timestamp,
          pending: true,
        })
      } else {
        setLastReplacement(null)
      }
    }

    if (existsExtra) {
      setExtraPendingChanges(prev => prev.filter(c => c.id !== id))
    }
    
    // 拒绝单条修改后也自动保存到磁盘
    setTimeout(() => { silentSaveToFile().catch(() => {}) }, 200)
  }, [pendingReplacements, extraPendingChanges, silentSaveToFile])

  const acceptAllChanges = useCallback(() => {
    confirmReplacement()
  }, [confirmReplacement])

  const rejectAllChanges = useCallback(() => {
    rejectReplacement()
  }, [rejectReplacement])
  
  // 插入内容到文档
  const insertInDocument = useCallback((position: string, content: string): { success: boolean; message: string } => {
    if (!content) {
      return { success: false, message: '插入内容不能为空' }
    }

    // 使用 ref 获取最新内容（避免连续工具调用时闭包拿到旧内容）
    let newContent = documentContentRef.current
    const diffId = generateDiffId()
    const now = Date.now()

    const trimmed = (content || '').trim()
    const looksLikeHtml = /^</.test(trimmed) && /<\/?[a-z][\s\S]*>/i.test(trimmed)
    const looksLikeBlockHtml = /<\s*(p|h[1-6]|ul|ol|table|blockquote|div|section|hr)\b/i.test(trimmed)

    // 最小转义 + 换行处理（纯文本插入时使用），避免传入 `<` 破坏结构
    const escapeHtmlText = (text: string) =>
      (text || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;')

    const sanitizeIncomingHtml = (html: string) => {
      // 只做轻量净化：去掉 script/style/iframe 等高风险标签 + 去掉 on* 事件属性 + javascript: URL
      try {
        const parser = new DOMParser()
        const doc = parser.parseFromString(`<div id="__wc_insert_root">${html}</div>`, 'text/html')
        const root = doc.getElementById('__wc_insert_root')
        if (!root) return html

        const banned = root.querySelectorAll('script, style, iframe, object, embed, link, meta')
        banned.forEach((n) => n.remove())

        const all = root.querySelectorAll<HTMLElement>('*')
        all.forEach((el) => {
          // remove on* handlers
          for (const attr of Array.from(el.attributes)) {
            const name = (attr.name || '').toLowerCase()
            const value = (attr.value || '').toLowerCase()
            if (name.startsWith('on')) el.removeAttribute(attr.name)
            if ((name === 'href' || name === 'src') && value.startsWith('javascript:')) el.removeAttribute(attr.name)
          }
        })

        return root.innerHTML
      } catch {
        return html
      }
    }

    let insertHtml = ''
    if (looksLikeHtml) {
      const safeHtml = sanitizeIncomingHtml(trimmed)
      if (looksLikeBlockHtml) {
        // 块级 HTML：用 data-diff-role="new" 进行块级修订标记（避免把 <p> 塞进 <span> 导致显示异常）
        const parser = new DOMParser()
        const doc = parser.parseFromString(`<div id="__wc_insert_root">${safeHtml}</div>`, 'text/html')
        const root = doc.getElementById('__wc_insert_root')
        const children = root ? Array.from(root.children) : []
        const blockTags = new Set([
          'P',
          'H1',
          'H2',
          'H3',
          'H4',
          'H5',
          'H6',
          'UL',
          'OL',
          'TABLE',
          'BLOCKQUOTE',
          'DIV',
          'SECTION',
          'HR',
        ])

        const marked: string[] = []
        for (const el of children) {
          const tag = el.tagName.toUpperCase()
          if (blockTags.has(tag)) {
            el.setAttribute('data-diff-id', diffId)
            el.setAttribute('data-diff-role', 'new')
            el.setAttribute('data-diff-kind', 'block')
            marked.push(el.outerHTML)
          } else {
            // 非块级顶层元素：包到一个段落里，仍按块级 diff 标记
            const p = doc.createElement('p')
            p.setAttribute('data-diff-id', diffId)
            p.setAttribute('data-diff-role', 'new')
            p.setAttribute('data-diff-kind', 'block')
            p.appendChild(el.cloneNode(true))
            marked.push(p.outerHTML)
          }
        }
        insertHtml = marked.length > 0 ? marked.join('') : `<p data-diff-id="${diffId}" data-diff-role="new" data-diff-kind="block">${safeHtml}</p>`
      } else {
        // 仅内联 HTML：用 span.diff-new 包裹即可（strong/em/span/img 等）
        insertHtml = `<p><span class="diff-new" data-diff-id="${diffId}">${safeHtml}</span></p>`
      }
    } else {
      const formatted = escapeHtmlText(content).replace(/\n/g, '<br>')
      insertHtml = `<p><span class="diff-new" data-diff-id="${diffId}">${formatted}</span></p>`
    }

    // 提取 before:/after: 锚点
    const extractAnchor = (pos: string): { mode: 'after' | 'before'; anchor: string } | null => {
      if (pos.startsWith('after:')) return { mode: 'after', anchor: pos.slice(6).trim() }
      if (pos.startsWith('before:')) return { mode: 'before', anchor: pos.slice(7).trim() }
      return null
    }

    // 去除外层引号（模型常见错误：after:"某段文字"）
    const unwrapQuotes = (s: string) => {
      const t = s.trim()
      if ((t.startsWith('"') && t.endsWith('"')) || (t.startsWith("'") && t.endsWith("'")) ||
          (t.startsWith('\u201c') && t.endsWith('\u201d')) || (t.startsWith('\u2018') && t.endsWith('\u2019')))
        return t.slice(1, -1).trim()
      return s
    }

    if (position === 'start') {
      newContent = insertHtml + newContent
    } else if (position === 'end') {
      newContent = newContent + insertHtml
    } else {
      const parsed = extractAnchor(position)
      if (!parsed) {
        return { success: false, message: `无效的位置参数: ${position}，支持 start / end / after:文字 / before:文字` }
      }

      let { mode, anchor } = parsed
      anchor = unwrapQuotes(anchor)
      if (!anchor) {
        return { success: false, message: '锚点文字不能为空' }
      }

      // 剥离 HTML 标签得到纯文本（与 replaceInDocument 一致）
      const textContent = newContent.replace(/<[^>]+>/g, '')
      let anchorIdx = textContent.indexOf(anchor)

      // 模糊匹配：忽略空白差异
      if (anchorIdx === -1) {
        const escapeRegex = (str: string) => str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
        const fuzzy = escapeRegex(anchor).replace(/\s+/g, '\\s*')
        const m = new RegExp(fuzzy).exec(textContent)
        if (m) anchorIdx = m.index
      }

      if (anchorIdx === -1) {
        return { success: false, message: `未找到「${anchor.slice(0, 40)}」，请检查文字是否与文档一致` }
      }

      // 将纯文本偏移映射回 HTML 偏移（遍历 HTML，跳过标签，计数纯文本字符）
      let textPos = 0
      let htmlPos = 0
      const targetTextPos = mode === 'after' ? anchorIdx + anchor.length : anchorIdx

      while (htmlPos < newContent.length && textPos < targetTextPos) {
        if (newContent[htmlPos] === '<') {
          const tagEnd = newContent.indexOf('>', htmlPos)
          htmlPos = tagEnd !== -1 ? tagEnd + 1 : htmlPos + 1
        } else {
          textPos++
          htmlPos++
        }
      }

      if (mode === 'after') {
        // 向后找到最近的闭合标签后插入
        const nextTagEnd = newContent.indexOf('>', htmlPos)
        const insertPos = nextTagEnd !== -1 && (nextTagEnd - htmlPos) < 200 ? nextTagEnd + 1 : htmlPos
        newContent = newContent.slice(0, insertPos) + insertHtml + newContent.slice(insertPos)
      } else {
        // before: 向前回退到最近的标签边界
        let insertPos = htmlPos
        while (insertPos > 0 && newContent[insertPos - 1] !== '>') {
          insertPos--
        }
        newContent = newContent.slice(0, insertPos) + insertHtml + newContent.slice(insertPos)
      }
    }

    // 同步更新 ref（关键！保证下一次工具调用拿到最新内容）
    documentContentRef.current = newContent
    setDocument(prev => ({
      ...prev,
      content: newContent,
      lastModified: new Date(),
    }))
    setDocxData(null)
    setHasUnsavedChanges(true)

    // 记录到待审阅修改（插入：仅 diff-new）
    setExtraPendingChanges(prev => [
      ...prev,
      {
        id: diffId,
        kind: 'structure_edit',
        scope: 'document',
        summary: `插入内容（${position === 'start' ? '开头' : position === 'end' ? '末尾' : position.startsWith('before:') ? '指定位置前' : '指定位置后'}）`,
        beforePreview: '—',
        afterPreview: (content || '').length > 180 ? (content || '').slice(0, 180) + '…' : (content || ''),
        stats: { matches: 1 },
        timestamp: now,
        meta: { position, contentLength: content.length },
      },
    ])

    return { success: true, message: `已插入内容（已生成修订，可在修订面板查看）` }
  }, []) // 使用 ref 后不依赖 document.content

  // 删除文档中的内容
  const deleteInDocument = useCallback((target: string): { success: boolean; count: number; message: string } => {
    if (!target) {
      return { success: false, count: 0, message: '删除目标不能为空' }
    }

    const content = documentContentRef.current
    const diffId = generateDiffId()
    const now = Date.now()

    // 用 DOM 方式只处理文本节点，避免误伤标签/属性；同时跳过已有 diff 区域
    const parser = new DOMParser()
    const doc = parser.parseFromString(content, 'text/html')
    const walker = doc.createTreeWalker(doc.body, NodeFilter.SHOW_TEXT)
    let count = 0

    const isInsideDiff = (n: Node) => {
      const el = (n.parentElement || null)
      return !!el?.closest?.('.diff-old, .diff-new')
    }

    const nodes: Text[] = []
    let n: Node | null
    while ((n = walker.nextNode())) {
      if (n.nodeType === Node.TEXT_NODE) nodes.push(n as Text)
    }

    for (const textNode of nodes) {
      if (isInsideDiff(textNode)) continue
      const text = textNode.nodeValue || ''
      if (!text || !text.includes(target)) continue

      const parts = text.split(target)
      if (parts.length <= 1) continue

      const frag = doc.createDocumentFragment()
      for (let i = 0; i < parts.length; i++) {
        if (parts[i]) frag.appendChild(doc.createTextNode(parts[i]))
        if (i < parts.length - 1) {
          const span = doc.createElement('span')
          span.setAttribute('class', 'diff-old')
          span.setAttribute('data-diff-id', diffId)
          span.textContent = target
          frag.appendChild(span)
          count++
        }
      }
      textNode.parentNode?.replaceChild(frag, textNode)
    }

    if (count === 0) {
      return { success: false, count: 0, message: `未找到「${target}」` }
    }

    const newContent = doc.body.innerHTML

    // 同步更新 ref
    documentContentRef.current = newContent
    setDocument(prev => ({
      ...prev,
      content: newContent,
      lastModified: new Date(),
    }))
    setDocxData(null)
    setHasUnsavedChanges(true)

    // 记录到待审阅修改（删除：仅 diff-old）
    setExtraPendingChanges(prev => [
      ...prev,
      {
        id: diffId,
        kind: 'structure_edit',
        scope: 'document',
        summary: `删除 ${count} 处`,
        beforePreview: target.length > 180 ? target.slice(0, 180) + '…' : target,
        afterPreview: '—',
        stats: { matches: count },
        timestamp: now,
        meta: { target, count },
      },
    ])

    return { success: true, count, message: `已标记删除 ${count} 处（可在修订面板接受/拒绝）` }
  }, []) // 使用 ref 后不依赖 document.content

  // ─── DSL 缓存 ───
  const dslCacheRef = useRef<{ html: string; dsl: DocDsl } | null>(null)
  const getCachedDsl = useCallback((): DocDsl => {
    const html = documentContentRef.current
    if (dslCacheRef.current?.html === html) return dslCacheRef.current.dsl
    const dsl = htmlToDsl(html)
    dslCacheRef.current = { html, dsl }
    return dsl
  }, [])

  // ─── DSL 工具：replaceViaDsl ───
  const replaceViaDsl = useCallback((
    search: string,
    replace: string,
    options?: { blockIndex?: number; format?: Partial<DslRun> }
  ): { success: boolean; count: number; message: string } => {
    if (!search) return { success: false, count: 0, message: '搜索内容不能为空' }

    // 保护：search 和 replace 相同时跳过
    if (search === replace) {
      return { success: true, count: 0, message: '内容相同，无需替换' }
    }

    try {
      const dsl = getCachedDsl()
      const blocks = dsl.blocks

      // 确定搜索范围
      const searchBlocks: { block: DslBlock; index: number }[] = []
      if (options?.blockIndex !== undefined) {
        const idx = options.blockIndex
        if (idx < 0 || idx >= blocks.length) {
          return { success: false, count: 0, message: `blockIndex ${idx} 超出范围（共 ${blocks.length} 块）` }
        }
        searchBlocks.push({ block: blocks[idx], index: idx })
      } else {
        blocks.forEach((block, i) => searchBlocks.push({ block, index: i }))
      }

      // 在块中搜索并替换
      let totalCount = 0
      const unifiedDiffId = `diff-${Date.now()}-dsl`
      for (const { block } of searchBlocks) {
        if (block.type !== 'heading' && block.type !== 'paragraph') continue
        const runs = normalizeContent(block.content)
        const plainText = runs.map(r => r.text).join('')
        if (!plainText.includes(search)) continue

        // 找到匹配 → 简化处理：合并为纯文本，替换后重建
        // 简化处理：将所有 runs 合并为纯文本，替换后重建
        const replaced = plainText.split(search)
        if (replaced.length <= 1) continue

        totalCount += replaced.length - 1

        // 构建新 content：细粒度 diff，只标记实际变化的部分
        const diffId = unifiedDiffId

        // 计算 search 和 replace 的公共前缀/后缀
        let prefixLen = 0
        const minLen = Math.min(search.length, replace.length)
        while (prefixLen < minLen && search[prefixLen] === replace[prefixLen]) prefixLen++
        let suffixLen = 0
        while (suffixLen < (minLen - prefixLen) && search[search.length - 1 - suffixLen] === replace[replace.length - 1 - suffixLen]) suffixLen++

        const commonPrefix = search.slice(0, prefixLen)
        const commonSuffix = suffixLen > 0 ? search.slice(search.length - suffixLen) : ''
        const oldDiff = search.slice(prefixLen, search.length - suffixLen)
        const newDiff = replace.slice(prefixLen, replace.length - suffixLen)

        const newContent: DslInline[] = []
        for (let i = 0; i < replaced.length; i++) {
          if (replaced[i]) {
            newContent.push(replaced[i])
          }
          if (i < replaced.length - 1) {
            // 细粒度：公共前缀 + diff-old + diff-new + 公共后缀
            if (commonPrefix) newContent.push(commonPrefix)
            if (oldDiff) {
              newContent.push({
                text: oldDiff,
                _meta: { diffType: 'old' as const, diffId },
              })
            }
            if (newDiff) {
              const newRun: DslRun = {
                text: newDiff,
                _meta: { diffType: 'new' as const, diffId },
              }
              if (options?.format) Object.assign(newRun, options.format)
              newContent.push(newRun)
            }
            if (commonSuffix) newContent.push(commonSuffix)
          }
        }

        // 更新块内容
        if (block.type === 'heading') {
          (block as any).content = newContent
        } else {
          (block as any).content = newContent
        }
      }

      if (totalCount === 0) {
        return { success: false, count: 0, message: `未找到「${search.slice(0, 40)}」` }
      }

      // DSL → HTML → 更新编辑器
      const newHtml = dslToHtml(dsl)
      dslCacheRef.current = null // 清除缓存
      documentContentRef.current = newHtml
      setDocument(prev => prev ? { ...prev, content: newHtml, lastModified: new Date() } : prev)
      setHasUnsavedChanges(true)

      // 注册到待审阅修改（使 diff 接受/拒绝按钮可用）
      setPendingReplacements(prev => ({
        items: [...prev.items, {
          id: unifiedDiffId,
          searchText: search,
          replaceText: replace,
          count: totalCount,
          timestamp: Date.now(),
        }],
        total: prev.total + totalCount,
      }))
      setLastReplacement({
        searchText: search,
        replaceText: replace,
        count: totalCount,
        timestamp: Date.now(),
        pending: true,
      })

      return { success: true, count: totalCount, message: `已替换 ${totalCount} 处` }
    } catch (e) {
      console.error('[replaceViaDsl] error:', e)
      return { success: false, count: 0, message: `DSL 替换失败: ${e}` }
    }
  }, [getCachedDsl])

  // ─── DSL 工具：insertViaDsl ───
  const insertViaDsl = useCallback((
    position: string,
    dslBlocks: DslBlock[]
  ): { success: boolean; message: string } => {
    if (!dslBlocks.length) return { success: false, message: '插入内容不能为空' }

    try {
      const dsl = getCachedDsl()

      // 给新块加 diff 标记
      const diffId = `diff-insert-${Date.now()}`
      const markedBlocks = dslBlocks.map(b => {
        if (b.type === 'heading' || b.type === 'paragraph') {
          return { ...b, _meta: { diffRole: 'new' as const, diffId } }
        }
        return b
      })

      // 解析 position
      if (position === 'start') {
        dsl.blocks.unshift(...markedBlocks)
      } else if (position === 'end') {
        dsl.blocks.push(...markedBlocks)
      } else if (position.startsWith('blockIndex:')) {
        const idx = parseInt(position.replace('blockIndex:', ''))
        if (isNaN(idx) || idx < 0 || idx >= dsl.blocks.length) {
          return { success: false, message: `无效的 blockIndex: ${position}` }
        }
        dsl.blocks.splice(idx + 1, 0, ...markedBlocks)
      } else if (position.startsWith('after:') || position.startsWith('before:')) {
        const isAfter = position.startsWith('after:')
        const anchor = position.slice(isAfter ? 6 : 7).trim()
        if (!anchor) return { success: false, message: '锚点文字不能为空' }

        // 在块中搜索锚点
        let foundIdx = -1
        for (let i = 0; i < dsl.blocks.length; i++) {
          const block = dsl.blocks[i]
          if (block.type === 'heading' || block.type === 'paragraph') {
            const text = extractPlainText(block.content)
            if (text.includes(anchor)) {
              foundIdx = i
              break
            }
          }
        }

        if (foundIdx === -1) {
          return { success: false, message: `未找到「${anchor.slice(0, 40)}」` }
        }

        const insertIdx = isAfter ? foundIdx + 1 : foundIdx
        dsl.blocks.splice(insertIdx, 0, ...markedBlocks)
      } else {
        return { success: false, message: `无效的位置参数: ${position}` }
      }

      // DSL → HTML → 更新编辑器
      const newHtml = dslToHtml(dsl)
      dslCacheRef.current = null
      documentContentRef.current = newHtml
      setDocument(prev => prev ? { ...prev, content: newHtml, lastModified: new Date() } : prev)

      // 注册为待确认修改，让 accept/reject 按钮出现
      setPendingReplacements(prev => ({
        items: [...prev.items, {
          id: diffId,
          searchText: '',
          replaceText: `[插入 ${dslBlocks.length} 个块]`,
          count: dslBlocks.length,
          timestamp: Date.now(),
        }],
        total: prev.total + dslBlocks.length,
      }))

      return { success: true, message: `已插入 ${dslBlocks.length} 个块` }
    } catch (e) {
      console.error('[insertViaDsl] error:', e)
      return { success: false, message: `DSL 插入失败: ${e}` }
    }
  }, [getCachedDsl])

  // ─── DSL 工具：deleteViaDsl ───
  const deleteViaDsl = useCallback((
    target: string,
    options?: { blockIndex?: number }
  ): { success: boolean; count: number; message: string } => {
    if (!target && options?.blockIndex === undefined) {
      return { success: false, count: 0, message: '删除目标不能为空' }
    }

    try {
      const dsl = getCachedDsl()
      let count = 0

      if (options?.blockIndex !== undefined) {
        // 按块索引删除整个块
        const idx = options.blockIndex
        if (idx < 0 || idx >= dsl.blocks.length) {
          return { success: false, count: 0, message: `blockIndex ${idx} 超出范围` }
        }
        // 标记为 diff-old 而不是直接删除（让用户审查）
        const block = dsl.blocks[idx]
        if (block.type === 'heading' || block.type === 'paragraph') {
          const runs = normalizeContent(block.content)
          const markedRuns: DslInline[] = runs.map(r => ({
            ...r,
            _meta: { ...r._meta, diffType: 'old' as const },
          }))
          ;(block as any).content = markedRuns
        }
        count = 1
      } else {
        // 按文本搜索删除
        for (const block of dsl.blocks) {
          if (block.type !== 'heading' && block.type !== 'paragraph') continue
          const runs = normalizeContent(block.content)
          const plainText = runs.map(r => r.text).join('')
          if (!plainText.includes(target)) continue

          // 标记匹配文本为 diff-old
          const newContent: DslInline[] = []
          const parts = plainText.split(target)
          for (let i = 0; i < parts.length; i++) {
            if (parts[i]) newContent.push(parts[i])
            if (i < parts.length - 1) {
              newContent.push({
                text: target,
                _meta: { diffType: 'old' as const },
              })
              count++
            }
          }
          ;(block as any).content = newContent
        }
      }

      if (count === 0) {
        return { success: false, count: 0, message: `未找到「${target?.slice(0, 40) || 'blockIndex:' + options?.blockIndex}」` }
      }

      const newHtml = dslToHtml(dsl)
      dslCacheRef.current = null
      documentContentRef.current = newHtml
      setDocument(prev => prev ? { ...prev, content: newHtml, lastModified: new Date() } : prev)

      return { success: true, count, message: `已标记删除 ${count} 处` }
    } catch (e) {
      console.error('[deleteViaDsl] error:', e)
      return { success: false, count: 0, message: `DSL 删除失败: ${e}` }
    }
  }, [getCachedDsl])

  // 滚动到指定文本
  const scrollToText = useCallback((text: string) => {
    setScrollTarget(text)
    // 触发一个自定义事件，让 WordEditor 处理滚动
    window.dispatchEvent(new CustomEvent('scroll-to-text', { detail: { text } }))
  }, [])

  const scrollToDiffId = useCallback((diffId: string) => {
    if (!diffId) return
    window.dispatchEvent(new CustomEvent('scroll-to-diff-id', { detail: { diffId } }))
  }, [])

  const escapeHtml = (text: string) => {
    return (text ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;')
  }

  const TEMPLATE_PLACEHOLDER_REGEX = /_{3,}|（\s*）|\(\s*\)|【\s*】|\[\s*\]|\{\s*\}/

  const normalizeText = (text: string) =>
    (text || '').replace(/\s+/g, ' ').trim()

  const normalizeLabelKey = (label: string) =>
    normalizeText(label).toLowerCase().replace(/[：:，,。\s]/g, '')

  const inferFieldType = (label: string, value?: string) => {
    const text = `${label || ''} ${value || ''}`.toLowerCase()
    if (/日期|时间|date/.test(text)) return 'date'
    if (/金额|费用|总价|¥|\$|usd|rmb/.test(text)) return 'amount'
    if (/姓名|联系人|负责人|作者|签名|name/.test(text)) return 'person'
    if (/单位|公司|机构|部门|org|company/.test(text)) return 'organization'
    if (/电话|手机|联系方式|tel|phone/.test(text)) return 'phone'
    if (/邮箱|邮件|email/.test(text)) return 'email'
    if (/地址|住址|address/.test(text)) return 'address'
    if (/编号|证件号|合同号|id/.test(text)) return 'id'
    if (/%/.test(text)) return 'percent'
    return 'text'
  }

  const buildBlockMeta = (blocks: HTMLElement[]) => {
    const meta: Array<{
      index: number
      tag: string
      path: string
      headingPath: string
      text: string
    }> = []
    const headingStack: Array<{ level: number; index: number; text: string }> = []
    const counters: Record<string, number> = { p: 0, h1: 0, h2: 0, h3: 0, h4: 0, h5: 0, h6: 0 }

    blocks.forEach((el, index) => {
      const tag = el.tagName.toLowerCase()
      if (/^h[1-6]$/.test(tag)) {
        const level = Number(tag[1])
        counters[tag] = (counters[tag] || 0) + 1
        while (headingStack.length && headingStack[headingStack.length - 1].level >= level) {
          headingStack.pop()
        }
        headingStack.push({
          level,
          index: counters[tag],
          text: normalizeText(el.textContent || ''),
        })
        const path = headingStack.map((h) => `h${h.level}[${h.index}]`).join('/')
        meta.push({
          index,
          tag,
          path,
          headingPath: path,
          text: normalizeText(el.textContent || ''),
        })
        return
      }

      if (tag === 'p') {
        counters.p += 1
        const headingPath = headingStack.map((h) => `h${h.level}[${h.index}]`).join('/')
        const path = headingPath ? `${headingPath}/p[${counters.p}]` : `p[${counters.p}]`
        meta.push({
          index,
          tag,
          path,
          headingPath,
          text: normalizeText(el.textContent || ''),
        })
      }
    })

    return meta
  }

  const detectTemplateFields = (doc: globalThis.Document, limit = 60): TemplateFieldCandidate[] => {
    const fields: TemplateFieldCandidate[] = []
    const blocks = Array.from(doc.body.querySelectorAll<HTMLElement>('p,h1,h2,h3,h4,h5,h6'))
    const blockMeta = buildBlockMeta(blocks)
    const groupCount = new Map<string, number>()

    blocks.forEach((el, index) => {
      if (fields.length >= limit) return
      const rawText = normalizeText(el.textContent || '')
      if (!rawText) return
      const path = blockMeta[index]?.path || ''

      // 1) 标签 + 冒号（右侧为空或占位符）
      const colonIndex = rawText.indexOf('：') >= 0 ? rawText.indexOf('：') : rawText.indexOf(':')
      if (colonIndex > -1) {
        const label = rawText.slice(0, colonIndex).trim()
        const tail = rawText.slice(colonIndex + 1).trim()
        if (label && (!tail || TEMPLATE_PLACEHOLDER_REGEX.test(tail))) {
          const labelWithColon = rawText.slice(0, colonIndex + 1)
          const groupKey = normalizeLabelKey(label)
          groupCount.set(groupKey, (groupCount.get(groupKey) || 0) + 1)
          fields.push({
            id: `p:${index}:colon`,
            label,
            kind: 'colon',
            context: rawText.slice(0, 80),
            path,
            currentValue: tail || '',
            fieldType: inferFieldType(label, tail || ''),
            groupKey,
            meta: { blockIndex: index, labelWithColon, placeholder: tail },
          })
        }
      }

      if (fields.length >= limit) return

      // 2) 纯占位符（下划线/括号空位）
      const blankMatch = rawText.match(TEMPLATE_PLACEHOLDER_REGEX)
      if (blankMatch) {
        const placeholder = blankMatch[0]
        const before = rawText.split(placeholder)[0].trim()
        const label = before || '未命名字段'
        const groupKey = normalizeLabelKey(label)
        groupCount.set(groupKey, (groupCount.get(groupKey) || 0) + 1)
        fields.push({
          id: `p:${index}:blank`,
          label,
          kind: 'blank',
          context: rawText.slice(0, 80),
          path,
          currentValue: '',
          fieldType: inferFieldType(label, ''),
          groupKey,
          meta: { blockIndex: index, placeholder },
        })
      }
    })

    if (fields.length >= limit) return fields

    // 3) 表格：左标签 + 右空位
    const tables = Array.from(doc.body.querySelectorAll<HTMLTableElement>('table'))
    tables.forEach((table, tableIndex) => {
      if (fields.length >= limit) return
      const rows = Array.from(table.querySelectorAll('tr'))
      rows.forEach((row, rowIndex) => {
        if (fields.length >= limit) return
        const cells = Array.from(row.querySelectorAll<HTMLElement>('th,td'))
        for (let col = 0; col < cells.length - 1; col++) {
          if (fields.length >= limit) break
          const labelCell = cells[col]
          const valueCell = cells[col + 1]
          const labelText = normalizeText(labelCell.textContent || '')
          if (!labelText) continue
          const valueText = normalizeText(valueCell.textContent || '')
          if (!valueText || TEMPLATE_PLACEHOLDER_REGEX.test(valueText)) {
            const groupKey = normalizeLabelKey(labelText)
            groupCount.set(groupKey, (groupCount.get(groupKey) || 0) + 1)
            fields.push({
              id: `t:${tableIndex}:r:${rowIndex}:c:${col + 1}`,
              label: labelText,
              kind: 'table',
              context: `表格${tableIndex + 1} 行${rowIndex + 1}`,
              path: `table[${tableIndex + 1}]/r[${rowIndex + 1}]/c[${col + 2}]`,
              currentValue: valueText || '',
              fieldType: inferFieldType(labelText, valueText || ''),
              groupKey,
              meta: { tableIndex, rowIndex, colIndex: col + 1 },
            })
          }
        }
      })
    })

    if (fields.length > 0) {
      const groupCounts = new Map<string, number>()
      fields.forEach((f) => {
        const key = f.groupKey || normalizeLabelKey(f.label)
        groupCounts.set(key, (groupCounts.get(key) || 0) + 1)
      })
      fields.forEach((f) => {
        const key = f.groupKey || normalizeLabelKey(f.label)
        const count = groupCounts.get(key) || 1
        f.groupKey = key
        f.meta = { ...(f.meta || {}), groupKey: key, groupCount: count }
      })
    }

    return fields
  }

  const previewWordOps = useCallback((ops: WordEditOp[]) => {
    try {
      if (!Array.isArray(ops) || ops.length === 0) {
        return { success: false, message: 'ops 为空或格式不正确' }
      }

      const content = documentContentRef.current || ''
      const parser = new DOMParser()
      const doc = parser.parseFromString(content, 'text/html')

      const lines: string[] = []
      let estimated = 0

      for (const op of ops) {
        if (!op || typeof op !== 'object') continue
        const type = op.type

        if (type === 'apply_style' || type === 'format_paragraph') {
          const anchor = op.target?.scope === 'anchor_text' ? (op.target?.text || '') : ''
          const blocks = Array.from(doc.body.querySelectorAll<HTMLElement>('p,h1,h2,h3'))
          const matched = anchor
            ? blocks.filter(b => (b.textContent || '').includes(anchor))
            : blocks
          estimated += matched.length
          lines.push(`${type}: 预计影响 ${matched.length} 个块`)
          continue
        }

        if (type === 'format_text') {
          const t = (op.target?.text || '').toString()
          if (!t) {
            lines.push('format_text: 缺少 target.text')
            continue
          }
          const plain = (doc.body.textContent || '')
          const count = plain.split(t).length - 1
          estimated += Math.max(0, count)
          lines.push(`format_text: 预计命中 ${count} 处 "${t.slice(0, 30)}${t.length > 30 ? '…' : ''}"`)
          continue
        }

        // 预览新增操作类型
        if (type === 'outline_summary') {
          const action = String(op.params?.action || 'extract_outline')
          const headings = Array.from(doc.body.querySelectorAll<HTMLElement>('h1,h2,h3,h4,h5,h6'))
          estimated += headings.length
          lines.push(`outline_summary (${action}): 文档含 ${headings.length} 个标题`)
          continue
        }

        if (type === 'template_fill') {
          const action = String(op.params?.action || 'detect_placeholders')

          if (action === 'detect_fields') {
            const fields = detectTemplateFields(doc)
            estimated += fields.length
            lines.push(`template_fill (detect_fields): 发现 ${fields.length} 个候选字段`)
            const previewLines = fields.slice(0, 12).map((f) => {
              const path = f.path ? ` | ${f.path}` : ''
              const value = f.currentValue ? ` | 当前值: ${f.currentValue}` : ''
              return `${f.id} | ${f.label}${path}${value}`
            })
            if (previewLines.length) {
              lines.push(...previewLines)
            }
            if (fields.length > previewLines.length) {
              lines.push(`... 还有 ${fields.length - previewLines.length} 个字段未展示`)
            }
            continue
          }

          if (action === 'apply' || action === 'apply_fields') {
            let assignments: Array<{ fieldId?: string; value?: string; label?: string }> = []
            const raw = op.params?.assignments
            if (Array.isArray(raw)) {
              assignments = raw as Array<{ fieldId?: string; value?: string; label?: string }>
            } else if (typeof raw === 'string') {
              try {
                assignments = JSON.parse(raw)
              } catch {
                assignments = []
              }
            }
            estimated += assignments.length
            lines.push(`template_fill (apply): 计划填充 ${assignments.length} 项`)
            assignments.slice(0, 12).forEach((item) => {
              const label = item.label || item.fieldId || '字段'
              const value = (item.value || '').toString().slice(0, 40)
              lines.push(`${label} → ${value}`)
            })
            if (assignments.length > 12) {
              lines.push(`... 还有 ${assignments.length - 12} 项未展示`)
            }
            continue
          }

          const plainText = doc.body.innerHTML
          // 简单统计占位符
          const placeholderCount = (plainText.match(/\{\{[^}]+\}\}|【[^】]+】|\[[^\]]+\]|_{3,}|<[^>]+>/g) || []).length
          estimated += placeholderCount
          lines.push(`template_fill (${action}): 预计 ${placeholderCount} 个占位符`)
          continue
        }

        if (type === 'citation_footnote') {
          const action = String(op.params?.action || 'insert_footnote')
          const existingNotes = doc.body.querySelectorAll('.footnote-ref, .endnote-ref, .bibliography-item')
          estimated += 1
          lines.push(`citation_footnote (${action}): 现有 ${existingNotes.length} 个注释/引用`)
          continue
        }
      }

      return {
        success: true,
        message: `word_edit_ops 预览：共 ${ops.length} 个操作，预计影响 ${estimated} 处/块。`,
        data: { lines, estimated, opCount: ops.length },
      }
    } catch (e) {
      return { success: false, message: `预览失败: ${(e as Error).message || String(e)}` }
    }
  }, [])

  const applyWordOps = useCallback((ops: WordEditOp[]) => {
    try {
      if (!Array.isArray(ops) || ops.length === 0) {
        return { success: false, message: 'ops 为空或格式不正确' }
      }

      const originalHtml = documentContentRef.current || ''
      let html = originalHtml
      const parser = new DOMParser()
      const doc = parser.parseFromString(html, 'text/html')

      const created: PendingChange[] = []
      const genId = () => `diff-${Date.now()}-${Math.random().toString(36).slice(2, 10)}`
      let templateFillReport: { applied: number; skipped: number; warnings: number } | null = null

      const markBlockPair = (oldEl: HTMLElement, newEl: HTMLElement, diffId: string) => {
        oldEl.setAttribute('data-diff-id', diffId)
        oldEl.setAttribute('data-diff-role', 'old')
        oldEl.setAttribute('data-diff-kind', 'block')
        newEl.setAttribute('data-diff-id', diffId)
        newEl.setAttribute('data-diff-role', 'new')
        newEl.setAttribute('data-diff-kind', 'block')
      }

      const buildDiffOld = (diffId: string, text: string) =>
        `<span class="diff-old" data-diff-id="${diffId}">${escapeHtml(text)}</span>`

      const buildDiffNew = (diffId: string, text: string) =>
        `<span class="diff-new" data-diff-id="${diffId}">${escapeHtml(text)}</span>`

      const buildDiffPair = (diffId: string, oldText: string, newText: string) =>
        `${buildDiffOld(diffId, oldText)}${buildDiffNew(diffId, newText)}`

      const extractFormatTags = (htmlFragment: string): { openTags: string[]; closeTags: string[] } => {
        const openTags: string[] = []
        const closeTags: string[] = []
        const formatTagRegex = /<(strong|em|u|s|b|i|sub|sup|span[^>]*|mark[^>]*)>/gi
        const closeTagRegex = /<\/(strong|em|u|s|b|i|sub|sup|span|mark)>/gi
        let match
        while ((match = formatTagRegex.exec(htmlFragment)) !== null) {
          openTags.push(match[0])
        }
        while ((match = closeTagRegex.exec(htmlFragment)) !== null) {
          closeTags.unshift(match[0])
        }
        return { openTags, closeTags }
      }

      const buildDiffPairWithFormat = (diffId: string, oldText: string, newText: string, originalHtml: string) => {
        const formatText = (text: string) => escapeHtml(text).replace(/\n/g, '<br>')
        const { openTags, closeTags } = extractFormatTags(originalHtml)
        const formattedNew = `${openTags.join('')}${formatText(newText)}${closeTags.join('')}`
        return `<span class="diff-old" data-diff-id="${diffId}">${originalHtml}</span>` +
          `<span class="diff-new" data-diff-id="${diffId}">${formattedNew}</span>`
      }

      const replaceFirstTextInHtml = (
        htmlFragment: string,
        searchText: string,
        buildReplacement: (matchedText: string, originalHtml: string) => string
      ) => {
        if (!searchText) return htmlFragment
        const parts: { type: 'text' | 'tag'; content: string; index: number }[] = []
        let lastIdx = 0
        const tagRegex = /<[^>]+>/g
        let tagMatch: RegExpExecArray | null
        while ((tagMatch = tagRegex.exec(htmlFragment)) !== null) {
          if (tagMatch.index > lastIdx) {
            parts.push({ type: 'text', content: htmlFragment.slice(lastIdx, tagMatch.index), index: lastIdx })
          }
          parts.push({ type: 'tag', content: tagMatch[0], index: tagMatch.index })
          lastIdx = tagMatch.index + tagMatch[0].length
        }
        if (lastIdx < htmlFragment.length) {
          parts.push({ type: 'text', content: htmlFragment.slice(lastIdx), index: lastIdx })
        }

        let pure = ''
        const map: number[] = []
        for (const p of parts) {
          if (p.type === 'text') {
            for (let i = 0; i < p.content.length; i++) {
              map.push(p.index + i)
              pure += p.content[i]
            }
          }
        }
        if (!pure) return htmlFragment
        const idx = pure.indexOf(searchText)
        if (idx < 0) return htmlFragment
        const htmlStart = map[idx]
        const htmlEnd = map[idx + searchText.length - 1] + 1
        const originalHtml = htmlFragment.slice(htmlStart, htmlEnd)
        const replacement = buildReplacement(searchText, originalHtml)
        return htmlFragment.slice(0, htmlStart) + replacement + htmlFragment.slice(htmlEnd)
      }

      const applyTemplateValueToElement = (el: HTMLElement, value: string, diffId: string, oldValue?: string) => {
        const rawText = normalizeText(el.textContent || '')
        if (!rawText) {
          el.innerHTML = buildDiffNew(diffId, value)
          return
        }

        if (oldValue && rawText.includes(oldValue)) {
          const updated = replaceFirstTextInHtml(el.innerHTML, oldValue, (_m, originalHtml) =>
            buildDiffPairWithFormat(diffId, oldValue, value, originalHtml)
          )
          el.innerHTML = updated
          return
        }

        const placeholderMatch = rawText.match(TEMPLATE_PLACEHOLDER_REGEX)
        if (placeholderMatch) {
          const placeholder = placeholderMatch[0]
          const updated = replaceFirstTextInHtml(el.innerHTML, placeholder, () =>
            buildDiffPair(diffId, placeholder, value)
          )
          el.innerHTML = updated
          return
        }

        const colonIndex = rawText.indexOf('：') >= 0 ? rawText.indexOf('：') : rawText.indexOf(':')
        if (colonIndex > -1) {
          const labelWithColon = rawText.slice(0, colonIndex + 1)
          const updated = replaceFirstTextInHtml(el.innerHTML, labelWithColon, (_m, originalHtml) =>
            `${originalHtml}${buildDiffNew(diffId, value)}`
          )
          el.innerHTML = updated
          return
        }

        el.innerHTML = `${el.innerHTML}${buildDiffNew(diffId, value)}`
      }

      for (const op of ops) {
        if (!op || typeof op !== 'object') continue
        const type = op.type

        if (type === 'apply_style' || type === 'format_paragraph') {
          const anchor = op.target?.scope === 'anchor_text' ? (op.target?.text || '') : ''
          const blocks = Array.from(doc.body.querySelectorAll<HTMLElement>('p,h1,h2,h3'))
          const matched = anchor
            ? blocks.filter(b => (b.textContent || '').includes(anchor))
            : blocks

          for (const el of matched) {
            const diffId = genId()
            const oldClone = el.cloneNode(true) as HTMLElement
            let newClone: HTMLElement

            if (type === 'apply_style') {
              const styleName = (op.params?.styleName || '').toString()
              const tag = styleName === 'Heading1'
                ? 'h1'
                : styleName === 'Heading2'
                  ? 'h2'
                  : styleName === 'Heading3'
                    ? 'h3'
                    : 'p'

              if (el.tagName.toLowerCase() === tag) {
                newClone = el.cloneNode(true) as HTMLElement
              } else {
                newClone = doc.createElement(tag)
                newClone.innerHTML = el.innerHTML
                const style = el.getAttribute('style') || ''
                if (style) newClone.setAttribute('style', style)
              }
            } else {
              // format_paragraph - 支持完整段落格式参数
              newClone = el.cloneNode(true) as HTMLElement
              const prevStyle = newClone.getAttribute('style') || ''
              
              // 需要清理的样式属性列表
              const stylePropsToClean = [
                'text-align',
                'line-height',
                'margin-top',
                'margin-bottom',
                'text-indent',
                'margin-left',
                'margin-right',
                'background-color',
                'background',
                'border',
                'padding'
              ]
              
              // 清理现有样式
              let cleaned = prevStyle
              for (const prop of stylePropsToClean) {
                cleaned = cleaned.replace(new RegExp(`${prop}\\s*:\\s*[^;]+;?`, 'gi'), '').trim()
              }
              
              // 构建新样式
              const newStyles: string[] = []
              
              // 对齐方式
              const alignment = (op.params?.alignment || '').toString()
              if (alignment) {
                newStyles.push(`text-align: ${alignment}`)
              }
              
              // 行距
              const lineHeight = op.params?.lineHeight
              if (lineHeight !== undefined && lineHeight !== null) {
                const lh = String(lineHeight)
                // 如果是纯数字（如 1.5, 2），直接使用；否则保留单位
                newStyles.push(`line-height: ${lh}`)
              }
              
              // 段前间距
              const spaceBefore = (op.params?.spaceBefore || '').toString()
              if (spaceBefore) {
                newStyles.push(`margin-top: ${spaceBefore}`)
              }
              
              // 段后间距
              const spaceAfter = (op.params?.spaceAfter || '').toString()
              if (spaceAfter) {
                newStyles.push(`margin-bottom: ${spaceAfter}`)
              }
              
              // 首行缩进
              const textIndent = (op.params?.textIndent || '').toString()
              if (textIndent) {
                newStyles.push(`text-indent: ${textIndent}`)
              }
              
              // 左边距
              const marginLeft = (op.params?.marginLeft || '').toString()
              if (marginLeft) {
                newStyles.push(`margin-left: ${marginLeft}`)
              }
              
              // 右边距
              const marginRight = (op.params?.marginRight || '').toString()
              if (marginRight) {
                newStyles.push(`margin-right: ${marginRight}`)
              }
              
              // 背景色
              const backgroundColor = (op.params?.backgroundColor || '').toString()
              if (backgroundColor) {
                newStyles.push(`background-color: ${backgroundColor}`)
              }
              
              // 边框
              const border = (op.params?.border || '').toString()
              if (border) {
                newStyles.push(`border: ${border}`)
              }
              
              // 内边距
              const padding = (op.params?.padding || '').toString()
              if (padding) {
                newStyles.push(`padding: ${padding}`)
              }
              
              // 合并样式
              const finalStyle = [cleaned, ...newStyles].filter(s => s.trim()).join('; ')
              if (finalStyle) {
                newClone.setAttribute('style', finalStyle + ';')
              }
            }

            markBlockPair(oldClone, newClone, diffId)

            el.replaceWith(oldClone)
            oldClone.insertAdjacentElement('afterend', newClone)

            created.push({
              id: diffId,
              kind: type === 'apply_style' ? 'apply_style' : 'format_paragraph',
              scope: op.target?.scope === 'anchor_text' ? 'selection' : 'document',
              summary: type === 'apply_style' ? '应用样式' : '段落格式调整',
              beforePreview: (oldClone.textContent || '').trim(),
              afterPreview: (newClone.textContent || '').trim(),
              stats: { matches: 1 },
              timestamp: Date.now(),
              meta: { op },
            })
          }

          continue
        }

        if (type === 'format_text') {
          const targetText = (op.target?.text || '').toString()
          if (!targetText) continue
          const diffId = genId()

          const makeStyled = (text: string) => {
            const escaped = escapeHtml(text)
            const styles: string[] = []
            const fontFamily = op.params?.fontFamily ? String(op.params.fontFamily) : ''
            const fontSize = op.params?.fontSize ? String(op.params.fontSize) : ''
            const color = op.params?.color ? String(op.params.color) : ''
            const highlight = op.params?.highlight ? String(op.params.highlight) : ''
            const letterSpacing = op.params?.letterSpacing ? String(op.params.letterSpacing) : ''
            
            if (fontFamily) styles.push(`font-family: ${fontFamily}`)
            if (fontSize) styles.push(`font-size: ${fontSize}`)
            if (color) styles.push(`color: ${color}`)
            if (highlight) styles.push(`background-color: ${highlight}`)
            if (letterSpacing) styles.push(`letter-spacing: ${letterSpacing}`)
            // 删除线通过样式实现
            if (op.params?.strikethrough) styles.push('text-decoration: line-through')

            let inner = escaped
            if (op.params?.bold) inner = `<strong>${inner}</strong>`
            if (op.params?.italic) inner = `<em>${inner}</em>`
            if (op.params?.underline) inner = `<u>${inner}</u>`
            if (op.params?.strikethrough) inner = `<s>${inner}</s>`
            if (op.params?.superscript) inner = `<sup>${inner}</sup>`
            if (op.params?.subscript) inner = `<sub>${inner}</sub>`

            if (styles.length > 0) {
              inner = `<span style="${styles.join('; ')}">${inner}</span>`
            }
            return inner
          }

          // 分段替换：跳过已有 diff span（避免嵌套）
          const diffPattern = /<span class="diff-(old|new)" data-diff-id="[^"]*"[^>]*>[^<]*<\/span>/g
          const segments: { type: 'normal' | 'diff'; content: string }[] = []
          let lastEnd = 0
          let m: RegExpExecArray | null
          while ((m = diffPattern.exec(html)) !== null) {
            const start = m.index
            const end = m.index + m[0].length
            if (start > lastEnd) segments.push({ type: 'normal', content: html.slice(lastEnd, start) })
            segments.push({ type: 'diff', content: m[0] })
            lastEnd = end
          }
          if (lastEnd < html.length) segments.push({ type: 'normal', content: html.slice(lastEnd) })

          const escapeRegex = (str: string) => str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')

          const replaceAllInFragment = (fragment: string) => {
            // 将 HTML 分解为文本与标签，定位纯文本匹配区间
            const parts: { type: 'text' | 'tag'; content: string; index: number }[] = []
            let lastIdx = 0
            const tagRegex = /<[^>]+>/g
            let tagMatch: RegExpExecArray | null
            while ((tagMatch = tagRegex.exec(fragment)) !== null) {
              if (tagMatch.index > lastIdx) {
                parts.push({ type: 'text', content: fragment.slice(lastIdx, tagMatch.index), index: lastIdx })
              }
              parts.push({ type: 'tag', content: tagMatch[0], index: tagMatch.index })
              lastIdx = tagMatch.index + tagMatch[0].length
            }
            if (lastIdx < fragment.length) {
              parts.push({ type: 'text', content: fragment.slice(lastIdx), index: lastIdx })
            }

            let pure = ''
            const map: number[] = []
            for (const p of parts) {
              if (p.type === 'text') {
                for (let i = 0; i < p.content.length; i++) {
                  map.push(p.index + i)
                  pure += p.content[i]
                }
              }
            }
            if (!pure) return { out: fragment, count: 0 }

            const re = new RegExp(escapeRegex(targetText), 'g')
            const matches: { start: number; end: number }[] = []
            let mm: RegExpExecArray | null
            while ((mm = re.exec(pure)) !== null) {
              matches.push({ start: mm.index, end: mm.index + mm[0].length })
            }
            if (matches.length === 0) return { out: fragment, count: 0 }

            let result = fragment
            for (let i = matches.length - 1; i >= 0; i--) {
              const match = matches[i]
              const htmlStart = map[match.start]
              const htmlEnd = map[match.end - 1] + 1
              const originalHtmlFragment = result.slice(htmlStart, htmlEnd)
              const originalText = originalHtmlFragment.replace(/<[^>]+>/g, '')
              const replacement =
                `<span class="diff-old" data-diff-id="${diffId}">${escapeHtml(originalText)}</span>` +
                `<span class="diff-new" data-diff-id="${diffId}">${makeStyled(originalText)}</span>`
              result = result.slice(0, htmlStart) + replacement + result.slice(htmlEnd)
            }
            return { out: result, count: matches.length }
          }

          let count = 0
          const merged = segments.map(seg => {
            if (seg.type === 'diff') return seg.content
            const r = replaceAllInFragment(seg.content)
            count += r.count
            return r.out
          }).join('')

          if (count > 0) {
            html = merged
            doc.body.innerHTML = html
            created.push({
              id: diffId,
              kind: 'format_text',
              scope: 'document',
              summary: '字符格式调整',
              beforePreview: targetText,
              afterPreview: targetText,
              stats: { matches: count },
              timestamp: Date.now(),
              meta: { op },
            })
          }
        }

        // clear_format - 清除格式
        if (type === 'clear_format') {
          const anchor = op.target?.scope === 'anchor_text' ? (op.target?.text || '') : ''
          const scopeType = (op.params?.scope || 'paragraph').toString() as 'selection' | 'paragraph' | 'document'
          
          if (scopeType === 'document') {
            // 清除整个文档的格式
            const diffId = genId()
            const allBlocks = Array.from(doc.body.querySelectorAll<HTMLElement>('p,h1,h2,h3,h4,h5,h6,div,span,strong,em,u,b,i'))
            
            // 对于整个文档，只做一次清理
            const originalHtml = doc.body.innerHTML
            
            // 移除所有 style 属性和格式标签
            const cleanContent = (node: HTMLElement) => {
              // 移除 style 属性
              node.removeAttribute('style')
              
              // 递归处理子节点
              Array.from(node.children).forEach(child => {
                if (child instanceof HTMLElement) {
                  cleanContent(child)
                }
              })
            }
            
            // 替换格式标签为纯文本
            const replaceFormattingTags = (html: string) => {
              return html
                .replace(/<strong>([^<]*)<\/strong>/gi, '$1')
                .replace(/<b>([^<]*)<\/b>/gi, '$1')
                .replace(/<em>([^<]*)<\/em>/gi, '$1')
                .replace(/<i>([^<]*)<\/i>/gi, '$1')
                .replace(/<u>([^<]*)<\/u>/gi, '$1')
                .replace(/<span[^>]*>([^<]*)<\/span>/gi, '$1')
                .replace(/\s*style="[^"]*"/gi, '')
            }
            
            const cleanedHtml = replaceFormattingTags(doc.body.innerHTML)
            doc.body.innerHTML = cleanedHtml
            html = cleanedHtml
            
            created.push({
              id: diffId,
              kind: 'clear_format',
              scope: 'document',
              summary: '清除全文格式',
              beforePreview: '（原格式）',
              afterPreview: '（纯文本）',
              stats: { matches: allBlocks.length },
              timestamp: Date.now(),
              meta: { op },
            })
          } else {
            // 清除特定段落/选区的格式
            const blocks = Array.from(doc.body.querySelectorAll<HTMLElement>('p,h1,h2,h3'))
            const matched = anchor
              ? blocks.filter(b => (b.textContent || '').includes(anchor))
              : blocks
            
            for (const el of matched) {
              const diffId = genId()
              const oldClone = el.cloneNode(true) as HTMLElement
              const newClone = doc.createElement('p')
              
              // 只保留纯文本
              newClone.textContent = el.textContent || ''
              
              markBlockPair(oldClone, newClone, diffId)
              el.replaceWith(oldClone)
              oldClone.insertAdjacentElement('afterend', newClone)
              
              created.push({
                id: diffId,
                kind: 'clear_format',
                scope: anchor ? 'selection' : 'document',
                summary: '清除格式',
                beforePreview: (oldClone.textContent || '').trim().slice(0, 50),
                afterPreview: (newClone.textContent || '').trim().slice(0, 50),
                stats: { matches: 1 },
                timestamp: Date.now(),
                meta: { op },
              })
            }
          }
          continue
        }

        // copy_format - 格式刷
        if (type === 'copy_format') {
          const sourceText = (op.params?.source || '').toString()
          const targetText = (op.params?.target || '').toString()
          
          if (!sourceText || !targetText) continue
          
          // 找到源元素
          const allBlocks = Array.from(doc.body.querySelectorAll<HTMLElement>('p,h1,h2,h3,h4,h5,h6'))
          const sourceEl = allBlocks.find(b => (b.textContent || '').includes(sourceText))
          const targetEls = allBlocks.filter(b => (b.textContent || '').includes(targetText))
          
          if (!sourceEl || targetEls.length === 0) continue
          
          // 获取源元素的样式和标签
          const sourceTag = sourceEl.tagName.toLowerCase()
          const sourceStyle = sourceEl.getAttribute('style') || ''
          
          for (const targetEl of targetEls) {
            const diffId = genId()
            const oldClone = targetEl.cloneNode(true) as HTMLElement
            
            // 创建与源相同标签的新元素
            const newClone = doc.createElement(sourceTag)
            newClone.innerHTML = targetEl.innerHTML
            if (sourceStyle) {
              newClone.setAttribute('style', sourceStyle)
            }
            
            markBlockPair(oldClone, newClone, diffId)
            targetEl.replaceWith(oldClone)
            oldClone.insertAdjacentElement('afterend', newClone)
            
            created.push({
              id: diffId,
              kind: 'copy_format',
              scope: 'selection',
              summary: `复制格式: ${sourceText.slice(0, 20)} → ${targetText.slice(0, 20)}`,
              beforePreview: (oldClone.textContent || '').trim().slice(0, 50),
              afterPreview: (newClone.textContent || '').trim().slice(0, 50),
              stats: { matches: 1 },
              timestamp: Date.now(),
              meta: { op, sourceTag, sourceStyle },
            })
          }
          continue
        }

        // list_edit - 列表操作
        if (type === 'list_edit') {
          const action = (op.params?.action || '').toString() as 'to_ordered_list' | 'to_unordered_list' | 'remove_list'
          const anchor = (op.params?.anchor || op.target?.text || '').toString()
          
          if (!action) continue
          
          // 找到包含锚点文本的段落
          const allBlocks = Array.from(doc.body.querySelectorAll<HTMLElement>('p,li,h1,h2,h3'))
          const matchedBlocks = anchor
            ? allBlocks.filter(b => (b.textContent || '').includes(anchor))
            : allBlocks.filter(b => b.tagName.toLowerCase() === 'p')
          
          if (matchedBlocks.length === 0) continue
          
          const diffId = genId()
          
          if (action === 'to_ordered_list' || action === 'to_unordered_list') {
            const listTag = action === 'to_ordered_list' ? 'ol' : 'ul'
            const list = doc.createElement(listTag)
            
            // 收集连续的段落转为列表项
            for (const block of matchedBlocks) {
              const li = doc.createElement('li')
              li.innerHTML = block.innerHTML
              list.appendChild(li)
            }
            
            // 标记为 diff
            list.setAttribute('data-diff-id', diffId)
            list.setAttribute('data-diff-role', 'new')
            list.setAttribute('data-diff-kind', 'block')
            
            // 替换第一个匹配的元素，删除其余的
            const firstBlock = matchedBlocks[0]
            const oldClone = firstBlock.cloneNode(true) as HTMLElement
            oldClone.setAttribute('data-diff-id', diffId)
            oldClone.setAttribute('data-diff-role', 'old')
            oldClone.setAttribute('data-diff-kind', 'block')
            
            firstBlock.replaceWith(oldClone)
            oldClone.insertAdjacentElement('afterend', list)
            
            // 移除其余的段落
            for (let i = 1; i < matchedBlocks.length; i++) {
              matchedBlocks[i].remove()
            }
            
            created.push({
              id: diffId,
              kind: 'list_edit',
              scope: 'selection',
              summary: action === 'to_ordered_list' ? '转为有序列表' : '转为无序列表',
              beforePreview: (oldClone.textContent || '').trim().slice(0, 50),
              afterPreview: `(${matchedBlocks.length} 项列表)`,
              stats: { matches: matchedBlocks.length },
              timestamp: Date.now(),
              meta: { op, action },
            })
          } else if (action === 'remove_list') {
            // 找到列表并转为段落
            const lists = Array.from(doc.body.querySelectorAll<HTMLElement>('ul,ol'))
            const targetLists = anchor
              ? lists.filter(l => (l.textContent || '').includes(anchor))
              : lists
            
            for (const list of targetLists) {
              const items = Array.from(list.querySelectorAll('li'))
              const oldClone = list.cloneNode(true) as HTMLElement
              oldClone.setAttribute('data-diff-id', diffId)
              oldClone.setAttribute('data-diff-role', 'old')
              oldClone.setAttribute('data-diff-kind', 'block')
              
              // 创建段落替换列表
              const container = doc.createDocumentFragment()
              for (const item of items) {
                const p = doc.createElement('p')
                p.innerHTML = item.innerHTML
                p.setAttribute('data-diff-id', diffId)
                p.setAttribute('data-diff-role', 'new')
                p.setAttribute('data-diff-kind', 'block')
                container.appendChild(p)
              }
              
              list.replaceWith(oldClone)
              oldClone.insertAdjacentElement('afterend', container.firstElementChild as Element)
              
              created.push({
                id: diffId,
                kind: 'list_edit',
                scope: 'selection',
                summary: '取消列表格式',
                beforePreview: `(${items.length} 项列表)`,
                afterPreview: `(${items.length} 个段落)`,
                stats: { matches: items.length },
                timestamp: Date.now(),
                meta: { op, action },
              })
            }
          }
          continue
        }

        // insert_page_break - 插入分页符
        if (type === 'insert_page_break') {
          const position = (op.params?.position || op.target?.text || '').toString()
          
          // 创建分页符元素
          const pageBreak = doc.createElement('div')
          pageBreak.className = 'page-break'
          pageBreak.setAttribute('style', 'page-break-before: always; border-top: 2px dashed var(--word-rule); margin: 20px 0; padding: 10px 0; text-align: center; color: var(--word-ink-muted); font-size: 12px;')
          pageBreak.textContent = '--- 分页符 ---'
          
          const diffId = genId()
          pageBreak.setAttribute('data-diff-id', diffId)
          pageBreak.setAttribute('data-diff-role', 'new')
          pageBreak.setAttribute('data-diff-kind', 'block')
          
          if (position.startsWith('before:')) {
            const anchorText = position.slice(7)
            const allBlocks = Array.from(doc.body.querySelectorAll<HTMLElement>('p,h1,h2,h3,h4,h5,h6'))
            const targetBlock = allBlocks.find(b => (b.textContent || '').includes(anchorText))
            
            if (targetBlock) {
              targetBlock.insertAdjacentElement('beforebegin', pageBreak)
              
              created.push({
                id: diffId,
                kind: 'insert_page_break',
                scope: 'selection',
                summary: `在"${anchorText.slice(0, 20)}"前插入分页符`,
                afterPreview: '分页符',
                stats: { matches: 1 },
                timestamp: Date.now(),
                meta: { op },
              })
            }
          } else if (position.startsWith('after:')) {
            const anchorText = position.slice(6)
            const allBlocks = Array.from(doc.body.querySelectorAll<HTMLElement>('p,h1,h2,h3,h4,h5,h6'))
            const targetBlock = allBlocks.find(b => (b.textContent || '').includes(anchorText))
            
            if (targetBlock) {
              targetBlock.insertAdjacentElement('afterend', pageBreak)
              
              created.push({
                id: diffId,
                kind: 'insert_page_break',
                scope: 'selection',
                summary: `在"${anchorText.slice(0, 20)}"后插入分页符`,
                afterPreview: '分页符',
                stats: { matches: 1 },
                timestamp: Date.now(),
                meta: { op },
              })
            }
          }
          continue
        }

        // structure_edit - 结构编辑（移动段落）
        if (type === 'structure_edit') {
          const action = (op.params?.action || '').toString() as 'move_block' | 'extract_outline'
          
          if (action === 'move_block') {
            const sourceText = (op.params?.source || '').toString()
            const targetPosition = (op.params?.target || '').toString()
            
            if (!sourceText || !targetPosition) continue
            
            const allBlocks = Array.from(doc.body.querySelectorAll<HTMLElement>('p,h1,h2,h3,h4,h5,h6'))
            const sourceBlock = allBlocks.find(b => (b.textContent || '').includes(sourceText))
            
            if (!sourceBlock) continue
            
            const diffId = genId()
            const movedClone = sourceBlock.cloneNode(true) as HTMLElement
            movedClone.setAttribute('data-diff-id', diffId)
            movedClone.setAttribute('data-diff-role', 'new')
            movedClone.setAttribute('data-diff-kind', 'block')
            
            // 标记原位置为删除
            sourceBlock.setAttribute('data-diff-id', diffId)
            sourceBlock.setAttribute('data-diff-role', 'old')
            sourceBlock.setAttribute('data-diff-kind', 'block')
            
            if (targetPosition.startsWith('before:')) {
              const targetText = targetPosition.slice(7)
              const targetBlock = allBlocks.find(b => (b.textContent || '').includes(targetText))
              if (targetBlock) {
                targetBlock.insertAdjacentElement('beforebegin', movedClone)
                
                created.push({
                  id: diffId,
                  kind: 'structure_edit',
                  scope: 'document',
                  summary: `移动段落到"${targetText.slice(0, 15)}"前`,
                  beforePreview: sourceText.slice(0, 50),
                  afterPreview: `移至"${targetText.slice(0, 15)}"前`,
                  stats: { matches: 1 },
                  timestamp: Date.now(),
                  meta: { op, action },
                })
              }
            } else if (targetPosition.startsWith('after:')) {
              const targetText = targetPosition.slice(6)
              const targetBlock = allBlocks.find(b => (b.textContent || '').includes(targetText))
              if (targetBlock) {
                targetBlock.insertAdjacentElement('afterend', movedClone)
                
                created.push({
                  id: diffId,
                  kind: 'structure_edit',
                  scope: 'document',
                  summary: `移动段落到"${targetText.slice(0, 15)}"后`,
                  beforePreview: sourceText.slice(0, 50),
                  afterPreview: `移至"${targetText.slice(0, 15)}"后`,
                  stats: { matches: 1 },
                  timestamp: Date.now(),
                  meta: { op, action },
                })
              }
            }
          } else if (action === 'extract_outline') {
            // 提取大纲不修改文档，只返回信息
            const headings = Array.from(doc.body.querySelectorAll<HTMLElement>('h1,h2,h3,h4,h5,h6'))
            const outline = headings.map(h => ({
              level: parseInt(h.tagName[1]),
              text: (h.textContent || '').trim()
            }))
            
            // 通过 meta 返回大纲信息
            created.push({
              id: genId(),
              kind: 'structure_edit',
              scope: 'document',
              summary: `提取大纲：${headings.length} 个标题`,
              afterPreview: outline.map(o => `${'  '.repeat(o.level - 1)}${o.text}`).join('\n').slice(0, 200),
              stats: { matches: headings.length },
              timestamp: Date.now(),
              meta: { op, action, outline },
            })
          }
          continue
        }

        // table_edit - 表格操作
        if (type === 'table_edit') {
          const action = (op.params?.action || '').toString()
          const tableAnchor = (op.params?.tableAnchor || '').toString()
          
          // 找到目标表格
          const tables = Array.from(doc.body.querySelectorAll<HTMLTableElement>('table'))
          const targetTable = tableAnchor
            ? tables.find(t => (t.textContent || '').includes(tableAnchor))
            : tables[0]
          
          if (action === 'insert_table') {
            const position = (op.params?.position || '').toString()
            const rows = parseInt(String(op.params?.rows || 3))
            const cols = parseInt(String(op.params?.cols || 3))
            const headers = op.params?.headers as string[] | undefined
            
            // 创建新表格
            const table = doc.createElement('table')
            table.setAttribute('style', 'border-collapse: collapse; width: 100%; margin: 10px 0;')
            
            const diffId = genId()
            table.setAttribute('data-diff-id', diffId)
            table.setAttribute('data-diff-role', 'new')
            table.setAttribute('data-diff-kind', 'block')
            
            for (let r = 0; r < rows; r++) {
              const tr = doc.createElement('tr')
              for (let c = 0; c < cols; c++) {
                const cell = doc.createElement(r === 0 && headers ? 'th' : 'td')
                cell.setAttribute('style', 'border: 0.5pt solid var(--word-rule); padding: 2pt 5pt;')
                if (r === 0 && headers && headers[c]) {
                  cell.textContent = headers[c]
                }
                tr.appendChild(cell)
              }
              table.appendChild(tr)
            }
            
            // 插入表格
            if (position.startsWith('after:')) {
              const anchorText = position.slice(6)
              const allBlocks = Array.from(doc.body.querySelectorAll<HTMLElement>('p,h1,h2,h3'))
              const targetBlock = allBlocks.find(b => (b.textContent || '').includes(anchorText))
              if (targetBlock) {
                targetBlock.insertAdjacentElement('afterend', table)
              }
            } else {
              doc.body.appendChild(table)
            }
            
            created.push({
              id: diffId,
              kind: 'table_edit',
              scope: 'document',
              summary: `插入 ${rows}×${cols} 表格`,
              afterPreview: headers ? headers.join(', ') : `${rows}行${cols}列表格`,
              stats: { matches: 1 },
              timestamp: Date.now(),
              meta: { op, action },
            })
          } else if (action === 'add_row' && targetTable) {
            const count = parseInt(String(op.params?.count || 1))
            const diffId = genId()
            
            const cols = targetTable.rows[0]?.cells.length || 3
            for (let i = 0; i < count; i++) {
              const tr = doc.createElement('tr')
              tr.setAttribute('data-diff-id', diffId)
              tr.setAttribute('data-diff-role', 'new')
              for (let c = 0; c < cols; c++) {
                const td = doc.createElement('td')
                td.setAttribute('style', 'border: 0.5pt solid var(--word-rule); padding: 2pt 5pt;')
                tr.appendChild(td)
              }
              targetTable.appendChild(tr)
            }
            
            created.push({
              id: diffId,
              kind: 'table_edit',
              scope: 'selection',
              summary: `添加 ${count} 行`,
              afterPreview: `新增 ${count} 行`,
              stats: { matches: count },
              timestamp: Date.now(),
              meta: { op, action },
            })
          } else if (action === 'add_column' && targetTable) {
            const count = parseInt(String(op.params?.count || 1))
            const diffId = genId()
            
            const rows = targetTable.rows
            for (let r = 0; r < rows.length; r++) {
              for (let i = 0; i < count; i++) {
                const cell = doc.createElement(r === 0 ? 'th' : 'td')
                cell.setAttribute('style', 'border: 0.5pt solid var(--word-rule); padding: 2pt 5pt;')
                cell.setAttribute('data-diff-id', diffId)
                cell.setAttribute('data-diff-role', 'new')
                rows[r].appendChild(cell)
              }
            }
            
            created.push({
              id: diffId,
              kind: 'table_edit',
              scope: 'selection',
              summary: `添加 ${count} 列`,
              afterPreview: `新增 ${count} 列`,
              stats: { matches: count },
              timestamp: Date.now(),
              meta: { op, action },
            })
          } else if (action === 'delete_row' && targetTable) {
            // 删除行
            const rowIndex = parseInt(String(op.params?.rowIndex ?? op.params?.row ?? -1))
            const count = parseInt(String(op.params?.count || 1))
            const diffId = genId()
            
            if (rowIndex >= 0 && rowIndex < targetTable.rows.length) {
              for (let i = 0; i < count && rowIndex < targetTable.rows.length; i++) {
                const row = targetTable.rows[rowIndex]
                row.setAttribute('data-diff-id', diffId)
                row.setAttribute('data-diff-role', 'old')
                row.style.textDecoration = 'line-through'
                row.style.opacity = '0.5'
              }
              
              created.push({
                id: diffId,
                kind: 'table_edit',
                scope: 'selection',
                summary: `删除第 ${rowIndex + 1} 行`,
                beforePreview: `第 ${rowIndex + 1} 行`,
                stats: { matches: count },
                timestamp: Date.now(),
                meta: { op, action },
              })
            }
          } else if (action === 'delete_column' && targetTable) {
            // 删除列
            const colIndex = parseInt(String(op.params?.colIndex ?? op.params?.column ?? -1))
            const diffId = genId()
            
            if (colIndex >= 0) {
              const rows = targetTable.rows
              let deletedCount = 0
              for (let r = 0; r < rows.length; r++) {
                if (colIndex < rows[r].cells.length) {
                  const cell = rows[r].cells[colIndex]
                  cell.setAttribute('data-diff-id', diffId)
                  cell.setAttribute('data-diff-role', 'old')
                  cell.style.textDecoration = 'line-through'
                  cell.style.opacity = '0.5'
                  deletedCount++
                }
              }
              
              created.push({
                id: diffId,
                kind: 'table_edit',
                scope: 'selection',
                summary: `删除第 ${colIndex + 1} 列`,
                beforePreview: `第 ${colIndex + 1} 列`,
                stats: { matches: deletedCount },
                timestamp: Date.now(),
                meta: { op, action },
              })
            }
          } else if (action === 'merge_cells' && targetTable) {
            // 合并单元格
            const startRow = parseInt(String(op.params?.startRow ?? 0))
            const startCol = parseInt(String(op.params?.startCol ?? 0))
            const rowSpan = parseInt(String(op.params?.rowSpan ?? 1))
            const colSpan = parseInt(String(op.params?.colSpan ?? 1))
            const diffId = genId()
            
            if (startRow < targetTable.rows.length && startCol < (targetTable.rows[startRow]?.cells.length || 0)) {
              const mainCell = targetTable.rows[startRow].cells[startCol]
              
              // 设置合并属性
              if (rowSpan > 1) mainCell.setAttribute('rowspan', String(rowSpan))
              if (colSpan > 1) mainCell.setAttribute('colspan', String(colSpan))
              mainCell.setAttribute('data-diff-id', diffId)
              mainCell.setAttribute('data-diff-role', 'new')
              
              // 标记要隐藏的单元格
              for (let r = startRow; r < startRow + rowSpan && r < targetTable.rows.length; r++) {
                for (let c = startCol; c < startCol + colSpan; c++) {
                  if (r === startRow && c === startCol) continue
                  if (c < targetTable.rows[r].cells.length) {
                    const cell = targetTable.rows[r].cells[c]
                    cell.setAttribute('data-diff-id', diffId)
                    cell.setAttribute('data-diff-role', 'old')
                    cell.style.display = 'none'
                  }
                }
              }
              
              created.push({
                id: diffId,
                kind: 'table_edit',
                scope: 'selection',
                summary: `合并单元格 (${rowSpan}x${colSpan})`,
                afterPreview: `从 (${startRow + 1},${startCol + 1}) 合并 ${rowSpan}行${colSpan}列`,
                stats: { matches: 1 },
                timestamp: Date.now(),
                meta: { op, action },
              })
            }
          } else if (action === 'set_cell_style' && targetTable) {
            // 设置单元格样式（边框、背景色等）
            const rowIndex = parseInt(String(op.params?.row ?? -1))
            const colIndex = parseInt(String(op.params?.col ?? -1))
            const backgroundColor = op.params?.backgroundColor as string | undefined
            const borderColor = op.params?.borderColor as string | undefined
            const borderWidth = op.params?.borderWidth as string | undefined
            const borderStyle = op.params?.borderStyle as string | undefined
            const diffId = genId()
            
            // 确定目标单元格（单个或范围）
            const targetCells: HTMLTableCellElement[] = []
            if (rowIndex >= 0 && colIndex >= 0) {
              // 单个单元格
              if (rowIndex < targetTable.rows.length && colIndex < targetTable.rows[rowIndex].cells.length) {
                targetCells.push(targetTable.rows[rowIndex].cells[colIndex])
              }
            } else if (rowIndex >= 0) {
              // 整行
              if (rowIndex < targetTable.rows.length) {
                targetCells.push(...Array.from(targetTable.rows[rowIndex].cells))
              }
            } else if (colIndex >= 0) {
              // 整列
              for (let r = 0; r < targetTable.rows.length; r++) {
                if (colIndex < targetTable.rows[r].cells.length) {
                  targetCells.push(targetTable.rows[r].cells[colIndex])
                }
              }
            }
            
            for (const cell of targetCells) {
              const prevStyle = cell.getAttribute('style') || ''
              let newStyle = prevStyle
              
              if (backgroundColor) {
                newStyle = newStyle.replace(/background-color\s*:\s*[^;]+;?/gi, '')
                newStyle += `; background-color: ${backgroundColor};`
              }
              if (borderColor || borderWidth || borderStyle) {
                const bw = borderWidth || '1px'
                const bs = borderStyle || 'solid'
                const bc = borderColor || '#000000'
                newStyle = newStyle.replace(/border\s*:\s*[^;]+;?/gi, '')
                newStyle += `; border: ${bw} ${bs} ${bc};`
              }
              
              cell.setAttribute('style', newStyle.replace(/^;?\s*/, ''))
              cell.setAttribute('data-diff-id', diffId)
              cell.setAttribute('data-diff-role', 'new')
            }
            
            if (targetCells.length > 0) {
              created.push({
                id: diffId,
                kind: 'table_edit',
                scope: 'selection',
                summary: `设置 ${targetCells.length} 个单元格样式`,
                afterPreview: backgroundColor ? `背景: ${backgroundColor}` : '边框样式',
                stats: { matches: targetCells.length },
                timestamp: Date.now(),
                meta: { op, action },
              })
            }
          } else if (action === 'set_table_border' && targetTable) {
            // 设置整个表格边框
            const borderColor = (op.params?.borderColor || '#000000').toString()
            const borderWidth = (op.params?.borderWidth || '1px').toString()
            const borderStyle = (op.params?.borderStyle || 'solid').toString()
            const diffId = genId()
            
            const allCells = Array.from(targetTable.querySelectorAll<HTMLTableCellElement>('td, th'))
            for (const cell of allCells) {
              const prevStyle = cell.getAttribute('style') || ''
              let newStyle = prevStyle.replace(/border\s*:\s*[^;]+;?/gi, '')
              newStyle += `; border: ${borderWidth} ${borderStyle} ${borderColor};`
              cell.setAttribute('style', newStyle.replace(/^;?\s*/, ''))
              cell.setAttribute('data-diff-id', diffId)
              cell.setAttribute('data-diff-role', 'new')
            }
            
            created.push({
              id: diffId,
              kind: 'table_edit',
              scope: 'selection',
              summary: `设置表格边框`,
              afterPreview: `${borderWidth} ${borderStyle} ${borderColor}`,
              stats: { matches: allCells.length },
              timestamp: Date.now(),
              meta: { op, action },
            })
          }
          continue
        }

        // image_edit - 图片操作
        if (type === 'image_edit') {
          const action = (op.params?.action || '').toString()
          
          if (action === 'insert_image') {
            const position = (op.params?.position || '').toString()
            const url = (op.params?.url || '').toString()
            const width = (op.params?.width || '300px').toString()
            const alignment = (op.params?.alignment || 'center').toString()
            
            if (!url) continue
            
            const diffId = genId()
            
            // 创建图片容器
            const container = doc.createElement('p')
            container.setAttribute('style', `text-align: ${alignment};`)
            container.setAttribute('data-diff-id', diffId)
            container.setAttribute('data-diff-role', 'new')
            container.setAttribute('data-diff-kind', 'block')
            
            const img = doc.createElement('img')
            img.setAttribute('src', url)
            img.setAttribute('style', `max-width: ${width}; height: auto;`)
            img.setAttribute('alt', '插入的图片')
            
            container.appendChild(img)
            
            // 插入图片
            if (position.startsWith('after:')) {
              const anchorText = position.slice(6)
              const allBlocks = Array.from(doc.body.querySelectorAll<HTMLElement>('p,h1,h2,h3'))
              const targetBlock = allBlocks.find(b => (b.textContent || '').includes(anchorText))
              if (targetBlock) {
                targetBlock.insertAdjacentElement('afterend', container)
              }
            } else {
              doc.body.appendChild(container)
            }
            
            created.push({
              id: diffId,
              kind: 'image_edit',
              scope: 'document',
              summary: '插入图片',
              afterPreview: url.slice(0, 50),
              stats: { matches: 1 },
              timestamp: Date.now(),
              meta: { op, action },
            })
          } else if (action === 'resize_image') {
            const anchor = (op.params?.anchor || '').toString()
            const newWidth = (op.params?.width || '').toString()
            const newHeight = (op.params?.height || '').toString()
            
            if (!newWidth && !newHeight) continue
            
            const images = Array.from(doc.body.querySelectorAll<HTMLImageElement>('img'))
            // 找到最近的图片（基于锚点或第一张）
            const targetImg = anchor
              ? images.find(img => {
                  const parent = img.parentElement
                  return parent && (parent.textContent || '').includes(anchor)
                })
              : images[0]
            
            if (targetImg) {
              const diffId = genId()
              let prevStyle = targetImg.getAttribute('style') || ''
              if (newWidth) {
                prevStyle = prevStyle.replace(/max-width\s*:\s*[^;]+;?/gi, '').replace(/width\s*:\s*[^;]+;?/gi, '')
                prevStyle += `; max-width: ${newWidth}; width: ${newWidth};`
              }
              if (newHeight) {
                prevStyle = prevStyle.replace(/height\s*:\s*[^;]+;?/gi, '')
                prevStyle += `; height: ${newHeight};`
              }
              
              targetImg.setAttribute('data-diff-id', diffId)
              targetImg.setAttribute('data-diff-role', 'new')
              targetImg.setAttribute('style', prevStyle.replace(/^;?\s*/, ''))
              
              created.push({
                id: diffId,
                kind: 'image_edit',
                scope: 'selection',
                summary: `调整图片大小`,
                afterPreview: newWidth ? `宽: ${newWidth}` : `高: ${newHeight}`,
                stats: { matches: 1 },
                timestamp: Date.now(),
                meta: { op, action },
              })
            }
          } else if (action === 'set_caption') {
            // 设置图片标题/说明
            const anchor = (op.params?.anchor || '').toString()
            const caption = (op.params?.caption || '').toString()
            
            if (!caption) continue
            
            const images = Array.from(doc.body.querySelectorAll<HTMLImageElement>('img'))
            const targetImg = anchor
              ? images.find(img => {
                  const parent = img.parentElement
                  return parent && (parent.textContent || '').includes(anchor)
                })
              : images[0]
            
            if (targetImg) {
              const diffId = genId()
              const parent = targetImg.parentElement
              
              // 检查是否已经在 figure 中
              if (parent?.tagName.toLowerCase() === 'figure') {
                // 更新或添加 figcaption
                let figcaption = parent.querySelector('figcaption')
                if (!figcaption) {
                  figcaption = doc.createElement('figcaption')
                  figcaption.setAttribute('style', 'font-size: 0.9em; color: #666; text-align: center; margin-top: 0.5em;')
                  parent.appendChild(figcaption)
                }
                figcaption.textContent = caption
                figcaption.setAttribute('data-diff-id', diffId)
                figcaption.setAttribute('data-diff-role', 'new')
              } else if (parent) {
                // 包装到 figure 中
                const figure = doc.createElement('figure')
                figure.setAttribute('style', 'margin: 1em 0; text-align: center;')
                figure.setAttribute('data-diff-id', diffId)
                figure.setAttribute('data-diff-role', 'new')
                figure.setAttribute('data-diff-kind', 'block')
                
                parent.insertBefore(figure, targetImg)
                figure.appendChild(targetImg)
                
                const figcaption = doc.createElement('figcaption')
                figcaption.setAttribute('style', 'font-size: 0.9em; color: #666; text-align: center; margin-top: 0.5em;')
                figcaption.textContent = caption
                figure.appendChild(figcaption)
              }
              
              created.push({
                id: diffId,
                kind: 'image_edit',
                scope: 'selection',
                summary: `设置图片标题`,
                afterPreview: caption.slice(0, 50),
                stats: { matches: 1 },
                timestamp: Date.now(),
                meta: { op, action },
              })
            }
          } else if (action === 'delete_image') {
            // 删除图片
            const anchor = (op.params?.anchor || '').toString()
            const index = parseInt(String(op.params?.index ?? -1))
            
            const images = Array.from(doc.body.querySelectorAll<HTMLImageElement>('img'))
            let targetImg: HTMLImageElement | undefined
            
            if (index >= 0 && index < images.length) {
              targetImg = images[index]
            } else if (anchor) {
              targetImg = images.find(img => {
                const parent = img.parentElement
                return parent && (parent.textContent || '').includes(anchor)
              })
            } else {
              targetImg = images[0]
            }
            
            if (targetImg) {
              const diffId = genId()
              const parent = targetImg.parentElement
              
              // 标记为删除
              if (parent?.tagName.toLowerCase() === 'figure') {
                parent.setAttribute('data-diff-id', diffId)
                parent.setAttribute('data-diff-role', 'old')
                parent.style.opacity = '0.3'
                parent.style.textDecoration = 'line-through'
              } else {
                targetImg.setAttribute('data-diff-id', diffId)
                targetImg.setAttribute('data-diff-role', 'old')
                targetImg.style.opacity = '0.3'
              }
              
              created.push({
                id: diffId,
                kind: 'image_edit',
                scope: 'selection',
                summary: `删除图片`,
                beforePreview: targetImg.alt || '图片',
                stats: { matches: 1 },
                timestamp: Date.now(),
                meta: { op, action },
              })
            }
          } else if (action === 'set_alignment') {
            // 设置图片对齐方式
            const anchor = (op.params?.anchor || '').toString()
            const alignment = (op.params?.alignment || 'center').toString()
            
            const images = Array.from(doc.body.querySelectorAll<HTMLImageElement>('img'))
            const targetImg = anchor
              ? images.find(img => {
                  const parent = img.parentElement
                  return parent && (parent.textContent || '').includes(anchor)
                })
              : images[0]
            
            if (targetImg) {
              const diffId = genId()
              const parent = targetImg.parentElement
              
              if (parent) {
                const prevStyle = parent.getAttribute('style') || ''
                let newStyle = prevStyle.replace(/text-align\s*:\s*[^;]+;?/gi, '')
                newStyle += `; text-align: ${alignment};`
                parent.setAttribute('style', newStyle.replace(/^;?\s*/, ''))
                parent.setAttribute('data-diff-id', diffId)
                parent.setAttribute('data-diff-role', 'new')
                
                created.push({
                  id: diffId,
                  kind: 'image_edit',
                  scope: 'selection',
                  summary: `设置图片对齐: ${alignment}`,
                  afterPreview: alignment,
                  stats: { matches: 1 },
                  timestamp: Date.now(),
                  meta: { op, action },
                })
              }
            }
          }
          continue
        }

        // page_setup - 页面设置（不修改文档内容，只修改页面设置状态）
        if (type === 'page_setup') {
          const diffId = genId()
          
          const newSetup: Partial<PageSetup> = {}
          
          // 纸张大小
          const paperSize = op.params?.paperSize as string
          if (paperSize && ['A4', 'A3', 'Letter', 'Legal', 'custom'].includes(paperSize)) {
            newSetup.paperSize = paperSize as PageSetup['paperSize']
          }
          
          // 页面方向
          const orientation = op.params?.orientation as string
          if (orientation && ['portrait', 'landscape'].includes(orientation)) {
            newSetup.orientation = orientation as PageSetup['orientation']
          }
          
          // 页边距
          const margins = op.params?.margins as { top?: string; bottom?: string; left?: string; right?: string } | undefined
          if (margins && typeof margins === 'object') {
            newSetup.margins = {
              top: margins.top || pageSetup.margins.top,
              bottom: margins.bottom || pageSetup.margins.bottom,
              left: margins.left || pageSetup.margins.left,
              right: margins.right || pageSetup.margins.right,
            }
          }
          
          // 自定义尺寸
          if (op.params?.customWidth) newSetup.customWidth = String(op.params.customWidth)
          if (op.params?.customHeight) newSetup.customHeight = String(op.params.customHeight)
          
          if (Object.keys(newSetup).length > 0) {
            setPageSetupState(prev => ({ ...prev, ...newSetup }))
            
            const changes: string[] = []
            if (newSetup.paperSize) changes.push(`纸张: ${newSetup.paperSize}`)
            if (newSetup.orientation) changes.push(`方向: ${newSetup.orientation === 'portrait' ? '纵向' : '横向'}`)
            if (newSetup.margins) changes.push('边距已更新')
            
            created.push({
              id: diffId,
              kind: 'page_setup',
              scope: 'document',
              summary: `页面设置: ${changes.join(', ')}`,
              afterPreview: changes.join(', '),
              stats: { matches: 1 },
              timestamp: Date.now(),
              meta: { op },
            })
          }
          continue
        }

        // define_style - 定义新样式
        if (type === 'define_style') {
          const diffId = genId()
          const styleName = String(op.params?.name || '')
          if (!styleName) continue
          
          const newStyle: CustomStyle = {
            name: styleName,
            basedOn: op.params?.basedOn as string | undefined,
            fontFamily: op.params?.fontFamily as string | undefined,
            fontSize: op.params?.fontSize as string | undefined,
            color: op.params?.color as string | undefined,
            bold: op.params?.bold as boolean | undefined,
            italic: op.params?.italic as boolean | undefined,
            underline: op.params?.underline as boolean | undefined,
            strikethrough: op.params?.strikethrough as boolean | undefined,
            letterSpacing: op.params?.letterSpacing as string | undefined,
            alignment: op.params?.alignment as 'left' | 'center' | 'right' | 'justify' | undefined,
            lineHeight: op.params?.lineHeight as string | undefined,
            spaceBefore: op.params?.spaceBefore as string | undefined,
            spaceAfter: op.params?.spaceAfter as string | undefined,
            textIndent: op.params?.textIndent as string | undefined,
            marginLeft: op.params?.marginLeft as string | undefined,
            marginRight: op.params?.marginRight as string | undefined,
            backgroundColor: op.params?.backgroundColor as string | undefined,
            border: op.params?.border as string | undefined,
          }
          
          // 如果基于其他样式继承
          if (newStyle.basedOn && customStyles[newStyle.basedOn]) {
            const baseStyle = customStyles[newStyle.basedOn]
            const styleKeys: (keyof CustomStyle)[] = [
              'fontFamily', 'fontSize', 'color', 'bold', 'italic', 'underline',
              'strikethrough', 'letterSpacing', 'alignment', 'lineHeight',
              'spaceBefore', 'spaceAfter', 'textIndent', 'marginLeft', 'marginRight',
              'backgroundColor', 'border'
            ]
            styleKeys.forEach(key => {
              if (newStyle[key] === undefined && baseStyle[key] !== undefined) {
                (newStyle[key] as typeof baseStyle[typeof key]) = baseStyle[key]
              }
            })
          }
          
          setCustomStyles(prev => ({ ...prev, [styleName]: newStyle }))
          
          created.push({
            id: diffId,
            kind: 'define_style',
            scope: 'document',
            summary: `定义样式: ${styleName}`,
            afterPreview: styleName,
            stats: { matches: 1 },
            timestamp: Date.now(),
            meta: { op },
          })
          continue
        }

        // modify_style - 修改现有样式
        if (type === 'modify_style') {
          const diffId = genId()
          const styleName = String(op.params?.name || '')
          if (!styleName || !customStyles[styleName]) continue
          
          const updates: Partial<CustomStyle> = {}
          if (op.params?.fontFamily !== undefined) updates.fontFamily = String(op.params.fontFamily)
          if (op.params?.fontSize !== undefined) updates.fontSize = String(op.params.fontSize)
          if (op.params?.color !== undefined) updates.color = String(op.params.color)
          if (op.params?.bold !== undefined) updates.bold = Boolean(op.params.bold)
          if (op.params?.italic !== undefined) updates.italic = Boolean(op.params.italic)
          if (op.params?.underline !== undefined) updates.underline = Boolean(op.params.underline)
          if (op.params?.strikethrough !== undefined) updates.strikethrough = Boolean(op.params.strikethrough)
          if (op.params?.letterSpacing !== undefined) updates.letterSpacing = String(op.params.letterSpacing)
          if (op.params?.alignment !== undefined) updates.alignment = op.params.alignment as CustomStyle['alignment']
          if (op.params?.lineHeight !== undefined) updates.lineHeight = String(op.params.lineHeight)
          if (op.params?.spaceBefore !== undefined) updates.spaceBefore = String(op.params.spaceBefore)
          if (op.params?.spaceAfter !== undefined) updates.spaceAfter = String(op.params.spaceAfter)
          if (op.params?.textIndent !== undefined) updates.textIndent = String(op.params.textIndent)
          if (op.params?.marginLeft !== undefined) updates.marginLeft = String(op.params.marginLeft)
          if (op.params?.marginRight !== undefined) updates.marginRight = String(op.params.marginRight)
          if (op.params?.backgroundColor !== undefined) updates.backgroundColor = String(op.params.backgroundColor)
          if (op.params?.border !== undefined) updates.border = String(op.params.border)
          
          setCustomStyles(prev => ({
            ...prev,
            [styleName]: { ...prev[styleName], ...updates }
          }))
          
          const changeList = Object.keys(updates).join(', ')
          created.push({
            id: diffId,
            kind: 'modify_style',
            scope: 'document',
            summary: `修改样式 ${styleName}: ${changeList}`,
            afterPreview: changeList,
            stats: { matches: 1 },
            timestamp: Date.now(),
            meta: { op },
          })
          continue
        }

        // columns - 分栏排版
        if (type === 'columns') {
          const diffId = genId()
          const columnCount = Number(op.params?.count) || 2
          const columnGap = String(op.params?.gap || '2em')
          const columnRule = op.params?.rule ? String(op.params.rule) : ''
          
          // 分栏通过 CSS multi-column 实现
          // 需要包裹整个内容或选定区域
          const anchor = (op.target?.text || '').toString()
          const blocks = anchor
            ? Array.from(doc.body.querySelectorAll<HTMLElement>('p,h1,h2,h3')).filter(b => (b.textContent || '').includes(anchor))
            : Array.from(doc.body.querySelectorAll<HTMLElement>('p,h1,h2,h3'))
          
          if (blocks.length > 0) {
            // 创建分栏容器
            const columnContainer = doc.createElement('div')
            columnContainer.className = 'column-layout'
            let columnStyle = `column-count: ${columnCount}; column-gap: ${columnGap};`
            if (columnRule) columnStyle += ` column-rule: ${columnRule};`
            columnContainer.setAttribute('style', columnStyle)
            columnContainer.setAttribute('data-diff-id', diffId)
            columnContainer.setAttribute('data-diff-role', 'new')
            columnContainer.setAttribute('data-diff-kind', 'block')
            
            // 将匹配的块移入分栏容器
            const firstBlock = blocks[0]
            firstBlock.parentNode?.insertBefore(columnContainer, firstBlock)
            blocks.forEach(block => {
              columnContainer.appendChild(block.cloneNode(true))
              block.remove()
            })
            
            created.push({
              id: diffId,
              kind: 'columns',
              scope: 'document',
              summary: `设置 ${columnCount} 栏排版`,
              afterPreview: `${columnCount} 栏，间距 ${columnGap}`,
              stats: { matches: blocks.length },
              timestamp: Date.now(),
              meta: { op },
            })
          }
          continue
        }

        // watermark - 水印
        if (type === 'watermark') {
          const diffId = genId()
          const text = String(op.params?.text || '')
          const imageUrl = String(op.params?.imageUrl || '')
          const opacity = Number(op.params?.opacity) || 0.15
          const angle = Number(op.params?.angle) || -30
          const fontSize = String(op.params?.fontSize || '48px')
          const color = String(op.params?.color || '#888888')
          
          if (!text && !imageUrl) continue
          
          // 创建水印元素
          const watermark = doc.createElement('div')
          watermark.className = 'document-watermark'
          watermark.setAttribute('data-diff-id', diffId)
          watermark.setAttribute('data-diff-role', 'new')
          watermark.setAttribute('data-diff-kind', 'block')
          
          if (text) {
            watermark.setAttribute('style', `
              position: fixed;
              top: 50%;
              left: 50%;
              transform: translate(-50%, -50%) rotate(${angle}deg);
              font-size: ${fontSize};
              color: ${color};
              opacity: ${opacity};
              pointer-events: none;
              z-index: 1000;
              white-space: nowrap;
              user-select: none;
            `)
            watermark.textContent = text
          } else if (imageUrl) {
            watermark.setAttribute('style', `
              position: fixed;
              top: 50%;
              left: 50%;
              transform: translate(-50%, -50%);
              opacity: ${opacity};
              pointer-events: none;
              z-index: 1000;
            `)
            const img = doc.createElement('img')
            img.src = imageUrl
            img.style.maxWidth = '300px'
            watermark.appendChild(img)
          }
          
          doc.body.insertBefore(watermark, doc.body.firstChild)
          
          created.push({
            id: diffId,
            kind: 'watermark',
            scope: 'document',
            summary: text ? `添加文字水印: ${text}` : '添加图片水印',
            afterPreview: text || '图片水印',
            stats: { matches: 1 },
            timestamp: Date.now(),
            meta: { op },
          })
          continue
        }

        // toc - 目录生成
        if (type === 'toc') {
          const diffId = genId()
          const maxLevel = Number(op.params?.maxLevel) || 3
          const position = String(op.params?.position || 'start')
          const title = String(op.params?.title || '目录')
          
          // 收集标题
          const headings = Array.from(doc.body.querySelectorAll<HTMLElement>('h1,h2,h3,h4,h5,h6'))
            .filter(h => {
              const level = parseInt(h.tagName.substring(1))
              return level <= maxLevel
            })
          
          if (headings.length === 0) continue
          
          // 创建目录容器
          const tocContainer = doc.createElement('div')
          tocContainer.className = 'table-of-contents'
          tocContainer.setAttribute('data-diff-id', diffId)
          tocContainer.setAttribute('data-diff-role', 'new')
          tocContainer.setAttribute('data-diff-kind', 'block')
          tocContainer.setAttribute('style', 'margin: 1em 0; padding: 1em; border: 1px solid var(--word-rule); background: #f9f9f9;')
          
          // 目录标题
          const tocTitle = doc.createElement('h2')
          tocTitle.textContent = title
          tocTitle.setAttribute('style', 'margin-bottom: 0.5em; font-size: 18px;')
          tocContainer.appendChild(tocTitle)
          
          // 目录列表
          const tocList = doc.createElement('ul')
          tocList.setAttribute('style', 'list-style: none; padding-left: 0; margin: 0;')
          
          headings.forEach((heading, index) => {
            const level = parseInt(heading.tagName.substring(1))
            const text = heading.textContent || `标题 ${index + 1}`
            
            const item = doc.createElement('li')
            item.setAttribute('style', `padding-left: ${(level - 1) * 1.5}em; margin: 0.3em 0;`)
            
            const link = doc.createElement('a')
            // 为标题添加 id
            const headingId = `heading-${index}`
            heading.id = headingId
            link.href = `#${headingId}`
            link.textContent = text
            link.setAttribute('style', 'color: #1976d2; text-decoration: none;')
            
            item.appendChild(link)
            tocList.appendChild(item)
          })
          
          tocContainer.appendChild(tocList)
          
          // 插入位置
          if (position === 'start') {
            doc.body.insertBefore(tocContainer, doc.body.firstChild)
          } else {
            // 在指定位置后插入
            const anchorBlock = Array.from(doc.body.querySelectorAll<HTMLElement>('p,h1,h2,h3'))
              .find(b => (b.textContent || '').includes(position))
            if (anchorBlock) {
              anchorBlock.insertAdjacentElement('afterend', tocContainer)
            } else {
              doc.body.insertBefore(tocContainer, doc.body.firstChild)
            }
          }
          
          created.push({
            id: diffId,
            kind: 'toc',
            scope: 'document',
            summary: `生成目录（${headings.length} 个标题）`,
            afterPreview: `${headings.length} 个条目`,
            stats: { matches: headings.length },
            timestamp: Date.now(),
            meta: { op },
          })
          continue
        }

        // header_footer - 页眉页脚设置
        if (type === 'header_footer') {
          const diffId = genId()
          
          const newSetup: Partial<HeaderFooterSetup> = {}
          
          // 页眉设置
          const headerParams = op.params?.header as { content?: string; alignment?: string; showOnFirstPage?: boolean } | undefined
          if (headerParams) {
            newSetup.header = {
              content: String(headerParams.content || ''),
              alignment: (headerParams.alignment || 'center') as 'left' | 'center' | 'right',
              showOnFirstPage: headerParams.showOnFirstPage !== false,
            }
          }
          
          // 页脚设置
          const footerParams = op.params?.footer as { content?: string; alignment?: string; showOnFirstPage?: boolean } | undefined
          if (footerParams) {
            newSetup.footer = {
              content: String(footerParams.content || ''),
              alignment: (footerParams.alignment || 'center') as 'left' | 'center' | 'right',
              showOnFirstPage: footerParams.showOnFirstPage !== false,
            }
          }
          
          // 页码设置
          const pageNumberParams = op.params?.pageNumber as { enabled?: boolean; position?: string; alignment?: string; format?: string; startFrom?: number } | undefined
          if (pageNumberParams) {
            newSetup.pageNumber = {
              enabled: pageNumberParams.enabled !== false,
              position: (pageNumberParams.position || 'footer') as 'header' | 'footer',
              alignment: (pageNumberParams.alignment || 'center') as 'left' | 'center' | 'right',
              format: (pageNumberParams.format || 'arabic') as 'arabic' | 'roman' | 'letter',
              startFrom: Number(pageNumberParams.startFrom) || 1,
            }
          }
          
          if (Object.keys(newSetup).length > 0) {
            setHeaderFooterSetupState(prev => ({ ...prev, ...newSetup }))
            
            const changes: string[] = []
            if (newSetup.header) changes.push('页眉')
            if (newSetup.footer) changes.push('页脚')
            if (newSetup.pageNumber) changes.push('页码')
            
            created.push({
              id: diffId,
              kind: 'header_footer',
              scope: 'document',
              summary: `设置 ${changes.join('、')}`,
              afterPreview: changes.join('、'),
              stats: { matches: 1 },
              timestamp: Date.now(),
              meta: { op },
            })
          }
          continue
        }

        // outline_summary - 大纲/摘要生成
        if (type === 'outline_summary') {
          const diffId = genId()
          const action = String(op.params?.action || 'extract_outline')
          
          if (action === 'extract_outline') {
            // 提取文档大纲
            const headings = Array.from(doc.body.querySelectorAll<HTMLElement>('h1,h2,h3,h4,h5,h6'))
            const outline: Array<{ level: number; text: string }> = []
            
            headings.forEach(h => {
              const level = parseInt(h.tagName.substring(1))
              const text = (h.textContent || '').trim()
              if (text) {
                outline.push({ level, text })
              }
            })
            
            // 生成大纲文本
            const outlineText = outline.map(item => {
              const indent = '  '.repeat(item.level - 1)
              const prefix = item.level === 1 ? '■' : item.level === 2 ? '●' : '○'
              return `${indent}${prefix} ${item.text}`
            }).join('\n')
            
            created.push({
              id: diffId,
              kind: 'outline_summary',
              scope: 'document',
              summary: `提取大纲（${outline.length} 个标题）`,
              afterPreview: outlineText.slice(0, 200) + (outlineText.length > 200 ? '...' : ''),
              stats: { matches: outline.length },
              timestamp: Date.now(),
              meta: { op, outline },
            })
            continue
          }
          
          if (action === 'generate_summary') {
            // 生成摘要 - 提取文档文本供 AI 使用
            const plainText = (doc.body.textContent || '').trim()
            const summaryLength = String(op.params?.summaryLength || 'medium')
            const targetLength = summaryLength === 'short' ? 100 : summaryLength === 'long' ? 300 : 200
            
            // 简单的摘要生成：提取前几个段落的内容
            const paragraphs = Array.from(doc.body.querySelectorAll<HTMLElement>('p'))
            let summaryText = ''
            for (const p of paragraphs) {
              const text = (p.textContent || '').trim()
              if (text && summaryText.length < targetLength) {
                summaryText += text + ' '
              }
            }
            summaryText = summaryText.trim().slice(0, targetLength) + '...'
            
            created.push({
              id: diffId,
              kind: 'outline_summary',
              scope: 'document',
              summary: `生成摘要（${targetLength}字）`,
              afterPreview: summaryText,
              stats: { matches: 1 },
              timestamp: Date.now(),
              meta: { op, plainText: plainText.slice(0, 2000), targetLength },
            })
            continue
          }
          
          if (action === 'insert_summary') {
            // 插入摘要到文档
            const summaryContent = String(op.params?.content || '')
            const position = String(op.params?.position || 'start')
            
            if (summaryContent) {
              const summaryDiv = doc.createElement('div')
              summaryDiv.className = 'document-summary'
              summaryDiv.setAttribute('data-diff-id', diffId)
              summaryDiv.setAttribute('data-diff-role', 'new')
              summaryDiv.setAttribute('data-diff-kind', 'block')
              summaryDiv.setAttribute('style', 'margin: 1em 0; padding: 1em; border-left: 4px solid var(--accent); background: #e3f2fd;')
              
              const summaryTitle = doc.createElement('h3')
              summaryTitle.textContent = '摘要'
              summaryTitle.setAttribute('style', 'margin: 0 0 0.5em 0; color: #1976d2;')
              summaryDiv.appendChild(summaryTitle)
              
              const summaryPara = doc.createElement('p')
              summaryPara.textContent = summaryContent
              summaryPara.setAttribute('style', 'margin: 0; line-height: 1.6;')
              summaryDiv.appendChild(summaryPara)
              
              if (position === 'start') {
                doc.body.insertBefore(summaryDiv, doc.body.firstChild)
              } else {
                doc.body.appendChild(summaryDiv)
              }
              
              created.push({
                id: diffId,
                kind: 'outline_summary',
                scope: 'document',
                summary: '插入摘要',
                afterPreview: summaryContent.slice(0, 100),
                stats: { matches: 1 },
                timestamp: Date.now(),
                meta: { op },
              })
            }
            continue
          }
        }

        // template_fill - 模板智能填充
        if (type === 'template_fill') {
          const diffId = genId()
          const action = String(op.params?.action || 'detect_placeholders')
          
          // 占位符模式：{{xxx}}, 【xxx】, [xxx], ___, <xxx>
          const placeholderPatterns = [
            /\{\{([^}]+)\}\}/g,           // {{公司名称}}
            /【([^】]+)】/g,               // 【日期】
            /\[([^\]]+)\]/g,              // [姓名]
            /_{3,}/g,                      // ___
            /<([^>]+)>/g,                  // <地址>
          ]
          
          if (action === 'detect_placeholders') {
            // 检测所有占位符
            const plainText = doc.body.innerHTML
            const foundPlaceholders: Array<{ pattern: string; text: string; count: number }> = []
            
            for (const pattern of placeholderPatterns) {
              const matches = plainText.match(pattern) || []
              if (matches.length > 0) {
                const uniqueMatches = [...new Set(matches)]
                uniqueMatches.forEach(m => {
                  const count = matches.filter(x => x === m).length
                  foundPlaceholders.push({ pattern: pattern.source, text: m, count })
                })
              }
            }
            
            created.push({
              id: diffId,
              kind: 'template_fill',
              scope: 'document',
              summary: `检测到 ${foundPlaceholders.length} 个占位符`,
              afterPreview: foundPlaceholders.map(p => `${p.text} (${p.count}处)`).join(', ').slice(0, 150),
              stats: { matches: foundPlaceholders.length },
              timestamp: Date.now(),
              meta: { op, placeholders: foundPlaceholders },
            })
            continue
          }

          if (action === 'detect_fields') {
            const fields = detectTemplateFields(doc)
            created.push({
              id: diffId,
              kind: 'template_fill',
              scope: 'document',
              summary: `检测到 ${fields.length} 个候选字段`,
              afterPreview: fields
                .slice(0, 6)
                .map(f => `${f.label}${f.path ? `(${f.path})` : ''}`)
                .join('，')
                .slice(0, 120),
              stats: { matches: fields.length },
              timestamp: Date.now(),
              meta: { op, fields },
            })
            continue
          }

          if (action === 'apply' || action === 'apply_fields') {
            let assignments: Array<{
              fieldId?: string
              label?: string
              value?: string
              path?: string
              oldValue?: string
              fieldType?: string
            }> = []
            const raw = op.params?.assignments
            if (Array.isArray(raw)) {
              assignments = raw as Array<{
                fieldId?: string
                label?: string
                value?: string
                path?: string
                oldValue?: string
                fieldType?: string
              }>
            } else if (typeof raw === 'string') {
              try {
                assignments = JSON.parse(raw)
              } catch {
                assignments = []
              }
            }

            const blocks = Array.from(doc.body.querySelectorAll<HTMLElement>('p,h1,h2,h3,h4,h5,h6'))
            const tables = Array.from(doc.body.querySelectorAll<HTMLTableElement>('table'))
            const fields = detectTemplateFields(doc, 200)
            const fieldsById = new Map(fields.map((f) => [f.id, f]))
            let appliedCount = 0
            const skipped: Array<{ fieldId?: string; label?: string; reason: string }> = []
            const warnings: string[] = []

            const validateFieldValue = (fieldType: string, value: string) => {
              if (!fieldType || !value) return true
              const v = value.trim()
              if (fieldType === 'date') {
                return /\d{4}[./\-年]\d{1,2}[./\-月]?\d{0,2}日?/.test(v)
              }
              if (fieldType === 'amount') {
                return /[\d,.]+/.test(v)
              }
              if (fieldType === 'phone') {
                return /[0-9]{6,}/.test(v)
              }
              if (fieldType === 'email') {
                return /@/.test(v)
              }
              if (fieldType === 'id') {
                return /[0-9A-Za-z]{4,}/.test(v)
              }
              return true
            }

            const resolveCellByPath = (path?: string) => {
              if (!path) return null
              const m = path.match(/table\[(\d+)\]\/r\[(\d+)\]\/c\[(\d+)\]/i)
              if (!m) return null
              const tableIndex = parseInt(m[1], 10) - 1
              const rowIndex = parseInt(m[2], 10) - 1
              const colIndex = parseInt(m[3], 10) - 1
              const table = tables[tableIndex]
              const row = table?.querySelectorAll('tr')?.[rowIndex]
              const cell = row?.querySelectorAll<HTMLElement>('th,td')?.[colIndex]
              return cell || null
            }

            const resolveBlockByPath = (path?: string) => {
              if (!path) return null
              const seg = path.split('/').pop() || path
              const m = seg.match(/(p|h[1-6])\[(\d+)\]/i)
              if (!m) return null
              const tag = m[1].toLowerCase()
              const idx = parseInt(m[2], 10) - 1
              if (Number.isNaN(idx) || idx < 0) return null
              const list = Array.from(doc.body.querySelectorAll<HTMLElement>(tag))
              return list[idx] || null
            }

            const applyByLabel = (label: string, value: string, diffIdForFill: string, oldValue?: string) => {
              const labelText = label.trim()
              if (!labelText) return false
              const target = blocks.find((b) => {
                const text = normalizeText(b.textContent || '')
                return text.includes(labelText)
              })
              if (!target) return false
              if (oldValue) {
                const text = normalizeText(target.textContent || '')
                if (!text.includes(oldValue)) return false
              }
              applyTemplateValueToElement(target, value, diffIdForFill, oldValue)
              return true
            }

            for (const item of assignments) {
              const value = (item.value || '').toString().trim()
              if (!value) continue
              const fieldId = (item.fieldId || '').toString().trim()
              let applied = false
              const itemDiffId = genId()
              const fieldInfo = fieldId ? fieldsById.get(fieldId) : undefined
              const resolvedPath = item.path || fieldInfo?.path || (fieldInfo?.meta as any)?.path
              const oldValue = (item.oldValue || fieldInfo?.currentValue || '').toString().trim()
              const fieldType = (item.fieldType || fieldInfo?.fieldType || '').toString().trim()

              if (fieldType && !validateFieldValue(fieldType, value)) {
                warnings.push(`${item.label || fieldId || resolvedPath || '字段'} 的值可能不符合类型(${fieldType})`)
              }

              if (fieldId.startsWith('p:')) {
                const parts = fieldId.split(':')
                const index = parseInt(parts[1], 10)
                if (!Number.isNaN(index) && blocks[index]) {
                  if (!oldValue || normalizeText(blocks[index].textContent || '').includes(oldValue)) {
                    applyTemplateValueToElement(blocks[index], value, itemDiffId, oldValue)
                    applied = true
                  }
                }
              } else if (fieldId.startsWith('t:')) {
                const parts = fieldId.split(':')
                const tableIndex = parseInt(parts[1], 10)
                const rowIndex = parseInt(parts[3], 10)
                const colIndex = parseInt(parts[5], 10)
                const table = tables[tableIndex]
                const row = table?.querySelectorAll('tr')?.[rowIndex]
                const cell = row?.querySelectorAll<HTMLElement>('th,td')?.[colIndex]
                if (cell) {
                  if (!oldValue || normalizeText(cell.textContent || '').includes(oldValue)) {
                    applyTemplateValueToElement(cell, value, itemDiffId, oldValue)
                    applied = true
                  }
                }
              } else if (resolvedPath) {
                const cell = resolveCellByPath(resolvedPath)
                if (cell) {
                  if (!oldValue || normalizeText(cell.textContent || '').includes(oldValue)) {
                    applyTemplateValueToElement(cell, value, itemDiffId, oldValue)
                    applied = true
                  }
                } else {
                  const block = resolveBlockByPath(resolvedPath)
                  if (block) {
                    if (!oldValue || normalizeText(block.textContent || '').includes(oldValue)) {
                      applyTemplateValueToElement(block, value, itemDiffId, oldValue)
                      applied = true
                    }
                  }
                }
              }

              if (!applied && item.label) {
                applied = applyByLabel(item.label, value, itemDiffId, oldValue)
              }

              if (applied) {
                appliedCount += 1
                created.push({
                  id: itemDiffId,
                  kind: 'template_fill',
                  scope: 'document',
                  summary: `填写 ${item.label || fieldId || '字段'}`,
                  beforePreview: item.label || fieldId,
                  afterPreview: value.slice(0, 80),
                  stats: { matches: 1 },
                  timestamp: Date.now(),
                  meta: { op, fieldId, label: item.label, path: resolvedPath },
                })
              } else {
                skipped.push({
                  fieldId,
                  label: item.label,
                  reason: resolvedPath ? '未匹配到定位路径或旧值不一致' : '未匹配到字段',
                })
              }
            }

            if (appliedCount > 0) {
              html = doc.body.innerHTML
            }

            templateFillReport = {
              applied: appliedCount,
              skipped: skipped.length,
              warnings: warnings.length,
            }

            if (skipped.length > 0) {
              created.push({
                id: diffId,
                kind: 'template_fill',
                scope: 'document',
                summary: `未匹配 ${skipped.length} 项字段`,
                afterPreview: skipped.slice(0, 6).map((s) => s.label || s.fieldId || '字段').join('，'),
                stats: { matches: 0 },
                timestamp: Date.now(),
                meta: { op, skipped },
              })
            }

            if (warnings.length > 0) {
              created.push({
                id: genId(),
                kind: 'template_fill',
                scope: 'document',
                summary: `字段类型校验提示 ${warnings.length} 条`,
                afterPreview: warnings.slice(0, 4).join('；'),
                stats: { matches: 0 },
                timestamp: Date.now(),
                meta: { op, warnings },
              })
            }

            continue
          }
          
          if (action === 'fill_single' || action === 'fill_all') {
            // 填充占位符
            const placeholder = String(op.params?.placeholder || '')
            const value = String(op.params?.value || '')
            
            if (placeholder && value) {
              let htmlContent = doc.body.innerHTML
              const escapedPlaceholder = placeholder.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
              const regex = new RegExp(escapedPlaceholder, 'g')
              const matches = htmlContent.match(regex) || []
              
              if (matches.length > 0) {
                // 创建带高亮的替换
                const highlightedValue = `<span style="background-color: #c8e6c9; padding: 0 2px;" data-diff-id="${diffId}" data-diff-role="new">${value}</span>`
                htmlContent = htmlContent.replace(regex, highlightedValue)
                doc.body.innerHTML = htmlContent
                
                created.push({
                  id: diffId,
                  kind: 'template_fill',
                  scope: 'document',
                  summary: `填充「${placeholder}」→「${value}」`,
                  beforePreview: placeholder,
                  afterPreview: value,
                  stats: { matches: matches.length },
                  timestamp: Date.now(),
                  meta: { op },
                })
              }
            }
            continue
          }
        }

        // citation_footnote - 引用/脚注管理
        if (type === 'citation_footnote') {
          const diffId = genId()
          const action = String(op.params?.action || 'insert_footnote')
          
          if (action === 'insert_footnote' || action === 'insert_endnote') {
            const content = String(op.params?.content || '')
            const anchorText = String(op.params?.anchorText || op.target?.text || '')
            
            if (content) {
              // 查找现有脚注数量
              const existingNotes = doc.body.querySelectorAll('.footnote-ref, .endnote-ref')
              const noteNumber = existingNotes.length + 1
              const isEndnote = action === 'insert_endnote'
              const noteClass = isEndnote ? 'endnote' : 'footnote'
              
              // 在锚点文本后插入脚注引用
              if (anchorText) {
                const walker = doc.createTreeWalker(doc.body, NodeFilter.SHOW_TEXT, null)
                let node: Node | null
                while ((node = walker.nextNode())) {
                  if (node.textContent?.includes(anchorText)) {
                    const text = node.textContent
                    const idx = text.indexOf(anchorText)
                    const before = text.slice(0, idx + anchorText.length)
                    const after = text.slice(idx + anchorText.length)
                    
                    const span = doc.createElement('span')
                    span.innerHTML = `${before}<sup class="${noteClass}-ref" style="color: #1976d2; cursor: pointer;" data-note-id="${noteNumber}">[${noteNumber}]</sup>${after}`
                    node.parentNode?.replaceChild(span, node)
                    break
                  }
                }
              }
              
              // 创建脚注/尾注内容区
              let notesContainer = doc.body.querySelector(`.${noteClass}-container`)
              if (!notesContainer) {
                notesContainer = doc.createElement('div')
                notesContainer.className = `${noteClass}-container`
                notesContainer.setAttribute('style', 'margin-top: 2em; padding-top: 1em; border-top: 1px solid var(--word-rule);')
                
                const title = doc.createElement('h4')
                title.textContent = isEndnote ? '尾注' : '脚注'
                title.setAttribute('style', 'margin: 0 0 0.5em 0; color: #666;')
                notesContainer.appendChild(title)
                
                doc.body.appendChild(notesContainer)
              }
              
              // 添加脚注内容
              const noteItem = doc.createElement('p')
              noteItem.className = `${noteClass}-item`
              noteItem.setAttribute('data-diff-id', diffId)
              noteItem.setAttribute('data-diff-role', 'new')
              noteItem.setAttribute('data-diff-kind', 'block')
              noteItem.setAttribute('style', 'margin: 0.3em 0; font-size: 0.9em; color: #555;')
              noteItem.innerHTML = `<sup style="color: #1976d2;">[${noteNumber}]</sup> ${content}`
              notesContainer.appendChild(noteItem)
              
              created.push({
                id: diffId,
                kind: 'citation_footnote',
                scope: 'document',
                summary: `插入${isEndnote ? '尾注' : '脚注'} [${noteNumber}]`,
                afterPreview: content.slice(0, 80),
                stats: { matches: 1 },
                timestamp: Date.now(),
                meta: { op, noteNumber },
              })
            }
            continue
          }
          
          if (action === 'add_citation') {
            // 添加引用
            const source = op.params?.source as { type?: string; title?: string; author?: string; year?: string; url?: string } | undefined
            const format = String(op.params?.format || 'gb7714')
            
            if (source?.title) {
              // 查找现有引用数量
              const existingCitations = doc.body.querySelectorAll('.citation-ref')
              const citationNumber = existingCitations.length + 1
              
              // 格式化引用文本
              let citationText = ''
              if (format === 'gb7714') {
                // GB/T 7714 格式
                if (source.author) citationText += `${source.author}. `
                citationText += `${source.title}`
                if (source.year) citationText += `[${source.type === 'website' ? 'EB/OL' : source.type === 'article' ? 'J' : 'M'}]. ${source.year}`
                if (source.url) citationText += `. ${source.url}`
              } else if (format === 'apa') {
                // APA 格式
                if (source.author) citationText += `${source.author} `
                if (source.year) citationText += `(${source.year}). `
                citationText += `${source.title}.`
                if (source.url) citationText += ` Retrieved from ${source.url}`
              } else {
                // MLA 格式
                if (source.author) citationText += `${source.author}. `
                citationText += `"${source.title}."`
                if (source.year) citationText += ` ${source.year}.`
              }
              
              // 创建或获取参考文献区
              let bibContainer = doc.body.querySelector('.bibliography-container')
              if (!bibContainer) {
                bibContainer = doc.createElement('div')
                bibContainer.className = 'bibliography-container'
                bibContainer.setAttribute('style', 'margin-top: 2em; padding-top: 1em; border-top: 2px solid #333;')
                
                const title = doc.createElement('h3')
                title.textContent = '参考文献'
                title.setAttribute('style', 'margin: 0 0 1em 0;')
                bibContainer.appendChild(title)
                
                doc.body.appendChild(bibContainer)
              }
              
              // 添加引用条目
              const bibItem = doc.createElement('p')
              bibItem.className = 'bibliography-item'
              bibItem.setAttribute('data-diff-id', diffId)
              bibItem.setAttribute('data-diff-role', 'new')
              bibItem.setAttribute('data-diff-kind', 'block')
              bibItem.setAttribute('style', 'margin: 0.5em 0; text-indent: -2em; padding-left: 2em;')
              bibItem.innerHTML = `[${citationNumber}] ${citationText}`
              bibContainer.appendChild(bibItem)
              
              created.push({
                id: diffId,
                kind: 'citation_footnote',
                scope: 'document',
                summary: `添加引用 [${citationNumber}]`,
                afterPreview: citationText.slice(0, 80),
                stats: { matches: 1 },
                timestamp: Date.now(),
                meta: { op, citationNumber },
              })
            }
            continue
          }
          
          if (action === 'generate_bibliography') {
            // 生成参考文献（如果已有引用，整理格式）
            const bibContainer = doc.body.querySelector('.bibliography-container')
            if (bibContainer) {
              // 重新编号所有引用
              const items = bibContainer.querySelectorAll('.bibliography-item')
              items.forEach((item, index) => {
                const text = item.textContent || ''
                const cleanText = text.replace(/^\[\d+\]\s*/, '')
                item.innerHTML = `[${index + 1}] ${cleanText}`
              })
              
              created.push({
                id: diffId,
                kind: 'citation_footnote',
                scope: 'document',
                summary: `整理参考文献（${items.length} 条）`,
                afterPreview: `共 ${items.length} 条引用`,
                stats: { matches: items.length },
                timestamp: Date.now(),
                meta: { op },
              })
            }
            continue
          }
        }
      }

      const newHtml = doc.body.innerHTML
      const htmlChanged = newHtml !== originalHtml
      if (htmlChanged) {
        documentContentRef.current = newHtml
        setDocument(prev => ({
          ...prev,
          content: newHtml,
          lastModified: new Date(),
        }))
        setDocxData(null)
        setHasUnsavedChanges(true)
      }

      if (created.length > 0) {
        setExtraPendingChanges(prev => [...prev, ...created])
      }

      if (!htmlChanged && created.length === 0) {
        return {
          success: false,
          message: 'No revisions were produced: no editable target matched current document for provided target/params',
          data: { created: 0, templateFillReport },
        }
      }

      const baseMessage = `已生成修订：${created.length} 条。请在底部或“修订面板”中逐条确认。`
      const reportMessage = templateFillReport
        ? `已生成修订：${created.length} 条（填充 ${templateFillReport.applied}，未匹配 ${templateFillReport.skipped}，校验提示 ${templateFillReport.warnings}）。请在底部或“修订面板”中逐条确认。`
        : baseMessage

      return {
        success: true,
        message: reportMessage,
        data: { created: created.length, templateFillReport },
      }
    } catch (e) {
      return { success: false, message: `应用失败: ${(e as Error).message || String(e)}` }
    }
  }, [])

  // 编辑器模式 - 默认使用 Tiptap（内置编辑器），更稳定可靠
  const [editorMode, setEditorMode] = useState<EditorMode>('tiptap')

  // ONLYOFFICE 专用操作 - 搜索替换
  const onlyOfficeReplace = useCallback(async (search: string, replace: string): Promise<ReplaceResult> => {
    if (!window.onlyOfficeConnector) {
      return { success: false, count: 0, message: 'ONLYOFFICE 编辑器未就绪' }
    }

    try {
      const result = await window.onlyOfficeConnector.searchAndReplace(search, replace, true)
      if (result) {
        return { 
          success: true, 
          count: 1, // ONLYOFFICE API 不返回替换次数，假设为1
          message: `已替换「${search}」→「${replace}」`,
          searchText: search,
          replaceText: replace
        }
      } else {
        return { success: false, count: 0, message: `未找到「${search}」` }
      }
    } catch (e) {
      console.error('ONLYOFFICE 替换失败:', e)
      return { success: false, count: 0, message: `替换失败: ${e}` }
    }
  }, [])

  // ONLYOFFICE 专用操作 - 插入文本
  const onlyOfficeInsert = useCallback(async (text: string): Promise<{ success: boolean; message: string }> => {
    if (!window.onlyOfficeConnector) {
      return { success: false, message: 'ONLYOFFICE 编辑器未就绪' }
    }

    try {
      const result = await window.onlyOfficeConnector.insertText(text)
      if (result) {
        return { success: true, message: '已插入文本' }
      } else {
        return { success: false, message: '插入失败' }
      }
    } catch (e) {
      console.error('ONLYOFFICE 插入失败:', e)
      return { success: false, message: `插入失败: ${e}` }
    }
  }, [])

  // ONLYOFFICE 专用操作 - 获取文档文本
  const onlyOfficeGetText = useCallback(async (): Promise<string> => {
    if (!window.onlyOfficeConnector) {
      return ''
    }

    try {
      return await window.onlyOfficeConnector.getDocumentText()
    } catch (e) {
      console.error('ONLYOFFICE 获取文本失败:', e)
      return ''
    }
  }, [])

  // ONLYOFFICE 专用操作 - 添加带格式的段落
  const onlyOfficeAddParagraph = useCallback(async (
    text: string, 
    options?: {
      fontSize?: number
      fontFamily?: string
      bold?: boolean
      italic?: boolean
      color?: string
      alignment?: 'left' | 'center' | 'right' | 'justify'
    }
  ): Promise<{ success: boolean; message: string }> => {
    if (!window.onlyOfficeConnector) {
      return { success: false, message: 'ONLYOFFICE 编辑器未就绪' }
    }

    try {
      const result = await window.onlyOfficeConnector.addFormattedParagraph(text, options)
      if (result) {
        return { success: true, message: '已添加段落' }
      } else {
        return { success: false, message: '添加段落失败' }
      }
    } catch (e) {
      console.error('ONLYOFFICE 添加段落失败:', e)
      return { success: false, message: `添加段落失败: ${e}` }
    }
  }, [])

  // ONLYOFFICE 专用操作 - 添加标题
  const onlyOfficeAddHeading = useCallback(async (
    text: string, 
    level: 1 | 2 | 3 | 4 | 5 | 6
  ): Promise<{ success: boolean; message: string }> => {
    if (!window.onlyOfficeConnector) {
      return { success: false, message: 'ONLYOFFICE 编辑器未就绪' }
    }

    try {
      const result = await window.onlyOfficeConnector.addHeading(text, level)
      if (result) {
        return { success: true, message: `已添加 ${level} 级标题` }
      } else {
        return { success: false, message: '添加标题失败' }
      }
    } catch (e) {
      console.error('ONLYOFFICE 添加标题失败:', e)
      return { success: false, message: `添加标题失败: ${e}` }
    }
  }, [])

  // ONLYOFFICE 专用操作 - 添加表格
  const onlyOfficeAddTable = useCallback(async (
    rows: number, 
    cols: number, 
    data?: string[][]
  ): Promise<{ success: boolean; message: string }> => {
    if (!window.onlyOfficeConnector) {
      return { success: false, message: 'ONLYOFFICE 编辑器未就绪' }
    }

    try {
      const result = await window.onlyOfficeConnector.addTable(rows, cols, data)
      if (result) {
        return { success: true, message: `已添加 ${rows}x${cols} 表格` }
      } else {
        return { success: false, message: '添加表格失败' }
      }
    } catch (e) {
      console.error('ONLYOFFICE 添加表格失败:', e)
      return { success: false, message: `添加表格失败: ${e}` }
    }
  }, [])

  // MCP Bridge: 让外部 agent 通过 MCP 协议操控文档
  useMcpBridge({
    getContent: () => documentContentRef.current,
    workspacePath,
    insertViaDsl,
    replaceViaDsl,
    deleteViaDsl,
    insertInDocument,
    silentSaveToFile,
    openFile,
    getTiptapDocumentStructure,
    currentFile,
  })

  return (
    <DocumentContext.Provider
      value={{
        document,
        files,
        currentFile,
        workspacePath,
        isElectron,
        hasUnsavedChanges,
        docxData,
        excelData,
        pptData,
        typographyProfile,
        refreshExcelData,
        lastReplacement,
        pendingChanges: [
          ...pendingReplacements.items.map((item) => {
            const stripHtml = (s: string) => (s || '').replace(/<[^>]+>/g, '').trim()
            const isReview = !!(item.reviewReason || item.reviewType)
            const typeLabels: Record<string, string> = { grammar: '语法', logic: '逻辑', style: '措辞', typo: '错别字', format: '格式' }
            const summaryText = isReview
              ? `[${typeLabels[item.reviewType || 'style'] || item.reviewType}] ${item.reviewReason || '审查修改'}`
              : `替换 ${item.count} 处`
            return ({
            id: item.id,
            kind: 'replace_text' as const,
            scope: 'document' as const,
            summary: summaryText,
            beforePreview: stripHtml(item.searchText),
            afterPreview: stripHtml(item.replaceText),
            stats: { matches: item.count },
            timestamp: item.timestamp,
            meta: {
              searchText: item.searchText,
              replaceText: item.replaceText,
              count: item.count,
            },
            reviewReason: item.reviewReason,
            reviewType: item.reviewType,
            })
          }),
          ...extraPendingChanges,
        ],
        pendingChangesTotal:
          pendingReplacements.total +
          extraPendingChanges.reduce((sum, c) => sum + (c.stats?.matches ?? 1), 0),
        editorMode,
        setEditorMode,
        setDocument,
        updateDocument,
        updateContent,
        updateStyles,
        setCurrentFile,
        addFile,
        createNewDocument,
        createDocumentFromDsl,
        uploadDocxFile,
        saveDocument,
        applyAIEdit,
        replaceInDocument,
        insertInDocument,
        deleteInDocument,
        replaceViaDsl,
        insertViaDsl,
        deleteViaDsl,
        scrollToText,
        scrollToDiffId,
        addPendingReplacementItem,
        previewWordOps,
        applyWordOps,
        confirmReplacement,
        rejectReplacement,
        acceptChange,
        rejectChange,
        acceptAllChanges,
        rejectAllChanges,
        openFolder,
        openFile,
        saveCurrentFile,
        silentSaveToFile,
        refreshFiles,
        onlyOfficeReplace,
        onlyOfficeInsert,
        onlyOfficeGetText,
        onlyOfficeAddParagraph,
        onlyOfficeAddHeading,
        onlyOfficeAddTable,
        getTiptapDocumentStructure,
        replaceWithFormat,
        docEntryAnimationKey,
        triggerDocEntryAnimation,
        getLatestContent: () => documentContentRef.current,
        pageSetup,
        setPageSetup: (setup: Partial<PageSetup>) => {
          setPageSetupState(prev => ({ ...prev, ...setup }))
          setHasUnsavedChanges(true)
        },
        headerFooterSetup,
        setHeaderFooterSetup: (setup: Partial<HeaderFooterSetup>) => {
          setHeaderFooterSetupState(prev => ({ ...prev, ...setup }))
          setHasUnsavedChanges(true)
        },
        customStyles,
        defineStyle: (style: CustomStyle) => {
          setCustomStyles(prev => ({ ...prev, [style.name]: style }))
          setHasUnsavedChanges(true)
        },
        modifyStyle: (name: string, updates: Partial<CustomStyle>) => {
          setCustomStyles(prev => {
            if (!prev[name]) return prev
            return { ...prev, [name]: { ...prev[name], ...updates } }
          })
          setHasUnsavedChanges(true)
        },
        deleteStyle: (name: string) => {
          // 不允许删除内置样式
          if (['Normal', 'Heading1', 'Heading2', 'Heading3'].includes(name)) return
          setCustomStyles(prev => {
            const newStyles = { ...prev }
            delete newStyles[name]
            return newStyles
          })
        },
        getStyleCSS: (styleName: string): string => {
          const style = customStyles[styleName]
          if (!style) return ''
          
          const rules: string[] = []
          if (style.fontFamily) rules.push(`font-family: ${style.fontFamily}`)
          if (style.fontSize) rules.push(`font-size: ${style.fontSize}`)
          if (style.color) rules.push(`color: ${style.color}`)
          if (style.bold) rules.push('font-weight: bold')
          if (style.italic) rules.push('font-style: italic')
          if (style.underline) rules.push('text-decoration: underline')
          if (style.strikethrough) rules.push('text-decoration: line-through')
          if (style.letterSpacing) rules.push(`letter-spacing: ${style.letterSpacing}`)
          if (style.alignment) rules.push(`text-align: ${style.alignment}`)
          if (style.lineHeight) rules.push(`line-height: ${style.lineHeight}`)
          if (style.spaceBefore) rules.push(`margin-top: ${style.spaceBefore}`)
          if (style.spaceAfter) rules.push(`margin-bottom: ${style.spaceAfter}`)
          if (style.textIndent) rules.push(`text-indent: ${style.textIndent}`)
          if (style.marginLeft) rules.push(`margin-left: ${style.marginLeft}`)
          if (style.marginRight) rules.push(`margin-right: ${style.marginRight}`)
          if (style.backgroundColor) rules.push(`background-color: ${style.backgroundColor}`)
          if (style.border) rules.push(`border: ${style.border}`)
          
          return rules.join('; ')
        },
      }}
    >
      {children}
    </DocumentContext.Provider>
  )
}

export function useDocument() {
  const context = useContext(DocumentContext)
  if (!context) {
    throw new Error('useDocument must be used within a DocumentProvider')
  }
  return context
}
