import { createContext, useContext, useState, useCallback, ReactNode, useRef, useEffect } from 'react'
import { DocumentContent, DocumentStyles, FileItem, PageSetup, HeaderFooterSetup, CustomStyle } from '../types'
import type { ExcelOpenResponse } from '../types/electron'
import { saveAs } from 'file-saver'
import { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell, WidthType, BorderStyle, UnderlineType } from 'docx'

// 将 ArrayBuffer 转换为 Base64（分块处理，避免大文件导致栈溢出）
function arrayBufferToBase64(buffer: ArrayBuffer): string {
  const bytes = new Uint8Array(buffer)
  const chunkSize = 8192 // 每次处理 8KB
  let binary = ''
  
  // #region agent log
  fetch('http://127.0.0.1:7242/ingest/65f1d8ba-6206-43cb-9f6f-22f7361d7de4',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({location:'DocumentContext.tsx:arrayBufferToBase64:entry',message:'base64 编码开始',data:{bufferSize:buffer.byteLength,bytesLength:bytes.length},timestamp:Date.now(),sessionId:'debug-session',hypothesisId:'H1'})}).catch(()=>{});
  // #endregion agent log
  
  for (let i = 0; i < bytes.length; i += chunkSize) {
    const chunk = bytes.subarray(i, Math.min(i + chunkSize, bytes.length))
    binary += String.fromCharCode.apply(null, Array.from(chunk))
  }
  
  const base64 = btoa(binary)
  
  // #region agent log
  fetch('http://127.0.0.1:7242/ingest/65f1d8ba-6206-43cb-9f6f-22f7361d7de4',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({location:'DocumentContext.tsx:arrayBufferToBase64:exit',message:'base64 编码完成',data:{binaryLength:binary.length,base64Length:base64.length,base64Preview:base64.slice(0,100)},timestamp:Date.now(),sessionId:'debug-session',hypothesisId:'H1'})}).catch(()=>{});
  // #endregion agent log
  
  return base64
}

interface ReplaceResult {
  success: boolean
  count: number
  message: string
  searchText?: string
  replaceText?: string
  positions?: number[]  // 替换发生的位置索引
}

// 单个替换记录
interface SingleReplacement {
  id: string  // 唯一标识
  searchText: string
  replaceText: string
  count: number
  timestamp: number
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
  createNewDocument: (title: string, content: string, elements?: FormattedElement[]) => void
  uploadDocxFile: (file: File) => Promise<void>
  saveDocument: () => Promise<void>
  applyAIEdit: (newContent: string) => void
  replaceInDocument: (search: string, replace: string) => ReplaceResult
  insertInDocument: (position: string, content: string) => { success: boolean; message: string }
  deleteInDocument: (target: string) => { success: boolean; count: number; message: string }
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
  }) => ReplaceResult
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
  fontFamily: '仿宋',
  lineHeight: 1.5,
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
    fontFamily: '仿宋',
    fontSize: '12pt',
    lineHeight: '1.5',
    textIndent: '2em',
  },
  'Heading1': {
    name: 'Heading1',
    fontFamily: '黑体',
    fontSize: '22pt',
    bold: true,
    alignment: 'center',
    spaceBefore: '12pt',
    spaceAfter: '6pt',
  },
  'Heading2': {
    name: 'Heading2',
    fontFamily: '黑体',
    fontSize: '16pt',
    bold: true,
    spaceBefore: '12pt',
    spaceAfter: '6pt',
  },
  'Heading3': {
    name: 'Heading3',
    fontFamily: '黑体',
    fontSize: '14pt',
    bold: true,
    spaceBefore: '6pt',
    spaceAfter: '3pt',
  },
  'Quote': {
    name: 'Quote',
    fontFamily: '楷体',
    fontSize: '12pt',
    italic: true,
    color: '#666666',
    marginLeft: '2em',
    marginRight: '2em',
    border: '1px solid #ddd',
    backgroundColor: '#f9f9f9',
  },
}

const defaultDocument: DocumentContent = {
  title: '新建文档',
  content: '',
  styles: defaultStyles,
  lastModified: new Date(),
}

const DocumentContext = createContext<DocumentContextType | undefined>(undefined)

// 检测是否在 Electron 环境
const isElectron = typeof window !== 'undefined' && !!window.electronAPI

// Markdown 转换为 docx 段落
function markdownToDocxParagraphs(content: string): Paragraph[] {
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
        children: parseInlineFormatting(trimmedLine.slice(4), true),
        heading: HeadingLevel.HEADING_3,
        spacing: { before: 200, after: 100 },
      }))
      continue
    }
    if (trimmedLine.startsWith('## ')) {
      flushList()
      paragraphs.push(new Paragraph({
        children: parseInlineFormatting(trimmedLine.slice(3), true),
        heading: HeadingLevel.HEADING_2,
        spacing: { before: 260, after: 130 },
      }))
      continue
    }
    if (trimmedLine.startsWith('# ')) {
      flushList()
      paragraphs.push(new Paragraph({
        children: parseInlineFormatting(trimmedLine.slice(2), true),
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
        children: parseInlineFormatting(trimmedLine.slice(2)),
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
        children: parseInlineFormatting(text),
        numbering: { reference: 'default-numbering', level: 0 },
        spacing: { after: 60 },
      }))
      continue
    }

    // 处理引用
    if (trimmedLine.startsWith('> ')) {
      flushList()
      paragraphs.push(new Paragraph({
        children: parseInlineFormatting(trimmedLine.slice(2)),
        indent: { left: 720 },
        border: { left: { style: 'single' as any, size: 12, space: 10, color: 'CCCCCC' } },
        spacing: { after: 100 },
      }))
      continue
    }

    // 处理普通段落
    flushList()
    paragraphs.push(new Paragraph({
      children: parseInlineFormatting(trimmedLine),
      indent: { firstLine: 480 }, // 首行缩进 2 字符
      spacing: { after: 120, line: 360 }, // 行距 1.5 倍
      alignment: AlignmentType.JUSTIFIED,
    }))
  }

  return paragraphs.length > 0 ? paragraphs : [new Paragraph({ children: [] })]
}

// 解析行内格式（粗体、斜体等）
function parseInlineFormatting(text: string, isHeading: boolean = false): TextRun[] {
  const runs: TextRun[] = []
  const fontSize = isHeading ? 28 : 28 // 四号字体 = 14pt = 28 half-points
  const fontName = isHeading ? '黑体' : '仿宋'
  
  // 简化处理：用正则分割文本
  const regex = /(\*\*\*.+?\*\*\*|\*\*.+?\*\*|\*.+?\*|__.+?__|_.+?_)/g
  let lastIndex = 0
  let match

  while ((match = regex.exec(text)) !== null) {
    // 添加匹配前的普通文本
    if (match.index > lastIndex) {
      runs.push(new TextRun({
        text: text.slice(lastIndex, match.index),
        font: fontName,
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
      font: fontName,
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
      font: fontName,
      size: fontSize,
    }))
  }

  return runs.length > 0 ? runs : [new TextRun({ text, font: fontName, size: fontSize })]
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
async function createDocxBlob(content: string, title: string): Promise<Blob> {
  // #region agent log
  fetch('http://127.0.0.1:7242/ingest/65f1d8ba-6206-43cb-9f6f-22f7361d7de4',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({location:'DocumentContext.tsx:createDocxBlob:entry',message:'createDocxBlob 开始',data:{title,contentLength:content.length,contentPreview:content.slice(0,100)},timestamp:Date.now(),sessionId:'debug-session',hypothesisId:'H2'})}).catch(()=>{});
  // #endregion agent log
  
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

  const htmlToDocxChildren = (html: string): (Paragraph | Table)[] => {
    const parser = new DOMParser()
    const doc = parser.parseFromString(html, 'text/html')

    const walkInline = (
      node: Node,
      inherited: {
        bold?: boolean
        italics?: boolean
        underline?: boolean
        color?: string
        font?: string
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
        if (fontFamily) next.font = fontFamily
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

    const processBlock = (el: HTMLElement) => {
      const tag = el.tagName.toLowerCase()

      // 忽略 old 块（导出默认“接受”）
      if (el.getAttribute('data-diff-role') === 'old') return
      if (tag === 'span' && el.classList.contains('diff-old')) return

      const style = el.getAttribute('style') || ''
      const { textAlign } = parseStyle(style)

      if (tag === 'h1' || tag === 'h2' || tag === 'h3') {
        const level =
          tag === 'h1' ? HeadingLevel.HEADING_1 : tag === 'h2' ? HeadingLevel.HEADING_2 : HeadingLevel.HEADING_3
        children.push(new Paragraph({
          heading: level,
          alignment: toAlignment(textAlign),
          children: walkInline(el, {}),
        }))
        return
      }

      if (tag === 'p') {
        children.push(new Paragraph({
          alignment: toAlignment(textAlign),
          children: walkInline(el, {}),
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

  const paragraphsOrTables = isHtml ? htmlToDocxChildren(content) : markdownToDocxParagraphs(content)
  
  // #region agent log
  fetch('http://127.0.0.1:7242/ingest/65f1d8ba-6206-43cb-9f6f-22f7361d7de4',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({location:'DocumentContext.tsx:createDocxBlob:paragraphs',message:'段落解析完成',data:{paragraphCount:paragraphsOrTables.length,isHtml},timestamp:Date.now(),sessionId:'debug-session',hypothesisId:'H2'})}).catch(()=>{});
  // #endregion agent log
  
  const doc = new Document({
    creator: 'Word-Cursor',
    title: title,
    description: 'Created by Word-Cursor',
    sections: [{
      properties: {
        page: {
          margin: {
            top: 1440,
            right: 1440,
            bottom: 1440,
            left: 1440,
          },
        },
      },
      children: paragraphsOrTables,
    }],
  })
  
  const blob = await Packer.toBlob(doc)
  
  // #region agent log
  fetch('http://127.0.0.1:7242/ingest/65f1d8ba-6206-43cb-9f6f-22f7361d7de4',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({location:'DocumentContext.tsx:createDocxBlob:exit',message:'Packer.toBlob 完成',data:{blobSize:blob.size,blobType:blob.type},timestamp:Date.now(),sessionId:'debug-session',hypothesisId:'H3'})}).catch(()=>{});
  // #endregion agent log
  
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
async function createFormattedDocxBlob(elements: FormattedElement[], title: string): Promise<Blob> {
  const children: (Paragraph | Table)[] = []
  
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
      
      children.push(new Paragraph({
        text: elem.content,
        heading: headingLevelMap[level] || HeadingLevel.HEADING_1,
        alignment: elem.alignment ? alignmentMap[elem.alignment] : AlignmentType.LEFT,
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
            size: (elem.fontSize || 12) * 2, // docx 使用半点
            font: elem.fontFamily || '宋体',
          }),
        ],
        alignment: elem.alignment ? alignmentMap[elem.alignment] : AlignmentType.LEFT,
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
                size: 24, // 12pt
                font: '宋体',
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
  
  const doc = new Document({
    sections: [{
      properties: {
        page: {
          margin: {
            top: 1440,
            right: 1440,
            bottom: 1440,
            left: 1440,
          },
        },
      },
      children,
    }],
  })
  
  return await Packer.toBlob(doc)
}

export function DocumentProvider({ children }: { children: ReactNode }) {
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

  const createNewDocument = useCallback(async (title: string, content: string, elements?: FormattedElement[]) => {
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
        
        // 如果有 elements，优先尝试使用 ONLYOFFICE Document Builder API
        if (elements && elements.length > 0) {
          console.log('尝试使用 ONLYOFFICE Document Builder API 创建文档，元素:', elements)
          
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
          
          // 如果 Document Builder 失败，回退到 docx 库
          if (!success) {
            console.log('使用 docx 库创建格式化文档')
            const blob = await createFormattedDocxBlob(elements, safeTitle)
            const arrayBuffer = await blob.arrayBuffer()
            // 使用分块方式将 ArrayBuffer 转换为 base64，避免大文件导致的栈溢出
            const base64 = arrayBufferToBase64(arrayBuffer)
            const result = await window.electronAPI.writeBinaryFile(filePath, base64)
            success = result.success
          }
        } else {
          // 纯文本文档，使用 docx 库
          console.log('使用纯文本方式创建文档')
          const blob = await createDocxBlob(content, safeTitle)
          const arrayBuffer = await blob.arrayBuffer()
          // 使用分块方式将 ArrayBuffer 转换为 base64，避免大文件导致的栈溢出
          const base64 = arrayBufferToBase64(arrayBuffer)
          
          // #region agent log
          fetch('http://127.0.0.1:7242/ingest/65f1d8ba-6206-43cb-9f6f-22f7361d7de4',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({location:'DocumentContext.tsx:createNewDocument:beforeWrite',message:'准备写入文件',data:{filePath,base64Length:base64.length,arrayBufferSize:arrayBuffer.byteLength},timestamp:Date.now(),sessionId:'debug-session',hypothesisId:'H4'})}).catch(()=>{});
          // #endregion agent log
          
          const result = await window.electronAPI.writeBinaryFile(filePath, base64)
          
          // #region agent log
          fetch('http://127.0.0.1:7242/ingest/65f1d8ba-6206-43cb-9f6f-22f7361d7de4',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({location:'DocumentContext.tsx:createNewDocument:afterWrite',message:'写入结果',data:{success:result.success,error:result.error,filePath},timestamp:Date.now(),sessionId:'debug-session',hypothesisId:'H4'})}).catch(()=>{});
          // #endregion agent log
          
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
          setPptData(null)
          setDocument({
            title: file.name.replace(/\.[^.]+$/, ''),
            content: '',
            styles: defaultStyles,
            lastModified: new Date(),
          })
        } else if (result.type === 'doc-html') {
          // .doc 文件 - 已经转换为 HTML
          setExcelData(null)
          setDocxData(null)
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
        const blob = await createDocxBlob(documentContentRef.current || document.content, document.title)
        const arrayBuffer = await blob.arrayBuffer()
        const base64 = arrayBufferToBase64(arrayBuffer)
        await window.electronAPI.writeBinaryFile(currentFile.path, base64)
      } else {
        await window.electronAPI.writeFile(currentFile.path, document.content)
      }
      
      // 保存成功后清除该文件的缓存
      fileContentCacheRef.current.delete(currentFile.path)
      setHasUnsavedChanges(false)
    } else {
      const blob = await createDocxBlob(documentContentRef.current || document.content, document.title)
      saveAs(blob, `${document.title}.docx`)
      // 保存成功后清除该文件的缓存
      if (currentFile) {
        fileContentCacheRef.current.delete(currentFile.path)
      }
      setHasUnsavedChanges(false)
    }
  }, [currentFile, document, pendingReplacements.total, extraPendingChanges])

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
      const blob = await createDocxBlob(documentContentRef.current || document.content, document.title)
      saveAs(blob, `${document.title}.docx`)
      setHasUnsavedChanges(false)
    }
  }, [currentFile, document, saveCurrentFile, pendingReplacements.total, extraPendingChanges])

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
  const getTiptapDocumentStructure = useCallback((): string => {
    const content = document.content
    if (!content) return ''
    
    // 解析 HTML 获取结构信息
    const parser = new DOMParser()
    const doc = parser.parseFromString(content, 'text/html')
    
    const elements: string[] = []
    elements.push('【文档结构 - 可用于精确替换的文字】')
    elements.push('⚠️ 替换时 search 必须与下面引号内的文字完全一致！\n')
    
    // 处理表格 - 单独提取，显示每个单元格
    const processTable = (table: HTMLTableElement, tableIndex: number) => {
      const rows = table.querySelectorAll('tr')
      const colCount = rows[0]?.querySelectorAll('td, th').length || 0
      elements.push(`\n📊 表格${tableIndex} (${rows.length}行×${colCount}列):`)
      
      rows.forEach((row, rowIdx) => {
        const cells = row.querySelectorAll('td, th')
        cells.forEach((cell, colIdx) => {
          const cellText = cell.textContent?.trim() || ''
          if (cellText) {
            // 获取单元格样式
            const style = (cell as HTMLElement).getAttribute('style') || ''
            const isBold = cell.querySelector('strong, b') !== null || style.includes('font-weight: bold')
            const bgColor = style.match(/background-color:\s*([^;]+)/)?.[1] || ''
            const borderInfo = style.includes('border') ? '有边框' : ''
            
            const formatInfo = []
            if (isBold) formatInfo.push('粗体')
            if (bgColor) formatInfo.push(`背景:${bgColor}`)
            if (borderInfo) formatInfo.push(borderInfo)
            
            const formatStr = formatInfo.length > 0 ? ` [${formatInfo.join(',')}]` : ''
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
        
        // 获取样式信息
        const style = el.getAttribute('style') || ''
        const isBold = tag === 'strong' || tag === 'b' || style.includes('font-weight: bold')
        const isItalic = tag === 'em' || tag === 'i' || style.includes('font-style: italic')
        const isUnderline = tag === 'u' || style.includes('text-decoration') && style.includes('underline')
        const alignment = style.match(/text-align:\s*(\w+)/)?.[1] || ''
        const fontSize = style.match(/font-size:\s*([^;]+)/)?.[1] || ''
        const fontFamily = style.match(/font-family:\s*([^;]+)/)?.[1] || ''
        const color = style.match(/(?:^|[^-])color:\s*([^;]+)/)?.[1] || ''
        
        if (tag === 'h1') {
          const text = el.textContent?.trim() || ''
          if (text) elements.push(`📌 标题1 [居中,大字]: "${text}"`)
        } else if (tag === 'h2') {
          const text = el.textContent?.trim() || ''
          if (text) elements.push(`📌 标题2: "${text}"`)
        } else if (tag === 'h3') {
          const text = el.textContent?.trim() || ''
          if (text) elements.push(`📌 标题3: "${text}"`)
        } else if (tag === 'p') {
          const text = el.textContent?.trim() || ''
          if (text) {
            const formatInfo = []
            if (isBold) formatInfo.push('粗体')
            if (isItalic) formatInfo.push('斜体')
            if (isUnderline) formatInfo.push('下划线')
            if (alignment && alignment !== 'left') formatInfo.push(alignment)
            if (fontSize) formatInfo.push(`字号:${fontSize}`)
            if (color) formatInfo.push(`颜色:${color}`)
            const formatStr = formatInfo.length > 0 ? ` [${formatInfo.join(',')}]` : ''
            elements.push(`📝 段落${formatStr}: "${text}"`)
          }
        } else if (tag === 'table') {
          processTable(el as HTMLTableElement, tableIndex++)
          processedTables.add(el as HTMLTableElement)
          return // 不再递归处理表格内部
        } else if (tag === 'ul') {
          const items = el.querySelectorAll(':scope > li')
          if (items.length > 0) {
            elements.push(`📋 无序列表 (${items.length}项):`)
            items.forEach((item, i) => {
              const text = item.textContent?.trim() || ''
              if (text) elements.push(`   • "${text}"`)
            })
          }
          return
        } else if (tag === 'ol') {
          const items = el.querySelectorAll(':scope > li')
          if (items.length > 0) {
            elements.push(`📋 有序列表 (${items.length}项):`)
            items.forEach((item, i) => {
              const text = item.textContent?.trim() || ''
              if (text) elements.push(`   ${i+1}. "${text}"`)
            })
          }
          return
        }
      }
      
      // 递归处理子节点
      node.childNodes.forEach(child => walkNodes(child))
    }
    
    walkNodes(doc.body)
    
    elements.push('\n【格式说明】')
    elements.push('- 替换时，search 参数必须从上面的引号内复制精确文字')
    elements.push('- 创建文档时可用的 HTML 格式：')
    elements.push('  - 标题: <h1>标题1</h1>, <h2>标题2</h2>, <h3>标题3</h3>')
    elements.push('  - 粗体: <strong>粗体文字</strong> 或 <b>粗体</b>')
    elements.push('  - 斜体: <em>斜体文字</em> 或 <i>斜体</i>')
    elements.push('  - 下划线: <u>下划线文字</u>')
    elements.push('  - 居中: <p style="text-align: center">居中文字</p>')
    elements.push('  - 颜色: <span style="color: red">红色文字</span>')
    elements.push('  - 表格: <table><tr><td>单元格1</td><td>单元格2</td></tr></table>')
    
    return elements.join('\n')
  }, [document.content])

  // 生成唯一 ID
  const generateDiffId = () => `diff-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`

  // 精准替换文档内容（支持格式保留，支持多个修改共存）
  // 使用 ref 来获取最新内容，解决连续调用时的闭包问题
  const replaceInDocument = useCallback((search: string, replace: string): ReplaceResult => {
    if (!search) {
      return { success: false, count: 0, message: '搜索内容不能为空' }
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
    const contentWithoutDiff = content.replace(/<span class="diff-(old|new)" data-diff-id="[^"]*"[^>]*>[^<]*<\/span>/g, '')
    
    let positions: number[] = []
    let match
    
    // 统计纯文本中的匹配（忽略 HTML 标签）
    const textContent = contentWithoutDiff.replace(/<[^>]+>/g, '')
    const textRegex = new RegExp(escapeRegex(search), 'g')
    while ((match = textRegex.exec(textContent)) !== null) {
      positions.push(match.index)
    }
    
    let count = positions.length
    let useFuzzy = false
    
    // 如果精确匹配没找到，尝试模糊匹配（忽略空格差异）
    if (count === 0) {
      const fuzzySearch = search.replace(/\s+/g, '\\s*')
      positions = []
      
      const fuzzyTextRegex = new RegExp(fuzzySearch, 'g')
      while ((match = fuzzyTextRegex.exec(textContent)) !== null) {
        positions.push(match.index)
      }
      count = positions.length
      useFuzzy = count > 0
    }

    if (count === 0) {
      // 提供更有帮助的错误信息
      const preview = search.length > 30 ? search.substring(0, 30) + '...' : search
      console.log(`[replaceInDocument] ❌ 未找到: "${preview}"`)
      console.log(`[replaceInDocument] 纯文本内容片段: "${textContent.slice(0, 200)}..."`)
      return { 
        success: false, 
        count: 0, 
        message: `未找到「${preview}」，请检查文字是否完全一致（包括标点和空格）` 
      }
    }
    
    console.log(`[replaceInDocument] ✓ 找到 ${count} 处匹配`)

    // 为这次替换生成唯一 ID
    const diffId = generateDiffId()
    
    // 从 HTML 片段中提取格式标签
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
    
    // 创建带唯一 ID 的 Diff 标记（保留原有格式）
    const createDiffHtml = (oldText: string, newText: string, originalHtml: string) => {
      // 将换行符转换为 <br> 标签，确保在 HTML 中正确显示
      const formatText = (text: string) => text.replace(/\n/g, '<br>')
      
      // 提取原有格式标签
      const { openTags, closeTags } = extractFormatTags(originalHtml)
      const openTagsStr = openTags.join('')
      const closeTagsStr = closeTags.join('')
      
      // 保留原有 HTML 中的格式标签用于旧内容显示
      // 新内容也应用相同的格式标签
      const formattedOld = originalHtml // 保留原有 HTML 结构
      const formattedNew = openTagsStr + formatText(newText) + closeTagsStr
      
      return `<span class="diff-old" data-diff-id="${diffId}" style="background-color: #fecaca; color: #b91c1c; text-decoration: line-through; padding: 1px 2px; border-radius: 2px;">${formattedOld}</span><span class="diff-new" data-diff-id="${diffId}" style="background-color: #bbf7d0; color: #15803d; padding: 1px 2px; border-radius: 2px;">${formattedNew}</span>`
    }
    
    // 分段替换策略：将内容按照已有的 diff 标记分割，只在非 diff 区域进行替换
    // 这样可以保留之前的修改标注（使用 [^<]* 代替 .*? 避免灾难性回溯）
    const diffPattern = /<span class="diff-(old|new)" data-diff-id="[^"]*"[^>]*>[^<]*<\/span>/g
    
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
    
    // 核心函数：在 HTML 中查找并替换文本（忽略标签，但保留格式）
    const replaceTextInHtml = (html: string, searchText: string, createReplacement: (matchedText: string, originalHtml: string) => string): string => {
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
      
      if (matches.length === 0) return html
      
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
      
      return result
    }
    
    if (diffMatches.length === 0) {
      // 没有已有标记，直接替换（保留原有格式）
      newContent = replaceTextInHtml(content, search, (matchedText, originalHtml) => createDiffHtml(matchedText, replace, originalHtml))
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
          return replaceTextInHtml(seg.content, search, (matchedText, originalHtml) => createDiffHtml(matchedText, replace, originalHtml))
        }
      }).join('')
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
      searchText: search,
      replaceText: replace,
      count,
      timestamp: Date.now()
    }
    
    setPendingReplacements(prev => ({
      items: [...prev.items, newReplacement],
      total: prev.total + count
    }))
    
    // 同时更新 lastReplacement 以保持向后兼容
    setLastReplacement({
      searchText: search,
      replaceText: replace,
      count,
      timestamp: Date.now(),
      pending: true
    })

    return { 
      success: true, 
      count, 
      message: `成功替换 ${count} 处`,
      searchText: search,
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
    }
  ): ReplaceResult => {
    if (!search) {
      return { success: false, count: 0, message: '搜索内容不能为空' }
    }

    const content = document.content
    const escapeRegex = (str: string) => str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
    const regex = new RegExp(escapeRegex(search), 'g')
    
    // 统计匹配
    const matches = content.match(regex)
    const count = matches ? matches.length : 0
    
    if (count === 0) {
      return { success: false, count: 0, message: `未找到「${search}」` }
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
    
    if (styles.length > 0) {
      formattedReplace = `<span style="${styles.join('; ')}">${formattedReplace}</span>`
    }

    const newContent = content.replace(regex, formattedReplace)
    
    setDocument(prev => ({
      ...prev,
      content: newContent,
      lastModified: new Date(),
    }))
    setDocxData(null)
    setHasUnsavedChanges(true)

    return { 
      success: true, 
      count, 
      message: `成功格式化替换 ${count} 处`,
      searchText: search,
      replaceText: replace
    }
  }, [document.content])
  
  
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
  }, [pendingReplacements, lastReplacement, extraPendingChanges])
  
  
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
  }, [pendingReplacements, lastReplacement, extraPendingChanges])

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
  }, [pendingReplacements, extraPendingChanges])

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
  }, [pendingReplacements, extraPendingChanges])

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

    let newContent = document.content
    const insertHtml = `<p>${content}</p>`

    if (position === 'start') {
      // 在开头插入
      newContent = insertHtml + newContent
    } else if (position === 'end') {
      // 在末尾插入
      newContent = newContent + insertHtml
    } else if (position.startsWith('after:')) {
      // 在指定文字后插入
      const anchor = position.slice(6)
      if (!anchor) {
        return { success: false, message: '锚点文字不能为空' }
      }
      
      // 查找锚点位置
      const anchorIndex = newContent.indexOf(anchor)
      if (anchorIndex === -1) {
        return { success: false, message: `未找到「${anchor}」` }
      }
      
      // 在锚点后插入（找到锚点所在标签的结束位置）
      const afterAnchor = anchorIndex + anchor.length
      // 查找下一个标签结束位置
      const nextTagEnd = newContent.indexOf('>', afterAnchor)
      const insertPos = nextTagEnd !== -1 ? nextTagEnd + 1 : afterAnchor
      
      newContent = newContent.slice(0, insertPos) + insertHtml + newContent.slice(insertPos)
    } else {
      return { success: false, message: `无效的位置参数: ${position}` }
    }

    setDocument(prev => ({
      ...prev,
      content: newContent,
      lastModified: new Date(),
    }))
    setDocxData(null)
    setHasUnsavedChanges(true)

    return { success: true, message: `已在 ${position === 'start' ? '开头' : position === 'end' ? '末尾' : position} 插入内容` }
  }, [document.content])

  // 删除文档中的内容
  const deleteInDocument = useCallback((target: string): { success: boolean; count: number; message: string } => {
    if (!target) {
      return { success: false, count: 0, message: '删除目标不能为空' }
    }

    const content = document.content
    
    // 统计匹配次数
    const regex = new RegExp(target.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g')
    const matches = content.match(regex)
    const count = matches ? matches.length : 0

    if (count === 0) {
      return { success: false, count: 0, message: `未找到「${target}」` }
    }

    // 执行删除
    const newContent = content.replace(regex, '')

    setDocument(prev => ({
      ...prev,
      content: newContent,
      lastModified: new Date(),
    }))
    setDocxData(null)
    setHasUnsavedChanges(true)

    return { success: true, count, message: `成功删除 ${count} 处「${target}」` }
  }, [document.content])
  
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

      let html = documentContentRef.current || ''
      const parser = new DOMParser()
      const doc = parser.parseFromString(html, 'text/html')

      const created: PendingChange[] = []
      const genId = () => `diff-${Date.now()}-${Math.random().toString(36).slice(2, 10)}`

      const markBlockPair = (oldEl: HTMLElement, newEl: HTMLElement, diffId: string) => {
        oldEl.setAttribute('data-diff-id', diffId)
        oldEl.setAttribute('data-diff-role', 'old')
        oldEl.setAttribute('data-diff-kind', 'block')
        newEl.setAttribute('data-diff-id', diffId)
        newEl.setAttribute('data-diff-role', 'new')
        newEl.setAttribute('data-diff-kind', 'block')
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
          pageBreak.setAttribute('style', 'page-break-before: always; border-top: 2px dashed #999; margin: 20px 0; padding: 10px 0; text-align: center; color: #999; font-size: 12px;')
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
                cell.setAttribute('style', 'border: 1px solid #ccc; padding: 8px;')
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
                td.setAttribute('style', 'border: 1px solid #ccc; padding: 8px;')
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
                cell.setAttribute('style', 'border: 1px solid #ccc; padding: 8px;')
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
          }
          // 其他表格操作（delete_row, delete_column, merge_cells）可以类似实现
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
            
            if (!newWidth) continue
            
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
              const prevStyle = targetImg.getAttribute('style') || ''
              const newStyle = prevStyle.replace(/max-width\s*:\s*[^;]+;?/gi, '') + `; max-width: ${newWidth};`
              
              targetImg.setAttribute('data-diff-id', diffId)
              targetImg.setAttribute('data-diff-role', 'new')
              targetImg.setAttribute('style', newStyle)
              
              created.push({
                id: diffId,
                kind: 'image_edit',
                scope: 'selection',
                summary: `调整图片大小为 ${newWidth}`,
                afterPreview: newWidth,
                stats: { matches: 1 },
                timestamp: Date.now(),
                meta: { op, action },
              })
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
          tocContainer.setAttribute('style', 'margin: 1em 0; padding: 1em; border: 1px solid #ddd; background: #f9f9f9;')
          
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
      }

      const newHtml = doc.body.innerHTML
      if (newHtml !== (documentContentRef.current || '')) {
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

      return {
        success: true,
        message: `已生成修订：${created.length} 条。请在底部或“修订面板”中逐条确认。`,
        data: { created: created.length },
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
        refreshExcelData,
        lastReplacement,
        pendingChanges: [
          ...pendingReplacements.items.map((item) => ({
            id: item.id,
            kind: 'replace_text' as const,
            scope: 'document' as const,
            summary: `替换 ${item.count} 处`,
            beforePreview: item.searchText,
            afterPreview: item.replaceText,
            stats: { matches: item.count },
            timestamp: item.timestamp,
            meta: {
              searchText: item.searchText,
              replaceText: item.replaceText,
              count: item.count,
            },
          })),
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
        uploadDocxFile,
        saveDocument,
        applyAIEdit,
        replaceInDocument,
        insertInDocument,
        deleteInDocument,
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
