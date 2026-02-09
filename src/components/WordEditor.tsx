import { useEffect, useState, useCallback, useRef, useMemo } from 'react'
import { useEditor, EditorContent } from '@tiptap/react'
import StarterKit from '@tiptap/starter-kit'
import { Paragraph } from '@tiptap/extension-paragraph'
import { Heading } from '@tiptap/extension-heading'
import { Underline } from '@tiptap/extension-underline'
import { TextAlign } from '@tiptap/extension-text-align'
import { TextStyle } from '@tiptap/extension-text-style'
import { Color } from '@tiptap/extension-color'
import { Highlight } from '@tiptap/extension-highlight'
import { Table } from '@tiptap/extension-table'
import { TableRow } from '@tiptap/extension-table-row'
import { TableHeader } from '@tiptap/extension-table-header'
import { TableCell } from '@tiptap/extension-table-cell'
import { DocxImage } from '../extensions/DocxImage'
import { Link } from '@tiptap/extension-link'
import { Placeholder } from '@tiptap/extension-placeholder'
import { FontFamily } from '@tiptap/extension-font-family'
import { Subscript } from '@tiptap/extension-subscript'
import { Superscript } from '@tiptap/extension-superscript'
import { DiffOld, DiffNew } from '../extensions/DiffMark'
import { DiffBlock } from '../extensions/DiffBlock'
import { TrackInsert, TrackDelete } from '../extensions/TrackChanges'
import { CommentMark } from '../extensions/CommentMark'
import { motion, AnimatePresence, animate } from 'framer-motion'
import { useDocument } from '../context/DocumentContext'
import { useAI } from '../context/AIContext'
import { useComments } from '../context/CommentContext'
import { 
  Save,
  Download,
  FileText,
  Loader2,
  Check,
  X,
  Sparkles,
  Eye,
  MessageSquare,
  Layout,
  Grid,
} from 'lucide-react'
import { parseDocxToHtml, parseDocxComplete, type DocxParseResult } from '../utils/docxParser'
import { parseDocxComments } from '../utils/docxCommentsParser'
import { commonSystemFonts, normalizeFontFamilyForDisplay } from '../fonts/fontManifest'
import { extractFontFamiliesFromHtml, loadFontsByNames } from '../fonts/loadWorkspaceFonts'
import InlineEditPopup from './InlineEditPopup'
import ContextMenu from './ContextMenu'
import ExcelPreview from './ExcelPreview'
import PptPreviewHtml from './PptPreviewHtml'
import RevisionPanel from './RevisionPanel'
import CommentPanel from './CommentPanel'
import WordRibbon, { type WordRibbonTab } from './WordRibbon'

// 控制条动画变体
const floatingBarVariants = {
  hidden: { 
    opacity: 0, 
    y: 20, 
    scale: 0.95,
    filter: 'blur(4px)'
  },
  visible: { 
    opacity: 1, 
    y: 0, 
    scale: 1,
    filter: 'blur(0px)',
    transition: { 
      duration: 0.25, 
      ease: [0.25, 0.46, 0.45, 0.94] as const // easeOut
    }
  },
  exit: { 
    opacity: 0, 
    y: -10, 
    scale: 0.95,
    filter: 'blur(4px)',
    transition: { 
      duration: 0.15, 
      ease: [0.55, 0.06, 0.68, 0.19] as const // easeIn
    }
  }
}

// Markdown 转 HTML 函数
function markdownToHtml(markdown: string): string {
  const lines = markdown.split('\n')
  const result: string[] = []
  let inList = false
  let listType = ''
  let currentParagraph: string[] = []

  const flushParagraph = () => {
    if (currentParagraph.length > 0) {
      const text = currentParagraph.join(' ')
      // 处理行内格式
      const formatted = formatInline(text)
      result.push(`<p style="text-indent: 2em; margin: 0.8em 0; line-height: 1.8;">${formatted}</p>`)
      currentParagraph = []
    }
  }

  const flushList = () => {
    if (inList) {
      result.push(listType === 'ul' ? '</ul>' : '</ol>')
      inList = false
      listType = ''
    }
  }

  const formatInline = (text: string): string => {
    // 粗斜体
    text = text.replace(/\*\*\*(.+?)\*\*\*/g, '<strong><em>$1</em></strong>')
    // 粗体
    text = text.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
    // 斜体
    text = text.replace(/\*(.+?)\*/g, '<em>$1</em>')
    // 下划线格式的粗体
    text = text.replace(/__(.+?)__/g, '<strong>$1</strong>')
    text = text.replace(/_(.+?)_/g, '<em>$1</em>')
    return text
  }

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i]
    const trimmedLine = line.trim()

    // 空行
    if (!trimmedLine) {
      flushParagraph()
      flushList()
      continue
    }

    // 分隔线
    if (/^(-{3,}|\*{3,}|_{3,})$/.test(trimmedLine)) {
      flushParagraph()
      flushList()
      result.push('<hr style="border: none; border-top: 1px solid var(--word-rule); margin: 1.5em 0;">')
      continue
    }

    // 标题
    if (trimmedLine.startsWith('### ')) {
      flushParagraph()
      flushList()
      const text = formatInline(trimmedLine.slice(4))
      result.push(`<h3 style="font-size: 14pt; font-weight: bold; margin: 1em 0 0.5em 0;">${text}</h3>`)
      continue
    }
    if (trimmedLine.startsWith('## ')) {
      flushParagraph()
      flushList()
      const text = formatInline(trimmedLine.slice(3))
      result.push(`<h2 style="font-size: 16pt; font-weight: bold; margin: 1.2em 0 0.6em 0;">${text}</h2>`)
      continue
    }
    if (trimmedLine.startsWith('# ')) {
      flushParagraph()
      flushList()
      const text = formatInline(trimmedLine.slice(2))
      result.push(`<h1 style="font-size: 22pt; font-weight: bold; text-align: center; margin: 0.5em 0 1em 0;">${text}</h1>`)
      continue
    }

    // 无序列表
    if (/^[-*] /.test(trimmedLine)) {
      flushParagraph()
      if (!inList || listType !== 'ul') {
        flushList()
        result.push('<ul style="margin: 0.5em 0; padding-left: 2em;">')
        inList = true
        listType = 'ul'
      }
      const text = formatInline(trimmedLine.slice(2))
      result.push(`<li style="margin: 0.3em 0;">${text}</li>`)
      continue
    }

    // 有序列表
    if (/^\d+\. /.test(trimmedLine)) {
      flushParagraph()
      if (!inList || listType !== 'ol') {
        flushList()
        result.push('<ol style="margin: 0.5em 0; padding-left: 2em;">')
        inList = true
        listType = 'ol'
      }
      const text = formatInline(trimmedLine.replace(/^\d+\. /, ''))
      result.push(`<li style="margin: 0.3em 0;">${text}</li>`)
      continue
    }

    // 引用
    if (trimmedLine.startsWith('> ')) {
      flushParagraph()
      flushList()
      const text = formatInline(trimmedLine.slice(2))
      result.push(`<blockquote style="border-left: 3px solid var(--word-rule); padding-left: 1em; margin: 0.5em 0; color: var(--word-ink-muted);">${text}</blockquote>`)
      continue
    }

    // 普通段落
    flushList()
    currentParagraph.push(trimmedLine)
  }

  flushParagraph()
  flushList()

  return result.join('\n')
}

// 自定义字体大小扩展
import { Extension } from '@tiptap/core'
import { Plugin, PluginKey } from 'prosemirror-state'
import { Decoration, DecorationSet } from 'prosemirror-view'
import { DOMSerializer } from 'prosemirror-model'

const FontSize = Extension.create({
  name: 'fontSize',
  addOptions() {
    return { types: ['textStyle'] }
  },
  addGlobalAttributes() {
    return [{
      types: this.options.types,
      attributes: {
        fontSize: {
          default: null,
          parseHTML: element => element.style.fontSize?.replace(/['"]+/g, ''),
          renderHTML: attributes => {
            if (!attributes.fontSize) return {}
            return { style: `font-size: ${attributes.fontSize}` }
          },
        },
      },
    }]
  },
})

const PT_PER_INCH = 72
const CM_PER_INCH = 2.54

const ptToCm = (pt: number) => (pt / PT_PER_INCH) * CM_PER_INCH

const formatCm = (pt: number) => {
  const cm = ptToCm(pt)
  const trimmed = cm.toFixed(4).replace(/\.?0+$/, '')
  return `${trimmed}cm`
}

const DocxTable = Table.extend({
  addAttributes() {
    return {
      ...this.parent?.(),
      style: {
        default: null,
        parseHTML: (element) => element.getAttribute('style'),
        renderHTML: (attrs) => (attrs.style ? { style: attrs.style } : {}),
      },
      'data-tbl-grid': {
        default: null,
        parseHTML: (element) => element.getAttribute('data-tbl-grid'),
        renderHTML: (attrs) => (attrs['data-tbl-grid'] ? { 'data-tbl-grid': attrs['data-tbl-grid'] } : {}),
      },
      'data-tbl-grid-total': {
        default: null,
        parseHTML: (element) => element.getAttribute('data-tbl-grid-total'),
        renderHTML: (attrs) =>
          attrs['data-tbl-grid-total'] ? { 'data-tbl-grid-total': attrs['data-tbl-grid-total'] } : {},
      },
      'data-tbl-layout': {
        default: null,
        parseHTML: (element) => element.getAttribute('data-tbl-layout'),
        renderHTML: (attrs) => (attrs['data-tbl-layout'] ? { 'data-tbl-layout': attrs['data-tbl-layout'] } : {}),
      },
    }
  },
})

const DocxTableRow = TableRow.extend({
  addAttributes() {
    return {
      ...this.parent?.(),
      style: {
        default: null,
        parseHTML: (element) => element.getAttribute('style'),
        renderHTML: (attrs) => (attrs.style ? { style: attrs.style } : {}),
      },
    }
  },
})

const DocxTableHeader = TableHeader.extend({
  addAttributes() {
    return {
      ...this.parent?.(),
      style: {
        default: null,
        parseHTML: (element) => element.getAttribute('style'),
        renderHTML: (attrs) => (attrs.style ? { style: attrs.style } : {}),
      },
    }
  },
})

const DocxTableCell = TableCell.extend({
  addAttributes() {
    return {
      ...this.parent?.(),
      style: {
        default: null,
        parseHTML: (element) => element.getAttribute('style'),
        renderHTML: (attrs) => (attrs.style ? { style: attrs.style } : {}),
      },
    }
  },
})

const DocxParagraph = Paragraph.extend({
  addAttributes() {
    return {
      ...this.parent?.(),
      style: {
        default: null,
        parseHTML: (element) => element.getAttribute('style'),
        renderHTML: (attrs) => (attrs.style ? { style: attrs.style } : {}),
      },
      class: {
        default: null,
        parseHTML: (element) => element.getAttribute('class'),
        renderHTML: (attrs) => (attrs.class ? { class: attrs.class } : {}),
      },
      'data-docx-tab-leader': {
        default: null,
        parseHTML: (element) => element.getAttribute('data-docx-tab-leader'),
        renderHTML: (attrs) =>
          attrs['data-docx-tab-leader'] ? { 'data-docx-tab-leader': attrs['data-docx-tab-leader'] } : {},
      },
    }
  },
})

const DocxHeading = Heading.extend({
  addAttributes() {
    return {
      ...this.parent?.(),
      style: {
        default: null,
        parseHTML: (element) => element.getAttribute('style'),
        renderHTML: (attrs) => (attrs.style ? { style: attrs.style } : {}),
      },
      class: {
        default: null,
        parseHTML: (element) => element.getAttribute('class'),
        renderHTML: (attrs) => (attrs.class ? { class: attrs.class } : {}),
      },
    }
  },
})

const detectPaperSize = (
  widthPt: number,
  heightPt: number
): 'A4' | 'A3' | 'Letter' | 'Legal' | 'custom' => {
  const shortSide = Math.min(widthPt, heightPt)
  const longSide = Math.max(widthPt, heightPt)
  const near = (a: number, b: number) => Math.abs(a - b) <= 5

  if (near(shortSide, 595) && near(longSide, 842)) return 'A4'
  if (near(shortSide, 842) && near(longSide, 1191)) return 'A3'
  if (near(shortSide, 612) && near(longSide, 792)) return 'Letter'
  if (near(shortSide, 612) && near(longSide, 1008)) return 'Legal'
  return 'custom'
}

// ========== Word 打印布局（分页视觉 + 每页顶端留白 + 页间距） ==========
// 说明：使用 ProseMirror Decorations 在不改动文档内容的情况下插入“分页占位块”，
// 让内容真实换页，并保持每页的上/下边距留白（更接近 Word/OnlyOffice 的打印布局）。
const PagedLayout = Extension.create({
  name: 'pagedLayout',
  addProseMirrorPlugins() {
    const key = new PluginKey('pagedLayout')
    const PX_PER_MM = 96 / 25.4
    const PAGE_HEIGHT_PX = 297 * PX_PER_MM
    const TOP_MARGIN_PX = 25.4 * PX_PER_MM // 2.54cm
    const BOTTOM_MARGIN_PX = 25.4 * PX_PER_MM // 2.54cm
    const PAGE_GAP_PX = 32 // 页间距（接近原版显示，可继续微调）

    const buildDecorations = (view: any) => {
      const root = view.dom as HTMLElement
      if (!root) return DecorationSet.empty
      const fonts = (document as any).fonts as FontFaceSet | undefined
      if (fonts && fonts.status !== 'loaded') return DecorationSet.empty
      const pageEl = root.closest('.word-page') as HTMLElement | null
      if (pageEl?.classList.contains('web')) return DecorationSet.empty

      const rootStyle = window.getComputedStyle(root)
      const pageStyle = pageEl ? window.getComputedStyle(pageEl) : null
      const parsePx = (value: string) => {
        const parsed = Number.parseFloat(value)
        return Number.isFinite(parsed) ? parsed : 0
      }
      const rootRect = root.getBoundingClientRect()
      const rawScale = root.offsetWidth > 0 ? rootRect.width / root.offsetWidth : 1
      const scale = Number.isFinite(rawScale) && rawScale > 0 ? rawScale : 1
      const epsilon = 0.5
      const topMarginPx = parsePx(rootStyle.paddingTop) || TOP_MARGIN_PX
      const bottomMarginPx = parsePx(rootStyle.paddingBottom) || BOTTOM_MARGIN_PX
      const pageHeightPx = pageStyle ? parsePx(pageStyle.minHeight) : PAGE_HEIGHT_PX
      const resolvedPageHeight = pageHeightPx || PAGE_HEIGHT_PX
      const contentHeightPx = Math.max(0, resolvedPageHeight - topMarginPx - bottomMarginPx)
      const lineHeightPx =
        parsePx(rootStyle.lineHeight) ||
        (parsePx(rootStyle.fontSize) ? parsePx(rootStyle.fontSize) * 1.5 : 0) ||
        20 // fallback ~15pt line

      const blocks = Array.from(root.children).filter((el) => {
        const rectHeight = (el as HTMLElement).getBoundingClientRect().height
        const h = rectHeight > 0 ? rectHeight / scale : (el as HTMLElement).offsetHeight
        if (h <= epsilon) return false
        return !(
          (el as HTMLElement).classList?.contains('pm-page-break') ||
          (el as HTMLElement).classList?.contains('pm-page-end')
        )
      }) as HTMLElement[]

      const decorations: any[] = []
      let y = 0 // 内容区内已占用高度（不含页边距）
      let prevMarginBottom = 0

      for (const el of blocks) {
        const rectHeight = el.getBoundingClientRect().height
        const h = rectHeight > 0 ? rectHeight / scale : el.offsetHeight
        if (h <= epsilon) continue
        const style = window.getComputedStyle(el)
        const marginTop = parsePx(style.marginTop)
        const marginBottom = parsePx(style.marginBottom)
        const collapsedMargin = Math.max(prevMarginBottom, marginTop)
        const blockHeight = collapsedMargin + h

        // 只要“下一个块”会进入底边距区域，就在它前面插入换页占位
        // 用每个段落自身的行高作为容差（中文28pt≈37px，英文12pt≈16px）
        // 允许约2行溢出进入下边距，更接近 Word 的分页行为
        const blockLineH = parsePx(style.lineHeight) || lineHeightPx
        const wouldOverflow = y + blockHeight > contentHeightPx + blockLineH * 2

        // 避免在一页开头反复插入空页（遇到超大表格/块时允许其自然溢出）
        const atPageStart = y <= epsilon

        if (wouldOverflow && !atPageStart) {
          // 用 pos-1 让 widget 放在段落节点外部（作为 .word-editor-content 的直接子节点）
          // 这样 CSS 的 calc(100% + margins) 才能基于正确的父宽度计算
          const rawPos = view.posAtDOM(el, 0)
          const pos = Math.max(0, rawPos - 1)
          // 正确计算剩余页面空间：内容区剩余 + 底边距，但要减去已溢出的部分
          const fillToEndOfPage = Math.max(0, contentHeightPx + bottomMarginPx - y)

          const widget = Decoration.widget(
            pos,
            () => {
              const dom = document.createElement('div')
              dom.className = 'pm-page-break'
              dom.style.setProperty('--fill', `${fillToEndOfPage}px`)
              dom.style.setProperty('--gap', `${PAGE_GAP_PX}px`)
              dom.style.setProperty('--top', `${topMarginPx}px`)
              dom.style.height = `${fillToEndOfPage + PAGE_GAP_PX + topMarginPx}px`
              return dom
            },
            { key: `pm-pb-${pos}`, side: -1 }
          )
          decorations.push(widget)

          // 新的一页
          y = 0
          prevMarginBottom = 0
          const nextMargin = Math.max(prevMarginBottom, marginTop)
          y = nextMargin + h
          prevMarginBottom = marginBottom
        } else {
          y += blockHeight
          prevMarginBottom = marginBottom
        }
      }

      y += prevMarginBottom

      const fillToEndOfContent = Math.max(0, contentHeightPx - y)
      if (fillToEndOfContent > epsilon) {
        const endPos = view.state.doc.content.size
        const endWidget = Decoration.widget(
          endPos,
          () => {
            const dom = document.createElement('div')
            dom.className = 'pm-page-end'
            dom.style.height = `${fillToEndOfContent}px`
            return dom
          },
          { key: `pm-pe-${endPos}` }
        )
        decorations.push(endWidget)
      }

      return DecorationSet.create(view.state.doc, decorations)
    }

    return [
      new Plugin({
        key,
        state: {
          init: () => DecorationSet.empty,
          apply(tr, old) {
            const meta = tr.getMeta(key)
            if (meta?.decorations) return meta.decorations
            return old.map(tr.mapping, tr.doc)
          },
        },
        props: {
          decorations(state) {
            return key.getState(state)
          },
        },
        view(view) {
          let raf = 0
          const schedule = () => {
            cancelAnimationFrame(raf)
            raf = requestAnimationFrame(() => {
              const next = buildDecorations(view)
              const current = key.getState(view.state)
              // 避免无意义 dispatch
              if (current && next && current.eq(next)) return
              view.dispatch(view.state.tr.setMeta(key, { decorations: next }))
            })
          }
          schedule()
          return {
            update: schedule,
            destroy() {
              cancelAnimationFrame(raf)
            },
          }
        },
      }),
    ]
  },
})

const DIFF_BLOCK_SELECTOR = 'p, li, div, td, th, h1, h2, h3, h4, h5, h6, blockquote, tr'
const REVEAL_BLOCK_SELECTOR = 'p, h1, h2, h3, h4, h5, h6, li, blockquote, table, ul, ol, pre, div'

export default function WordEditor() {
  const {
    document,
    currentFile,
    docxData,
    excelData,
    pptData,
    hasUnsavedChanges,
    saveDocument,
    updateContent,
    lastReplacement,
    pendingChanges,
    pendingChangesTotal,
    acceptAllChanges,
    rejectAllChanges,
    addPendingReplacementItem,
    docEntryAnimationKey,
    refreshExcelData,
    pageSetup,
    setPageSetup,
    editorMode,
    setEditorMode,
    headerFooterSetup,
    typographyProfile,
  } = useDocument()
  const { getCompletion, cancelCompletion, isCompleting, settings } = useAI()
  const { setComments, addComment, replyToComment, setSelectedCommentId } = useComments()
  const [isLoading, setIsLoading] = useState(false)
  const [currentFontSize, setCurrentFontSize] = useState('12pt')
  const [currentFontFamily, setCurrentFontFamily] = useState('仿宋')
  const [isConfirming, setIsConfirming] = useState(false)
  const [showRevisionPanel, setShowRevisionPanel] = useState(false)
  const [showCommentPanel, setShowCommentPanel] = useState(false)
  
  // 补全相关状态
  const [completionSuggestion, setCompletionSuggestion] = useState<string | null>(null)
  const [showCompletion, setShowCompletion] = useState(false)
  const lastTextRef = useRef<string>('')
  const animatedDiffIdsRef = useRef<Set<string>>(new Set())
  const docEntryAnimationRef = useRef<number>(0)
  
  // 跟踪上一次同步的文档内容，避免重复同步
  const lastSyncedContentRef = useRef<string>('')
  const lastLoadedCommentsKeyRef = useRef<string>('')
  
  // 内联编辑和右键菜单状态
  const [showInlineEdit, setShowInlineEdit] = useState(false)
  const [inlineEditPosition, setInlineEditPosition] = useState({ x: 0, y: 0 })
  const [selectedTextForEdit, setSelectedTextForEdit] = useState('')
  const [selectedHtmlForEdit, setSelectedHtmlForEdit] = useState('') // 保存选区的 HTML 格式
  const [inlineEditRange, setInlineEditRange] = useState<{ from: number; to: number } | null>(null)
  const [showContextMenu, setShowContextMenu] = useState(false)
  const [contextMenuPosition, setContextMenuPosition] = useState({ x: 0, y: 0 })
  const [isAIProcessing, setIsAIProcessing] = useState(false)
  
  // 缩放状态
  const [zoomLevel, setZoomLevel] = useState(100) // 百分比
  const editorContainerRef = useRef<HTMLDivElement>(null)
  
  // DOCX 页眉/页脚数据（从解析器获取，单独渲染）
  const [docxHeaderFooter, setDocxHeaderFooter] = useState<{
    headerHtml: string
    footerHtml: string
    headerStyle?: { fontSize?: string; alignment?: string; color?: string }
    footerStyle?: { fontSize?: string; alignment?: string; color?: string }
  } | null>(null)

  // --- Docx floating image positioning (wp:anchor) ---
  // NOTE: implementation is defined after `const editor = useEditor(...)` to avoid TDZ issues.
  const positionDocxFloatingImages = useCallback(() => {}, [])
  const [viewMode, setViewMode] = useState<'print' | 'web'>('print')
  const [ribbonTab, setRibbonTab] = useState<WordRibbonTab>('home')
  const [docStats, setDocStats] = useState(() => ({
    words: 0,
    chars: 0,
    selectionChars: 0,
    selectionFrom: 0,
    selectionTo: 0,
  }))
  const [pageStats, setPageStats] = useState(() => ({ current: 1, total: 1 }))

  const pageSize = useMemo(() => {
    let width = '210mm'
    let height = '297mm'

    switch (pageSetup.paperSize) {
      case 'A3':
        width = '297mm'
        height = '420mm'
        break
      case 'Letter':
        width = '8.5in'
        height = '11in'
        break
      case 'Legal':
        width = '8.5in'
        height = '14in'
        break
      case 'custom':
        width = pageSetup.customWidth || width
        height = pageSetup.customHeight || height
        break
      case 'A4':
      default:
        break
    }

    if (pageSetup.orientation === 'landscape') {
      return { width: height, height: width }
    }

    return { width, height }
  }, [pageSetup])

  const [bundledFontFamilies, setBundledFontFamilies] = useState<string[]>(() => {
    try {
      const fonts = (window as any)?.__onlyofficeBundledFonts as Array<{ family?: string }> | undefined
      if (!fonts) return []
      return Array.from(new Set(fonts.map((f) => (f?.family || '').trim()).filter(Boolean)))
    } catch {
      return []
    }
  })

  useEffect(() => {
    const handler = (event: Event) => {
      const detail = (event as CustomEvent<{ fonts?: Array<{ family?: string }> }>).detail
      const families = Array.isArray(detail?.fonts)
        ? detail!.fonts.map((f) => (f?.family || '').trim()).filter(Boolean)
        : []
      if (families.length) {
        setBundledFontFamilies(Array.from(new Set(families)))
      }
    }
    window.addEventListener('onlyoffice-fonts-loaded', handler as EventListener)
    return () => window.removeEventListener('onlyoffice-fonts-loaded', handler as EventListener)
  }, [])

  const docDefaultFontFamily = useMemo(() => {
    const fromProfile =
      typographyProfile?.normal?.fontEastAsia ||
      typographyProfile?.normal?.fontAscii ||
      typographyProfile?.normal?.fontHAnsi ||
      ''
    return (fromProfile || '').trim()
  }, [typographyProfile])

  const fontFamilyOptions = useMemo(() => {
    const set = new Set<string>()
    if (docDefaultFontFamily) set.add(docDefaultFontFamily)
    bundledFontFamilies.forEach((f) => set.add(f))
    commonSystemFonts.forEach((f) => set.add(f))
    return Array.from(set).filter(Boolean)
  }, [docDefaultFontFamily, bundledFontFamilies])

  const escapeHtml = useCallback((text: string) => {
    return (text ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;')
  }, [])

  const textToHtml = useCallback((text: string) => {
    return escapeHtml(text).replace(/\n/g, '<br>')
  }, [escapeHtml])

  // 从选区获取当前字体信息
  const updateCurrentStyles = (editor: any) => {
    if (!editor) return

    try {
      // 获取当前选区位置
      const { from } = editor.state.selection
      const resolvedPos = editor.state.doc.resolve(from)
      const node = resolvedPos.parent
      
      // 检查是否是标题，设置对应字号
      if (node.type.name === 'heading') {
        const level = node.attrs.level
        if (level === 1) {
          setCurrentFontSize('22pt')
          setCurrentFontFamily('黑体')
        } else if (level === 2) {
          setCurrentFontSize('16pt')
          setCurrentFontFamily('黑体')
        } else if (level === 3) {
          setCurrentFontSize('14pt')
          setCurrentFontFamily('黑体')
        }
        return
      }

      // 尝试从 textStyle 属性获取
      const textStyleAttrs = editor.getAttributes('textStyle')
      
      // 辅助函数：从 DOM 元素向上查找字体名
      // 优先使用 data-font-name 属性（原始字体名），其次使用 font-family 样式
      const findFontFromElement = (startElement: HTMLElement | null): string | null => {
        let element = startElement
        while (element && element !== window.document.body) {
          // 优先检查 data-font-name 属性（DOCX 解析时存储的原始字体名）
          const fontName = element.getAttribute('data-font-name')
          if (fontName) {
            return fontName
          }
          
          // 其次检查内联 style 属性中的 font-family
          const inlineStyle = element.getAttribute('style')
          if (inlineStyle) {
            const match = inlineStyle.match(/font-family:\s*([^;]+)/i)
            if (match && match[1]) {
              const fontValue = match[1].trim()
              // 排除 CSS 变量
              if (!fontValue.startsWith('var(')) {
                return normalizeFontFamilyForDisplay(fontValue)
              }
            }
          }
          element = element.parentElement
        }
        return null
      }
      
      // 字体大小
      if (textStyleAttrs.fontSize) {
        setCurrentFontSize(textStyleAttrs.fontSize)
      } else {
        // 尝试从 DOM 获取实际渲染的字体大小
        const domSelection = window.getSelection()
        if (domSelection && domSelection.rangeCount > 0) {
          const range = domSelection.getRangeAt(0)
          let element = range.startContainer as HTMLElement
          if (element.nodeType === Node.TEXT_NODE) {
            element = element.parentElement as HTMLElement
          }
          if (element) {
            // 首先尝试从内联样式获取
            let fontSize: string | null = null
            let el: HTMLElement | null = element
            while (el && el !== window.document.body) {
              const inlineStyle = el.getAttribute('style')
              if (inlineStyle) {
                const match = inlineStyle.match(/font-size:\s*([^;]+)/i)
                if (match && match[1]) {
                  fontSize = match[1].trim()
                  break
                }
              }
              el = el.parentElement
            }
            
            if (fontSize) {
              setCurrentFontSize(fontSize)
            } else {
              // 回退到 computedStyle
              const computedStyle = window.getComputedStyle(element)
              const computedFontSize = computedStyle.fontSize
              if (computedFontSize) {
                // 转换 px 到 pt (1pt = 1.333px)
                const pxValue = parseFloat(computedFontSize)
                const ptValue = Math.round(pxValue / 1.333)
                setCurrentFontSize(`${ptValue}pt`)
              }
            }
          }
        } else {
          setCurrentFontSize('12pt')
        }
      }

      // 字体
      if (textStyleAttrs.fontFamily) {
        const normalized = normalizeFontFamilyForDisplay(textStyleAttrs.fontFamily)
        setCurrentFontFamily(normalized || docDefaultFontFamily || '宋体')
      } else {
        // 尝试从 DOM 获取 - 优先从内联 style 获取，避免继承问题
        const domSelection = window.getSelection()
        if (domSelection && domSelection.rangeCount > 0) {
          const range = domSelection.getRangeAt(0)
          let element = range.startContainer as HTMLElement
          if (element.nodeType === Node.TEXT_NODE) {
            element = element.parentElement as HTMLElement
          }
          if (element) {
            // 优先从 data-font-name 或内联样式获取字体（更准确）
            const fontFromElement = findFontFromElement(element)
            if (fontFromElement) {
              setCurrentFontFamily(fontFromElement)
            } else {
              // 回退到 computedStyle
              const computedStyle = window.getComputedStyle(element)
              const fontFamily = computedStyle.fontFamily
              if (fontFamily && !fontFamily.startsWith('var(')) {
                // 提取第一个字体名称
                const firstFont = normalizeFontFamilyForDisplay(fontFamily)
                setCurrentFontFamily(firstFont || docDefaultFontFamily || '宋体')
              } else {
                setCurrentFontFamily(docDefaultFontFamily || '宋体')
              }
            }
          }
        } else {
          setCurrentFontFamily(docDefaultFontFamily || '宋体')
        }
      }
    } catch (error) {
      // 忽略错误，使用默认值
      setCurrentFontSize('12pt')
      setCurrentFontFamily(docDefaultFontFamily || '宋体')
    }
  }

  const editor = useEditor({
    extensions: [
      StarterKit.configure({
        paragraph: false,
        heading: false,
      }),
      DocxParagraph,
      DocxHeading.configure({ levels: [1, 2, 3, 4, 5, 6] }),
      Underline,
      TextAlign.configure({
        types: ['heading', 'paragraph'],
      }),
      TextStyle,
      FontFamily,
      FontSize,
      Color,
      Highlight.configure({ multicolor: true }),
      Subscript,
      Superscript,
      DocxTable.configure({
        resizable: true,
      }),
      DocxTableRow,
      DocxTableHeader,
      DocxTableCell,
      DocxImage,
      Link.configure({
        openOnClick: false,
      }),
      Placeholder.configure({
        placeholder: '开始输入内容...',
      }),
      // 分页打印布局（多页间距 + 每页顶部留白）
      PagedLayout,
      // Diff 标记扩展 - 用于显示 AI 修改的差异
      DiffOld,
      DiffNew,
      // Diff 块级属性（段落/标题样式修订）
      DiffBlock,
      // Word 原生修订（Track Changes）
      TrackInsert,
      TrackDelete,
      // Word 批注范围
      CommentMark,
    ],
    content: '',
    editorProps: {
      attributes: {
        class: 'word-editor-content',
      },
    },
    onUpdate: ({ editor }) => {
      // 同步内容到 context
      const html = editor.getHTML()
      
      // 更新 lastSyncedContentRef 以防止无限循环
      // （Tiptap 会标准化 HTML，导致输入输出不一致）
      lastSyncedContentRef.current = html
      
      updateContent(html)
      
      // 清除补全建议（用户继续输入时）
      if (completionSuggestion) {
        setCompletionSuggestion(null)
        setShowCompletion(false)
      }
      
      // 获取当前文本，用于触发补全
      const text = editor.getText()
      lastTextRef.current = text

      // 文档统计：字数/字符数
      // 中文场景下“字数”更接近“非空白字符数”，这里同时提供 words(按空白分词) 与 chars(总字符)
      const trimmed = (text || '').trim()
      const words = trimmed ? trimmed.split(/\s+/).filter(Boolean).length : 0
      const chars = (text || '').length
      setDocStats((prev) => ({ ...prev, words, chars }))
    },
    onSelectionUpdate: ({ editor }) => {
      // 选区变化时更新工具栏显示
      updateCurrentStyles(editor)

      const { from, to } = editor.state.selection
      const selectionChars = Math.max(0, to - from)
      setDocStats((prev) => ({
        ...prev,
        selectionFrom: from,
        selectionTo: to,
        selectionChars,
      }))
    },
  })

  // Parse DOCX comments (word/comments.xml) when a docx is opened
  useEffect(() => {
    if (!docxData) {
      setComments([])
      lastLoadedCommentsKeyRef.current = ''
      return
    }
    const key = currentFile?.path || `web:${document.title || ''}`
    if (key && lastLoadedCommentsKeyRef.current === key) return
    lastLoadedCommentsKeyRef.current = key

    ;(async () => {
      const raw = await parseDocxComments(docxData)
      // Map to CommentContext type (keep ids stable)
      setComments(
        raw.map(c => ({
          id: String(c.id),
          author: c.author,
          date: c.date,
          text: c.text,
          parentId: c.parentId,
          resolved: false,
        }))
      )
    })()
  }, [docxData, currentFile?.path, document.title, setComments])

  // 迁移旧缓存：把段落/标题的 block 级 font-family/font-size/color 下沉到 inline span
  // 目的：Tiptap 通常不会保留 block 的 style，导致“原 docx 字体”在预览中丢失。
  const migrateBlockFontStylesToInlineSpan = useCallback((html: string) => {
    try {
      if (!html || typeof html !== 'string') return html
      // quick reject
      if (!/style\s*=\s*["'][^"']*(font-family|font-size|color)\s*:/i.test(html)) return html

      const parser = new DOMParser()
      const doc = parser.parseFromString(html, 'text/html')
      const targets = doc.querySelectorAll<HTMLElement>('p[style], h1[style], h2[style], h3[style], h4[style], h5[style], h6[style]')
      let changed = false

      targets.forEach((el) => {
        const fontFamily = el.style.fontFamily
        const fontSize = el.style.fontSize
        const color = el.style.color
        if (!fontFamily && !fontSize && !color) return

        // avoid repeated wrap if already migrated
        const directWrap = el.firstElementChild as HTMLElement | null
        if (
          el.childNodes.length === 1 &&
          directWrap?.tagName === 'SPAN' &&
          directWrap.getAttribute('data-para-font') === '1'
        ) {
          // still remove styles from block to prevent overriding
          el.style.fontFamily = ''
          el.style.fontSize = ''
          el.style.color = ''
          changed = true
          return
        }

        el.style.fontFamily = ''
        el.style.fontSize = ''
        el.style.color = ''

        const span = doc.createElement('span')
        span.setAttribute('data-para-font', '1')
        if (fontFamily) span.style.fontFamily = fontFamily
        if (fontSize) span.style.fontSize = fontSize
        if (color) span.style.color = color

        while (el.firstChild) span.appendChild(el.firstChild)
        el.appendChild(span)
        changed = true
      })

      return changed ? doc.body.innerHTML : html
    } catch {
      return html
    }
  }, [])

  // --- Docx floating image positioning (wp:anchor) ---
  const positionDocxFloatingImagesImpl = useCallback(() => {
    if (!editor) return
    const root = editor.view.dom as HTMLElement
    if (!root) return
    const pageEl = root.closest('.word-page') as HTMLElement | null

    const imgs = Array.from(root.querySelectorAll<HTMLImageElement>('img[data-floating="1"]'))
    if (!imgs.length) return

    const computed = window.getComputedStyle(root)
    const paddingLeft = parseFloat(computed.paddingLeft || '0') || 0
    const paddingTop = parseFloat(computed.paddingTop || '0') || 0

    // Page starts: [0, ...] based on inserted page-break blocks
    const pageStarts: number[] = [0]
    const breaks = Array.from(root.querySelectorAll<HTMLElement>('.pm-page-break'))
    for (const br of breaks) {
      pageStarts.push(br.offsetTop + br.offsetHeight)
    }
    pageStarts.sort((a, b) => a - b)

    const findPageIndexByTop = (top: number) => {
      let idx = 0
      for (let i = 0; i < pageStarts.length; i++) {
        if (pageStarts[i] <= top) idx = i
        else break
      }
      return idx
    }

    const emuToPx = (emu: number) => emu / 9525

    for (const img of imgs) {
      const ds = img.dataset as any
      const xEmu = Number(ds.xEmu || 0) || 0
      const yEmu = Number(ds.yEmu || 0) || 0
      const relH = String(ds.relH || 'margin').toLowerCase()
      const relV = String(ds.relV || 'margin').toLowerCase()
      const wrapType = String(ds.wrap || '').toLowerCase()
      const distL = Number(ds.distL || 0) || 0
      const distR = Number(ds.distR || 0) || 0
      const distT = Number(ds.distT || 0) || 0
      const distB = Number(ds.distB || 0) || 0
      const distLeft = emuToPx(distL)
      const distRight = emuToPx(distR)
      const distTop = emuToPx(distT)
      const distBottom = emuToPx(distB)

      // Determine which page the anchor belongs to using the image's current block position
      const block = img.closest('p, li, div, td, th, h1, h2, h3, h4, h5, h6, blockquote, tr') as HTMLElement | null
      const anchorTop = block ? block.offsetTop : img.offsetTop
      const pageIndex = findPageIndexByTop(anchorTop)
      const pageTop = pageStarts[pageIndex] || 0

      const originX = relH.includes('page') ? 0 : paddingLeft
      const originY = relV.includes('page') ? 0 : paddingTop

      const left = originX + emuToPx(xEmu)
      const top = pageTop + originY + emuToPx(yEmu)

      if (wrapType === 'square' || wrapType === 'tight') {
        const pageWidth = pageEl?.clientWidth || root.clientWidth || 0
        const floatSide = left < pageWidth / 2 ? 'left' : 'right'
        img.style.position = 'relative'
        img.style.left = ''
        img.style.top = ''
        img.style.float = floatSide
        img.style.clear = 'none'
        img.style.margin = `${distTop}px ${distRight}px ${distBottom}px ${distLeft}px`
        img.style.visibility = 'visible'
        img.style.zIndex = '5'
      } else if (wrapType === 'topandbottom') {
        img.style.position = 'relative'
        img.style.left = ''
        img.style.top = ''
        img.style.float = 'none'
        img.style.display = 'block'
        img.style.clear = 'both'
        img.style.margin = `${distTop}px 0 ${distBottom}px 0`
        img.style.visibility = 'visible'
        img.style.zIndex = '5'
      } else {
        img.style.position = 'absolute'
        img.style.left = `${left}px`
        img.style.top = `${top}px`
        img.style.float = 'none'
        img.style.margin = '0'
        img.style.visibility = 'visible'
        img.style.zIndex = '5'
      }
    }
  }, [editor])

  const fontApplyTimerRef = useRef<number | null>(null)
  const scheduleApplyFontFamily = useCallback((fontFamily: string) => {
    if (!editor) return
    const v = (fontFamily || '').trim()
    if (!v) return
    if (fontApplyTimerRef.current) window.clearTimeout(fontApplyTimerRef.current)
    fontApplyTimerRef.current = window.setTimeout(() => {
      editor.chain().focus().setFontFamily(v).run()
    }, 150)
  }, [editor])

  useEffect(() => {
    return () => {
      if (fontApplyTimerRef.current) window.clearTimeout(fontApplyTimerRef.current)
    }
  }, [])

  // 使用 ProseMirror 的 DOMSerializer 从选区获取 HTML（保留格式）
  const getSelectionHtml = useCallback(() => {
    if (!editor) return ''
    
    const { from, to } = editor.state.selection
    if (from === to) return ''
    
    try {
      const { doc, schema } = editor.state
      const slice = doc.slice(from, to)
      const serializer = DOMSerializer.fromSchema(schema)
      const fragment = serializer.serializeFragment(slice.content)
      const tempDiv = window.document.createElement('div')
      tempDiv.appendChild(fragment)
      return tempDiv.innerHTML
    } catch (err) {
      console.warn('获取选区 HTML 失败:', err)
      return ''
    }
  }, [editor])

  // 选区修订（将选中文本替换为 AI 生成的新文本，显示为可接受/拒绝的修订）
  // isHtml: 如果为 true，则 newText 已经是 HTML 格式（AI 返回的带格式内容）
  const applySelectionRevision = useCallback((oldText: string, newText: string, isHtml?: boolean) => {
    if (!editor) return
    const diffId = `diff-${Date.now()}-${Math.random().toString(36).slice(2, 10)}`

    // 使用保存的 HTML 格式（如果有）作为原内容
    const oldHtml = selectedHtmlForEdit || textToHtml(oldText)
    
    // 新内容：如果 AI 返回的是 HTML 格式则直接使用，否则转换为 HTML
    const newHtml = isHtml ? newText : textToHtml(newText)

    const html = `<span class="diff-old" data-diff-id="${diffId}">${oldHtml}</span><span class="diff-new" data-diff-id="${diffId}">${newHtml}</span>`

    editor.chain().focus().deleteSelection().insertContent(html).run()

    addPendingReplacementItem({
      id: diffId,
      searchText: oldText,
      replaceText: newText,
      count: 1,
      timestamp: Date.now(),
    })
    
    // 清空已使用的 HTML
    setSelectedHtmlForEdit('')
  }, [editor, addPendingReplacementItem, textToHtml, selectedHtmlForEdit])

  const findBlockContainer = useCallback((node: HTMLElement) => {
    return (node.closest(DIFF_BLOCK_SELECTOR) as HTMLElement) || node
  }, [])

  const animateBlockReveal = useCallback((element: HTMLElement, delay: number) => {
    element.style.opacity = '0'
    element.style.transform = 'translateY(32px)'
    element.style.filter = 'blur(12px)'
    element.style.willChange = 'transform, opacity, filter'
    animate(
      element,
      { opacity: 1, y: 0, filter: 'blur(0px)' },
      {
        delay,
        type: 'spring',
        stiffness: 110,
        damping: 16,
        mass: 0.9,
      }
    ).finished.finally(() => {
      element.style.willChange = ''
    })
  }, [])

  const animateDiffMark = useCallback((element: HTMLElement, delay: number) => {
    element.style.opacity = '0'
    element.style.filter = 'blur(6px)'
    element.style.willChange = 'opacity, filter, background-color'
    animate(
      element,
      { opacity: 1, filter: 'blur(0px)' },
      {
        delay,
        duration: 0.4,
        ease: 'easeOut',
      }
    ).finished.finally(() => {
      element.style.willChange = ''
    })
  }, [])

  const runDiffInsertionAnimation = useCallback(() => {
    if (!editor) return
    const root = editor.view.dom as HTMLElement
    const diffNodes = Array.from(root.querySelectorAll<HTMLElement>('span.diff-new'))
    const animatedBlocks = new WeakSet<HTMLElement>()
    let newIndex = 0
    diffNodes.forEach(node => {
      const diffId = node.getAttribute('data-diff-id') || `${node.textContent}-${animatedDiffIdsRef.current.size}`
      if (diffId && animatedDiffIdsRef.current.has(diffId)) return
      newIndex++
      if (diffId) {
        animatedDiffIdsRef.current.add(diffId)
      }
      const delay = Math.min(newIndex - 1, 8) * 0.08
      requestAnimationFrame(() => {
        const block = findBlockContainer(node)
        if (block && !animatedBlocks.has(block)) {
          animatedBlocks.add(block)
          animateBlockReveal(block, delay)
        }
        animateDiffMark(node, delay + 0.04)
      })
    })
  }, [editor, findBlockContainer, animateBlockReveal, animateDiffMark])

  const runDocumentRevealAnimation = useCallback(() => {
    if (!editor) return
    const root = editor.view.dom as HTMLElement
    const blockNodes = Array.from(root.querySelectorAll<HTMLElement>(REVEAL_BLOCK_SELECTOR))
    if (!blockNodes.length) return
    blockNodes.slice(0, 240).forEach((node, index) => {
      const delay = Math.min(index, 15) * 0.06
      animateBlockReveal(node, delay)
    })
  }, [editor, animateBlockReveal])

  const playAcceptAnimation = useCallback(() => {
    return new Promise<void>(resolve => {
      if (!editor) {
        resolve()
        return
      }
      const root = editor.view.dom as HTMLElement
      const newNodes = Array.from(root.querySelectorAll<HTMLElement>('span.diff-new'))
      const oldNodes = Array.from(root.querySelectorAll<HTMLElement>('span.diff-old'))
      if (newNodes.length === 0 && oldNodes.length === 0) {
        resolve()
        return
      }
      newNodes.forEach(node => node.classList.add('diff-accept-new'))
      oldNodes.forEach(node => node.classList.add('diff-accept-old'))
      setTimeout(() => {
        newNodes.forEach(node => node.classList.remove('diff-accept-new'))
        oldNodes.forEach(node => node.classList.remove('diff-accept-old'))
        resolve()
      }, 320)
    })
  }, [editor])

  const handleConfirmClick = useCallback(async () => {
    if (isConfirming) return
    setIsConfirming(true)
    try {
      await playAcceptAnimation()
      acceptAllChanges()
    } finally {
      setIsConfirming(false)
    }
  }, [isConfirming, playAcceptAnimation, acceptAllChanges])

  // 触发补全的函数
  const triggerCompletion = useCallback(async () => {
    if (!editor) return
    
    const text = editor.getText()
    if (!text || text.length < 10) return  // 至少10个字符才触发
    
    // 获取光标位置之前的文本
    const { from } = editor.state.selection
    const textBefore = editor.state.doc.textBetween(0, from, '\n')
    
    if (!textBefore || textBefore.length < 10) return
    
    console.log('触发补全，上下文长度:', textBefore.length)
    
    const suggestion = await getCompletion(textBefore)
    if (suggestion) {
      setCompletionSuggestion(suggestion)
      setShowCompletion(true)
      console.log('获取到补全建议:', suggestion)
    }
  }, [editor, getCompletion])
  
  // 接受补全建议
  const acceptCompletion = useCallback(() => {
    if (!editor || !completionSuggestion) return
    
    // 在光标位置插入补全内容
    editor.commands.insertContent(completionSuggestion)
    
    // 清除补全状态
    setCompletionSuggestion(null)
    setShowCompletion(false)
  }, [editor, completionSuggestion])
  
  // 拒绝补全建议
  const dismissCompletion = useCallback(() => {
    setCompletionSuggestion(null)
    setShowCompletion(false)
    cancelCompletion()
  }, [cancelCompletion])
  
  // 打开内联编辑弹窗 (Ctrl+K)
  const openInlineEdit = useCallback(() => {
    if (!editor) return
    
    const { from, to } = editor.state.selection
    if (from === to) {
      // 没有选中文本，提示用户
      console.log('请先选中要编辑的文本')
      return
    }
    
    const selectedText = editor.state.doc.textBetween(from, to, ' ')
    if (!selectedText.trim()) return
    
    setSelectedTextForEdit(selectedText)
    setInlineEditRange({ from, to })
    
    // 使用 ProseMirror 获取选区的 HTML（保留格式）
    const selectionHtml = getSelectionHtml()
    setSelectedHtmlForEdit(selectionHtml || escapeHtml(selectedText))
    
    // 获取选区位置（用于弹窗定位）
    const domSelection = window.getSelection()
    if (domSelection && domSelection.rangeCount > 0) {
      const range = domSelection.getRangeAt(0)
      const rect = range.getBoundingClientRect()
      setInlineEditPosition({
        x: rect.left,
        y: rect.bottom
      })
    }
    
    setShowInlineEdit(true)
    setShowContextMenu(false)
  }, [editor, escapeHtml, getSelectionHtml])
  
  // 应用内联编辑结果
  const applyInlineEdit = useCallback((newText: string, isHtml?: boolean) => {
    if (!editor) return
    
    const oldText = selectedTextForEdit
    if (!oldText.trim()) return

    // 尽量把选区恢复到打开弹窗时的位置，避免用户移动光标导致应用到错误位置
    if (inlineEditRange) {
      try {
        editor.chain().focus().setTextSelection(inlineEditRange).run()
      } catch {
        // ignore
      }
    }
    applySelectionRevision(oldText, newText, isHtml)
  }, [editor, selectedTextForEdit, applySelectionRevision, inlineEditRange])
  
  // 处理右键菜单操作
  const handleContextMenuAction = useCallback(async (action: string, instruction?: string) => {
    if (!editor) return
    
    setShowContextMenu(false)
    
    if (action === 'copy') {
      // 复制
      const { from, to } = editor.state.selection
      const text = editor.state.doc.textBetween(from, to, ' ')
      navigator.clipboard.writeText(text)
      return
    }
    
    if (action === 'delete') {
      // 删除
      editor.chain().focus().deleteSelection().run()
      return
    }
    
    if (action === 'custom') {
      // 打开自定义编辑弹窗
      openInlineEdit()
      return
    }
    
    if (action === 'ai' && instruction) {
      // AI 操作
      setIsAIProcessing(true)
      
      try {
        // 正确构建 API URL
        const apiUrl = settings.apiUrl || `${settings.baseUrl}/chat/completions`
        
        // 检查是否有 HTML 格式可用
        const hasHtmlFormat = selectedHtmlForEdit && selectedHtmlForEdit !== selectedTextForEdit && selectedHtmlForEdit.includes('<')
        const contentToProcess = hasHtmlFormat ? selectedHtmlForEdit : selectedTextForEdit
        const formatHint = hasHtmlFormat ? '（HTML格式，请保留格式标签）' : ''
        
        // 生成 system prompt
        const systemPrompt = hasHtmlFormat 
          ? `你是一位专业的文档编辑助手。请严格按照用户的修改指令处理文字。

【输入格式】
用户会提供原文的 HTML 格式，其中包含格式标签（如 <strong>粗体</strong>、<em>斜体</em>、<u>下划线</u>、<span style="color:xxx">颜色</span> 等）。

【输出规则】
1. 输出格式为 HTML，保留并合理应用原文的格式标签
2. 只输出修改后的 HTML 内容，不要有任何解释、引号、前缀或后缀
3. 不要输出完整的 HTML 文档结构（如 <html>、<body> 等），只输出内容片段
4. 如果原文有粗体/斜体/颜色等格式，修改后的对应内容也应保持相同格式
5. 你可以根据内容语义决定是否调整格式（如重点内容可以加粗）`
          : `你是一位专业的文档编辑助手。请严格按照用户的修改指令处理文字。

【输出规则】
- 只输出修改后的文字，不要有任何解释、引号、前缀或后缀
- 保持原文的格式风格（标点、换行、缩进等）
- 如果是翻译任务，保持原文的语义和风格`
        
        const response = await fetch(apiUrl, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${settings.apiKey}`,
          },
          body: JSON.stringify({
            model: settings.model,
            messages: [
              {
                role: 'system',
                content: systemPrompt
              },
              {
                role: 'user',
                content: `【原文${formatHint}】
${contentToProcess}

【修改指令】${instruction}

请直接输出修改后的${hasHtmlFormat ? 'HTML' : '文字'}：`
              }
            ],
            temperature: 0.7,
            max_tokens: 4000,
          }),
        })

        if (!response.ok) {
          const errorText = await response.text().catch(() => '')
          throw new Error(`AI 请求失败 (${response.status})${errorText ? ': ' + errorText.slice(0, 100) : ''}`)
        }

        const data = await response.json()
        let newText = data.choices?.[0]?.message?.content?.trim()
        
        if (newText) {
          // 清理可能的 markdown 代码块包裹
          if (newText.startsWith('```html')) {
            newText = newText.slice(7)
          } else if (newText.startsWith('```')) {
            newText = newText.slice(3)
          }
          if (newText.endsWith('```')) {
            newText = newText.slice(0, -3)
          }
          newText = newText.trim()
          
          // 修订式应用（可接受/拒绝）
          const oldText = selectedTextForEdit
          applySelectionRevision(oldText, newText, hasHtmlFormat ? true : undefined)
        } else {
          throw new Error('AI 返回内容为空')
        }
      } catch (err) {
        console.error('AI 处理失败:', err)
        // 可以在这里添加用户友好的错误提示
        alert(`AI 处理失败：${(err as Error).message || '未知错误'}`)
      } finally {
        setIsAIProcessing(false)
      }
    }
  }, [editor, selectedTextForEdit, selectedHtmlForEdit, settings.apiUrl, settings.baseUrl, settings.apiKey, settings.model, openInlineEdit, applySelectionRevision])
  
  // 监听键盘事件
  useEffect(() => {
    if (!editor) return
    
    const handleKeyDown = (event: KeyboardEvent) => {
      // Ctrl+K 打开内联编辑
      if ((event.ctrlKey || event.metaKey) && event.key === 'k') {
        event.preventDefault()
        openInlineEdit()
        return
      }
      
      // Tab 键触发补全
      if (event.key === 'Tab' && !event.shiftKey) {
        // 如果有补全建议，接受它
        if (showCompletion && completionSuggestion) {
          event.preventDefault()
          acceptCompletion()
          return
        }
        
        // 否则触发新的补全请求
        event.preventDefault()
        triggerCompletion()
        return
      }
      
      // Escape 键取消补全或关闭弹窗，或拒绝待确认修改
      if (event.key === 'Escape') {
        // 优先处理待确认修改
        if (pendingChangesTotal > 0) {
          event.preventDefault()
          rejectAllChanges()
          return
        }
        if (showCompletion) {
          event.preventDefault()
          dismissCompletion()
          return
        }
        if (showInlineEdit) {
          event.preventDefault()
          setShowInlineEdit(false)
          return
        }
        if (showContextMenu) {
          event.preventDefault()
          setShowContextMenu(false)
          return
        }
      }
      
      // Enter 键接受待确认修改（仅在没有输入焦点时）
      if (event.key === 'Enter' && pendingChangesTotal > 0 && !event.shiftKey && !event.ctrlKey && !event.metaKey) {
        // 检查是否在编辑器中有选区或正在输入
        const selection = editor.state.selection
        if (selection.empty) {
          event.preventDefault()
          handleConfirmClick()
          return
        }
      }
      
      // 其他按键时清除补全建议
      if (showCompletion && !['Tab', 'Shift', 'Control', 'Alt', 'Meta'].includes(event.key)) {
        dismissCompletion()
      }
    }
    
    // 添加到编辑器的 DOM 元素（使用 window.document 避免与 context 的 document 冲突）
    const editorElement = window.document.querySelector('.word-editor-content')
    if (editorElement) {
      editorElement.addEventListener('keydown', handleKeyDown as EventListener)
    }
    
    return () => {
      if (editorElement) {
        editorElement.removeEventListener('keydown', handleKeyDown as EventListener)
      }
    }
  }, [editor, showCompletion, completionSuggestion, acceptCompletion, dismissCompletion, triggerCompletion, openInlineEdit, showInlineEdit, showContextMenu, handleConfirmClick, pendingChangesTotal, rejectAllChanges])

  // 编辑器加载后也更新一次
  useEffect(() => {
    if (editor) {
      updateCurrentStyles(editor)
    }
  }, [editor])
  
  // 监听右键菜单事件
  useEffect(() => {
    if (!editor) return
    
    const handleContextMenu = (event: MouseEvent) => {
      const { from, to } = editor.state.selection
      
      // 只有选中文本时才显示自定义右键菜单
      if (from !== to) {
        event.preventDefault()
        
        const selectedText = editor.state.doc.textBetween(from, to, ' ')
        setSelectedTextForEdit(selectedText)
        
        // 使用 ProseMirror 获取选区的 HTML（保留格式）
        const selectionHtml = getSelectionHtml()
        setSelectedHtmlForEdit(selectionHtml || selectedText)
        
        setContextMenuPosition({ x: event.clientX, y: event.clientY })
        setShowContextMenu(true)
        setShowInlineEdit(false)
      }
    }
    
    const editorElement = window.document.querySelector('.word-editor-content')
    if (editorElement) {
      editorElement.addEventListener('contextmenu', handleContextMenu as EventListener)
    }
    
    return () => {
      if (editorElement) {
        editorElement.removeEventListener('contextmenu', handleContextMenu as EventListener)
      }
    }
  }, [editor, getSelectionHtml])

  // 当 docx 数据或文档内容变化时，解析并加载到编辑器
  useEffect(() => {
    const loadContent = async () => {
      if (!editor) return

      // 如果有 docx 二进制数据，需要解析
      if (docxData) {
        // 如果 document.content 已经有内容了，说明已经解析过了
        const hasFontInfo =
          /font-family\s*:\s*[^;]+/i.test(document.content || '') ||
          document.content.includes('data-para-font="1"') ||
          document.content.includes('data-font-name=')
        // 检查是否有旧格式的内嵌页眉/页脚（需要重新解析）
        const hasOldEmbeddedHeaderFooter = 
          document.content.includes('class="docx-header"') || 
          document.content.includes('class="docx-footer"')
        const hasValidCachedContent = document.content && 
          document.content.length > 100 && 
          !document.content.includes('blob:') &&  // blob: 图片重启后失效，需要重新解析
          !hasOldEmbeddedHeaderFooter &&  // 旧格式内嵌页眉/页脚，需要重新解析
          hasFontInfo  // 没有字体信息说明旧缓存，强制重新解析
        
        if (hasValidCachedContent) {
          // 检查是否需要同步（与上次同步的内容不同）
          if (document.content !== lastSyncedContentRef.current) {
            // 提取文档中使用的字体并加载
            const usedFonts = extractFontFamiliesFromHtml(document.content)
            if (usedFonts.length > 0) {
              // 加载字体，完成后触发重渲染
              loadFontsByNames(usedFonts).then(() => {
                window.document.fonts.ready.then(() => {
                  if (editor && !editor.isDestroyed) {
                    requestAnimationFrame(() => {
                      editor.view.dispatch(editor.state.tr)
                    })
                  }
                })
              }).catch(err => {
                console.warn('Font loading failed, using fallback fonts:', err)
              })
            }
            
            // 强制清空再设置，确保完全刷新
            editor.commands.clearContent()
            const migrated = migrateBlockFontStylesToInlineSpan(document.content)
            editor.commands.setContent(migrated)
            // 注意：lastSyncedContentRef 会在 onUpdate 中被更新为 editor.getHTML()
            // 重算浮动图片位置（如果存在）
            requestAnimationFrame(() => positionDocxFloatingImagesImpl())
          }
          return
        }
        
        // 首次解析 docx（或缓存无效需要重新解析）
        setIsLoading(true)
        try {
          // 使用完整解析器，获取正文 + 页眉/页脚分离的数据
          const parseResult = await parseDocxComplete(docxData)
          const pageSettings = parseResult.pageSettings
          if (pageSettings) {
            const paperSize = detectPaperSize(pageSettings.width, pageSettings.height)
            const orientation =
              pageSettings.orientation ||
              (pageSettings.width > pageSettings.height ? 'landscape' : 'portrait')

            const nextSetup = {
              paperSize,
              orientation,
              margins: {
                top: formatCm(pageSettings.marginTop),
                bottom: formatCm(pageSettings.marginBottom),
                left: formatCm(pageSettings.marginLeft),
                right: formatCm(pageSettings.marginRight),
              },
            } as const

            const customSize =
              paperSize === 'custom'
                ? {
                    customWidth: formatCm(pageSettings.width),
                    customHeight: formatCm(pageSettings.height),
                  }
                : undefined

            setPageSetup({ ...nextSetup, ...(customSize || {}) })
          }
          const htmlContent = parseResult.bodyHtml
          
          // 保存页眉/页脚数据，供单独渲染使用
          if (parseResult.headerHtml || parseResult.footerHtml) {
            setDocxHeaderFooter({
              headerHtml: parseResult.headerHtml || '',
              footerHtml: parseResult.footerHtml || '',
              headerStyle: parseResult.headerStyle,
              footerStyle: parseResult.footerStyle,
            })
            console.log('[WordEditor] 页眉页脚数据已保存:', {
              hasHeader: !!parseResult.headerHtml,
              hasFooter: !!parseResult.footerHtml,
              headerPreview: parseResult.headerHtml?.slice(0, 100),
              footerPreview: parseResult.footerHtml?.slice(0, 100),
            })
          } else {
            setDocxHeaderFooter(null)
          }
          
          // 提取文档中使用的字体并加载
          const usedFonts = extractFontFamiliesFromHtml(htmlContent)
          
          if (usedFonts.length > 0) {
            // 先加载字体，然后触发重渲染以确保字体正确显示
            loadFontsByNames(usedFonts).then(() => {
              // 字体加载完成后，等待 window.document.fonts.ready 确保所有字体都已渲染就绪
              window.document.fonts.ready.then(() => {
                // 触发编辑器视图更新，强制重新渲染以应用新加载的字体
                if (editor && !editor.isDestroyed) {
                  // 使用 requestAnimationFrame 确保在下一帧渲染
                  requestAnimationFrame(() => {
                    // 触发 ProseMirror 视图更新
                    editor.view.dispatch(editor.state.tr)
                  })
                }
              })
            }).catch(err => {
              console.warn('Font loading failed, using fallback fonts:', err)
            })
          }
          
          editor.commands.setContent(htmlContent)
          
          // 直接更新 lastSyncedContentRef，避免 onUpdate 触发后再次同步
          lastSyncedContentRef.current = editor.getHTML()
          requestAnimationFrame(() => positionDocxFloatingImagesImpl())
          // 同步内容到 document.content，这样 AI 替换才能正常工作
          // 使用 setTimeout 避免在 effect 中直接触发状态更新
          setTimeout(() => {
            updateContent(lastSyncedContentRef.current)
          }, 0)
        } catch (err) {
          console.error('Failed to parse docx:', err)
          // 显示错误提示
          editor.commands.setContent(`
            <div style="padding: 40px; text-align: center; color: var(--word-ink-muted);">
              <p style="font-size: 18px; margin-bottom: 10px;">⚠️ 文档解析失败</p>
              <p style="font-size: 14px;">无法解析此 Word 文档，可能是文件损坏或格式不兼容。</p>
              <p style="font-size: 12px; margin-top: 10px; color: var(--word-ink-muted);">错误信息: ${(err as Error).message}</p>
            </div>
          `)
        } finally {
          setIsLoading(false)
        }
        return
      }
      
      // 否则使用文档内容
      if (document.content) {
        // 检查编辑器当前内容是否已经是最新的（避免重复设置导致光标跳动）
        const currentHtml = editor.getHTML()
        
        if (document.content.startsWith('<')) {
          // HTML 内容
          // 检查是否包含 diff 标记 - 如果有，强制更新以确保 diff 显示
          const hasDiffMarkers = document.content.includes('diff-old') || document.content.includes('diff-new')
          const currentHasDiff = currentHtml.includes('diff-old') || currentHtml.includes('diff-new')
          
          // 强制更新的条件：
          // 1. 内容不同
          // 2. 或者新内容有 diff 标记但编辑器当前没有
          // 3. 或者新内容没有 diff 标记但编辑器当前有（确认后清理）
          const needsUpdate = currentHtml !== document.content || 
                             (hasDiffMarkers && !currentHasDiff) ||
                             (!hasDiffMarkers && currentHasDiff)
          
          if (needsUpdate) {
            const migrated = migrateBlockFontStylesToInlineSpan(document.content)
            editor.commands.setContent(migrated)
            requestAnimationFrame(() => positionDocxFloatingImagesImpl())
          }
        } else {
          // Markdown 转 HTML
          const html = markdownToHtml(document.content)
          if (currentHtml !== html) {
            editor.commands.setContent(html)
            requestAnimationFrame(() => positionDocxFloatingImagesImpl())
          }
        }
      } else {
        editor.commands.setContent('')
      }
    }

    loadContent()
  }, [docxData, document.content, editor, updateContent, positionDocxFloatingImagesImpl, migrateBlockFontStylesToInlineSpan])

  const locateCommentInDom = useCallback((commentId: string) => {
    try {
      setSelectedCommentId(commentId)
      // Find first mark span that includes this commentId
      const el = window.document.querySelector(`.word-editor-content span.docx-comment[data-comment-ids*="${commentId}"]`) as HTMLElement | null
      if (el) {
        el.scrollIntoView({ behavior: 'smooth', block: 'center' })
      }
    } catch {
      // ignore
    }
  }, [setSelectedCommentId])

  const applyCommentToSelection = useCallback((text: string) => {
    if (!editor) return
    const { from, to } = editor.state.selection
    if (from === to) return
    const id = addComment(text, { author: 'User' })
    // best-effort: set mark on selection
    editor.commands.setMark('comment', { commentIds: [id] })
    setShowCommentPanel(true)
    requestAnimationFrame(() => locateCommentInDom(id))
  }, [editor, addComment, locateCommentInDom])

  const replyTo = useCallback((parentId: string, text: string) => {
    replyToComment(parentId, text, { author: 'User' })
    setShowCommentPanel(true)
  }, [replyToComment])

  useEffect(() => {
    animatedDiffIdsRef.current.clear()
  }, [currentFile?.path])

  // Reposition floating images on zoom/resize
  useEffect(() => {
    if (!editor) return
    const onResize = () => positionDocxFloatingImagesImpl()
    window.addEventListener('resize', onResize)
    // zoom changes trigger rerender; do a best-effort reposition
    requestAnimationFrame(() => positionDocxFloatingImagesImpl())
    return () => window.removeEventListener('resize', onResize)
  }, [editor, zoomLevel, positionDocxFloatingImagesImpl])

  useEffect(() => {
    runDiffInsertionAnimation()
  }, [document.content, runDiffInsertionAnimation])

  useEffect(() => {
    if (!docEntryAnimationKey) return
    if (docEntryAnimationRef.current === docEntryAnimationKey) return
    docEntryAnimationRef.current = docEntryAnimationKey
    requestAnimationFrame(() => {
      runDocumentRevealAnimation()
    })
  }, [docEntryAnimationKey, runDocumentRevealAnimation])

  // Ctrl + 滚轮缩放功能
  useEffect(() => {
    const container = editorContainerRef.current
    if (!container) return

    const handleWheel = (e: WheelEvent) => {
      // 只在按住 Ctrl 键时触发缩放
      if (!e.ctrlKey) return
      
      e.preventDefault()
      
      // 计算新的缩放级别
      const delta = e.deltaY > 0 ? -10 : 10
      setZoomLevel(prev => {
        const newZoom = prev + delta
        // 限制缩放范围：25% - 500%
        return Math.min(500, Math.max(25, newZoom))
      })
    }

    // 键盘快捷键: Ctrl+0 重置, Ctrl++ 放大, Ctrl+- 缩小
    const handleKeyDown = (e: KeyboardEvent) => {
      if (!e.ctrlKey) return
      
      if (e.key === '0') {
        e.preventDefault()
        setZoomLevel(100)
      } else if (e.key === '=' || e.key === '+') {
        e.preventDefault()
        setZoomLevel(prev => Math.min(500, prev + 10))
      } else if (e.key === '-') {
        e.preventDefault()
        setZoomLevel(prev => Math.max(25, prev - 10))
      }
    }

    container.addEventListener('wheel', handleWheel, { passive: false })
    window.document.addEventListener('keydown', handleKeyDown)
    
    return () => {
      container.removeEventListener('wheel', handleWheel)
      window.document.removeEventListener('keydown', handleKeyDown)
    }
  }, [])

  const recomputePageStats = useCallback(() => {
    const container = editorContainerRef.current
    if (!container) return

    // 总页数：分页占位块数量 + 1（视觉页）
    const breaks = window.document.querySelectorAll('.word-editor-content .pm-page-break')
    const total = Math.max(1, (breaks?.length || 0) + 1)

    // 当前页数：按滚动位置粗略估算（匹配 PagedLayout 的 PAGE_HEIGHT_PX + PAGE_GAP_PX）
    // 注意：zoom 通过 transform 实现，不影响布局尺寸；这里用未缩放的滚动高度估算即可
    const PAGE_HEIGHT_PX = 1122.5
    const PAGE_GAP_PX = 32
    const step = PAGE_HEIGHT_PX + PAGE_GAP_PX
    const current = Math.min(total, Math.max(1, Math.floor((container.scrollTop + 120) / step) + 1))

    setPageStats({ current, total })
  }, [])

  useEffect(() => {
    recomputePageStats()
  }, [recomputePageStats, document.content, docEntryAnimationKey, zoomLevel])

  useEffect(() => {
    const container = editorContainerRef.current
    if (!container) return
    const onScroll = () => recomputePageStats()
    container.addEventListener('scroll', onScroll, { passive: true })
    return () => container.removeEventListener('scroll', onScroll)
  }, [recomputePageStats])

  // 监听滚动到文本的事件
  useEffect(() => {
    const handleScrollToText = (event: CustomEvent<{ text: string }>) => {
      const { text } = event.detail
      if (!text) return
      
      console.log('WordEditor received scroll-to-text:', text)
      
      // 获取编辑器 DOM
      const editorElement = window.document.querySelector('.word-editor-content')
      if (!editorElement) {
        console.log('Editor element not found')
        return
      }
      
      // 高亮函数 - 添加临时高亮效果
      const highlightElement = (el: HTMLElement) => {
        el.scrollIntoView({ behavior: 'smooth', block: 'center' })
        
        // 保存原始样式
        const originalBg = el.style.backgroundColor
        const originalBoxShadow = el.style.boxShadow
        const originalTransition = el.style.transition
        
        // 添加高亮效果
        el.style.transition = 'all 0.3s ease'
        el.style.backgroundColor = 'rgb(var(--success-rgb) / 0.22)'
        el.style.boxShadow = '0 0 0 3px rgb(var(--success-rgb) / 0.32)'
        
        // 2秒后恢复
        setTimeout(() => {
          el.style.backgroundColor = originalBg
          el.style.boxShadow = originalBoxShadow
          el.style.transition = originalTransition
        }, 2000)
      }
      
      // 1. 优先查找 diff-new 中包含该文本的元素（绿色新增内容）
      const diffNewElements = editorElement.querySelectorAll('.diff-new, span[class*="diff-new"]')
      for (const el of diffNewElements) {
        if (el.textContent?.includes(text)) {
          console.log('Found in diff-new element')
          highlightElement(el as HTMLElement)
          return
        }
      }
      
      // 2. 查找带有内联 diff 样式的元素（兼容旧版本：不要写死颜色值）
      const allSpans = editorElement.querySelectorAll('span')
      for (const span of allSpans) {
        const style = span.getAttribute('style') || ''
        // 兼容：旧实现可能用绿色背景标记“新增内容”
        // 这里用更宽松的判断：包含 background 且包含绿色系 rgba/rgb 片段
        const looksLikeAddHighlight =
          /background/i.test(style) && /(rgb|rgba)\(/i.test(style) && /(34,\s*197,\s*94|52,\s*199,\s*89)/.test(style)
        if (looksLikeAddHighlight && span.textContent?.includes(text)) {
          console.log('Found in inline styled span')
          highlightElement(span as HTMLElement)
          return
        }
      }
      
      // 3. 查找任何包含该文本的元素（不限于 diff）
      const walker = window.document.createTreeWalker(
        editorElement,
        NodeFilter.SHOW_TEXT,
        null
      )
      
      let node: Text | null
      while ((node = walker.nextNode() as Text | null)) {
        if (node.textContent?.includes(text)) {
          const parentElement = node.parentElement
          if (parentElement) {
            console.log('Found in text node, parent:', parentElement.tagName)
            highlightElement(parentElement)
            return
          }
        }
      }
      
      console.log('Text not found in editor:', text)
    }
    
    window.addEventListener('scroll-to-text', handleScrollToText as EventListener)
    return () => {
      window.removeEventListener('scroll-to-text', handleScrollToText as EventListener)
    }
  }, [])

  // 监听滚动到 diffId 的事件（RevisionPanel 定位）
  useEffect(() => {
    const handleScrollToDiffId = (event: CustomEvent<{ diffId: string }>) => {
      const diffId = event.detail?.diffId
      if (!diffId) return

      const editorElement = window.document.querySelector('.word-editor-content')
      if (!editorElement) return

      const el =
        editorElement.querySelector(`[data-diff-id="${diffId}"].diff-new`) ||
        editorElement.querySelector(`[data-diff-id="${diffId}"].diff-old`) ||
        editorElement.querySelector(`[data-diff-id="${diffId}"]`)

      if (!el) return

      const node = el as HTMLElement
      node.scrollIntoView({ behavior: 'smooth', block: 'center' })

      // 临时高亮
      const originalOutline = node.style.outline
      const originalOutlineOffset = node.style.outlineOffset
      node.style.outline = '2px solid rgb(var(--accent-rgb) / 0.9)'
      node.style.outlineOffset = '2px'
      setTimeout(() => {
        node.style.outline = originalOutline
        node.style.outlineOffset = originalOutlineOffset
      }, 1200)
    }

    window.addEventListener('scroll-to-diff-id', handleScrollToDiffId as EventListener)
    return () => {
      window.removeEventListener('scroll-to-diff-id', handleScrollToDiffId as EventListener)
    }
  }, [])

  if (!editor) {
    return (
      <div className="flex items-center justify-center h-full bg-background">
        <Loader2 className="w-6 h-6 animate-spin text-text-muted" />
      </div>
    )
  }

  if (excelData?.sheets?.length) {
    return (
      <div className="word-editor h-full w-full overflow-hidden">
        <ExcelPreview 
          sheets={excelData.sheets} 
          filePath={currentFile?.path}
          onRefresh={refreshExcelData}
        />
      </div>
    )
  }

  if (pptData?.pptxBase64) {
    return (
      <div className="word-editor h-full w-full overflow-hidden">
        <PptPreviewHtml 
          title={document.title || '未命名演示文稿'} 
          pptxBase64={pptData.pptxBase64}
          pptxPath={currentFile?.path}
        />
      </div>
    )
  }

  return (
    <div className="flex flex-col h-full bg-background overflow-hidden relative">
      {/* RibbonShell：固定在顶部（不随文档滚动） */}
      <div className="flex-shrink-0">
        {/* 顶部信息栏 - 柔和玻璃态 */}
        <div className="flex items-center justify-between px-4 py-2 glass border-b border-border z-20 relative">
          <div className="flex items-center gap-2.5">
            <div className="p-1.5 rounded-lg bg-accent/12 text-accent border border-accent/20">
              <FileText className="w-4 h-4" />
            </div>
            <div className="flex flex-col">
              <span className="text-sm font-medium text-text">{document.title || '未命名文档'}</span>
              <span className="text-[10px] text-text-muted flex items-center gap-1.5">
                {hasUnsavedChanges ? (
                  <>
                    <span className="w-1.5 h-1.5 rounded-full bg-warning animate-pulse" />
                    未保存
                  </>
                ) : (
                  <>
                    <span className="w-1.5 h-1.5 rounded-full bg-success" />
                    已保存
                  </>
                )}
                <span className="opacity-50">|</span>
                Word 模式
              </span>
            </div>
          </div>
          <div className="flex items-center gap-2">
            <button
              onClick={() => saveDocument()}
              className="flex items-center gap-1.5 px-3 py-1.5 bg-accent/90 hover:bg-accent text-white text-xs font-medium rounded-lg shadow-sm shadow-accent/20 transition-all hover:-translate-y-0.5"
            >
              <Save className="w-3.5 h-3.5" />
              <span>保存</span>
            </button>
            <button
              onClick={() => saveDocument()}
              className="flex items-center gap-1.5 px-3 py-1.5 bg-white/5 hover:bg-white/10 text-text border border-white/10 text-xs font-medium rounded-lg transition-all"
            >
              <Download className="w-3.5 h-3.5" />
              <span>导出</span>
            </button>
          </div>
        </div>

        {/* Word 风格 Ribbon（多标签 + 分组） */}
        <WordRibbon
          tab={ribbonTab}
          setTab={setRibbonTab}
          editor={editor}
          currentFontFamily={currentFontFamily}
          setCurrentFontFamily={setCurrentFontFamily}
          fontFamilyOptions={fontFamilyOptions}
          currentFontSize={currentFontSize}
          setCurrentFontSize={setCurrentFontSize}
          applyFontFamily={scheduleApplyFontFamily}
          applyFontSize={(v) => editor.chain().focus().setMark('textStyle', { fontSize: v }).run()}
          viewMode={viewMode}
          setViewMode={setViewMode}
          zoomLevel={zoomLevel}
          setZoomLevel={setZoomLevel}
          pageSetup={pageSetup}
          setPageSetup={setPageSetup}
          pendingChangesTotal={pendingChangesTotal}
          onOpenRevisionPanel={() => setShowRevisionPanel(true)}
          acceptAllChanges={acceptAllChanges}
          rejectAllChanges={rejectAllChanges}
          editorMode={editorMode}
          setEditorMode={setEditorMode}
        />
      </div>

      {/* 编辑区域 - Word 风格界面 */}
      <div 
        ref={editorContainerRef}
        className={`flex-1 overflow-auto word-editor-container relative bg-[var(--word-canvas-bg)] ${viewMode === 'web' ? 'web-mode' : 'print-mode'}`}
      >
        {isLoading && (
          <div className="absolute inset-0 flex items-center justify-center bg-black/25 dark:bg-black/45 backdrop-blur-sm z-10">
            <div className="flex items-center gap-2 text-text-secondary">
              <Loader2 className="w-5 h-5 animate-spin" />
              <span>正在加载文档...</span>
            </div>
          </div>
        )}

        {/* 缩放容器 */}
        <div 
          className="zoom-container"
          style={{
            transform: `scale(${zoomLevel / 100})`,
            transformOrigin: 'top center',
            transition: 'transform 0.1s ease-out',
          }}
        >
          <div 
            className={`word-page relative ${pageSetup.orientation === 'landscape' ? 'landscape' : ''} ${viewMode === 'web' ? 'web' : ''}`}
            style={{
              // 根据页面设置动态调整边距
              '--page-margin-top': pageSetup.margins.top,
              '--page-margin-bottom': pageSetup.margins.bottom,
              '--page-margin-left': pageSetup.margins.left,
              '--page-margin-right': pageSetup.margins.right,
              '--page-width': pageSize.width,
              '--page-height': pageSize.height,
            } as React.CSSProperties}
          >
            {/* 页眉 - 优先显示 DOCX 解析的页眉，其次是用户设置的页眉 */}
            {viewMode === 'print' && (docxHeaderFooter?.headerHtml || headerFooterSetup.header?.content) && (
              <div 
                className="page-header docx-header-footer"
                style={{ 
                  textAlign: (docxHeaderFooter?.headerStyle?.alignment || headerFooterSetup.header?.alignment || 'right') as 'left' | 'center' | 'right',
                  fontSize: docxHeaderFooter?.headerStyle?.fontSize || '9pt',
                  color: docxHeaderFooter?.headerStyle?.color || 'var(--word-ink-muted)',
                }}
              >
                {docxHeaderFooter?.headerHtml ? (
                  <div 
                    className="docx-header-content"
                    dangerouslySetInnerHTML={{ __html: docxHeaderFooter.headerHtml.replace(/\{PAGE\}/g, '1').replace(/\{NUMPAGES\}/g, '1') }} 
                  />
                ) : (
                  headerFooterSetup.header?.content
                )}
              </div>
            )}
            
          <EditorContent editor={editor} />

            {/* 页脚 - 优先显示 DOCX 解析的页脚，其次是用户设置的页脚 */}
            {viewMode === 'print' && (docxHeaderFooter?.footerHtml || headerFooterSetup.footer?.content || headerFooterSetup.pageNumber?.enabled) && (
              <div 
                className="page-footer docx-header-footer"
                style={{ 
                  textAlign: (docxHeaderFooter?.footerStyle?.alignment || headerFooterSetup.footer?.alignment || headerFooterSetup.pageNumber?.alignment || 'center') as 'left' | 'center' | 'right',
                  fontSize: docxHeaderFooter?.footerStyle?.fontSize || '9pt',
                  color: docxHeaderFooter?.footerStyle?.color || 'var(--word-ink-muted)',
                }}
              >
                {docxHeaderFooter?.footerHtml ? (
                  <div 
                    className="docx-footer-content"
                    dangerouslySetInnerHTML={{ __html: docxHeaderFooter.footerHtml.replace(/\{PAGE\}/g, '1').replace(/\{NUMPAGES\}/g, '1') }} 
                  />
                ) : (
                  <>
                    {headerFooterSetup.footer?.content}
                    {headerFooterSetup.pageNumber?.enabled && headerFooterSetup.pageNumber.position === 'footer' && (
                      <span className="page-number">第 1 页</span>
                    )}
                  </>
                )}
              </div>
            )}
          </div>
        </div>
        
        {/*（缩放控制已迁移到底部状态栏，避免遮挡内容）*/}
        
        {/* AI 补全建议浮窗 */}
        {showCompletion && completionSuggestion && (
          <div className="fixed bottom-20 left-1/2 -translate-x-1/2 z-50 max-w-xl">
            <div className="glass-card border border-border rounded-xl shadow-2xl overflow-hidden">
              {/* 标题栏 */}
              <div className="flex items-center justify-between px-3 py-2 bg-black/10 dark:bg-white/10 backdrop-blur-md border-b border-border">
                <div className="flex items-center gap-2">
                  <Sparkles className="w-4 h-4 text-accent" />
                  <span className="text-xs text-text-secondary">AI 补全建议</span>
                </div>
                <div className="flex items-center gap-1 text-[10px] text-text-dim">
                  <kbd className="px-1.5 py-0.5 bg-black/10 dark:bg-white/10 rounded text-text-dim border border-black/10 dark:border-white/10">Tab</kbd>
                  <span>接受</span>
                  <kbd className="px-1.5 py-0.5 bg-black/10 dark:bg-white/10 rounded text-text-dim border border-black/10 dark:border-white/10 ml-2">Esc</kbd>
                  <span>取消</span>
                </div>
              </div>
              {/* 补全内容预览 */}
              <div className="px-4 py-3">
                <p className="text-sm text-text-secondary leading-relaxed whitespace-pre-wrap">
                  {completionSuggestion}
                </p>
              </div>
              {/* 操作按钮 */}
              <div className="flex items-center gap-2 px-3 py-2 bg-black/10 dark:bg-white/10 backdrop-blur-md border-t border-border">
                <button
                  onClick={acceptCompletion}
                  className="flex items-center gap-1.5 px-3 py-1.5 bg-success/16 text-success border border-success/25 hover:bg-success/22 text-xs rounded-lg transition-colors"
                >
                  <Check className="w-3 h-3" />
                  <span>接受</span>
                </button>
                <button
                  onClick={dismissCompletion}
                  className="flex items-center gap-1.5 px-3 py-1.5 bg-black/10 dark:bg-white/10 hover:bg-black/15 dark:hover:bg-white/15 text-text-secondary text-xs rounded-lg transition-colors border border-black/10 dark:border-white/10"
                >
                  <X className="w-3 h-3" />
                  <span>取消</span>
                </button>
              </div>
            </div>
          </div>
        )}
        
        {/* 补全加载指示器 */}
        {isCompleting && (
          <div className="fixed bottom-20 left-1/2 -translate-x-1/2 z-50">
            <div className="flex items-center gap-2 px-4 py-2 glass-card border border-border rounded-lg shadow-xl">
              <Loader2 className="w-4 h-4 text-accent animate-spin" />
              <span className="text-xs text-text-secondary">AI 正在思考...</span>
            </div>
          </div>
        )}
        
        {/* 待确认修改的操作栏 - 使用 Framer Motion 动画 */}
        <AnimatePresence>
          {pendingChangesTotal > 0 && (
            <motion.div 
              className="fixed bottom-4 left-1/2 -translate-x-1/2 z-50 max-w-md"
              variants={floatingBarVariants}
              initial="hidden"
              animate="visible"
              exit="exit"
              style={{ x: '-50%' }}
            >
              <div className="glass-card border border-border rounded-xl shadow-2xl overflow-hidden">
                {/* 顶部蓝色指示条 - 表示这是 AI 建议 */}
                <div className="h-0.5 bg-gradient-to-r from-accent via-system-purple to-accent-hover" />
                
                <div className="px-4 py-2.5 border-b border-border flex items-center justify-between gap-4 bg-black/10 dark:bg-white/10 backdrop-blur-md">
                  <div className="flex items-center gap-2">
                    <div className="w-1.5 h-1.5 rounded-full bg-accent animate-pulse" />
                    <span className="text-xs text-text-secondary font-medium">待确认修改</span>
                    <span className="text-[10px] text-text-dim bg-black/10 dark:bg-white/10 px-1.5 py-0.5 rounded border border-black/10 dark:border-white/10">
                      ×{pendingChangesTotal}
                    </span>
                  </div>
                  <div className="flex items-center gap-2">
                    <motion.button
                      onClick={rejectAllChanges}
                      className="flex items-center gap-1.5 px-3 py-1.5 bg-black/10 dark:bg-white/10 hover:bg-error/12 text-text-muted hover:text-error text-xs rounded-lg transition-all border border-black/10 dark:border-white/10 hover:border-error/30"
                      whileHover={{ scale: 1.02 }}
                      whileTap={{ scale: 0.98 }}
                    >
                      <X className="w-3.5 h-3.5" />
                      <span>拒绝</span>
                    </motion.button>
                    <motion.button
                      onClick={handleConfirmClick}
                      disabled={isConfirming}
                      className={`flex items-center gap-1.5 px-3 py-1.5 bg-success/18 text-success border border-success/25 text-xs rounded-lg transition-all shadow-sm disabled:opacity-60 disabled:cursor-not-allowed ${isConfirming ? '' : 'hover:bg-success/24'}`}
                      whileHover={isConfirming ? undefined : { scale: 1.02 }}
                      whileTap={isConfirming ? undefined : { scale: 0.98 }}
                    >
                      <Check className="w-3.5 h-3.5" />
                      <span>{isConfirming ? '应用中…' : '接受'}</span>
                    </motion.button>
                    <motion.button
                      onClick={() => setShowRevisionPanel(true)}
                      className="flex items-center gap-1.5 px-3 py-1.5 bg-black/10 dark:bg-white/10 hover:bg-black/15 dark:hover:bg-white/15 text-text-secondary text-xs rounded-lg transition-all border border-black/10 dark:border-white/10"
                      whileHover={{ scale: 1.02 }}
                      whileTap={{ scale: 0.98 }}
                      title="查看全部修改"
                    >
                      <Eye className="w-3.5 h-3.5" />
                      <span>查看全部</span>
                    </motion.button>
                  </div>
                </div>
                
                {/* 修改预览 */}
                <div className="px-4 py-3 max-h-28 overflow-y-auto bg-black/5 dark:bg-white/5">
                  <div className="flex items-start gap-3 text-xs">
                    <div className="flex-1 min-w-0">
                      <div className="text-[10px] text-text-dim mb-1">删除</div>
                      <div className="text-red-400/90 line-through break-all bg-red-500/10 px-2 py-1 rounded border border-red-500/20" title={(pendingChanges[pendingChanges.length - 1]?.meta as any)?.searchText || ''}>
                        {(() => {
                          const searchText = ((pendingChanges[pendingChanges.length - 1]?.meta as any)?.searchText || '') as string
                          return searchText.length > 60 ? searchText.slice(0, 60) + '...' : searchText
                        })()}
                      </div>
                    </div>
                    <div className="text-text-dim pt-5">→</div>
                    <div className="flex-1 min-w-0">
                      <div className="text-[10px] text-text-dim mb-1">替换为</div>
                      <div className="text-success/90 break-all bg-success/10 px-2 py-1 rounded border border-success/20" title={(pendingChanges[pendingChanges.length - 1]?.meta as any)?.replaceText || ''}>
                        {(() => {
                          const replaceText = ((pendingChanges[pendingChanges.length - 1]?.meta as any)?.replaceText || '') as string
                          return replaceText.length > 60 ? replaceText.slice(0, 60) + '...' : replaceText
                        })()}
                      </div>
                    </div>
                  </div>
                </div>
                
                {/* 快捷键提示 */}
                <div className="px-4 py-1.5 bg-black/10 dark:bg-white/10 backdrop-blur-md border-t border-border flex items-center justify-center gap-4 text-[10px] text-text-dim">
                  <span><kbd className="px-1 py-0.5 bg-black/10 dark:bg-white/10 rounded text-text-dim border border-black/10 dark:border-white/10">Enter</kbd> 接受</span>
                  <span><kbd className="px-1 py-0.5 bg-black/10 dark:bg-white/10 rounded text-text-dim border border-black/10 dark:border-white/10">Esc</kbd> 拒绝</span>
                </div>
              </div>
            </motion.div>
          )}
        </AnimatePresence>
        
        {/* AI 处理中指示器 */}
        {isAIProcessing && (
          <div className="fixed bottom-20 left-1/2 -translate-x-1/2 z-50">
            <div className="flex items-center gap-2 px-4 py-2 glass-card border border-border rounded-lg shadow-xl">
              <Loader2 className="w-4 h-4 text-accent animate-spin" />
              <span className="text-xs text-text-secondary">AI 正在处理...</span>
            </div>
          </div>
        )}
      </div>
      
      {/* 底部状态栏（Word 风格：页码/字数/缩放/视图） */}
      <div className="h-10 flex-shrink-0 glass border-t border-border flex items-center justify-between px-3 gap-3 select-none">
        {/* 左侧：页码、字数 */}
        <div className="flex items-center gap-3 text-[11px] text-text-secondary min-w-0">
          <span className="truncate">
            第 {pageStats.current} 页 / 共 {pageStats.total} 页
          </span>
          <span className="opacity-50">|</span>
          <span className="truncate">
            字 {docStats.chars} · 词 {docStats.words}
          </span>
          {docStats.selectionChars > 0 && (
            <>
              <span className="opacity-50">|</span>
              <span className="truncate">选中 {docStats.selectionChars}</span>
            </>
          )}
        </div>

        {/* 中间：视图模式 */}
        <div className="flex items-center gap-1 bg-black/5 dark:bg-white/5 border border-black/10 dark:border-white/10 rounded-xl p-1">
          <button
            onClick={() => setViewMode('print')}
            className={`px-2.5 py-1 rounded-lg text-[11px] transition-colors ${
              viewMode === 'print'
                ? 'bg-accent/14 text-accent border border-accent/25'
                : 'text-text-muted hover:text-text hover:bg-white/5'
            }`}
            title="打印布局"
          >
            <Layout className="w-3.5 h-3.5 inline -mt-0.5 mr-1" />
            打印
          </button>
          <button
            onClick={() => setViewMode('web')}
            className={`px-2.5 py-1 rounded-lg text-[11px] transition-colors ${
              viewMode === 'web'
                ? 'bg-accent/14 text-accent border border-accent/25'
                : 'text-text-muted hover:text-text hover:bg-white/5'
            }`}
            title="网页布局"
          >
            <Grid className="w-3.5 h-3.5 inline -mt-0.5 mr-1" />
            网页
          </button>
          {/* 专注模式已移除：保持 Word 顶部功能栏始终可见 */}
        </div>

        {/* 右侧：缩放 */}
        <div className="flex items-center gap-2 text-[11px] text-text-secondary">
          <button
            onClick={() => setZoomLevel((prev) => Math.max(25, prev - 10))}
            className="px-2 py-1 rounded-lg hover:bg-white/5 border border-black/10 dark:border-white/10"
            title="缩小 (Ctrl + -)"
          >
            −
          </button>
          <input
            type="range"
            min={25}
            max={500}
            step={5}
            value={zoomLevel}
            onChange={(e) => setZoomLevel(Number(e.target.value))}
            className="w-28 accent-[var(--accent)]"
            title="缩放"
          />
          <button
            onClick={() => setZoomLevel((prev) => Math.min(500, prev + 10))}
            className="px-2 py-1 rounded-lg hover:bg-white/5 border border-black/10 dark:border-white/10"
            title="放大 (Ctrl + +)"
          >
            +
          </button>
          <button
            onClick={() => setZoomLevel(100)}
            className="px-2 py-1 rounded-lg hover:bg-white/5 border border-black/10 dark:border-white/10"
            title="重置为 100% (Ctrl + 0)"
          >
            {zoomLevel}%
          </button>
          <button
            onClick={() => setShowRevisionPanel(true)}
            className="px-2 py-1 rounded-lg hover:bg-white/5 border border-black/10 dark:border-white/10"
            title="修订/修改记录"
          >
            <Eye className="w-3.5 h-3.5" />
          </button>
          <button
            onClick={() => setShowCommentPanel(true)}
            className="px-2 py-1 rounded-lg hover:bg-white/5 border border-black/10 dark:border-white/10"
            title="批注"
          >
            <MessageSquare className="w-3.5 h-3.5" />
          </button>
        </div>
      </div>

      {/* 底部裁剪标记容器 */}
      <div className={`crop-marks-bottom pointer-events-none absolute inset-0 z-10 ${viewMode === 'print' ? '' : 'hidden'}`}></div>
      
      {/* Ctrl+K 内联编辑弹窗 */}
      {showInlineEdit && (
        <InlineEditPopup
          selectedText={selectedTextForEdit}
          selectedHtml={selectedHtmlForEdit}
          position={inlineEditPosition}
          onClose={() => setShowInlineEdit(false)}
          onApply={applyInlineEdit}
        />
      )}
      
      {/* 右键菜单 */}
      {showContextMenu && (
        <ContextMenu
          position={contextMenuPosition}
          selectedText={selectedTextForEdit}
          onClose={() => setShowContextMenu(false)}
          onAction={handleContextMenuAction}
        />
      )}

      <RevisionPanel open={showRevisionPanel} onClose={() => setShowRevisionPanel(false)} />
      <CommentPanel
        open={showCommentPanel}
        onClose={() => setShowCommentPanel(false)}
        onAddComment={applyCommentToSelection}
        onReply={replyTo}
        onLocate={locateCommentInDom}
      />

      <style>{`
        /* 编辑器容器背景 */
        .word-editor-container {
          background-color: var(--word-canvas-bg);
          padding: 20px;
          display: flex;
          flex-direction: column;
          align-items: center;
          justify-content: flex-start;
          min-height: 100%;
        }

        .word-editor-container.print-mode {
          padding-bottom: calc(20px + 24px);
        }

        /* 视图：网页布局（更像“无纸张”的阅读/编辑） */
        .word-editor-container.web-mode {
          padding: 16px;
        }


        /* 缩放容器 */
        .zoom-container {
          display: flex;
          flex-direction: column;
          align-items: center;
          width: 100%;
        }

        /* 缩放指示器 - 已迁移到底部状态栏（保留类名但隐藏，避免遮挡） */
        .zoom-indicator {
          display: none;
        }

        .zoom-btn {
          display: none;
        }

        .zoom-level {
          display: none;
        }

        /* 页面包裹容器：支持动态页边距 */
        .word-page {
          --page-margin-top: 2.54cm;
          --page-margin-bottom: 2.54cm;
          --page-margin-left: 3.17cm;
          --page-margin-right: 3.17cm;
          flex: 0 0 auto;
          width: var(--page-width, var(--a4-width));
          max-width: var(--page-width, var(--a4-width));
          min-width: var(--page-width, var(--a4-width));
          min-height: var(--page-height, var(--a4-height));
          margin: 0 auto 24px;
          background: var(--word-page-bg);
          border: 1px solid transparent;
          padding: 0;
          box-shadow: var(--word-page-shadow);
          position: relative;
          /* 不使用 overflow:hidden，让分页占位块自然展示 */
        }

        /* 网页布局：不显示“纸张”、阴影、裁剪标记 */
        .word-page.web {
          width: min(1100px, 100%);
          max-width: 1100px;
          min-height: auto;
          margin: 0 auto 16px;
          background: transparent;
          box-shadow: none;
        }

        .word-page.web::before,
        .word-page.web::after {
          display: none;
        }

        .word-page.web .word-editor-content {
          width: 100%;
          min-height: auto;
          background: transparent;
          padding: 18px 22px;
          margin: 0;
          box-shadow: none;
          border-radius: 12px;
          border: 1px solid var(--word-page-border);
        }

        /* 横向页面 */
        .word-page.landscape {
          width: var(--page-width, var(--a4-height));
          max-width: var(--page-width, var(--a4-height));
          min-height: var(--page-height, var(--a4-width));
        }

        .word-page.landscape .word-editor-content {
          width: var(--page-width, 297mm);
          min-height: var(--page-height, 210mm);
        }

        /* 页眉样式 - 参考 Word 的浅色显示 */
        .page-header {
          position: absolute;
          top: 0.8cm;
          left: var(--page-margin-left);
          right: var(--page-margin-right);
          padding-bottom: 0.3cm;
          border-bottom: 1px dashed var(--word-rule);
          font-size: 9pt;
          color: var(--word-ink-dim);
          opacity: 0.7;
          z-index: 10;
        }

        /* 页脚样式 - 参考 Word 的浅色显示 */
        .page-footer {
          position: absolute;
          bottom: 0.8cm;
          left: var(--page-margin-left);
          right: var(--page-margin-right);
          padding-top: 0.3cm;
          border-top: 1px dashed var(--word-rule);
          font-size: 9pt;
          color: var(--word-ink-dim);
          opacity: 0.7;
          z-index: 10;
        }

        .page-number {
          margin-left: 0.5em;
        }

        /* DOCX 页眉/页脚内容样式 */
        .docx-header-footer {
          line-height: 1.5;
        }
        
        .docx-header-footer .header-footer-para {
          margin: 0;
          padding: 0;
        }
        
        .docx-header-footer .hf-tabs {
          display: flex;
          justify-content: space-between;
          align-items: center;
        }
        
        .docx-header-footer .hf-left {
          text-align: left;
        }
        
        .docx-header-footer .hf-center {
          text-align: center;
          flex: 1;
        }
        
        .docx-header-footer .hf-right {
          text-align: right;
        }

        /* 裁剪标记 (Crop Marks) - 使用动态边距 */
        .word-page::before,
        .word-page::after {
          content: '';
          position: absolute;
          width: 20px;
          height: 20px;
          border: 1px solid rgb(255 255 255 / 0.22);
          pointer-events: none;
        }

        /* 左上角标记 */
        .word-page::before {
          top: var(--page-margin-top);
          left: var(--page-margin-left);
          border-right: none;
          border-bottom: none;
        }

        /* 右上角标记 */
        .word-page::after {
          top: var(--page-margin-top);
          right: var(--page-margin-right);
          border-left: none;
          border-bottom: none;
        }

        /* 添加额外的标记容器用于底部标记 */
        .crop-marks-bottom::before,
        .crop-marks-bottom::after {
          content: '';
          position: absolute;
          width: 20px;
          height: 20px;
          border: 1px solid rgb(255 255 255 / 0.22);
          pointer-events: none;
        }

        /* 左下角标记 */
        .crop-marks-bottom::before {
          bottom: var(--page-margin-bottom, 2.54cm);
          left: var(--page-margin-left, 3.17cm);
          border-right: none;
          border-top: none;
        }

        /* 右下角标记 */
        .crop-marks-bottom::after {
          bottom: var(--page-margin-bottom, 2.54cm);
          right: var(--page-margin-right, 3.17cm);
          border-left: none;
          border-top: none;
        }

        /* 正文样式 - 精确模拟 Word 默认样式 + 多页白纸背景（关键） */
        .word-editor-content {
          outline: none;
          /* 让"白纸背景"跟随内容无限延伸 */
          width: var(--page-width, 210mm);
          max-width: none;
          min-width: var(--page-width, 210mm);
          margin: 0 auto 24px auto;
          box-sizing: border-box;
          background: var(--word-page-bg);
          /* 使用 CSS 变量支持动态页边距 */
          padding: var(--page-margin-top, 2.54cm) var(--page-margin-right, 3.17cm) var(--page-margin-bottom, 2.54cm) var(--page-margin-left, 3.17cm);
          min-height: var(--page-height, 297mm); /* 至少一页高度 */
          box-shadow: none;
          position: relative;
          /* 默认字体/字号/行距：由 CSS 变量控制，便于根据 docx 的 Normal/Theme 动态覆盖 */
          font-family: var(--word-font-family-cn);
          font-size: var(--word-font-size); /* 10.5pt @ 96 DPI */
          line-height: var(--word-line-height);
          color: var(--word-ink);
          text-align: left; /* Word 默认左对齐 */
        }

        /* 分页占位块：包含（本页剩余白底）+（页间灰色间距）+（下一页顶端留白） */
        .pm-page-break {
          /* 扩展到整页宽度（含左右页边距） */
          margin-left: calc(var(--page-margin-left, 3.17cm) * -1);
          margin-right: calc(var(--page-margin-right, 3.17cm) * -1);
          width: calc(100% + var(--page-margin-left, 3.17cm) + var(--page-margin-right, 3.17cm));
          display: block;
          box-sizing: content-box;
          /* 用实色三段拼接：白纸底部 + 灰色间距 + 白纸顶部 */
          background: linear-gradient(
            to bottom,
            var(--word-page-bg) 0,
            var(--word-page-bg) var(--fill, 0px),
            var(--word-canvas-bg) var(--fill, 0px),
            var(--word-canvas-bg) calc(var(--fill, 0px) + var(--gap, 32px)),
            var(--word-page-bg) calc(var(--fill, 0px) + var(--gap, 32px)),
            var(--word-page-bg) 100%
          );
          position: relative;
          z-index: 1;
        }

        /* 末页填充占位块：仅补足到页底 */
        .pm-page-end {
          margin-left: calc(var(--page-margin-left, 3.17cm) * -1);
          margin-right: calc(var(--page-margin-right, 3.17cm) * -1);
          width: calc(100% + var(--page-margin-left, 3.17cm) + var(--page-margin-right, 3.17cm));
          display: block;
          box-sizing: content-box;
          background: transparent;
        }

        /* 取消页间分隔线，避免出现细白线 */
        .pm-page-break::before {
          content: '';
          display: none;
        }

        .word-editor-content:focus {
          outline: none;
        }

        /* 段落样式 - 不添加默认缩进，由解析的内联样式控制 */
        .word-editor-content p {
          margin: 0;
          padding: 0;
          /* 不设置默认 text-indent，由 docx 解析的样式控制 */
        }

        /* 空段落保持行高 */
        .word-editor-content p:empty::after {
          content: '\\00a0';
        }

        /* TOC 段落：点引导线 */
        .docx-toc {
          width: 100%;
          max-width: var(--docx-tab-pos, 100%);
          display: flex;
          align-items: baseline;
          overflow: hidden;
          min-width: 0;
        }

        .word-editor-content .docx-toc a {
          color: inherit !important;
          text-decoration: none;
        }

        .docx-toc > a {
          display: flex;
          align-items: baseline;
          flex: 1 1 auto;
          overflow: hidden;
          min-width: 0;
        }

        .docx-toc-inner {
          display: flex;
          align-items: baseline;
          flex: 1 1 auto;
          overflow: hidden;
        }

        .docx-toc-inner > a {
          display: flex;
          align-items: baseline;
          flex: 1 1 auto;
          overflow: hidden;
          color: inherit;
          text-decoration: none;
        }

        .docx-toc-left {
          flex: 0 0 auto;
          white-space: nowrap;
        }

        .docx-toc-leader {
          flex: 1 1 auto;
          overflow: hidden;
          white-space: nowrap;
          letter-spacing: 0.15em;
          min-width: 1em;
        }

        .docx-toc-right {
          flex: 0 0 auto;
          white-space: nowrap;
        }

        .docx-tab {
          display: inline-block;
          width: 36pt; /* Word 默认制表位 0.5 inch */
        }

        .docx-list-marker {
          display: inline-block;
          white-space: pre;
        }

        .docx-image-only img {
          display: block;
          margin: 0 auto;
        }

        /* 标题样式 - 让内联样式优先，这里只设置基础样式 */
        .word-editor-content h1,
        .word-editor-content h2,
        .word-editor-content h3,
        .word-editor-content h4,
        .word-editor-content h5,
        .word-editor-content h6 {
          margin: 0;
          padding: 0;
          font-weight: bold;
        }

        /* 如果标题没有内联样式，使用这些默认值 */
        .word-editor-content h1:not([style*="font-size"]) {
          font-size: 22pt;
        }
        .word-editor-content h2:not([style*="font-size"]) {
          font-size: 16pt;
        }
        .word-editor-content h3:not([style*="font-size"]) {
          font-size: 14pt;
        }

        /* 无序列表 */
        .word-editor-content ul {
          padding-left: 0;
          margin: 0 0 0 1.5em;
          list-style-type: disc;
          list-style-position: outside;
        }

        /* 有序列表 */
        .word-editor-content ol {
          padding-left: 0;
          margin: 0 0 0 1.5em;
          list-style-type: decimal;
          list-style-position: outside;
        }

        .word-editor-content li {
          margin: 0;
          padding: 0 0 0 0.3em;
        }

        .word-editor-content li p {
          margin: 0;
          display: inline;
        }

        /* 嵌套列表 */
        .word-editor-content li > ul,
        .word-editor-content li > ol {
          margin-left: 1.5em;
        }

        /* 引用块 */
        .word-editor-content blockquote {
          border-left: 3px solid var(--word-rule);
          padding-left: 1em;
          margin: 0 0 0 2em;
          color: var(--word-ink-muted);
        }

        .word-editor-content blockquote p {
          text-indent: 0 !important;
        }

        /* 分隔线 */
        .word-editor-content hr {
          border: none;
          border-top: 1px solid var(--word-rule);
          margin: 6pt 0;
        }

        /* 文本格式 */
        .word-editor-content strong,
        .word-editor-content b {
          font-weight: bold;
        }

        .word-editor-content em,
        .word-editor-content i {
          font-style: italic;
        }

        .word-editor-content u {
          text-decoration: underline;
        }

        .word-editor-content s,
        .word-editor-content del {
          text-decoration: line-through;
        }

        /* 表格 */
        .word-editor-content table {
          border-collapse: collapse;
          width: auto;
          margin: 0;
          table-layout: fixed;
        }

        .word-editor-content table[data-tbl-grid] {
          width: var(--table-width, auto);
          table-layout: fixed;
        }

        .word-editor-content table[data-tbl-layout="autofit"] {
          table-layout: auto;
        }

        .word-editor-content th,
        .word-editor-content td {
          border: 0.5pt solid var(--word-rule);
          padding: 2pt 5pt;
          text-align: left;
          vertical-align: top;
        }

        .word-editor-content th {
          font-weight: bold;
        }

        /* 图片 */
        .word-editor-content img {
          max-width: 100%;
          height: auto;
        }

        /* 链接 - Word 默认超链接样式 */
        .word-editor-content a {
          color: var(--accent);
          text-decoration: underline;
        }

        /* 上下标 */
        .word-editor-content sup {
          font-size: 0.65em;
          vertical-align: super;
          line-height: 0;
        }

        .word-editor-content sub {
          font-size: 0.65em;
          vertical-align: sub;
          line-height: 0;
        }

        /* span 样式继承 */
        .word-editor-content span {
          /* 保持内联样式 */
        }

        /* 块级修订：old/new */
        .word-editor-content [data-diff-role="old"] {
          background: rgba(239, 68, 68, 0.10);
          text-decoration: line-through;
        }

        .word-editor-content [data-diff-role="new"] {
          background: rgba(34, 197, 94, 0.10);
        }

        .word-editor-content .ProseMirror-selectednode {
          outline: 2px solid rgb(var(--accent-rgb) / 0.9);
        }

        /* 占位符 */
        .word-editor-content p.is-editor-empty:first-child::before {
          color: var(--word-ink-muted);
          content: attr(data-placeholder);
          float: left;
          height: 0;
          pointer-events: none;
        }

        /* 表格选中 */
        .word-editor-content .selectedCell:after {
          background: rgb(var(--accent-rgb) / 0.10);
          content: "";
          position: absolute;
          left: 0;
          right: 0;
          top: 0;
          bottom: 0;
          pointer-events: none;
          z-index: 2;
        }

        .word-editor-content .column-resize-handle {
          background-color: rgb(var(--accent-rgb) / 0.85);
          position: absolute;
          right: -2px;
          top: 0;
          bottom: -2px;
          width: 4px;
          pointer-events: none;
        }

        .word-editor-content .tableWrapper {
          overflow-x: auto;
        }

        .word-editor-content .resize-cursor {
          cursor: col-resize;
        }
      `}</style>
    </div>
  )
}
