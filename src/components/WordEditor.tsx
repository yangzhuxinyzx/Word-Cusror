import { useEffect, useState, useCallback, useRef } from 'react'
import { useEditor, EditorContent } from '@tiptap/react'
import StarterKit from '@tiptap/starter-kit'
import { Underline } from '@tiptap/extension-underline'
import { TextAlign } from '@tiptap/extension-text-align'
import { TextStyle } from '@tiptap/extension-text-style'
import { Color } from '@tiptap/extension-color'
import { Highlight } from '@tiptap/extension-highlight'
import { Table } from '@tiptap/extension-table'
import { TableRow } from '@tiptap/extension-table-row'
import { TableHeader } from '@tiptap/extension-table-header'
import { TableCell } from '@tiptap/extension-table-cell'
import { Image } from '@tiptap/extension-image'
import { Link } from '@tiptap/extension-link'
import { Placeholder } from '@tiptap/extension-placeholder'
import { FontFamily } from '@tiptap/extension-font-family'
import { Subscript } from '@tiptap/extension-subscript'
import { Superscript } from '@tiptap/extension-superscript'
import { DiffOld, DiffNew } from '../extensions/DiffMark'
import { DiffBlock } from '../extensions/DiffBlock'
import { motion, AnimatePresence, animate } from 'framer-motion'
import { useDocument } from '../context/DocumentContext'
import { useAI } from '../context/AIContext'
import { 
  Bold, Italic, Underline as UnderlineIcon, Strikethrough,
  AlignLeft, AlignCenter, AlignRight, AlignJustify,
  List, ListOrdered, Undo, Redo, 
  Save, Download, FileText, Loader2,
  Table as TableIcon, Link as LinkIcon,
  Heading1, Heading2, Heading3, Minus, Quote,
  Check, X, Sparkles, Eye,
  Superscript as SuperscriptIcon, Subscript as SubscriptIcon
} from 'lucide-react'
import { parseDocxToHtml } from '../utils/docxParser'
import InlineEditPopup from './InlineEditPopup'
import ContextMenu from './ContextMenu'
import ExcelPreview from './ExcelPreview'
import PptPreviewHtml from './PptPreviewHtml'
import RevisionPanel from './RevisionPanel'

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
      result.push('<hr style="border: none; border-top: 1px solid #ccc; margin: 1.5em 0;">')
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
      result.push(`<blockquote style="border-left: 3px solid #ccc; padding-left: 1em; margin: 0.5em 0; color: #666;">${text}</blockquote>`)
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

      const blocks = Array.from(root.children).filter((el) => {
        const h = (el as HTMLElement).offsetHeight
        if (h <= 0) return false
        return !(el as HTMLElement).classList?.contains('pm-page-break')
      }) as HTMLElement[]

      const decorations: any[] = []
      let y = TOP_MARGIN_PX // 第一页从上边距之后开始（对应 .word-editor-content 的 padding-top）

      for (const el of blocks) {
        const h = el.getBoundingClientRect().height || el.offsetHeight || 0
        if (h <= 0) continue

        // 只要“下一个块”会进入底边距区域，就在它前面插入换页占位
        const contentBottomLimit = PAGE_HEIGHT_PX - BOTTOM_MARGIN_PX
        const wouldOverflow = y + h > contentBottomLimit

        // 避免在一页开头反复插入空页（遇到超大表格/块时允许其自然溢出）
        const atPageStart = Math.abs(y - TOP_MARGIN_PX) < 1

        if (wouldOverflow && !atPageStart) {
          const pos = view.posAtDOM(el, 0)
          const fillToEndOfPage = Math.max(0, PAGE_HEIGHT_PX - y) // 包含底边距的白色区域
          const widget = Decoration.widget(
            pos,
            () => {
              const dom = document.createElement('div')
              dom.className = 'pm-page-break'
              dom.style.setProperty('--fill', `${fillToEndOfPage}px`)
              dom.style.setProperty('--gap', `${PAGE_GAP_PX}px`)
              dom.style.setProperty('--top', `${TOP_MARGIN_PX}px`)
              dom.style.height = `${fillToEndOfPage + PAGE_GAP_PX + TOP_MARGIN_PX}px`
              return dom
            },
            { key: `pm-pb-${pos}` }
          )
          decorations.push(widget)

          // 新的一页：从顶端留白后开始
          y = TOP_MARGIN_PX + h
        } else {
          y += h
        }
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
    headerFooterSetup,
  } = useDocument()
  const { getCompletion, cancelCompletion, isCompleting, settings } = useAI()
  const [isLoading, setIsLoading] = useState(false)
  const [currentFontSize, setCurrentFontSize] = useState('12pt')
  const [currentFontFamily, setCurrentFontFamily] = useState('仿宋')
  const [isConfirming, setIsConfirming] = useState(false)
  const [showRevisionPanel, setShowRevisionPanel] = useState(false)
  
  // 补全相关状态
  const [completionSuggestion, setCompletionSuggestion] = useState<string | null>(null)
  const [showCompletion, setShowCompletion] = useState(false)
  const lastTextRef = useRef<string>('')
  const animatedDiffIdsRef = useRef<Set<string>>(new Set())
  const docEntryAnimationRef = useRef<number>(0)
  
  // 跟踪上一次同步的文档内容，避免重复同步
  const lastSyncedContentRef = useRef<string>('')
  
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
            const computedStyle = window.getComputedStyle(element)
            const fontSize = computedStyle.fontSize
            if (fontSize) {
              // 转换 px 到 pt (1pt = 1.333px)
              const pxValue = parseFloat(fontSize)
              const ptValue = Math.round(pxValue / 1.333)
              setCurrentFontSize(`${ptValue}pt`)
            }
          }
        } else {
          setCurrentFontSize('12pt')
        }
      }

      // 字体
      if (textStyleAttrs.fontFamily) {
        setCurrentFontFamily(textStyleAttrs.fontFamily)
      } else {
        // 尝试从 DOM 获取
        const domSelection = window.getSelection()
        if (domSelection && domSelection.rangeCount > 0) {
          const range = domSelection.getRangeAt(0)
          let element = range.startContainer as HTMLElement
          if (element.nodeType === Node.TEXT_NODE) {
            element = element.parentElement as HTMLElement
          }
          if (element) {
            const computedStyle = window.getComputedStyle(element)
            const fontFamily = computedStyle.fontFamily
            if (fontFamily) {
              // 提取第一个字体名称
              const firstFont = fontFamily.split(',')[0].replace(/['"]/g, '').trim()
              if (firstFont) {
                setCurrentFontFamily(firstFont)
              }
            }
          }
        } else {
          setCurrentFontFamily('仿宋')
        }
      }
    } catch (error) {
      // 忽略错误，使用默认值
      setCurrentFontSize('12pt')
      setCurrentFontFamily('仿宋')
    }
  }

  const editor = useEditor({
    extensions: [
      StarterKit.configure({
        heading: { levels: [1, 2, 3] },
      }),
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
      Table.configure({
        resizable: true,
      }),
      TableRow,
      TableHeader,
      TableCell,
      Image,
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
      
      // 调试：检查编辑器输出的 HTML 是否包含 diff 标记
      const hasDiffOld = html.includes('diff-old')
      const hasDiffNew = html.includes('diff-new')
      console.log('[WordEditor onUpdate] editor.getHTML() length:', html.length, 'hasDiffOld:', hasDiffOld, 'hasDiffNew:', hasDiffNew)
      
      updateContent(html)
      
      // 清除补全建议（用户继续输入时）
      if (completionSuggestion) {
        setCompletionSuggestion(null)
        setShowCompletion(false)
      }
      
      // 获取当前文本，用于触发补全
      const text = editor.getText()
      lastTextRef.current = text
    },
    onSelectionUpdate: ({ editor }) => {
      // 选区变化时更新工具栏显示
      updateCurrentStyles(editor)
    },
  })

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
          applySelectionRevision(oldText, newText, hasHtmlFormat)
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
        if (document.content && document.content.length > 100) {
          // 检查是否需要同步（与上次同步的内容不同）
          if (document.content !== lastSyncedContentRef.current) {
            console.log('[WordEditor] Content changed, syncing to editor')
            console.log(`  document.content length: ${document.content.length}`)
            console.log(`  lastSynced length: ${lastSyncedContentRef.current.length}`)
            console.log(`  has diff marks: ${document.content.includes('diff-')}`)
            
            // 强制清空再设置，确保完全刷新
            editor.commands.clearContent()
            editor.commands.setContent(document.content)
            lastSyncedContentRef.current = document.content
          }
          return
        }
        
        // 首次解析 docx
        setIsLoading(true)
        try {
          // 使用自定义解析器，保留字体大小等样式
          const htmlContent = await parseDocxToHtml(docxData)
          editor.commands.setContent(htmlContent)
          lastSyncedContentRef.current = htmlContent
          // 同步内容到 document.content，这样 AI 替换才能正常工作
          // 使用 setTimeout 避免在 effect 中直接触发状态更新
          setTimeout(() => {
            updateContent(htmlContent)
          }, 0)
        } catch (err) {
          console.error('Failed to parse docx:', err)
          // 显示错误提示
          editor.commands.setContent(`
            <div style="padding: 40px; text-align: center; color: #888;">
              <p style="font-size: 18px; margin-bottom: 10px;">⚠️ 文档解析失败</p>
              <p style="font-size: 14px;">无法解析此 Word 文档，可能是文件损坏或格式不兼容。</p>
              <p style="font-size: 12px; margin-top: 10px; color: #666;">错误信息: ${(err as Error).message}</p>
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
            console.log('[WordEditor] Updating editor content')
            console.log(`  hasDiffMarkers: ${hasDiffMarkers}, currentHasDiff: ${currentHasDiff}`)
            console.log(`  document.content length: ${document.content.length}`)
            editor.commands.setContent(document.content)
          }
        } else {
          // Markdown 转 HTML
          const html = markdownToHtml(document.content)
          if (currentHtml !== html) {
            editor.commands.setContent(html)
          }
        }
      } else {
        editor.commands.setContent('')
      }
    }

    loadContent()
  }, [docxData, document.content, editor, updateContent])

  useEffect(() => {
    animatedDiffIdsRef.current.clear()
  }, [currentFile?.path])

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
        el.style.backgroundColor = 'rgba(34, 197, 94, 0.4)'
        el.style.boxShadow = '0 0 0 3px rgba(34, 197, 94, 0.6)'
        
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
      
      // 2. 查找带有内联 diff 样式的元素
      const allSpans = editorElement.querySelectorAll('span')
      for (const span of allSpans) {
        const style = span.getAttribute('style') || ''
        if (style.includes('#bbf7d0') && span.textContent?.includes(text)) {
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
      node.style.outline = '2px solid rgba(59, 130, 246, 0.9)'
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

  // 工具栏按钮
  const ToolbarButton = ({ 
    onClick, 
    active, 
    disabled, 
    children, 
    title 
  }: { 
    onClick: () => void
    active?: boolean
    disabled?: boolean
    children: React.ReactNode
    title: string
  }) => (
    <button
      onClick={onClick}
      disabled={disabled}
      title={title}
      className={`p-1.5 rounded transition-all ${
        active 
          ? 'bg-primary text-white' 
          : 'text-text-muted hover:bg-surface-hover hover:text-text'
      } ${disabled ? 'opacity-50 cursor-not-allowed' : ''}`}
    >
      {children}
    </button>
  )

  const ToolbarDivider = () => (
    <div className="w-px h-6 bg-border mx-1" />
  )

  if (!editor) {
    return (
      <div className="flex items-center justify-center h-full bg-neutral-200">
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
    <div className="flex flex-col h-full bg-neutral-300 overflow-hidden">
      {/* 顶部信息栏 */}
      <div className="flex items-center justify-between px-4 py-2 bg-background border-b border-border">
        <div className="flex items-center gap-2">
          <FileText className="w-4 h-4 text-primary" />
          <span className="text-sm font-medium text-text">{document.title || '未命名文档'}</span>
          {hasUnsavedChanges && (
            <span className="text-xs text-amber-400">• 未保存</span>
          )}
        </div>
        <div className="flex items-center gap-2">
          <button
            onClick={() => saveDocument()}
            className="flex items-center gap-1.5 px-3 py-1.5 bg-primary text-white text-xs rounded-md hover:bg-primary-hover transition-all"
          >
            <Save className="w-3.5 h-3.5" />
            <span>保存</span>
          </button>
          <button
            onClick={() => saveDocument()}
            className="flex items-center gap-1.5 px-3 py-1.5 bg-surface border border-border text-text text-xs rounded-md hover:bg-surface-hover transition-all"
          >
            <Download className="w-3.5 h-3.5" />
            <span>导出</span>
          </button>
        </div>
      </div>

      {/* 工具栏 */}
      <div className="flex flex-wrap items-center gap-0.5 px-4 py-2 bg-surface border-b border-border">
        {/* 字体选择 */}
        <select
          value={currentFontFamily}
          onChange={(e) => {
            setCurrentFontFamily(e.target.value)
            editor.chain().focus().setFontFamily(e.target.value).run()
          }}
          className="h-7 px-2 text-xs bg-background border border-border rounded text-text"
        >
          <option value="仿宋">仿宋</option>
          <option value="宋体">宋体</option>
          <option value="黑体">黑体</option>
          <option value="楷体">楷体</option>
          <option value="微软雅黑">微软雅黑</option>
          <option value="Times New Roman">Times New Roman</option>
          <option value="Arial">Arial</option>
        </select>

        {/* 字号选择 */}
        <select
          value={currentFontSize}
          onChange={(e) => {
            setCurrentFontSize(e.target.value)
            editor.chain().focus().setMark('textStyle', { fontSize: e.target.value }).run()
          }}
          className="h-7 px-2 text-xs bg-background border border-border rounded text-text ml-1"
        >
          <option value="10pt">五号</option>
          <option value="10.5pt">五号半</option>
          <option value="12pt">小四</option>
          <option value="14pt">四号</option>
          <option value="15pt">小三</option>
          <option value="16pt">三号</option>
          <option value="18pt">小二</option>
          <option value="22pt">二号</option>
          <option value="26pt">小一</option>
          <option value="36pt">一号</option>
        </select>

        <ToolbarDivider />

        {/* 撤销/重做 */}
        <ToolbarButton
          onClick={() => editor.chain().focus().undo().run()}
          disabled={!editor.can().undo()}
          title="撤销"
        >
          <Undo className="w-4 h-4" />
        </ToolbarButton>
        <ToolbarButton
          onClick={() => editor.chain().focus().redo().run()}
          disabled={!editor.can().redo()}
          title="重做"
        >
          <Redo className="w-4 h-4" />
        </ToolbarButton>

        <ToolbarDivider />

        {/* 标题 */}
        <ToolbarButton
          onClick={() => editor.chain().focus().toggleHeading({ level: 1 }).run()}
          active={editor.isActive('heading', { level: 1 })}
          title="标题1"
        >
          <Heading1 className="w-4 h-4" />
        </ToolbarButton>
        <ToolbarButton
          onClick={() => editor.chain().focus().toggleHeading({ level: 2 }).run()}
          active={editor.isActive('heading', { level: 2 })}
          title="标题2"
        >
          <Heading2 className="w-4 h-4" />
        </ToolbarButton>
        <ToolbarButton
          onClick={() => editor.chain().focus().toggleHeading({ level: 3 }).run()}
          active={editor.isActive('heading', { level: 3 })}
          title="标题3"
        >
          <Heading3 className="w-4 h-4" />
        </ToolbarButton>

        <ToolbarDivider />

        {/* 文本格式 */}
        <ToolbarButton
          onClick={() => editor.chain().focus().toggleBold().run()}
          active={editor.isActive('bold')}
          title="粗体"
        >
          <Bold className="w-4 h-4" />
        </ToolbarButton>
        <ToolbarButton
          onClick={() => editor.chain().focus().toggleItalic().run()}
          active={editor.isActive('italic')}
          title="斜体"
        >
          <Italic className="w-4 h-4" />
        </ToolbarButton>
        <ToolbarButton
          onClick={() => editor.chain().focus().toggleUnderline().run()}
          active={editor.isActive('underline')}
          title="下划线"
        >
          <UnderlineIcon className="w-4 h-4" />
        </ToolbarButton>
        <ToolbarButton
          onClick={() => editor.chain().focus().toggleStrike().run()}
          active={editor.isActive('strike')}
          title="删除线"
        >
          <Strikethrough className="w-4 h-4" />
        </ToolbarButton>
        <ToolbarButton
          onClick={() => editor.chain().focus().toggleSuperscript().run()}
          active={editor.isActive('superscript')}
          title="上标 (如 X²)"
        >
          <SuperscriptIcon className="w-4 h-4" />
        </ToolbarButton>
        <ToolbarButton
          onClick={() => editor.chain().focus().toggleSubscript().run()}
          active={editor.isActive('subscript')}
          title="下标 (如 H₂O)"
        >
          <SubscriptIcon className="w-4 h-4" />
        </ToolbarButton>

        <ToolbarDivider />

        {/* 对齐 */}
        <ToolbarButton
          onClick={() => editor.chain().focus().setTextAlign('left').run()}
          active={editor.isActive({ textAlign: 'left' })}
          title="左对齐"
        >
          <AlignLeft className="w-4 h-4" />
        </ToolbarButton>
        <ToolbarButton
          onClick={() => editor.chain().focus().setTextAlign('center').run()}
          active={editor.isActive({ textAlign: 'center' })}
          title="居中"
        >
          <AlignCenter className="w-4 h-4" />
        </ToolbarButton>
        <ToolbarButton
          onClick={() => editor.chain().focus().setTextAlign('right').run()}
          active={editor.isActive({ textAlign: 'right' })}
          title="右对齐"
        >
          <AlignRight className="w-4 h-4" />
        </ToolbarButton>
        <ToolbarButton
          onClick={() => editor.chain().focus().setTextAlign('justify').run()}
          active={editor.isActive({ textAlign: 'justify' })}
          title="两端对齐"
        >
          <AlignJustify className="w-4 h-4" />
        </ToolbarButton>

        <ToolbarDivider />

        {/* 列表 */}
        <ToolbarButton
          onClick={() => editor.chain().focus().toggleBulletList().run()}
          active={editor.isActive('bulletList')}
          title="无序列表"
        >
          <List className="w-4 h-4" />
        </ToolbarButton>
        <ToolbarButton
          onClick={() => editor.chain().focus().toggleOrderedList().run()}
          active={editor.isActive('orderedList')}
          title="有序列表"
        >
          <ListOrdered className="w-4 h-4" />
        </ToolbarButton>

        <ToolbarDivider />

        {/* 引用和分隔线 */}
        <ToolbarButton
          onClick={() => editor.chain().focus().toggleBlockquote().run()}
          active={editor.isActive('blockquote')}
          title="引用"
        >
          <Quote className="w-4 h-4" />
        </ToolbarButton>
        <ToolbarButton
          onClick={() => editor.chain().focus().setHorizontalRule().run()}
          title="分隔线"
        >
          <Minus className="w-4 h-4" />
        </ToolbarButton>

        <ToolbarDivider />

        {/* 表格 */}
        <ToolbarButton
          onClick={() => editor.chain().focus().insertTable({ rows: 3, cols: 3, withHeaderRow: true }).run()}
          title="插入表格"
        >
          <TableIcon className="w-4 h-4" />
        </ToolbarButton>

        {/* 链接 */}
        <ToolbarButton
          onClick={() => {
            const url = window.prompt('输入链接地址:')
            if (url) {
              editor.chain().focus().setLink({ href: url }).run()
            }
          }}
          active={editor.isActive('link')}
          title="插入链接"
        >
          <LinkIcon className="w-4 h-4" />
        </ToolbarButton>
      </div>

      {/* 编辑区域 - Word 风格界面 */}
      <div 
        ref={editorContainerRef}
        className="flex-1 overflow-auto word-editor-container relative"
      >
        {isLoading && (
          <div className="absolute inset-0 flex items-center justify-center bg-white/80 z-10">
            <div className="flex items-center gap-2 text-gray-500">
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
            className={`word-page relative ${pageSetup.orientation === 'landscape' ? 'landscape' : ''}`}
            style={{
              // 根据页面设置动态调整边距
              '--page-margin-top': pageSetup.margins.top,
              '--page-margin-bottom': pageSetup.margins.bottom,
              '--page-margin-left': pageSetup.margins.left,
              '--page-margin-right': pageSetup.margins.right,
            } as React.CSSProperties}
          >
            {/* 页眉 */}
            {headerFooterSetup.header?.content && (
              <div 
                className="page-header"
                style={{ textAlign: headerFooterSetup.header.alignment || 'center' }}
              >
                {headerFooterSetup.header.content}
              </div>
            )}
            
          <EditorContent editor={editor} />

            {/* 页脚 */}
            {(headerFooterSetup.footer?.content || headerFooterSetup.pageNumber?.enabled) && (
              <div 
                className="page-footer"
                style={{ textAlign: headerFooterSetup.footer?.alignment || headerFooterSetup.pageNumber?.alignment || 'center' }}
              >
                {headerFooterSetup.footer?.content}
                {headerFooterSetup.pageNumber?.enabled && headerFooterSetup.pageNumber.position === 'footer' && (
                  <span className="page-number">第 1 页</span>
                )}
              </div>
            )}
          </div>
        </div>
        
        {/* 缩放指示器 */}
        <div className="zoom-indicator" title="Ctrl + 滚轮缩放 | Ctrl+0 重置 | Ctrl++ 放大 | Ctrl+- 缩小">
          <button 
            onClick={() => setZoomLevel(prev => Math.max(25, prev - 10))}
            className="zoom-btn"
            title="缩小 (Ctrl + -)"
          >
            −
          </button>
          <span 
            className="zoom-level"
            onClick={() => setZoomLevel(100)}
            title="点击重置为 100% (Ctrl + 0)"
          >
            {zoomLevel}%
          </span>
          <button 
            onClick={() => setZoomLevel(prev => Math.min(500, prev + 10))}
            className="zoom-btn"
            title="放大 (Ctrl + +)"
          >
            +
          </button>
        </div>
        
        {/* AI 补全建议浮窗 */}
        {showCompletion && completionSuggestion && (
          <div className="fixed bottom-20 left-1/2 -translate-x-1/2 z-50 max-w-xl">
            <div className="bg-zinc-900 border border-zinc-700 rounded-xl shadow-2xl overflow-hidden">
              {/* 标题栏 */}
              <div className="flex items-center justify-between px-3 py-2 bg-zinc-800 border-b border-zinc-700">
                <div className="flex items-center gap-2">
                  <Sparkles className="w-4 h-4 text-violet-400" />
                  <span className="text-xs text-zinc-300">AI 补全建议</span>
                </div>
                <div className="flex items-center gap-1 text-[10px] text-zinc-500">
                  <kbd className="px-1.5 py-0.5 bg-zinc-700 rounded text-zinc-400">Tab</kbd>
                  <span>接受</span>
                  <kbd className="px-1.5 py-0.5 bg-zinc-700 rounded text-zinc-400 ml-2">Esc</kbd>
                  <span>取消</span>
                </div>
              </div>
              {/* 补全内容预览 */}
              <div className="px-4 py-3">
                <p className="text-sm text-green-400 leading-relaxed whitespace-pre-wrap">
                  {completionSuggestion}
                </p>
              </div>
              {/* 操作按钮 */}
              <div className="flex items-center gap-2 px-3 py-2 bg-zinc-800/50 border-t border-zinc-700">
                <button
                  onClick={acceptCompletion}
                  className="flex items-center gap-1.5 px-3 py-1.5 bg-green-600 hover:bg-green-500 text-white text-xs rounded-lg transition-colors"
                >
                  <Check className="w-3 h-3" />
                  <span>接受</span>
                </button>
                <button
                  onClick={dismissCompletion}
                  className="flex items-center gap-1.5 px-3 py-1.5 bg-zinc-700 hover:bg-zinc-600 text-zinc-300 text-xs rounded-lg transition-colors"
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
            <div className="flex items-center gap-2 px-4 py-2 bg-zinc-900 border border-zinc-700 rounded-lg shadow-xl">
              <Loader2 className="w-4 h-4 text-violet-400 animate-spin" />
              <span className="text-xs text-zinc-300">AI 正在思考...</span>
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
              <div className="bg-zinc-900/95 backdrop-blur-md border border-zinc-700/50 rounded-xl shadow-2xl overflow-hidden">
                {/* 顶部蓝色指示条 - 表示这是 AI 建议 */}
                <div className="h-0.5 bg-gradient-to-r from-blue-500 via-violet-500 to-purple-500" />
                
                <div className="px-4 py-2.5 border-b border-zinc-700/50 flex items-center justify-between gap-4">
                  <div className="flex items-center gap-2">
                    <div className="w-1.5 h-1.5 rounded-full bg-blue-400 animate-pulse" />
                    <span className="text-xs text-zinc-300 font-medium">待确认修改</span>
                    <span className="text-[10px] text-zinc-500 bg-zinc-800 px-1.5 py-0.5 rounded">
                      ×{pendingChangesTotal}
                    </span>
                  </div>
                  <div className="flex items-center gap-2">
                    <motion.button
                      onClick={rejectAllChanges}
                      className="flex items-center gap-1.5 px-3 py-1.5 bg-zinc-800 hover:bg-red-600/20 text-zinc-400 hover:text-red-400 text-xs rounded-lg transition-all border border-transparent hover:border-red-500/30"
                      whileHover={{ scale: 1.02 }}
                      whileTap={{ scale: 0.98 }}
                    >
                      <X className="w-3.5 h-3.5" />
                      <span>拒绝</span>
                    </motion.button>
                    <motion.button
                      onClick={handleConfirmClick}
                      disabled={isConfirming}
                      className={`flex items-center gap-1.5 px-3 py-1.5 bg-green-600 text-white text-xs rounded-lg transition-all shadow-lg shadow-green-600/20 disabled:opacity-60 disabled:cursor-not-allowed ${isConfirming ? '' : 'hover:bg-green-500'}`}
                      whileHover={isConfirming ? undefined : { scale: 1.02 }}
                      whileTap={isConfirming ? undefined : { scale: 0.98 }}
                    >
                      <Check className="w-3.5 h-3.5" />
                      <span>{isConfirming ? '应用中…' : '接受'}</span>
                    </motion.button>
                    <motion.button
                      onClick={() => setShowRevisionPanel(true)}
                      className="flex items-center gap-1.5 px-3 py-1.5 bg-zinc-800 hover:bg-zinc-700 text-zinc-200 text-xs rounded-lg transition-all border border-zinc-700"
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
                <div className="px-4 py-3 max-h-28 overflow-y-auto bg-zinc-900/50">
                  <div className="flex items-start gap-3 text-xs">
                    <div className="flex-1 min-w-0">
                      <div className="text-[10px] text-zinc-500 mb-1">删除</div>
                      <div className="text-red-400/90 line-through break-all bg-red-500/10 px-2 py-1 rounded border border-red-500/20" title={(pendingChanges[pendingChanges.length - 1]?.meta as any)?.searchText || ''}>
                        {(() => {
                          const searchText = ((pendingChanges[pendingChanges.length - 1]?.meta as any)?.searchText || '') as string
                          return searchText.length > 60 ? searchText.slice(0, 60) + '...' : searchText
                        })()}
                      </div>
                    </div>
                    <div className="text-zinc-600 pt-5">→</div>
                    <div className="flex-1 min-w-0">
                      <div className="text-[10px] text-zinc-500 mb-1">替换为</div>
                      <div className="text-green-400/90 break-all bg-green-500/10 px-2 py-1 rounded border border-green-500/20" title={(pendingChanges[pendingChanges.length - 1]?.meta as any)?.replaceText || ''}>
                        {(() => {
                          const replaceText = ((pendingChanges[pendingChanges.length - 1]?.meta as any)?.replaceText || '') as string
                          return replaceText.length > 60 ? replaceText.slice(0, 60) + '...' : replaceText
                        })()}
                      </div>
                    </div>
                  </div>
                </div>
                
                {/* 快捷键提示 */}
                <div className="px-4 py-1.5 bg-zinc-800/50 border-t border-zinc-700/30 flex items-center justify-center gap-4 text-[10px] text-zinc-500">
                  <span><kbd className="px-1 py-0.5 bg-zinc-700 rounded text-zinc-400">Enter</kbd> 接受</span>
                  <span><kbd className="px-1 py-0.5 bg-zinc-700 rounded text-zinc-400">Esc</kbd> 拒绝</span>
                </div>
              </div>
            </motion.div>
          )}
        </AnimatePresence>
        
        {/* AI 处理中指示器 */}
        {isAIProcessing && (
          <div className="fixed bottom-20 left-1/2 -translate-x-1/2 z-50">
            <div className="flex items-center gap-2 px-4 py-2 bg-zinc-900 border border-zinc-700 rounded-lg shadow-xl">
              <Loader2 className="w-4 h-4 text-violet-400 animate-spin" />
              <span className="text-xs text-zinc-300">AI 正在处理...</span>
            </div>
          </div>
        )}
      </div>
      
      {/* 底部裁剪标记容器 */}
      <div className="crop-marks-bottom pointer-events-none absolute inset-0 z-10"></div>
      
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

      <style>{`
        /* 编辑器容器背景 */
        .word-editor-container {
          background-color: #f3f3f3; /* Word 默认深灰色背景 */
          padding: 20px;
          display: flex;
          flex-direction: column;
          align-items: center;
          justify-content: flex-start;
          min-height: 100%;
        }

        /* 缩放容器 */
        .zoom-container {
          display: flex;
          flex-direction: column;
          align-items: center;
          width: 100%;
        }

        /* 缩放指示器 - 放在编辑器左下角 */
        .zoom-indicator {
          position: absolute;
          bottom: 12px;
          left: 12px;
          display: flex;
          align-items: center;
          gap: 4px;
          background: rgba(255, 255, 255, 0.95);
          border: 1px solid #ddd;
          border-radius: 6px;
          padding: 4px 8px;
          box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
          z-index: 50;
          font-size: 12px;
          user-select: none;
        }

        .zoom-btn {
          width: 24px;
          height: 24px;
          display: flex;
          align-items: center;
          justify-content: center;
          background: #f5f5f5;
          border: 1px solid #ddd;
          border-radius: 4px;
          cursor: pointer;
          font-size: 16px;
          font-weight: bold;
          color: #666;
          transition: all 0.15s;
        }

        .zoom-btn:hover {
          background: #e8e8e8;
          color: #333;
        }

        .zoom-btn:active {
          background: #ddd;
        }

        .zoom-level {
          min-width: 48px;
          text-align: center;
          font-weight: 500;
          color: #333;
          cursor: pointer;
          padding: 2px 4px;
          border-radius: 4px;
          transition: background 0.15s;
        }

        .zoom-level:hover {
          background: #f0f0f0;
        }

        /* 页面包裹容器：支持动态页边距 */
        .word-page {
          --page-margin-top: 2.54cm;
          --page-margin-bottom: 2.54cm;
          --page-margin-left: 3.17cm;
          --page-margin-right: 3.17cm;
          width: 100%;
          max-width: 100%;
          margin: 0;
          background: transparent;
          padding: 0;
          box-shadow: none;
          position: relative;
        }

        /* 横向页面 */
        .word-page.landscape .word-editor-content {
          width: 297mm;
          min-height: 210mm;
        }

        /* 页眉样式 */
        .page-header {
          position: absolute;
          top: 1cm;
          left: var(--page-margin-left);
          right: var(--page-margin-right);
          padding-bottom: 0.5cm;
          border-bottom: 1px solid #ddd;
          font-size: 10pt;
          color: #666;
          z-index: 10;
        }

        /* 页脚样式 */
        .page-footer {
          position: absolute;
          bottom: 1cm;
          left: var(--page-margin-left);
          right: var(--page-margin-right);
          padding-top: 0.5cm;
          border-top: 1px solid #ddd;
          font-size: 10pt;
          color: #666;
          z-index: 10;
        }

        .page-number {
          margin-left: 0.5em;
        }

        /* 裁剪标记 (Crop Marks) - 使用动态边距 */
        .word-page::before,
        .word-page::after {
          content: '';
          position: absolute;
          width: 20px;
          height: 20px;
          border: 1px solid #a0a0a0;
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
          border: 1px solid #a0a0a0;
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
          width: 210mm;
          max-width: 100%;
          margin: 0 auto 24px auto;
          box-sizing: border-box;
          background: #ffffff;
          /* 使用 CSS 变量支持动态页边距 */
          padding: var(--page-margin-top, 2.54cm) var(--page-margin-right, 3.17cm) var(--page-margin-bottom, 2.54cm) var(--page-margin-left, 3.17cm);
          min-height: 297mm; /* 至少一页高度 */
          box-shadow: 0 0 10px rgba(0, 0, 0, 0.10);
          position: relative;
          /* Word 中文默认字体：等线/宋体，英文：Calibri */
          font-family: "等线", "DengXian", "宋体", "SimSun", "Songti SC", "Microsoft YaHei", serif;
          font-size: 10.5pt; /* Word 默认五号字 */
          line-height: 1.15; /* Word 默认单倍行距 */
          color: #000;
          text-align: left; /* Word 默认左对齐 */
        }

        /* 分页占位块：包含（本页剩余白底）+（页间灰色间距）+（下一页顶端留白） */
        .pm-page-break {
          /* 把占位块扩展到整页宽度（包含左右页边距） */
          margin-left: -3.17cm;
          margin-right: -3.17cm;
          width: calc(100% + 6.34cm);
          /* 透明区域=白纸；中间一段=灰色页间距 */
          background: linear-gradient(
            to bottom,
            transparent 0,
            transparent var(--fill, 0px),
            #f3f3f3 var(--fill, 0px),
            #f3f3f3 calc(var(--fill, 0px) + var(--gap, 32px)),
            transparent calc(var(--fill, 0px) + var(--gap, 32px)),
            transparent 100%
          );
          position: relative;
        }

        /* 在页间距处加一点阴影/分隔感，更接近原版 */
        .pm-page-break::before {
          content: '';
          position: absolute;
          left: 0;
          right: 0;
          top: var(--fill, 0px);
          height: calc(var(--gap, 32px));
          box-shadow: inset 0 12px 16px rgba(0, 0, 0, 0.06), inset 0 -12px 16px rgba(0, 0, 0, 0.05);
          pointer-events: none;
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
          border-left: 3px solid #ccc;
          padding-left: 1em;
          margin: 0 0 0 2em;
          color: #333;
        }

        .word-editor-content blockquote p {
          text-indent: 0 !important;
        }

        /* 分隔线 */
        .word-editor-content hr {
          border: none;
          border-top: 1px solid #000;
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
          table-layout: auto;
        }

        .word-editor-content th,
        .word-editor-content td {
          border: 1px solid #000;
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
          color: #0563c1;
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
          outline: 2px solid #0066cc;
        }

        /* 占位符 */
        .word-editor-content p.is-editor-empty:first-child::before {
          color: #aaa;
          content: attr(data-placeholder);
          float: left;
          height: 0;
          pointer-events: none;
        }

        /* 表格选中 */
        .word-editor-content .selectedCell:after {
          background: rgba(0, 102, 204, 0.1);
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
          background-color: #0066cc;
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
