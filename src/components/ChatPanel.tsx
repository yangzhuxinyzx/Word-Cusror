import { useState, useRef, useEffect, useCallback } from 'react'
import { 
  Send, 
  Trash2, 
  FileText,
  X,
  Paperclip,
  CheckCircle,
  FileEdit,
  FilePlus,
  Eye,
  Loader2,
  CheckCircle2,
  Circle,
  Bot,
  Table
} from 'lucide-react'
import ReactMarkdown from 'react-markdown'
import { motion, AnimatePresence } from 'framer-motion'
import { useAI, ToolResult } from '../context/AIContext'
import { useDocument } from '../context/DocumentContext'
import { FileItem, AgentStep, AgentFileChange } from '../types'
import { runWebSearch, WebSearchResponse } from '../utils/webSearch'
import CinematicTyper from './CinematicTyper'

type PptOutlineSlideDraft = {
  pageNumber: number
  pageType?: string
  headline: string
  subheadline?: string
  bullets?: string[]
  footerNote?: string
  layoutIntent?: string
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
  status: 'running' | 'success' | 'error'
  detail?: string
}

const truncateLabel = (text: string, limit = 32) => {
  if (!text) return ''
  return text.length > limit ? `${text.slice(0, limit)}…` : text
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
  const { messages, isLoading, streamingContent, settings, addMessage, sendAgentMessage, clearMessages } = useAI()
  const { 
    document, 
    createNewDocument, 
    isElectron, 
    currentFile, 
    replaceInDocument, 
    insertInDocument, 
    deleteInDocument, 
    openFile, 
    files, 
    workspacePath,
    editorMode,
    setEditorMode,
    refreshFiles,
    getTiptapDocumentStructure,
    replaceWithFormat,
    excelData,
    refreshExcelData,
    previewWordOps,
    applyWordOps,
    getLatestContent
  } = useDocument()
  const [input, setInput] = useState('')
  const [attachedFiles, setAttachedFiles] = useState<FileItem[]>([])
  const [isDragOver, setIsDragOver] = useState(false)
  const messagesEndRef = useRef<HTMLDivElement>(null)
  const inputRef = useRef<HTMLTextAreaElement>(null)
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
      const filePath = `${workspacePath}\\${fileName}`
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

  const resetToolActivity = useCallback(() => {
    setToolActivity([])
  }, [])

  const registerToolActivity = useCallback((tool: string, label: string) => {
    const id = `${tool}-${Date.now()}-${Math.random().toString(16).slice(2)}`
    setToolActivity(prev => [...prev, { id, tool, label, status: 'running' }])
    return id
  }, [])

  const completeToolActivity = useCallback((id: string, status: 'success' | 'error', detail?: string) => {
    setToolActivity(prev =>
      prev.map(item =>
        item.id === id ? { ...item, status, detail: detail ?? item.detail } : item
      )
    )
  }, [])

  // 更新思考时间
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
    resetToolActivity()
  }, [resetToolActivity])

  // ========== 直接执行 PPT 生成（确认按钮用） ==========
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

      if (!isElectron || !window.electronAPI?.pptGenerateDeck) {
        throw new Error('PPT 生成仅支持桌面版（Electron）')
      }

      // 输出路径
      const dir = currentFile?.path
        ? currentFile.path.substring(0, currentFile.path.lastIndexOf('\\'))
        : (workspacePath || null)

      if (!dir) {
        throw new Error('缺少工作区路径，请先打开一个文件夹')
      }

      const safeTitle = String(title).replace(/[<>:"/\\|?*]/g, '_').slice(0, 60) || '新建演示文稿'
      const pptxName = safeTitle.toLowerCase().endsWith('.pptx') ? safeTitle : `${safeTitle}.pptx`
      const outputPath = `${dir}\\${pptxName}`

      // 获取 API Keys
      const openRouterApiKey = settings?.openRouterApiKey || ''
      // 优先使用专门的 DashScope API Key，否则回退到主模型 API Key
      const dashscopeApiKey = settings?.dashscopeApiKey || settings?.apiKey || ''

      // 如果没有 DashScope API Key，提示用户配置
      if (!dashscopeApiKey) {
        throw new Error('缺少 DashScope API Key。请在设置中配置阿里云百炼 API Key')
      }

      const estimatedSlideCount = draft.slides?.length || 3

      // ========== 阶段1：调用 Gemini 生成文生图提示词 ==========
      updateAgentAction(`正在让 Gemini 设计视觉风格...`)
      addAgentFileOperation(`PPT: 正在设计 ${estimatedSlideCount} 页视觉`)

      const geminiResult = await window.electronAPI.openrouterGeminiPptPrompts({
        apiKey: openRouterApiKey,
        outline,
        slideCount: estimatedSlideCount,
        theme,
        style: draft.styleHint || '',
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

      // ========== 阶段2：调用 DashScope 生成图片 ==========
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
      
      // 根据用户选择的模型决定分辨率（默认使用 Gemini 生图）
      const pptImageModel = settings?.pptImageModel || 'gemini-image'
      const imageSize = pptImageModel === 'z-image-turbo' ? '2048*1152' : '1664*928'
      console.log(`[PPT] 使用生图模型: ${pptImageModel}`)

      const result = await window.electronAPI.pptGenerateDeck({
        outputPath,
        slides: slidesWithContent,
        // 主模型 API Key（用于 Gemini 生图）
        mainApiKey: settings?.apiKey || '',
        dashscope: {
          apiKey: dashscopeApiKey,
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

      await refreshFiles()

      // 打开新生成的 PPT
      await openFile({ name: pptxName, path: result.path, type: 'file' as const })

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
  }, [pptGenerating, isElectron, currentFile, workspacePath, settings, addMessage, registerToolActivity, completeToolActivity, updateAgentAction, completeAgentStep, updateAgentFile, addAgentFileOperation, finishAgentProgress, refreshFiles, openFile])

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
    if (pptGenerating || !isElectron) return
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
      if (!dashscopeApiKey) {
        throw new Error('请先在 AI 设置中配置 DashScope API Key（阿里云百炼）')
      }

      updateAgentAction(`正在${modeLabel}：${pagesLabel}...`)
      addAgentFileOperation(`PPT: ${modeLabel} ${pagesLabel}`)

      const result = await window.electronAPI!.pptEditSlides({
        pptxPath,
        pageNumbers,
        feedback,
        mode,
        openRouterApiKey,
        dashscopeApiKey,
        mainApiKey: settings.apiKey || '',
        pptImageModel: settings.pptImageModel || 'gemini-image',
      })

      if (!result.success) {
        throw new Error(result.error || '编辑失败')
      }

      await refreshFiles()

      // 重新打开 PPT 以刷新预览，并跳转到被编辑的页面
      const pptxName = pptxPath.split(/[\\/]/).pop() || 'output.pptx'
      const firstEditedPage = (result.editedPages && result.editedPages.length > 0) ? result.editedPages[0] : pageNumbers[0]
      
      // 触发自定义事件通知 PptPreviewHtml 跳转到指定页
      window.dispatchEvent(new CustomEvent('ppt-jump-to-page', {
        detail: { pageNumber: firstEditedPage }
      }))
      
      await openFile({ name: pptxName, path: result.path || pptxPath, type: 'file' as const })

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
  }, [pptGenerating, isElectron, settings, addMessage, registerToolActivity, completeToolActivity, updateAgentAction, addAgentFileOperation, finishAgentProgress, refreshFiles, openFile])

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
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' })
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
    const createKeywords = ['创建', '新建', '生成', '写一份', '帮我写', '起草', '总结文档', '做一个总结', '做个总结', '写总结', '生成总结', '/会议纪要']
    // 编辑类关键词 - 包含快捷命令
    const editKeywords = [
      '修改', '编辑', '润色', '优化', '改成', '替换', '删除', '添加', '扩展', '精简', '翻译', '重写',
      '格式化', '统一格式', '编号', '标题编号', '公文格式', '转换为公文',
      '/润色', '/精简', '/翻译', '/格式化', '/编号', '/公文', '/总结'
    ]
    const analyzeKeywords = ['分析', '解释', '什么意思', '有哪些', '告诉我', '是什么', '检查', '论文检查']
    
    // 优先检测创建类（包括总结文档）
    if (createKeywords.some(k => text.includes(k))) return 'create'
    if (editKeywords.some(k => text.includes(k))) return 'edit'
    if (analyzeKeywords.some(k => text.includes(k))) return 'analyze'
    return 'chat'
  }

  // 获取文件内容
  const getFileContent = useCallback(async (file: FileItem): Promise<string> => {
    if (isElectron && window.electronAPI) {
      const result = await window.electronAPI.readFile(file.path)
      if (result.success && result.data) {
        return result.type === 'docx' ? `[Word文档: ${file.name}]` : result.data
      }
    }
    return file.content || `[文件: ${file.name}]`
  }, [isElectron])

  // 构建文件上下文
  const buildFilesContext = useCallback(async () => {
    if (attachedFiles.length === 0) return ''
    const contents: string[] = []
    for (const file of attachedFiles) {
      const content = await getFileContent(file)
      contents.push(`=== ${file.name} ===\n${content}`)
    }
    return contents.join('\n\n')
  }, [attachedFiles, getFileContent])

  const handleSend = useCallback(async () => {
    if (!input.trim() || isLoading) return

    const userMessage = input.trim()
    setInput('')
    resetToolActivity()
    
    // 保存 PPT 编辑上下文（如果有）并清除状态
    const currentPptEditContext = pptEditContext
    if (pptEditContext) {
      setPptEditContext(null)
    }
    
    const operation = detectOperation(userMessage)
    const fileNames = attachedFiles.map(f => f.name).join(', ')
    
    // 构建用户消息内容（包含 PPT 编辑上下文标记）
    let displayMessage = userMessage
    if (currentPptEditContext) {
      displayMessage = `🖼️ [第 ${currentPptEditContext.pageNumber} 页${currentPptEditContext.isRegion ? '（框选区域）' : ''}] ${userMessage}`
    } else if (attachedFiles.length > 0) {
      displayMessage = `${userMessage}\n📎 ${fileNames}`
    }
    
    // 添加用户消息
    addMessage({ 
      role: 'user', 
      content: displayMessage
    })

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
    let fullContext = attachedContext || ''
    
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

❌ 不要使用 replace/delete/insert 这些 Word 文档工具！`
        
        fullContext = fullContext ? `${excelContext}\n\n${fullContext}` : excelContext
      } else {
        // Word/文本文档处理
        let docContent = document.content
        let docStructure = ''
        
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
        
        if (docContent) {
          // 发送内容，让 AI 能看到原文档
          // AI 编辑始终使用内置编辑器，ONLYOFFICE 仅用于预览
          const formatNote = '\n\n[提示：AI 编辑使用内置编辑器。支持 HTML 格式，可使用 <h1>/<h2>/<strong>/<em> 等标签。' +
            (editorMode === 'onlyoffice' ? ' 当前预览模式为 ONLYOFFICE。]' : ']')
          const currentFileContext = `=== ${currentFile.name} (当前编辑) ===\n${docContent}${docStructure}${formatNote}`
          fullContext = fullContext ? `${currentFileContext}\n\n${fullContext}` : currentFileContext
        }
      }
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
🚫 禁止：replace / insert / delete / create / create_from_template（这些是 Word/Excel 工具，会导致错误操作）
根据用户的描述判断：
- 如果用户对整体不满意（太丑、换风格、重做等），使用 mode="regenerate"
- 如果用户只想修改局部细节（改颜色、换文字、调整位置等），使用 mode="partial_edit"
`
      fullContext = fullContext ? `${pptEditInfo}\n\n${fullContext}` : pptEditInfo
    }

    // 使用 Agent 模式发送消息
    await sendAgentMessage(
      userMessage,
      document.content,
      fullContext || undefined,
      {
        // 工具调用处理
        onToolCall: async (tool, args): Promise<ToolResult> => {
          if (tool === 'replace') {
            const search = args.search || ''
            const replaceText = args.replace || ''
            
            if (!search) {
              return { tool, success: false, message: '缺少 search 参数' }
            }

            const activityId = registerToolActivity('replace', `替换：${truncateLabel(search, 24)}`)

            // 如果当前是 ONLYOFFICE 模式，自动切换到内置编辑器以显示 diff 标记
            if (editorMode === 'onlyoffice') {
              setEditorMode('tiptap')
              // 等待编辑器切换完成
              await new Promise(resolve => setTimeout(resolve, 100))
            }

            // 解析格式化参数
            const format = {
              bold: args.bold === 'true',
              italic: args.italic === 'true',
              underline: args.underline === 'true',
              color: args.color || undefined,
              backgroundColor: args.backgroundColor || undefined,
              fontSize: args.fontSize || undefined
            }
            const hasFormat = format.bold || format.italic || format.underline || 
                             format.color || format.backgroundColor || format.fontSize

            // 更新 Agent 进度 - 显示正在执行替换
            const formatInfo = hasFormat ? ' (带格式)' : ''
            updateAgentAction(`正在替换「${search.slice(0, 20)}${search.length > 20 ? '...' : ''}」${formatInfo}`)
            completeAgentStep()
            updateAgentFile({ status: 'writing', name: fileName })
            addAgentFileOperation(`替换: "${search.slice(0, 15)}..." → "${replaceText.slice(0, 15)}..."`)

            // AI 编辑始终使用内置编辑器（Tiptap）的方法
            // ONLYOFFICE 仅用于预览，不参与 AI 编辑
            let result
            if (hasFormat) {
              result = replaceWithFormat(search, replaceText, format)
            } else {
              result = replaceInDocument(search, replaceText)
            }
            
            if (result.success && result.count > 0) {
              totalReplacements += result.count
              updateAgentFile({ additions: result.count, status: 'writing', name: fileName })
              completeToolActivity(activityId, 'success', `${result.count} 处`)
              return { 
                tool, 
                success: true, 
                message: `成功替换 ${result.count} 处：「${search}」→「${replaceText}」`,
                data: { 
                  count: result.count,
                  searchText: search,
                  replaceText: replaceText,
                  positions: result.positions
                }
              }
            } else {
              completeToolActivity(activityId, 'error', '未找到匹配')
              return { 
                tool, 
                success: false, 
                message: `未找到「${search}」，请检查是否与文档内容完全匹配` 
              }
            }
          }

          if (tool === 'word_edit_ops') {
            // 统一格式/样式/字符格式的结构化操作：支持 dryRun 预览 → 用户确认 → 应用修订
            const rawOps = args.ops || ''
            const dryRunTop = (args.dryRun || '').toLowerCase() === 'true'

            let ops: any[] = []
            if (rawOps) {
              try {
                ops = JSON.parse(rawOps)
              } catch (e) {
                return { tool, success: false, message: 'ops 解析失败：不是合法 JSON 数组' }
              }
            }

            if (!Array.isArray(ops) || ops.length === 0) {
              return { tool, success: false, message: '缺少 ops 或 ops 为空（必须是 JSON 数组）' }
            }

            const inferredDryRun = ops.some((op) => op?.dryRun === true)
            const isDryRun = dryRunTop || inferredDryRun

            if (isDryRun) {
              const preview = previewWordOps(ops)
              const lines = (preview.data?.lines as string[] | undefined) || []
              setPendingWordOps({
                ops,
                previewMessage: preview.message,
                previewLines: lines,
              })
              return {
                tool,
                success: preview.success,
                message: preview.success
                  ? `${preview.message}\n${lines.length ? '- ' + lines.join('\n- ') : ''}\n\n请在下方点击「应用修订」以执行。`
                  : preview.message,
                data: preview.data,
              }
            }

            const result = applyWordOps(ops)
            return {
              tool,
              success: result.success,
              message: result.message,
              data: result.data,
            }
          }
          
          if (tool === 'create') {
            const title = args.title || '新文档'
            const content = args.content || ''
            const activityId = registerToolActivity('create', `创建：${truncateLabel(title, 24)}`)
            
            // 检查是否有 elements 参数（带格式创建）
            let elements: Array<{
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
            }> = []
            
            if (args.elements) {
              try {
                elements = JSON.parse(args.elements)
              } catch (e) {
                console.error('解析 elements 失败:', e)
                // 继续使用 content 方式
              }
            }

            // 更新 Agent 进度
            updateAgentAction(`正在创建「${title}.docx」`)
            completeAgentStep()
            updateAgentFile({ status: 'writing', name: `${title}.docx` })

            try {
              console.log('create 工具参数:', { title, content: content.slice(0, 100), elements, rawArgs: args })
              
              // 如果有 elements，使用带格式创建（直接用 docx 库生成文件）
              if (elements.length > 0) {
                console.log('使用 elements 创建带格式文档:', elements)
                await createNewDocument(title, content, elements)
                 completeToolActivity(activityId, 'success', `${elements.length} 段`)
                finishAgentProgress()
                return {
                  tool,
                  success: true,
                  message: `已创建文档：${title}.docx（包含 ${elements.length} 个格式化元素）`,
                  data: { fileName: `${title}.docx`, elementCount: elements.length }
                }
              }
              
              // 普通方式创建（纯文本内容）
              console.log('使用纯文本创建文档')
              await createNewDocument(title, content)
              const lineCount = content.split('\n').length
              completeToolActivity(activityId, 'success', `${lineCount} 行`)
              finishAgentProgress()
              
              return { 
                tool, 
                success: true, 
                message: `已创建文档：${title}.docx`,
                data: { fileName: `${title}.docx`, lines: lineCount }
              }
            } catch (e) {
              console.error('创建文档失败:', e)
              completeToolActivity(activityId, 'error', '创建失败')
              return { tool, success: false, message: `创建失败: ${e}` }
            }
          }

          if (tool === 'ppt_create') {
            const title = args.title || '新建演示文稿'
            const theme = args.theme || ''
            const style = args.style || ''
            const outline = args.outline || ''
            const activityId = registerToolActivity('ppt_create', `PPT：${truncateLabel(title, 24)}`)

            if (!isElectron || !window.electronAPI?.pptGenerateDeck) {
              completeToolActivity(activityId, 'error', '不支持')
              return { tool, success: false, message: 'PPT 生成仅支持桌面版（Electron）' }
            }

            if (!outline || outline.trim().length < 10) {
              completeToolActivity(activityId, 'error', '缺少大纲')
              return { tool, success: false, message: '缺少 outline 参数（需要 PPT 大纲内容）' }
            }

            // 输出路径：优先当前文件目录，其次工作区根目录
            const dir = currentFile?.path
              ? currentFile.path.substring(0, currentFile.path.lastIndexOf('\\'))
              : (workspacePath || null)

            if (!dir) {
              completeToolActivity(activityId, 'error', '缺少工作区')
              return { tool, success: false, message: '缺少工作区路径，请先打开一个文件夹' }
            }

            const safeTitle = String(title).replace(/[<>:"/\\|?*]/g, '_').slice(0, 60) || '新建演示文稿'
            const pptxName = safeTitle.toLowerCase().endsWith('.pptx') ? safeTitle : `${safeTitle}.pptx`
            const outputPath = `${dir}\\${pptxName}`

            // 获取 API Keys
            const openRouterApiKey = settings?.openRouterApiKey || ''
            // 优先使用专门的 DashScope API Key
            const dashscopeApiKey = settings?.dashscopeApiKey || settings?.apiKey || ''

            // 计算大概的页数
            const slideCountMatch = outline.match(/第\s*(\d+)\s*页/g)
            const estimatedSlideCount = slideCountMatch ? slideCountMatch.length : 3

            try {
              // ========== 阶段1：调用 Gemini 生成文生图提示词 ==========
              updateAgentAction(`正在让 Gemini 设计视觉风格...`)
              completeAgentStep()
              updateAgentFile({ status: 'writing', name: pptxName })
              addAgentFileOperation(`PPT: 正在设计 ${estimatedSlideCount} 页视觉`)

              let slides: Array<{ prompt: string; negativePrompt?: string }> = []
              let deckDesignConcept = ''
              let deckColorPalette = ''

              if (window.electronAPI?.openrouterGeminiPptPrompts) {
                const geminiResult = await window.electronAPI.openrouterGeminiPptPrompts({
                  apiKey: openRouterApiKey,
                  outline,
                  slideCount: estimatedSlideCount,
                  theme,
                  style,
                  // 主模型回退参数（当没有 OpenRouter API Key 时使用）
                  mainApiKey: settings?.apiKey || '',
                  mainBaseUrl: settings?.baseUrl || '',
                  mainModel: settings?.model || '',
                })

                if (!geminiResult.success || !geminiResult.slides) {
                  completeToolActivity(activityId, 'error', '设计生成失败')
                  return { tool, success: false, message: `设计提示词生成失败: ${geminiResult.error || '未知错误'}` }
                }

                deckDesignConcept = geminiResult.designConcept || ''
                deckColorPalette = geminiResult.colorPalette || ''

                slides = geminiResult.slides.map((s) => ({
                  prompt: s.prompt,
                  negativePrompt: s.negativePrompt,
                }))

                updateAgentAction(`设计完成，共 ${slides.length} 页，开始生成图片...`)
              } else {
                completeToolActivity(activityId, 'error', '缺少 API')
                return { 
                  tool, 
                  success: false, 
                  message: '缺少可用的 API。请在设置中配置 OpenRouter API Key 或主模型 API Key。' 
                }
              }

              // ========== 阶段2：调用 DashScope 生成图片 ==========
              updateAgentAction(`正在生成「${pptxName}」(${slides.length} 页，两张两张生图)...`)
              addAgentFileOperation(`PPT: 生成 ${slides.length} 页图片`)

              // 注意：负面词用于“去廉价/去AI味”，避免过强霓虹、塑料感、模板化等距城市
              const negativeDefault =
                'watermark, logo, brand name text, badge, QR code, UI, screenshot, HUD, sci-fi interface, holographic UI, futuristic dashboard, neon cyberpunk, neon cyan, bright cyan, fluorescent cyan, neon teal, cheap turquoise, generic isometric city, isometric cityscape, circuit-board city, lowres, blurry, garbled Chinese, wrong characters, text distortion, misspelling, random letters, gibberish, extra text, english text, ugly typography, amateur layout, noisy background, oversaturated, cheap plastic, toy-like, glossy, harsh specular, overbloom, stock 3d icons, generic template, ai artifacts, uncanny'

              // 根据用户选择的模型决定分辨率（默认使用 Gemini 生图）
              const pptImageModel = settings?.pptImageModel || 'gemini-image'
              const imageSize = pptImageModel === 'z-image-turbo' ? '2048*1152' : '1664*928'
              console.log(`[PPT Tool] 使用生图模型: ${pptImageModel}`)

              const result = await window.electronAPI.pptGenerateDeck({
                outputPath,
                slides: slides.map((s) => ({
                  prompt: s.prompt,
                  negativePrompt: s.negativePrompt || negativeDefault,
                })),
                // 主模型 API Key（用于 Gemini 生图）
                mainApiKey: settings?.apiKey || '',
                dashscope: {
                  apiKey: dashscopeApiKey,
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
                    designConcept: deckDesignConcept,
                    colorPalette: deckColorPalette,
                  },
                },
              })

              if (!result.success || !result.path) {
                completeToolActivity(activityId, 'error', result.error || '失败')
                return { tool, success: false, message: `PPT 生成失败: ${result.error || '未知错误'}` }
              }

              await refreshFiles()

              // 打开新生成的 PPT
              await openFile({ name: pptxName, path: result.path, type: 'file' as const })

              updateAgentFile({ additions: slides.length, status: 'done', name: pptxName })
              finishAgentProgress()
              completeToolActivity(activityId, 'success', `${slides.length} 页`)

              return {
                tool,
                success: true,
                message: `已生成 PPT：${pptxName}（${slides.length} 页，由 Gemini 设计 + DashScope 生图，已导出到工作区）`,
                data: { fileName: pptxName, path: result.path, slideCount: slides.length },
              }
            } catch (e) {
              console.error('PPT 生成失败:', e)
              completeToolActivity(activityId, 'error', '异常')
              return { tool, success: false, message: `PPT 生成失败: ${e}` }
            }
          }
          
          // PPT 编辑工具（拖拽/框选触发）
          if (tool === 'ppt_edit') {
            const pageNumber = Number(args.pageNumber) || 1
            const mode = args.mode === 'partial_edit' ? 'partial_edit' : 'regenerate'
            const feedback = args.feedback || ''
            const pptxPath = args.pptxPath || currentPptEditContext?.pptxPath || ''
            
            // 注意：Agent 参数解析默认都是 string，这里做一次安全解析
            let regionRect: { x: number; y: number; w: number; h: number } | undefined = currentPptEditContext?.regionRect
            if (typeof args.regionRect === 'string' && args.regionRect.trim()) {
              try {
                regionRect = JSON.parse(args.regionRect)
              } catch {
                // ignore
              }
            }
            const regionScreenshot =
              (typeof args.regionScreenshot === 'string' && args.regionScreenshot.trim())
                ? args.regionScreenshot
                : currentPptEditContext?.imageBase64
            
            const modeLabel = mode === 'regenerate' ? '整页重做' : '局部编辑'
            const activityId = registerToolActivity('ppt_edit', `PPT ${modeLabel}：第 ${pageNumber} 页`)
            
            if (!isElectron || !window.electronAPI?.pptEditSlides) {
              completeToolActivity(activityId, 'error', '不支持')
              return { tool, success: false, message: 'PPT 编辑仅支持桌面版（Electron）' }
            }
            
            if (!pptxPath) {
              completeToolActivity(activityId, 'error', '缺少路径')
              return { tool, success: false, message: '缺少 PPTX 文件路径' }
            }
            
            try {
              updateAgentAction(`正在${modeLabel}第 ${pageNumber} 页...`)
              
              const openRouterApiKey = settings?.openRouterApiKey || ''
              // 优先使用专门的 DashScope API Key
              const dashscopeApiKey = settings?.dashscopeApiKey || settings?.apiKey || ''
              
              const result = await window.electronAPI.pptEditSlides({
                pptxPath,
                pageNumbers: [pageNumber],
                mode,
                feedback,
                regionScreenshot,
                regionRect,
                openRouterApiKey,
                dashscopeApiKey,
                mainApiKey: settings?.apiKey || '',
                pptImageModel: settings?.pptImageModel || 'gemini-image',
              })
              
              if (!result.success) {
                completeToolActivity(activityId, 'error', result.error || '失败')
                return { tool, success: false, message: `PPT 编辑失败: ${result.error || '未知错误'}` }
              }
              
              // 刷新文件并跳转到编辑的页面
              await refreshFiles()
              
              // 触发跳转事件
              window.dispatchEvent(new CustomEvent('ppt-jump-to-page', {
                detail: { pageNumber }
              }))
              
              // 重新打开文件以刷新预览
              if (currentFile?.path === pptxPath) {
                await openFile({ name: currentFile.name, path: pptxPath, type: 'file' as const })
              }
              
              completeToolActivity(activityId, 'success', modeLabel)
              
              return {
                tool,
                success: true,
                message: `已完成第 ${pageNumber} 页的${modeLabel}`,
                data: { pageNumber, mode, fileName: (pptxPath.split(/[\\/]/).pop() || ''), pptxPath },
              }
            } catch (e) {
              console.error('PPT 编辑失败:', e)
              completeToolActivity(activityId, 'error', '异常')
              return { tool, success: false, message: `PPT 编辑失败: ${e}` }
            }
          }
          
          if (tool === 'insert') {
            const position = args.position || 'end'
            const content = args.content || ''
            
            if (!content) {
              return { tool, success: false, message: '缺少 content 参数' }
            }

            const activityId = registerToolActivity('insert', `插入：${position}`)

            // 如果当前是 ONLYOFFICE 模式，自动切换到内置编辑器
            if (editorMode === 'onlyoffice') {
              setEditorMode('tiptap')
              await new Promise(resolve => setTimeout(resolve, 100))
            }
            
            // 更新 Agent 进度
            updateAgentAction(`正在插入内容到 ${position === 'start' ? '开头' : position === 'end' ? '末尾' : position}`)
            completeAgentStep()
            addAgentFileOperation(`插入: ${content.slice(0, 30)}...`)
            
            // AI 编辑始终使用内置编辑器（Tiptap）的方法
            const result = insertInDocument(position, content)
            
            if (result.success) {
              updateAgentFile({ additions: 1, status: 'writing', name: fileName })
              completeToolActivity(activityId, 'success')
              return { 
                tool, 
                success: true, 
                message: result.message,
                data: { position, contentLength: content.length }
              }
            } else {
              completeToolActivity(activityId, 'error', result.message)
              return { tool, success: false, message: result.message }
            }
          }
          
          if (tool === 'delete') {
            const target = args.target || ''
            
            if (!target) {
              return { tool, success: false, message: '缺少 target 参数' }
            }

            const activityId = registerToolActivity('delete', `删除：${truncateLabel(target, 24)}`)

            // 如果当前是 ONLYOFFICE 模式，自动切换到内置编辑器
            if (editorMode === 'onlyoffice') {
              setEditorMode('tiptap')
              await new Promise(resolve => setTimeout(resolve, 100))
            }
            
            // 更新 Agent 进度
            updateAgentAction(`正在删除「${target.slice(0, 20)}${target.length > 20 ? '...' : ''}」`)
            completeAgentStep()
            addAgentFileOperation(`删除: "${target.slice(0, 30)}..."`)
            
            const result = deleteInDocument(target)
            
            if (result.success) {
              updateAgentFile({ deletions: result.count, status: 'writing', name: fileName })
              completeToolActivity(activityId, 'success', `${result.count} 处`)
              return { 
                tool, 
                success: true, 
                message: result.message,
                data: { count: result.count, target }
              }
            } else {
              completeToolActivity(activityId, 'error', result.message)
              return { tool, success: false, message: result.message }
            }
          }

          // 复制模板并自动替换内容
          // 方案：先复制文件，再用 ONLYOFFICE 在编辑器中执行替换
          if (tool === 'copy_template' || tool === 'create_from_template') {
            const newTitle = args.newTitle || '新文档'
            let replacements: Array<{search: string, replace: string}> = []
            const activityId = registerToolActivity(tool, `模板：${truncateLabel(newTitle, 24)}`)
            
            if (args.replacements) {
              try {
                replacements = JSON.parse(args.replacements)
              } catch (e) {
                console.error('解析替换数据失败:', e)
              }
            }

            if (!currentFile) {
              completeToolActivity(activityId, 'error', '缺少模板')
              return { tool, success: false, message: '没有打开的文档作为模板' }
            }

            updateAgentAction(`正在基于模板创建「${newTitle}.docx」`)
            completeAgentStep()

            try {
              if (isElectron && window.electronAPI) {
                const sourcePath = currentFile.path
                const dir = sourcePath.substring(0, sourcePath.lastIndexOf('\\'))
                const newPath = `${dir}\\${newTitle}.docx`
                
                // 第一步：复制文件
                updateAgentAction(`正在复制模板...`)
                const sourceContent = await window.electronAPI.readFile(sourcePath)
                if (!sourceContent.success) {
                  return { tool, success: false, message: '读取模板文件失败' }
                }
                
                if (sourceContent.type === 'docx') {
                  await window.electronAPI.writeBinaryFile(newPath, sourceContent.data!)
                } else {
                  await window.electronAPI.writeFile(newPath, sourceContent.data!)
                }
                
                // 刷新文件列表
                await refreshFiles()
                
                // 第二步：打开新文件
                updateAgentAction(`正在打开新文档...`)
                const newFile = { name: `${newTitle}.docx`, path: newPath, type: 'file' as const }
                await openFile(newFile)
                
                // 第三步：等待 ONLYOFFICE 加载完成并执行替换
                if (replacements.length > 0) {
                  updateAgentAction(`等待编辑器加载...`)
                  
                  // 等待 connector 就绪
                  let connectorReady = false
                  for (let retry = 0; retry < 40; retry++) {
                    await new Promise(resolve => setTimeout(resolve, 500))
                    
                    if (window.onlyOfficeConnector?.searchAndReplace) {
                      try {
                        const testText = await window.onlyOfficeConnector.getDocumentText()
                        if (testText && testText.length > 10) {
                          connectorReady = true
                          console.log('✓ ONLYOFFICE connector 已就绪')
                          break
                        }
                      } catch (e) {
                        console.log('等待 connector...', retry)
                      }
                    }
                  }
                  
                  if (!connectorReady) {
                    updateAgentFile({ additions: 0, status: 'done', name: `${newTitle}.docx` })
                    finishAgentProgress()
                    completeToolActivity(activityId, 'success', '已创建')
                    return { 
                      tool, 
                      success: true, 
                      message: `已创建「${newTitle}.docx」，但编辑器未就绪，请手动替换内容`
                    }
                  }
                  
                  // 执行替换
                  await new Promise(resolve => setTimeout(resolve, 1000))
                  
                  let successCount = 0
                  updateAgentAction(`正在替换内容 (0/${replacements.length})...`)
                  
                  for (let i = 0; i < replacements.length; i++) {
                    const item = replacements[i]
                    updateAgentAction(`替换 (${i+1}/${replacements.length}): ${item.search.slice(0, 20)}...`)
                    
                    try {
                      console.log(`尝试替换: "${item.search}" -> "${item.replace}"`)
                      const result = await window.onlyOfficeConnector!.searchAndReplace(item.search, item.replace, true)
                      if (result) {
                        successCount++
                        console.log(`✓ 替换成功`)
                      } else {
                        console.log(`✗ 未找到匹配`)
                      }
                    } catch (e) {
                      console.error(`替换失败:`, e)
                    }
                    
                    await new Promise(resolve => setTimeout(resolve, 300))
                  }
                  
                  updateAgentFile({ additions: successCount, status: 'done', name: `${newTitle}.docx` })
                  finishAgentProgress()
                  completeToolActivity(activityId, 'success', `${successCount}/${replacements.length}`)
                  
                  const resultMsg = successCount > 0
                    ? `已创建「${newTitle}.docx」，成功替换 ${successCount}/${replacements.length} 处内容`
                    : `已创建「${newTitle}.docx」，但替换未成功（可能是搜索文字不精确）`
                  
                  return { 
                    tool, 
                    success: true, 
                    message: resultMsg,
                    data: { 
                      fileName: `${newTitle}.docx`,
                      totalReplacements: replacements.length,
                      successfulReplacements: successCount
                    }
                  }
                } else {
                  updateAgentFile({ additions: 0, status: 'done', name: `${newTitle}.docx` })
                  finishAgentProgress()
                  completeToolActivity(activityId, 'success')
                  
                  return { 
                    tool, 
                    success: true, 
                    message: `已复制创建「${newTitle}.docx」`
                  }
                }
              } else {
                completeToolActivity(activityId, 'error', '仅支持桌面')
                return { tool, success: false, message: '此功能需要在桌面应用中使用' }
              }
            } catch (e) {
              console.error('复制模板失败:', e)
              completeToolActivity(activityId, 'error', '复制失败')
              return { tool, success: false, message: `复制模板失败: ${e}` }
            }
          }

          if (tool === 'web_search') {
            const query = (args.query || args.q || args.keyword || '').trim()
            if (!query) {
              return { tool, success: false, message: '缺少 query 参数' }
            }
            const locale = args.hl || args.locale || 'zh-CN'
            const region = args.gl || args.region || 'cn'
            const num = args.num ? parseInt(args.num, 10) || 5 : 5

            const activityId = registerToolActivity('web_search', `搜索：${truncateLabel(query, 28)}`)
            updateAgentAction(`正在检索外部信息：${truncateLabel(query, 28)}`)

            const searchResponse = await runWebSearch(query, { locale, region, num, braveApiKey: settings.braveApiKey })

            const webResults = searchResponse.results ?? []
            if (!searchResponse.success || webResults.length === 0) {
              completeToolActivity(activityId, 'error', searchResponse.message || '0 条结果')
              return { 
                tool, 
                success: false, 
                message: searchResponse.message || '未获取到搜索结果，请稍后重试' 
              }
            }

            const extraTotal = (searchResponse.sections?.faq?.length ?? 0)
              + (searchResponse.sections?.news?.length ?? 0)
              + (searchResponse.sections?.videos?.length ?? 0)
              + (searchResponse.sections?.discussions?.length ?? 0)
            const summaryLabel = `${webResults.length}${extraTotal ? `+${extraTotal}` : ''} 条`

            completeToolActivity(activityId, 'success', summaryLabel)
            const formatted = formatSearchResults(searchResponse, query)

            return {
              tool,
              success: true,
              message: formatted,
              data: {
                query,
                locale,
                region,
                results: webResults,
                sections: searchResponse.sections,
                summarizerKey: searchResponse.summarizerKey
              }
            }
          }

          // ==================== Excel 工具处理 ====================
          
          // 检查是否有打开的 Excel 文件
          const isExcelFile = currentFile?.name?.toLowerCase().endsWith('.xlsx') || currentFile?.name?.toLowerCase().endsWith('.xls')
          const excelFilePath = currentFile?.path
          
          if (tool === 'excel_read') {
            if (!isExcelFile || !excelFilePath) {
              return { tool, success: false, message: '请先打开一个 Excel 文件（.xlsx）' }
            }
            const sheet = args.sheet || 'Sheet1'
            const range = args.range || 'A1'
            const activityId = registerToolActivity('excel_read', `读取：${sheet}!${range}`)
            
            try {
              const result = await window.electronAPI!.excelReadCells(excelFilePath, sheet, range)
              if (result.success && result.cells) {
                const cellsInfo = result.cells.map(c => `${c.address}: ${c.text || c.value || '(空)'}`).join('\n')
                completeToolActivity(activityId, 'success', `${result.cells.length} 个单元格`)
                return {
                  tool,
                  success: true,
                  message: `读取 ${sheet}!${range} 成功：\n${cellsInfo}`,
                  data: { cells: result.cells, count: result.cells.length }
                }
              } else {
                completeToolActivity(activityId, 'error', result.error)
                return { tool, success: false, message: result.error || '读取失败' }
              }
            } catch (e) {
              completeToolActivity(activityId, 'error', '读取失败')
              return { tool, success: false, message: `读取失败: ${e}` }
            }
          }

          if (tool === 'excel_search') {
            if (!isExcelFile || !excelFilePath) {
              return { tool, success: false, message: '请先打开一个 Excel 文件（.xlsx）' }
            }
            const sheet = args.sheet || 'Sheet1'
            const text = args.text || args.searchText || ''
            if (!text) {
              return { tool, success: false, message: '缺少搜索文本' }
            }
            const activityId = registerToolActivity('excel_search', `搜索：${truncateLabel(text, 20)}`)
            
            try {
              const result = await window.electronAPI!.excelSearch(excelFilePath, sheet, text)
              if (result.success) {
                const count = result.count || 0
                if (count === 0) {
                  completeToolActivity(activityId, 'success', '未找到')
                  return { tool, success: true, message: `在 ${sheet} 中未找到 "${text}"` }
                }
                const cellsInfo = result.results?.slice(0, 10).map(c => `${c.address}: ${c.text}`).join('\n')
                completeToolActivity(activityId, 'success', `${count} 处`)
                return {
                  tool,
                  success: true,
                  message: `在 ${sheet} 中找到 ${count} 处 "${text}"：\n${cellsInfo}${count > 10 ? `\n...还有 ${count - 10} 处` : ''}`,
                  data: { results: result.results, count }
                }
              } else {
                completeToolActivity(activityId, 'error', result.error)
                return { tool, success: false, message: result.error || '搜索失败' }
              }
            } catch (e) {
              completeToolActivity(activityId, 'error', '搜索失败')
              return { tool, success: false, message: `搜索失败: ${e}` }
            }
          }

          if (tool === 'excel_write') {
            if (!isExcelFile || !excelFilePath) {
              return { tool, success: false, message: '请先打开一个 Excel 文件（.xlsx）' }
            }
            const sheet = args.sheet || 'Sheet1'
            let updates: Array<{address: string, value?: any, style?: any}> = []
            
            if (args.updates) {
              try {
                updates = JSON.parse(args.updates)
              } catch (e) {
                return { tool, success: false, message: '无效的 updates 参数格式' }
              }
            }
            
            if (updates.length === 0) {
              return { tool, success: false, message: '缺少要更新的单元格数据' }
            }
            
            const activityId = registerToolActivity('excel_write', `写入：${sheet}`)
            updateAgentAction(`正在写入 ${updates.length} 个单元格...`)
            
            try {
              const result = await window.electronAPI!.excelWriteCells(excelFilePath, sheet, updates)
              if (result.success) {
                // 刷新预览
                await refreshExcelData()
                completeToolActivity(activityId, 'success', `${result.count} 个`)
                return {
                  tool,
                  success: true,
                  message: `成功写入 ${result.count} 个单元格：${result.updatedCells?.join(', ')}`,
                  data: { updatedCells: result.updatedCells, count: result.count }
                }
              } else {
                completeToolActivity(activityId, 'error', result.error)
                return { tool, success: false, message: result.error || '写入失败' }
              }
            } catch (e) {
              completeToolActivity(activityId, 'error', '写入失败')
              return { tool, success: false, message: `写入失败: ${e}` }
            }
          }

          if (tool === 'excel_insert_rows') {
            if (!isExcelFile || !excelFilePath) {
              return { tool, success: false, message: '请先打开一个 Excel 文件（.xlsx）' }
            }
            const sheet = args.sheet || 'Sheet1'
            const startRow = parseInt(args.startRow, 10) || 1
            const count = parseInt(args.count, 10) || 1
            let data: any[][] | undefined
            
            if (args.data) {
              try {
                data = JSON.parse(args.data)
              } catch (e) {
                // 忽略解析错误，data 可选
              }
            }
            
            const activityId = registerToolActivity('excel_insert_rows', `插入行：${startRow}`)
            
            try {
              const result = await window.electronAPI!.excelInsertRows(excelFilePath, sheet, startRow, count, data)
              if (result.success) {
                await refreshExcelData()
                completeToolActivity(activityId, 'success', `${count} 行`)
                return {
                  tool,
                  success: true,
                  message: `成功在第 ${startRow} 行插入 ${count} 行`,
                  data: { insertedAt: result.insertedAt, count: result.count }
                }
              } else {
                completeToolActivity(activityId, 'error', result.error)
                return { tool, success: false, message: result.error || '插入失败' }
              }
            } catch (e) {
              completeToolActivity(activityId, 'error', '插入失败')
              return { tool, success: false, message: `插入失败: ${e}` }
            }
          }

          if (tool === 'excel_insert_columns') {
            if (!isExcelFile || !excelFilePath) {
              return { tool, success: false, message: '请先打开一个 Excel 文件（.xlsx）' }
            }
            const sheet = args.sheet || 'Sheet1'
            const startCol = parseInt(args.startCol, 10) || 1
            const count = parseInt(args.count, 10) || 1
            
            const activityId = registerToolActivity('excel_insert_columns', `插入列：${startCol}`)
            
            try {
              const result = await window.electronAPI!.excelInsertColumns(excelFilePath, sheet, startCol, count)
              if (result.success) {
                await refreshExcelData()
                completeToolActivity(activityId, 'success', `${count} 列`)
                return {
                  tool,
                  success: true,
                  message: `成功在第 ${startCol} 列插入 ${count} 列`,
                  data: { insertedAt: result.insertedAt, count: result.count }
                }
              } else {
                completeToolActivity(activityId, 'error', result.error)
                return { tool, success: false, message: result.error || '插入失败' }
              }
            } catch (e) {
              completeToolActivity(activityId, 'error', '插入失败')
              return { tool, success: false, message: `插入失败: ${e}` }
            }
          }

          if (tool === 'excel_delete_rows') {
            if (!isExcelFile || !excelFilePath) {
              return { tool, success: false, message: '请先打开一个 Excel 文件（.xlsx）' }
            }
            const sheet = args.sheet || 'Sheet1'
            const startRow = parseInt(args.startRow, 10) || 1
            const count = parseInt(args.count, 10) || 1
            
            const activityId = registerToolActivity('excel_delete_rows', `删除行：${startRow}`)
            
            try {
              const result = await window.electronAPI!.excelDeleteRows(excelFilePath, sheet, startRow, count)
              if (result.success) {
                await refreshExcelData()
                completeToolActivity(activityId, 'success', `${count} 行`)
                return {
                  tool,
                  success: true,
                  message: `成功删除第 ${startRow} 行开始的 ${count} 行`,
                  data: { deletedFrom: result.deletedFrom, count: result.count }
                }
              } else {
                completeToolActivity(activityId, 'error', result.error)
                return { tool, success: false, message: result.error || '删除失败' }
              }
            } catch (e) {
              completeToolActivity(activityId, 'error', '删除失败')
              return { tool, success: false, message: `删除失败: ${e}` }
            }
          }

          if (tool === 'excel_delete_columns') {
            if (!isExcelFile || !excelFilePath) {
              return { tool, success: false, message: '请先打开一个 Excel 文件（.xlsx）' }
            }
            const sheet = args.sheet || 'Sheet1'
            const startCol = parseInt(args.startCol, 10) || 1
            const count = parseInt(args.count, 10) || 1
            
            const activityId = registerToolActivity('excel_delete_columns', `删除列：${startCol}`)
            
            try {
              const result = await window.electronAPI!.excelDeleteColumns(excelFilePath, sheet, startCol, count)
              if (result.success) {
                await refreshExcelData()
                completeToolActivity(activityId, 'success', `${count} 列`)
                return {
                  tool,
                  success: true,
                  message: `成功删除第 ${startCol} 列开始的 ${count} 列`,
                  data: { deletedFrom: result.deletedFrom, count: result.count }
                }
              } else {
                completeToolActivity(activityId, 'error', result.error)
                return { tool, success: false, message: result.error || '删除失败' }
              }
            } catch (e) {
              completeToolActivity(activityId, 'error', '删除失败')
              return { tool, success: false, message: `删除失败: ${e}` }
            }
          }

          if (tool === 'excel_add_sheet') {
            if (!isExcelFile || !excelFilePath) {
              return { tool, success: false, message: '请先打开一个 Excel 文件（.xlsx）' }
            }
            const name = args.name || args.sheetName || '新工作表'
            
            const activityId = registerToolActivity('excel_add_sheet', `新建：${name}`)
            
            try {
              const result = await window.electronAPI!.excelAddSheet(excelFilePath, name)
              if (result.success) {
                await refreshExcelData()
                completeToolActivity(activityId, 'success')
                return {
                  tool,
                  success: true,
                  message: `成功创建工作表 "${name}"`,
                  data: { sheetName: result.sheetName }
                }
              } else {
                completeToolActivity(activityId, 'error', result.error)
                return { tool, success: false, message: result.error || '创建失败' }
              }
            } catch (e) {
              completeToolActivity(activityId, 'error', '创建失败')
              return { tool, success: false, message: `创建失败: ${e}` }
            }
          }

          if (tool === 'excel_delete_sheet') {
            if (!isExcelFile || !excelFilePath) {
              return { tool, success: false, message: '请先打开一个 Excel 文件（.xlsx）' }
            }
            const name = args.name || args.sheetName || ''
            if (!name) {
              return { tool, success: false, message: '缺少工作表名称' }
            }
            
            const activityId = registerToolActivity('excel_delete_sheet', `删除：${name}`)
            
            try {
              const result = await window.electronAPI!.excelDeleteSheet(excelFilePath, name)
              if (result.success) {
                await refreshExcelData()
                completeToolActivity(activityId, 'success')
                return {
                  tool,
                  success: true,
                  message: `成功删除工作表 "${name}"`,
                  data: { deletedSheet: result.deletedSheet }
                }
              } else {
                completeToolActivity(activityId, 'error', result.error)
                return { tool, success: false, message: result.error || '删除失败' }
              }
            } catch (e) {
              completeToolActivity(activityId, 'error', '删除失败')
              return { tool, success: false, message: `删除失败: ${e}` }
            }
          }

          if (tool === 'excel_merge') {
            if (!isExcelFile || !excelFilePath) {
              return { tool, success: false, message: '请先打开一个 Excel 文件（.xlsx）' }
            }
            const sheet = args.sheet || 'Sheet1'
            const range = args.range || ''
            if (!range) {
              return { tool, success: false, message: '缺少合并范围 range（如 A1:C1）' }
            }
            
            const activityId = registerToolActivity('excel_merge', `合并：${range}`)
            
            try {
              const result = await window.electronAPI!.excelMergeCells(excelFilePath, sheet, range)
              if (result.success) {
                await refreshExcelData()
                completeToolActivity(activityId, 'success')
                return {
                  tool,
                  success: true,
                  message: `成功合并单元格 ${range}`,
                  data: { mergedRange: result.mergedRange }
                }
              } else {
                completeToolActivity(activityId, 'error', result.error)
                return { tool, success: false, message: result.error || '合并失败' }
              }
            } catch (e) {
              completeToolActivity(activityId, 'error', '合并失败')
              return { tool, success: false, message: `合并失败: ${e}` }
            }
          }

          if (tool === 'excel_unmerge') {
            if (!isExcelFile || !excelFilePath) {
              return { tool, success: false, message: '请先打开一个 Excel 文件（.xlsx）' }
            }
            const sheet = args.sheet || 'Sheet1'
            const range = args.range || ''
            if (!range) {
              return { tool, success: false, message: '缺少取消合并范围 range（如 A1:C1）' }
            }
            
            const activityId = registerToolActivity('excel_unmerge', `取消合并：${range}`)
            
            try {
              const result = await window.electronAPI!.excelUnmergeCells(excelFilePath, sheet, range)
              if (result.success) {
                await refreshExcelData()
                completeToolActivity(activityId, 'success')
                return {
                  tool,
                  success: true,
                  message: `成功取消合并单元格 ${range}`,
                  data: { unmergedRange: result.unmergedRange }
                }
              } else {
                completeToolActivity(activityId, 'error', result.error)
                return { tool, success: false, message: result.error || '取消合并失败' }
              }
            } catch (e) {
              completeToolActivity(activityId, 'error', '取消合并失败')
              return { tool, success: false, message: `取消合并失败: ${e}` }
            }
          }

          // 创建新 Excel 文件
          if (tool === 'excel_create') {
            // 检查是否有工作区
            if (!workspacePath) {
              return { 
                tool, 
                success: false, 
                message: '请先在左侧点击"打开文件夹"选择一个工作区，然后再创建 Excel 文件' 
              }
            }
            
            const filename = args.filename || args.name || '新建表格.xlsx'
            let sheets: Array<{ name?: string; data?: any[][]; columnWidths?: number[]; merges?: string[] }> = []
            
            // 解析 sheets 参数
            if (args.sheets) {
              try {
                sheets = JSON.parse(args.sheets)
              } catch (e) {
                // 如果解析失败，尝试简单数据格式
              }
            }
            
            // 如果没有 sheets，使用简单数据格式
            if (sheets.length === 0 && args.data) {
              try {
                const data = JSON.parse(args.data)
                sheets = [{ name: args.sheetName || 'Sheet1', data }]
              } catch (e) {
                return { tool, success: false, message: '无效的数据格式，请提供有效的 JSON 数组' }
              }
            }
            
            // 如果还是没有数据，创建空表格
            if (sheets.length === 0) {
              sheets = [{ name: 'Sheet1', data: [] }]
            }
            
            // 构建文件路径 - 保存到工作区
            let finalFilename = filename
            // 确保文件名以 .xlsx 结尾
            if (!finalFilename.toLowerCase().endsWith('.xlsx')) {
              finalFilename += '.xlsx'
            }
            // 使用工作区路径
            const filePath = `${workspacePath}/${finalFilename}`
            
            const activityId = registerToolActivity('excel_create', `创建：${finalFilename}`)
            
            try {
              const result = await window.electronAPI!.excelCreate(filePath, { sheets, openAfterCreate: true })
              if (result.success) {
                completeToolActivity(activityId, 'success')
                
                // 刷新文件列表，让新文件出现在左侧
                await refreshFiles()
                
                // 自动打开创建的文件
                if (result.openAfterCreate && result.filePath) {
                  const newFile = {
                    name: finalFilename,
                    path: result.filePath,
                    type: 'file' as const
                  }
                  await openFile(newFile)
                }
                
                return {
                  tool,
                  success: true,
                  message: `成功创建 Excel 文件：${result.filePath}\n工作表：${result.sheetsCreated?.join(', ')}\n文件已保存到工作区并自动打开`,
                  data: { filePath: result.filePath, fileName: finalFilename, sheetsCreated: result.sheetsCreated }
                }
              } else {
                completeToolActivity(activityId, 'error', result.error)
                return { tool, success: false, message: result.error || '创建失败' }
              }
            } catch (e) {
              completeToolActivity(activityId, 'error', '创建失败')
              return { tool, success: false, message: `创建失败: ${e}` }
            }
          }

          // Excel 公式设置
          if (tool === 'excel_formula') {
            if (!isExcelFile || !excelFilePath) {
              return { tool, success: false, message: '请先打开一个 Excel 文件（.xlsx）' }
            }
            const sheet = args.sheet || 'Sheet1'
            let formulas: Array<{ address: string; formula: string; numberFormat?: string }> = []
            
            try {
              if (args.formulas) {
                formulas = JSON.parse(args.formulas)
              } else if (args.address && args.formula) {
                formulas = [{ address: args.address, formula: args.formula, numberFormat: args.numberFormat }]
              }
            } catch {
              return { tool, success: false, message: '无效的公式格式' }
            }
            
            if (formulas.length === 0) {
              return { tool, success: false, message: '缺少公式参数' }
            }
            
            const activityId = registerToolActivity('excel_formula', `设置 ${formulas.length} 个公式`)
            
            try {
              const result = await window.electronAPI!.excelSetFormula(excelFilePath, sheet, formulas)
              if (result.success) {
                await refreshExcelData()
                completeToolActivity(activityId, 'success')
                return {
                  tool,
                  success: true,
                  message: `成功设置 ${result.count} 个公式`,
                  data: { formulas: result.formulas }
                }
              } else {
                completeToolActivity(activityId, 'error', result.error)
                return { tool, success: false, message: result.error || '设置公式失败' }
              }
            } catch (e) {
              completeToolActivity(activityId, 'error', '设置公式失败')
              return { tool, success: false, message: `设置公式失败: ${e}` }
            }
          }

          // Excel 排序
          if (tool === 'excel_sort') {
            if (!isExcelFile || !excelFilePath) {
              return { tool, success: false, message: '请先打开一个 Excel 文件（.xlsx）' }
            }
            const sheet = args.sheet || 'Sheet1'
            const range = args.range || ''
            const column = args.column || 'A'
            const ascending = args.ascending !== 'false'
            const hasHeader = args.hasHeader !== 'false'
            
            if (!range) {
              return { tool, success: false, message: '缺少排序范围 range（如 A1:D10）' }
            }
            
            const activityId = registerToolActivity('excel_sort', `排序 ${range} 按列 ${column}`)
            
            try {
              const result = await window.electronAPI!.excelSort(excelFilePath, sheet, {
                range, column, ascending, hasHeader
              })
              if (result.success) {
                await refreshExcelData()
                completeToolActivity(activityId, 'success')
                return {
                  tool,
                  success: true,
                  message: `成功排序 ${result.sortedRows} 行，按列 ${column} ${ascending ? '升序' : '降序'}`,
                  data: result
                }
              } else {
                completeToolActivity(activityId, 'error', result.error)
                return { tool, success: false, message: result.error || '排序失败' }
              }
            } catch (e) {
              completeToolActivity(activityId, 'error', '排序失败')
              return { tool, success: false, message: `排序失败: ${e}` }
            }
          }

          // Excel 自动填充
          if (tool === 'excel_autofill') {
            if (!isExcelFile || !excelFilePath) {
              return { tool, success: false, message: '请先打开一个 Excel 文件（.xlsx）' }
            }
            const sheet = args.sheet || 'Sheet1'
            const sourceRange = args.sourceRange || args.source || ''
            const targetRange = args.targetRange || args.target || ''
            const fillType = (args.fillType || args.type || 'copy') as 'copy' | 'series' | 'formula'
            
            if (!sourceRange || !targetRange) {
              return { tool, success: false, message: '缺少源范围或目标范围' }
            }
            
            const activityId = registerToolActivity('excel_autofill', `从 ${sourceRange} 填充到 ${targetRange}`)
            
            try {
              const result = await window.electronAPI!.excelAutoFill(excelFilePath, sheet, {
                sourceRange, targetRange, fillType
              })
              if (result.success) {
                await refreshExcelData()
                completeToolActivity(activityId, 'success')
                return {
                  tool,
                  success: true,
                  message: `成功填充 ${result.filledCells} 个单元格（${fillType} 模式）`,
                  data: result
                }
              } else {
                completeToolActivity(activityId, 'error', result.error)
                return { tool, success: false, message: result.error || '自动填充失败' }
              }
            } catch (e) {
              completeToolActivity(activityId, 'error', '自动填充失败')
              return { tool, success: false, message: `自动填充失败: ${e}` }
            }
          }

          // Excel 设置列宽/行高
          if (tool === 'excel_dimensions') {
            if (!isExcelFile || !excelFilePath) {
              return { tool, success: false, message: '请先打开一个 Excel 文件（.xlsx）' }
            }
            const sheet = args.sheet || 'Sheet1'
            let columns: Array<{ column: string | number; width?: number; hidden?: boolean }> = []
            let rows: Array<{ row: number; height?: number; hidden?: boolean }> = []
            
            try {
              if (args.columns) columns = JSON.parse(args.columns)
              if (args.rows) rows = JSON.parse(args.rows)
            } catch {
              return { tool, success: false, message: '无效的列宽/行高格式' }
            }
            
            const activityId = registerToolActivity('excel_dimensions', `设置 ${columns.length} 列宽, ${rows.length} 行高`)
            
            try {
              const result = await window.electronAPI!.excelSetDimensions(excelFilePath, sheet, { columns, rows })
              if (result.success) {
                await refreshExcelData()
                completeToolActivity(activityId, 'success')
                return {
                  tool,
                  success: true,
                  message: `成功设置 ${result.columnsSet} 列宽, ${result.rowsSet} 行高`,
                  data: result
                }
              } else {
                completeToolActivity(activityId, 'error', result.error)
                return { tool, success: false, message: result.error || '设置失败' }
              }
            } catch (e) {
              completeToolActivity(activityId, 'error', '设置失败')
              return { tool, success: false, message: `设置失败: ${e}` }
            }
          }

          // Excel 条件格式
          if (tool === 'excel_conditional_format') {
            if (!isExcelFile || !excelFilePath) {
              return { tool, success: false, message: '请先打开一个 Excel 文件（.xlsx）' }
            }
            const sheet = args.sheet || 'Sheet1'
            const range = args.range || ''
            let rules: Array<{ type: string; operator?: string; value?: string | number | string[]; fill?: { bgColor: string } | string; font?: object }> = []
            
            if (!range) {
              return { tool, success: false, message: '缺少范围 range' }
            }
            
            try {
              if (args.rules) {
                rules = JSON.parse(args.rules)
              } else if (args.type) {
                // 简单格式
                rules = [{
                  type: args.type,
                  operator: args.operator,
                  value: args.value,
                  fill: args.fill ? { bgColor: args.fill } : undefined
                }]
              }
            } catch {
              return { tool, success: false, message: '无效的规则格式' }
            }
            
            const activityId = registerToolActivity('excel_conditional_format', `设置 ${rules.length} 条条件格式`)
            
            try {
              const result = await window.electronAPI!.excelConditionalFormat(excelFilePath, sheet, { range, rules })
              if (result.success) {
                await refreshExcelData()
                completeToolActivity(activityId, 'success')
                return {
                  tool,
                  success: true,
                  message: `成功设置 ${result.rulesApplied} 条条件格式规则`,
                  data: result
                }
              } else {
                completeToolActivity(activityId, 'error', result.error)
                return { tool, success: false, message: result.error || '设置条件格式失败' }
              }
            } catch (e) {
              completeToolActivity(activityId, 'error', '设置条件格式失败')
              return { tool, success: false, message: `设置条件格式失败: ${e}` }
            }
          }

          // Excel 获取计算结果
          if (tool === 'excel_calculate') {
            if (!isExcelFile || !excelFilePath) {
              return { tool, success: false, message: '请先打开一个 Excel 文件（.xlsx）' }
            }
            const sheet = args.sheet || 'Sheet1'
            let addresses: string[] = []
            
            try {
              if (args.addresses) {
                addresses = JSON.parse(args.addresses)
              } else if (args.address) {
                addresses = [args.address]
              }
            } catch {
              return { tool, success: false, message: '无效的地址格式' }
            }
            
            if (addresses.length === 0) {
              return { tool, success: false, message: '缺少单元格地址' }
            }
            
            try {
              const result = await window.electronAPI!.excelCalculate(excelFilePath, sheet, addresses)
              if (result.success) {
                return {
                  tool,
                  success: true,
                  message: `获取了 ${result.results?.length || 0} 个单元格的值`,
                  data: { results: result.results }
                }
              } else {
                return { tool, success: false, message: result.error || '获取计算结果失败' }
              }
            } catch (e) {
              return { tool, success: false, message: `获取计算结果失败: ${e}` }
            }
          }

          // 【新增】Excel 自动筛选
          if (tool === 'excel_filter') {
            if (!isExcelFile || !excelFilePath) {
              return { tool, success: false, message: '请先打开一个 Excel 文件（.xlsx）' }
            }
            const sheet = args.sheet || 'Sheet1'
            const range = args.range || ''
            const action = (args.action || 'set').toLowerCase()
            
            try {
              const result = await window.electronAPI!.excelSetFilter(excelFilePath, sheet, {
                range: range,
                remove: action === 'remove'
              })
              if (result.success) {
                await refreshExcelData()
                return { tool, success: true, message: result.message || '已设置自动筛选' }
              } else {
                return { tool, success: false, message: result.error || '设置自动筛选失败' }
              }
            } catch (e) {
              return { tool, success: false, message: `设置自动筛选失败: ${e}` }
            }
          }

          // 【新增】Excel 数据验证
          if (tool === 'excel_validation') {
            if (!isExcelFile || !excelFilePath) {
              return { tool, success: false, message: '请先打开一个 Excel 文件（.xlsx）' }
            }
            const sheet = args.sheet || 'Sheet1'
            const range = args.range || ''
            const type = args.type || 'list'
            const action = (args.action || 'set').toLowerCase()
            
            if (!range) {
              return { tool, success: false, message: '请指定单元格范围 (range)' }
            }
            
            let values: string[] = []
            if (args.values) {
              try {
                values = JSON.parse(args.values)
              } catch {
                // 如果不是 JSON，尝试按逗号分割
                values = args.values.split(',').map((v: string) => v.trim())
              }
            }
            
            try {
              const result = await window.electronAPI!.excelSetValidation(excelFilePath, sheet, {
                range,
                type: type as 'list' | 'whole' | 'decimal',
                values,
                min: args.min ? parseFloat(args.min) : undefined,
                max: args.max ? parseFloat(args.max) : undefined,
                remove: action === 'remove'
              })
              if (result.success) {
                await refreshExcelData()
                return { tool, success: true, message: result.message || '已设置数据验证' }
              } else {
                return { tool, success: false, message: result.error || '设置数据验证失败' }
              }
            } catch (e) {
              return { tool, success: false, message: `设置数据验证失败: ${e}` }
            }
          }

          // 【新增】Excel 超链接
          if (tool === 'excel_hyperlink') {
            if (!isExcelFile || !excelFilePath) {
              return { tool, success: false, message: '请先打开一个 Excel 文件（.xlsx）' }
            }
            const sheet = args.sheet || 'Sheet1'
            const cell = args.cell || ''
            const url = args.url || ''
            const text = args.text || url
            const action = (args.action || 'set').toLowerCase()
            
            if (!cell) {
              return { tool, success: false, message: '请指定单元格地址 (cell)' }
            }
            
            try {
              const result = await window.electronAPI!.excelSetHyperlink(excelFilePath, sheet, {
                cell,
                url,
                text,
                tooltip: args.tooltip,
                remove: action === 'remove'
              })
              if (result.success) {
                await refreshExcelData()
                return { tool, success: true, message: result.message || '已设置超链接' }
              } else {
                return { tool, success: false, message: result.error || '设置超链接失败' }
              }
            } catch (e) {
              return { tool, success: false, message: `设置超链接失败: ${e}` }
            }
          }

          // 【新增】Excel 查找替换
          if (tool === 'excel_find_replace') {
            if (!isExcelFile || !excelFilePath) {
              return { tool, success: false, message: '请先打开一个 Excel 文件（.xlsx）' }
            }
            const sheet = args.sheet || 'Sheet1'
            const find = args.find || ''
            const replace = args.replace || ''
            
            if (!find) {
              return { tool, success: false, message: '请指定要查找的内容 (find)' }
            }
            
            try {
              const result = await window.electronAPI!.excelFindReplace(excelFilePath, sheet, {
                find,
                replace,
                matchCase: args.matchCase === 'true',
                matchWholeCell: args.matchWholeCell === 'true',
                allSheets: args.allSheets === 'true'
              })
              if (result.success) {
                await refreshExcelData()
                return { 
                  tool, 
                  success: true, 
                  message: result.message || `已替换 ${result.count || 0} 处`,
                  data: { count: result.count, details: result.details }
                }
              } else {
                return { tool, success: false, message: result.error || '查找替换失败' }
              }
            } catch (e) {
              return { tool, success: false, message: `查找替换失败: ${e}` }
            }
          }

          // 【新增】Excel 图表
          if (tool === 'excel_chart') {
            if (!isExcelFile || !excelFilePath) {
              return { tool, success: false, message: '请先打开一个 Excel 文件（.xlsx）' }
            }
            const sheet = args.sheet || 'Sheet1'
            const type = args.type || 'column'
            const dataRange = args.dataRange || ''
            const title = args.title || ''
            const position = args.position || 'E1'
            
            if (!dataRange) {
              return { tool, success: false, message: '请指定数据范围 (dataRange)' }
            }
            
            try {
              const result = await window.electronAPI!.excelInsertChart(excelFilePath, sheet, {
                type: type as 'column' | 'bar' | 'line' | 'pie',
                dataRange,
                title,
                position,
                width: args.width ? parseInt(args.width) : 500,
                height: args.height ? parseInt(args.height) : 300
              })
              if (result.success) {
                await refreshExcelData()
                return { 
                  tool, 
                  success: true, 
                  message: result.message || '已添加图表配置',
                  data: { chartConfig: result.chartConfig }
                }
              } else {
                return { tool, success: false, message: result.error || '添加图表失败' }
              }
            } catch (e) {
              return { tool, success: false, message: `添加图表失败: ${e}` }
            }
          }

          return { tool, success: false, message: `未知工具: ${tool}` }
        },

        // 完成时的处理
        onComplete: (content, toolResults) => {
          // 完成 Agent 进度
          finishAgentProgress()
          
          console.log('[onComplete] content:', content?.substring(0, 200))
          console.log('[onComplete] toolResults:', toolResults.length)
          
          // 如果有工具调用结果，显示统计
          if (toolResults.length > 0) {
            const successCount = toolResults.filter(r => r.success).length
            const replaceResults = toolResults.filter(r => r.tool === 'replace' && r.success)
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
            } else if (replaceResults.length > 0) {
              const diffChanges = replaceResults.map(r => ({
                searchText: r.data?.searchText as string || '',
                replaceText: r.data?.replaceText as string || '',
                count: (r.data?.count as number) || 0
              }))
              const totalCount = diffChanges.reduce((sum, d) => sum + d.count, 0)
              statusBadge = `\n\n---\n✅ **已更新文档** 📄 \`${fileName}\` (~${totalCount} 处修改)`
              
              // 替换操作保留 diffChanges
              addMessage({
                role: 'assistant',
                content: (content?.trim() ? content : '已按你的要求完成修改，下面是变更结果：') + statusBadge,
                diffChanges,
                fileName
              })
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
              // 所有工具调用都失败了
              addMessage({
                role: 'assistant',
                content: content || '操作未能完成，请检查文档内容是否匹配'
              })
            } else {
              // 显示 AI 的总结内容 + 状态标签
              // 如果 content 为空，至少显示操作结果
              const finalContent = content?.trim() 
                ? content + statusBadge 
                : (statusBadge ? `任务已完成！${statusBadge}` : '任务已完成')
              
              console.log('[onComplete] finalContent:', finalContent?.substring(0, 200))
              
              addMessage({
                role: 'assistant',
                content: finalContent,
                fileName: resultFileName
              })
            }
          } else {
            // 没有工具调用，普通对话
            addMessage({
              role: 'assistant',
              content: content || '完成'
            })
          }
        },
        
        // 获取最新文档内容（用于在工具调用后让 AI 知道文档已更新）
        // 使用 getLatestContent() 而不是 document.content 避免闭包问题
        getLatestDocument: () => {
          return getLatestContent()
        }
      }
    )
  }, [
    input,
    isLoading,
    pptEditContext,
    attachedFiles,
    addMessage,
    sendAgentMessage,
    document.content,
    buildFilesContext,
    createNewDocument,
    currentFile?.name,
    replaceInDocument,
    startAgentProgress,
    updateAgentAction,
    completeAgentStep,
    updateAgentFile,
    addAgentFileOperation,
    finishAgentProgress,
    insertInDocument,
    deleteInDocument,
    currentFile?.path,
    resetToolActivity,
    registerToolActivity,
    completeToolActivity,
    excelData,
    refreshExcelData,
    settings,
    refreshFiles,
    openFile,
    workspacePath,
    getLatestContent
  ])

  const handleKeyDown = useCallback((e: React.KeyboardEvent) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault()
      handleSend()
    }
  }, [handleSend])

  // 拖拽处理
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
        }
      }
    } catch (error) {
      console.error('Drop error:', error)
    }
  }

  const removeAttachedFile = (path: string) => {
    setAttachedFiles(prev => prev.filter(f => f.path !== path))
  }

  // 快捷命令
  const quickCommands = [
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

  return (
    <div 
      className={`flex flex-col h-full bg-[#1e1e1e] border-l border-[#2d2d2d] ${isDragOver ? 'ring-2 ring-primary ring-inset' : ''}`}
      onDragOver={handleDragOver}
      onDragLeave={handleDragLeave}
      onDrop={handleDrop}
    >
      {/* 头部 - Cursor 风格 */}
      <div className="flex items-center justify-between px-3 py-2.5 border-b border-[#2d2d2d] bg-[#252526]">
        <div className="flex items-center gap-2">
          <div className="w-6 h-6 rounded-md bg-gradient-to-br from-violet-500 to-fuchsia-500 flex items-center justify-center">
            <Bot className="w-3.5 h-3.5 text-white" />
          </div>
          <span className="text-[13px] font-medium text-[#cccccc]">AI 助手</span>
        </div>
        <button
          onClick={clearMessages}
          className="p-1.5 rounded-md text-[#858585] hover:text-[#cccccc] hover:bg-[#2d2d2d] transition-colors"
          title="清空对话"
        >
          <Trash2 className="w-3.5 h-3.5" />
        </button>
      </div>

      {/* 快捷命令 - 更紧凑 */}
      <div className="px-3 py-2 border-b border-[#2d2d2d] flex gap-1.5 overflow-x-auto scrollbar-none">
        {quickCommands.map((cmd, i) => (
          <button
            key={i}
            onClick={() => setInput(cmd.command)}
            className="flex items-center gap-1 px-2 py-1 bg-[#2d2d2d] hover:bg-[#3c3c3c] text-[11px] text-[#858585] hover:text-[#cccccc] rounded-md transition-colors whitespace-nowrap"
          >
            {cmd.icon}
            <span>{cmd.label}</span>
          </button>
        ))}
      </div>

      {/* 拖拽提示 */}
      {isDragOver && (
        <div className="absolute inset-0 z-50 flex items-center justify-center bg-[#1e1e1e]/90 backdrop-blur-sm">
          <div className="flex flex-col items-center gap-2 p-6 bg-[#2d2d2d] border border-[#3c3c3c] rounded-lg">
            <Paperclip className="w-8 h-8 text-violet-400" />
            <p className="text-sm text-[#cccccc]">释放以添加文件</p>
          </div>
        </div>
      )}

      {/* 消息列表 - Cursor 风格 + Framer Motion */}
      <div className="flex-1 overflow-y-auto px-3 py-3 space-y-4 scrollbar-thin">
        <AnimatePresence mode="popLayout">
        {displayMessages.map((message) => (
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
                <div className="bg-gradient-to-b from-[#0e639c]/35 to-[#0e639c]/20 border border-[#0e639c]/35 text-[#e6f1ff] rounded-2xl rounded-tr-sm px-3 py-2 shadow-[0_6px_20px_rgba(14,99,156,0.12)]">
                  <p className="text-[13px] leading-relaxed whitespace-pre-wrap">{message.content}</p>
                </div>
                <span className="text-[10px] text-[#5a5a5a] mt-1 block text-right pr-1">
                  {message.timestamp.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
                </span>
              </div>
            ) : message.content.includes('\n---\n✅') ? (
              /* 操作完成消息 - 显示 AI 总结 + 状态卡片 */
              <div className="w-full space-y-3">
                {/* AI 总结内容 */}
                {(() => {
                  const parts = message.content.split('\n---\n')
                  const summaryContent = parts[0]
                  const statusContent = parts.slice(1).join('\n---\n')
                  return (
                    <>
                      {summaryContent && (
                        <div className="text-[13px] leading-relaxed text-[#cccccc] prose prose-invert prose-sm max-w-none">
                          <ReactMarkdown
                            components={{
                              p: ({ children }) => <p className="mb-2 last:mb-0">{children}</p>,
                              ul: ({ children }) => <ul className="list-disc pl-4 mb-2 space-y-1">{children}</ul>,
                              ol: ({ children }) => <ol className="list-decimal pl-4 mb-2 space-y-1">{children}</ol>,
                              li: ({ children }) => <li className="text-[13px]">{children}</li>,
                              strong: ({ children }) => <strong className="font-semibold text-[#e5c07b]">{children}</strong>,
                              code: ({ children }) => <code className="bg-[#2d2d2d] px-1 py-0.5 rounded text-[#e06c75] text-[12px]">{children}</code>,
                            }}
                          >
                            {summaryContent}
                          </ReactMarkdown>
                        </div>
                      )}
                      {/* 状态卡片 */}
                      {statusContent && (
                        <div className="bg-[#252526] border border-[#2d2d2d] rounded-lg overflow-hidden">
                          <div className="flex items-center gap-2 px-3 py-2 bg-[#1e3a29] border-b border-[#2d4a39]">
                            <CheckCircle className="w-3.5 h-3.5 text-[#4ec9b0]" />
                            <span className="text-[12px] font-medium text-[#4ec9b0]">
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
                                    className="w-full flex items-center justify-between gap-2 py-1 hover:bg-[#2d2d2d] cursor-pointer rounded"
                                  >
                                    <div className="flex items-center gap-2 min-w-0">
                                      {emoji === '📊' ? (
                                        <Table className="w-3.5 h-3.5 text-[#4ec9b0] flex-shrink-0" />
                                      ) : (
                                        <FileText className="w-3.5 h-3.5 text-[#75beff] flex-shrink-0" />
                                      )}
                                      <span className="text-[12px] text-[#cccccc] font-mono truncate">{fileNamePart}</span>
                                    </div>
                                    {stats && (
                                      <span className="text-[10px] font-mono text-[#4ec9b0]">{stats}</span>
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
                <span className="text-[10px] text-[#5a5a5a] mt-1 block">
                  {message.timestamp.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
                </span>
              </div>
            ) : message.content.startsWith('✅') ? (
              /* 简单操作完成消息 - Cursor 风格卡片 */
              <div className="w-full">
                <div className="bg-[#252526] border border-[#2d2d2d] rounded-lg overflow-hidden">
                  {/* 成功标题栏 */}
                  <div className="flex items-center gap-2 px-3 py-2 bg-[#1e3a29] border-b border-[#2d4a39]">
                    <CheckCircle className="w-3.5 h-3.5 text-[#4ec9b0]" />
                    <span className="text-[12px] font-medium text-[#4ec9b0]">
                      {message.content.includes('表格') ? '表格已创建' : message.content.includes('创建') ? '文档已创建' : '文档已更新'}
                    </span>
                  </div>
                  {/* 文件信息 */}
                  <div className="px-3 py-2">
                    {message.content.split('\n').slice(1).map((line, i) => {
                      if (line.startsWith('📄') || line.startsWith('📊')) {
                        const emoji = line.startsWith('📊') ? '📊' : '📄'
                        const parts = line.replace(/^(📄|📊)\s*/, '').split(/\s+/)
                        const fileNamePart = parts[0]?.replace(/`/g, '')
                        const stats = parts.slice(1).join(' ')
                        const isCreateMessage = message.content.includes('创建')
                        return (
                          <button
                            key={i}
                            onClick={() => {
                              if (isCreateMessage && fileNamePart) {
                                openCreatedFile(fileNamePart)
                              }
                            }}
                            className={`w-full flex items-center justify-between gap-2 py-1 ${isCreateMessage ? 'hover:bg-[#2d2d2d] cursor-pointer rounded' : ''}`}
                          >
                            <div className="flex items-center gap-2 min-w-0">
                              {emoji === '📊' ? (
                                <Table className="w-3.5 h-3.5 text-[#4ec9b0] flex-shrink-0" />
                              ) : (
                                <FileText className="w-3.5 h-3.5 text-[#75beff] flex-shrink-0" />
                              )}
                              <span className="text-[12px] text-[#cccccc] font-mono truncate">{fileNamePart}</span>
                            </div>
                            <div className="flex items-center gap-1 flex-shrink-0">
                              {stats.includes('+') && (
                                <span className="text-[10px] font-mono text-[#4ec9b0]">
                                  {stats.match(/\+\d+/)?.[0]}
                                </span>
                              )}
                              {stats.includes('-') && (
                                <span className="text-[10px] font-mono text-[#f14c4c]">
                                  {stats.match(/-\d+/)?.[0]}
                                </span>
                              )}
                              {stats.includes('~') && (
                                <span className="text-[10px] font-mono text-[#cca700]">
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
                    <div className="border-t border-[#2d2d2d] px-3 py-2">
                      <div className="text-[10px] text-[#858585] mb-2">修改详情</div>
                      <div className="space-y-1">
                        {message.diffChanges.slice(0, 5).map((diff, i) => (
                          <button
                            key={i}
                            onClick={() => scrollToChange(diff.replaceText)}
                            className="w-full text-left px-2 py-1.5 rounded bg-[#1e1e1e] hover:bg-[#2d2d2d] transition-colors"
                          >
                            <div className="flex items-center gap-2 text-[11px]">
                              <span className="text-[#f14c4c] line-through truncate flex-1" title={diff.searchText}>
                                {diff.searchText.slice(0, 25)}{diff.searchText.length > 25 ? '...' : ''}
                              </span>
                              <span className="text-[#5a5a5a]">→</span>
                              <span className="text-[#4ec9b0] truncate flex-1" title={diff.replaceText}>
                                {diff.replaceText.slice(0, 25)}{diff.replaceText.length > 25 ? '...' : ''}
                              </span>
                            </div>
                          </button>
                        ))}
                        {message.diffChanges.length > 5 && (
                          <div className="text-[10px] text-[#858585] text-center py-1">
                            还有 {message.diffChanges.length - 5} 处修改...
                          </div>
                        )}
                      </div>
                    </div>
                  )}
                </div>
                <span className="text-[10px] text-[#5a5a5a] mt-1 block pl-1">
                  {message.timestamp.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
                </span>
              </div>
            ) : (
              /* AI 普通消息 - 使用 Markdown 渲染 */
              <div className="w-full">
                <div className="bg-[#252526] border border-[#2d2d2d] rounded-lg rounded-tl-sm px-3 py-2">
                  <div className="ai-markdown text-[13px] text-[#d4d4d4] leading-relaxed">
                    {(() => {
                      const parsed = tryParsePptOutlineDraft(message.content)
                      const cleanedText = parsed ? stripPptOutlineJsonFromText(message.content) : message.content
                      const jsonOpen = !!outlineJsonOpen[message.id]

                      return (
                        <>
                          {parsed && (
                            <div className="mb-3 bg-[#1e1e1e] border border-[#2d2d2d] rounded-lg overflow-hidden">
                              <div className="flex items-center justify-between px-3 py-2 bg-[#252526] border-b border-[#2d2d2d]">
                                <div className="min-w-0">
                                  <div className="text-[12px] text-[#cccccc] truncate">
                                    PPT 大纲：{parsed.draft.title || '未命名'}（{parsed.draft.slides.length} 页）
                                  </div>
                                  <div className="text-[10px] text-[#858585] truncate">
                                    {parsed.draft.theme ? `主题：${parsed.draft.theme}  ` : ''}{parsed.draft.styleHint ? `风格：${parsed.draft.styleHint}` : ''}
                                  </div>
                                </div>
                                <button
                                  onClick={() =>
                                    setOutlineJsonOpen((prev) => ({ ...prev, [message.id]: !prev[message.id] }))
                                  }
                                  className="px-2 py-1 text-[10px] rounded bg-[#2d2d2d] hover:bg-[#3c3c3c] text-[#cccccc] transition-colors flex-shrink-0"
                                  title={jsonOpen ? '收起 JSON' : '展开 JSON'}
                                >
                                  {jsonOpen ? '收起 JSON' : '展开 JSON'}
                                </button>
                              </div>

                              <div className="px-3 py-2 space-y-2">
                                {parsed.draft.slides.map((s, idx) => (
                                  <div key={`${s.pageNumber}-${idx}`} className="border border-[#2d2d2d] rounded-md bg-[#252526]">
                                    <div className="px-2.5 py-2 border-b border-[#2d2d2d] flex items-center justify-between gap-2">
                                      <div className="min-w-0">
                                        <div className="text-[12px] text-[#e1e1e1] truncate">
                                          第{s.pageNumber || idx + 1}页：{s.headline || '（未填写标题）'}
                                        </div>
                                        {s.subheadline && (
                                          <div className="text-[10px] text-[#9cdcfe] truncate">{s.subheadline}</div>
                                        )}
                                      </div>
                                      {s.layoutIntent && (
                                        <div className="text-[10px] text-[#858585] flex-shrink-0 truncate max-w-[45%]" title={s.layoutIntent}>
                                          {s.layoutIntent}
                                        </div>
                                      )}
                                    </div>
                                    {(s.bullets?.length || s.footerNote) && (
                                      <div className="px-2.5 py-2">
                                        {s.bullets?.length ? (
                                          <ul className="space-y-1">
                                            {s.bullets.slice(0, 8).map((b, bi) => (
                                              <li key={bi} className="text-[12px] text-[#d4d4d4] leading-relaxed flex items-start gap-1.5">
                                                <span className="text-[#858585] mt-0.5">•</span>
                                                <span className="flex-1">{b}</span>
                                              </li>
                                            ))}
                                          </ul>
                                        ) : null}
                                        {s.footerNote && (
                                          <div className="mt-2 text-[10px] text-[#858585] border-t border-[#2d2d2d] pt-2">
                                            页脚：{s.footerNote}
                                          </div>
                                        )}
                                      </div>
                                    )}
                                  </div>
                                ))}

                                {jsonOpen && (
                                  <pre className="mt-2 bg-[#0f0f10] border border-[#2d2d2d] rounded-md p-2 text-[11px] text-[#d4d4d4] overflow-x-auto">
                                    {parsed.rawJson}
                                  </pre>
                                )}
                              </div>
                            </div>
                          )}

                          {cleanedText && (
                    <ReactMarkdown
                      components={{
                        h1: ({children}) => <h1 className="text-[15px] font-semibold text-[#e1e1e1] mt-3 mb-2 pb-1 border-b border-[#3c3c3c]">{children}</h1>,
                        h2: ({children}) => <h2 className="text-[14px] font-semibold text-[#e1e1e1] mt-3 mb-1.5 flex items-center gap-1.5">{children}</h2>,
                        h3: ({children}) => <h3 className="text-[13px] font-medium text-[#cccccc] mt-2 mb-1">{children}</h3>,
                        p: ({children}) => <p className="mb-2 last:mb-0">{children}</p>,
                        ul: ({children}) => <ul className="list-none ml-0 mb-2 space-y-1">{children}</ul>,
                        ol: ({children}) => <ol className="list-decimal ml-4 mb-2 space-y-1">{children}</ol>,
                        li: ({children}) => <li className="text-[13px] leading-relaxed flex items-start gap-1.5"><span className="text-[#858585] mt-0.5">•</span><span className="flex-1">{children}</span></li>,
                        strong: ({children}) => <strong className="font-semibold text-[#e1e1e1]">{children}</strong>,
                        em: ({children}) => <em className="italic text-[#9cdcfe]">{children}</em>,
                        code: ({children, className}) => {
                          const isBlock = className?.includes('language-')
                          if (isBlock) {
                            return <code className="block bg-[#1e1e1e] text-[#ce9178] p-2 rounded text-[12px] font-mono overflow-x-auto my-2">{children}</code>
                          }
                          return <code className="bg-[#1e1e1e] text-[#ce9178] px-1 py-0.5 rounded text-[12px] font-mono">{children}</code>
                        },
                        pre: ({children}) => <pre className="bg-[#1e1e1e] rounded-md overflow-hidden my-2">{children}</pre>,
                        a: ({href, children}) => <a href={href} className="text-[#75beff] hover:underline" target="_blank" rel="noopener noreferrer">{children}</a>,
                        blockquote: ({children}) => <blockquote className="border-l-2 border-[#0e639c] pl-3 my-2 text-[#9a9a9a] italic">{children}</blockquote>,
                        hr: () => <hr className="border-[#3c3c3c] my-3" />,
                      }}
                    >
                              {cleanedText}
                    </ReactMarkdown>
                          )}
                        </>
                      )
                    })()}
                  </div>
                </div>
                <span className="text-[10px] text-[#5a5a5a] mt-1 block pl-1">
                  {message.timestamp.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
                </span>
              </div>
            )}
          </motion.div>
        ))}
        </AnimatePresence>

        {/* 流式输出 - 实时显示 AI 响应 (使用 Framer Motion) */}
        <AnimatePresence mode="wait">
          {isLoading && (
            <motion.div 
              className="w-full"
              layout
              variants={streamingVariants}
              initial="hidden"
              animate="visible"
              exit="exit"
            >
              <div className="bg-[#252526] border border-[#2d2d2d] rounded-lg rounded-tl-sm px-3 py-2">
                <motion.div className="streaming-container" layout>
                  <CinematicTyper text={streamingContent} isStreaming={isLoading} />
                </motion.div>
              </div>
              {/* 状态指示 */}
              <div className="flex items-center gap-1.5 mt-1.5 pl-1">
                <div className="flex gap-0.5">
                  <span className="w-1 h-1 rounded-full bg-violet-400 animate-pulse" style={{ animationDelay: '0ms' }} />
                  <span className="w-1 h-1 rounded-full bg-violet-400 animate-pulse" style={{ animationDelay: '150ms' }} />
                  <span className="w-1 h-1 rounded-full bg-violet-400 animate-pulse" style={{ animationDelay: '300ms' }} />
                </div>
                <span className="text-[10px] text-[#5a5a5a]">AI 正在生成...</span>
              </div>
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
              <div className="bg-[#252526] border border-[#2d2d2d] rounded-lg px-3 py-2">
                <div className="flex items-center gap-2">
                  <Loader2 className="w-3.5 h-3.5 text-violet-400 animate-spin flex-shrink-0" />
                  <span className="text-[12px] text-[#cccccc] flex-1 truncate">
                    {agentProgress.currentAction}
                  </span>
                  {agentProgress.thinkingTime > 0 && (
                    <span className="text-[10px] text-[#858585] flex-shrink-0">
                      {agentProgress.thinkingTime}s
                    </span>
                  )}
                </div>
                {toolActivity.length > 0 && (
                  <div className="mt-2 border-t border-[#2d2d2d] pt-2">
                    <div className="text-[10px] text-[#5a5a5a] uppercase tracking-wider mb-1">工具调用</div>
                    <div className="space-y-1">
                      {toolActivity.slice(-4).map(activity => (
                        <div key={activity.id} className="flex items-center gap-1.5 text-[11px] text-[#cccccc]">
                          {activity.status === 'running' ? (
                            <Loader2 className="w-3 h-3 text-violet-400 animate-spin flex-shrink-0" />
                          ) : activity.status === 'success' ? (
                            <CheckCircle2 className="w-3 h-3 text-[#4ec9b0] flex-shrink-0" />
                          ) : (
                            <X className="w-3 h-3 text-[#f14c4c] flex-shrink-0" />
                          )}
                          <span className="truncate flex-1">{activity.label}</span>
                          {activity.detail && (
                            <span className="text-[10px] text-[#858585] flex-shrink-0">{activity.detail}</span>
                          )}
                        </div>
                      ))}
                    </div>
                  </div>
                )}
              </div>
            </motion.div>
          )}
        </AnimatePresence>

        <div ref={messagesEndRef} />
      </div>

      {/* 上下文文件显示 - Cursor 风格 */}
      <div className="px-3 py-2 border-t border-[#2d2d2d] bg-[#252526]">
        <div className="flex items-center gap-1.5 flex-wrap">
          <span className="text-[10px] text-[#858585]">上下文:</span>
          
          {/* 当前编辑的文档 */}
          {currentFile && (
            <div className="flex items-center gap-1 px-1.5 py-0.5 bg-[#1e3a29] text-[#4ec9b0] text-[10px] rounded">
              <FileText className="w-2.5 h-2.5" />
              <span className="max-w-[80px] truncate">{currentFile.name}</span>
            </div>
          )}
          
          {/* 用户拖拽的附加文件 */}
          {attachedFiles.map((file) => (
            <div 
              key={file.path}
              className="flex items-center gap-1 px-1.5 py-0.5 bg-[#0e639c]/30 text-[#75beff] text-[10px] rounded"
            >
              <FileText className="w-2.5 h-2.5" />
              <span className="max-w-[60px] truncate">{file.name}</span>
              <button onClick={() => removeAttachedFile(file.path)} className="hover:bg-[#0e639c]/50 rounded p-0.5 -mr-0.5">
                <X className="w-2.5 h-2.5" />
              </button>
            </div>
          ))}
          
          {!currentFile && attachedFiles.length === 0 && (
            <span className="text-[10px] text-[#5a5a5a]">拖拽文件添加上下文</span>
          )}
        </div>
      </div>

      {/* AI 处理中状态指示器 - Cursor 风格 */}
      {isLoading && (
        <div className="px-3 py-2 border-t border-[#2d2d2d] bg-[#1e1e1e]">
          <div className="flex items-center gap-2">
            <div className="relative w-5 h-5">
              <div className="absolute inset-0 rounded-full border border-violet-500/30"></div>
              <div className="absolute inset-0 rounded-full border border-transparent border-t-violet-500 animate-spin"></div>
            </div>
            <div className="flex-1 min-w-0">
              <span className="text-[12px] text-[#cccccc]">
                {agentProgress.currentAction || '正在处理...'}
              </span>
            </div>
            {agentProgress.thinkingTime > 0 && (
              <span className="text-[10px] text-[#858585] flex-shrink-0">
                {agentProgress.thinkingTime}s
              </span>
            )}
          </div>
          
          {/* 进度步骤 - 更紧凑 */}
          {agentProgress.steps.length > 0 && (
            <div className="mt-2 pl-7 space-y-0.5">
              {agentProgress.steps.map((step) => (
                <div key={step.id} className="flex items-center gap-1.5">
                  {step.status === 'completed' ? (
                    <CheckCircle2 className="w-3 h-3 text-[#4ec9b0]" />
                  ) : step.status === 'running' ? (
                    <Loader2 className="w-3 h-3 text-violet-400 animate-spin" />
                  ) : (
                    <Circle className="w-3 h-3 text-[#5a5a5a]" />
                  )}
                  <span className={`text-[11px] ${
                    step.status === 'completed' ? 'text-[#858585]' :
                    step.status === 'running' ? 'text-[#cccccc]' : 'text-[#5a5a5a]'
                  }`}>
                    {step.description}
                  </span>
                </div>
              ))}
            </div>
          )}
          
          {toolActivity.length > 0 && (
            <div className="mt-2 pl-7 space-y-0.5">
              <div className="text-[10px] text-[#5a5a5a] uppercase tracking-wider">工具调用</div>
              {toolActivity.slice(-4).map(activity => (
                <div key={activity.id} className="flex items-center gap-1.5 text-[11px] text-[#cccccc]">
                  {activity.status === 'running' ? (
                    <Loader2 className="w-3 h-3 text-violet-400 animate-spin" />
                  ) : activity.status === 'success' ? (
                    <CheckCircle2 className="w-3 h-3 text-[#4ec9b0]" />
                  ) : (
                    <X className="w-3 h-3 text-[#f14c4c]" />
                  )}
                  <span className="truncate flex-1">{activity.label}</span>
                  {activity.detail && (
                    <span className="text-[10px] text-[#858585]">{activity.detail}</span>
                  )}
                </div>
              ))}
            </div>
          )}
        </div>
      )}

      {/* 快捷命令提示 - Cursor 风格 */}
      {input.startsWith('/') && !isLoading && (
        <div className="px-3 py-2 border-t border-[#2d2d2d] bg-[#252526]">
          <div className="space-y-0.5">
            {[
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
                className="w-full flex items-center justify-between px-2 py-1.5 hover:bg-[#2d2d2d] rounded text-left"
              >
                <span className="text-[12px] text-violet-400">{item.cmd}</span>
                <span className="text-[10px] text-[#858585]">{item.desc}</span>
              </button>
            ))}
          </div>
        </div>
      )}

      {/* Word 格式操作确认条（dryRun → apply） */}
      {pendingWordOps && !isLoading && (
        <div className="px-3 py-2 border-t border-[#2d2d2d] bg-[#252526]">
          <div className="flex items-center gap-2">
            <div className="flex-1 min-w-0">
              <div className="text-[12px] text-[#cccccc] truncate">
                {pendingWordOps.previewMessage || '已生成格式修改预览'}
              </div>
              <div className="text-[10px] text-[#858585] truncate">
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
              className="flex items-center gap-1.5 px-2.5 py-1.5 bg-gradient-to-b from-[#0e639c]/35 to-[#0e639c]/20 border border-[#0e639c]/35 hover:from-[#0e639c]/45 hover:to-[#0e639c]/25 text-[#e6f1ff] text-[11px] rounded-md transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
              title="应用修订"
            >
              <CheckCircle2 className="w-3.5 h-3.5" />
              应用修订
            </button>
            <button
              disabled={wordOpsApplying}
              onClick={() => setPendingWordOps(null)}
              className="p-1.5 rounded-md text-[#858585] hover:text-[#cccccc] hover:bg-[#2d2d2d] transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
              title="取消"
            >
              <X className="w-4 h-4" />
            </button>
          </div>
        </div>
      )}

      {/* PPT 大纲确认条（阶段1 → 阶段2） */}
      {pendingPptOutline && !pptGenerating && (
        <div className="px-3 py-2 border-t border-[#2d2d2d] bg-[#252526]">
          <div className="flex items-center gap-2">
            <div className="flex-1 min-w-0">
              <div className="text-[12px] text-[#cccccc] truncate">
                已检测到 PPT 大纲：{pendingPptOutline.draft.title || '未命名'}（{pendingPptOutline.draft.slides?.length || 0} 页）
              </div>
              <div className="text-[10px] text-[#858585] truncate">
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
              className="flex items-center gap-1.5 px-2.5 py-1.5 bg-gradient-to-b from-[#0e639c]/35 to-[#0e639c]/20 border border-[#0e639c]/35 hover:from-[#0e639c]/45 hover:to-[#0e639c]/25 text-[#e6f1ff] text-[11px] rounded-md transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
              title="确认大纲并开始生成 PPT"
            >
              <CheckCircle2 className="w-3.5 h-3.5" />
              确认并开始生成
            </button>
            <button
              disabled={isLoading || pptGenerating}
              onClick={() => setPendingPptOutline(null)}
              className="p-1.5 rounded-md text-[#858585] hover:text-[#cccccc] hover:bg-[#2d2d2d] transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
              title="关闭提示"
            >
              <X className="w-4 h-4" />
            </button>
          </div>
        </div>
      )}

      {/* PPT 编辑反馈输入区域 */}
      {pptEditPending && !pptGenerating && (
        <div className="px-3 py-2 border-t border-[#2d2d2d] bg-[#252526]">
          <div className="flex flex-col gap-2">
            <div className="flex items-center gap-2">
              <div className="flex-1 min-w-0">
                <div className="text-[12px] text-[#cccccc]">
                  {pptEditPending.mode === 'regenerate' ? '🔄 整页重做' : '🎨 局部编辑'}：
                  {pptEditPending.pageNumbers.length === 1 
                    ? `第 ${pptEditPending.pageNumbers[0]} 页`
                    : `${pptEditPending.pageNumbers.length} 页（${pptEditPending.pageNumbers.join(', ')}）`
                  }
                </div>
                <div className="text-[10px] text-[#858585]">
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
                className="p-1.5 rounded-md text-[#858585] hover:text-[#cccccc] hover:bg-[#2d2d2d] transition-colors"
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
                className="flex-1 bg-[#2d2d2d] border border-[#3c3c3c] rounded-md px-3 py-1.5 text-[12px] text-[#d4d4d4] placeholder-[#5a5a5a] focus:outline-none focus:border-[#0e639c]"
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
                className="flex items-center gap-1.5 px-3 py-1.5 bg-gradient-to-b from-[#0e639c]/35 to-[#0e639c]/20 border border-[#0e639c]/35 hover:from-[#0e639c]/45 hover:to-[#0e639c]/25 text-[#e6f1ff] text-[11px] rounded-md transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
              >
                <Send className="w-3.5 h-3.5" />
                开始{pptEditPending.mode === 'regenerate' ? '重做' : '编辑'}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* 输入区域 - Cursor 风格 + PPT 拖拽支持 */}
      <div 
        className={`p-3 bg-[#1e1e1e] border-t transition-colors ${
          isPptDragOver ? 'border-[#0e639c] bg-[#0e639c]/10' : 'border-[#2d2d2d]'
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
          <div className="mb-2 p-2 bg-[#2d2d2d] rounded-lg border border-[#3c3c3c] flex items-start gap-3">
            <div className="relative flex-shrink-0">
              <img
                src={`data:image/png;base64,${pptEditContext.imageBase64}`}
                alt={`第 ${pptEditContext.pageNumber} 页${pptEditContext.isRegion ? '（框选区域）' : ''}`}
                className="w-[100px] h-[62px] object-contain rounded border border-[#4a4a4a] bg-black"
              />
              <div className="absolute -top-1 -left-1 bg-[#0e639c] text-[9px] text-white px-1.5 py-0.5 rounded">
                {pptEditContext.isRegion ? '框选' : `第 ${pptEditContext.pageNumber} 页`}
              </div>
            </div>
            <div className="flex-1 min-w-0">
              <div className="text-[11px] text-[#cccccc] mb-1">
                {pptEditContext.isRegion ? (
                  <>已框选第 <span className="text-[#0e639c] font-medium">{pptEditContext.pageNumber}</span> 页的区域</>
                ) : (
                  <>已选择第 <span className="text-[#0e639c] font-medium">{pptEditContext.pageNumber}</span> 页</>
                )}
              </div>
              <div className="text-[10px] text-[#888]">
                输入修改要求，AI 将自动判断是整页重做还是局部调整
              </div>
            </div>
            <button
              onClick={() => setPptEditContext(null)}
              className="p-1 text-[#888] hover:text-white hover:bg-[#3c3c3c] rounded transition-colors"
              title="移除"
            >
              <X className="w-3.5 h-3.5" />
            </button>
          </div>
        )}
        
        {/* 拖拽提示 */}
        {isPptDragOver && (
          <div className="mb-2 p-3 border-2 border-dashed border-[#0e639c] rounded-lg bg-[#0e639c]/10 text-center">
            <div className="text-[12px] text-[#0e639c]">松开鼠标，将 PPT 页面添加到对话</div>
          </div>
        )}
        
        <div className="relative">
          <textarea
            ref={inputRef}
            value={input}
            onChange={(e) => setInput(e.target.value)}
            onKeyDown={handleKeyDown}
            placeholder={
              pptEditContext 
                ? `描述如何修改第 ${pptEditContext.pageNumber} 页...` 
                : isLoading 
                  ? "AI 正在处理中..." 
                  : "输入消息或 / 查看命令..."
            }
            className={`w-full bg-[#2d2d2d] border rounded-lg pl-3 pr-10 py-2.5 text-[13px] text-[#d4d4d4] placeholder-[#5a5a5a] focus:outline-none transition-colors resize-none scrollbar-none ${
              isLoading ? 'border-violet-500/30' : pptEditContext ? 'border-[#0e639c]/50 focus:border-[#0e639c]' : 'border-[#3c3c3c] focus:border-[#0e639c]'
            }`}
            rows={2}
            disabled={isLoading}
          />
          <button
            onClick={handleSend}
            disabled={isLoading || !input.trim()}
            className={`absolute right-2 bottom-2 p-1.5 rounded-md transition-colors disabled:cursor-not-allowed ${
              isLoading 
                ? 'text-violet-400' 
                : 'text-[#858585] hover:text-[#cccccc] hover:bg-[#3c3c3c] disabled:opacity-30'
            }`}
          >
            {isLoading ? (
              <Loader2 className="w-4 h-4 animate-spin" />
            ) : (
              <Send className="w-4 h-4" />
            )}
          </button>
        </div>
        
        <p className="text-[10px] text-[#5a5a5a] text-center mt-1.5">
          {isLoading ? (
            <span className="text-violet-400">处理中...</span>
          ) : pptEditContext ? (
            <span className="text-[#0e639c]">输入修改要求后按 Enter 发送</span>
          ) : (
            <>按 <kbd className="px-1 py-0.5 bg-[#2d2d2d] rounded text-[9px]">Enter</kbd> 发送 · <span className="text-violet-400">/</span> 快捷命令 · 拖拽 PPT 页面到此处编辑</>
          )}
        </p>
      </div>
    </div>
  )
}
