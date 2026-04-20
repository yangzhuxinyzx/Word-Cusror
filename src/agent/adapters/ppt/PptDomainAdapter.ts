import { BUILTIN_KEYS } from '../../../config/models'
import type {
  AISettings,
  FileItem,
  PptBlendStrategy,
  PptCleanupStrategy,
  PptDetectedTextBox,
  PptEditCandidate,
  PptTextEditOperation,
} from '../../../types'

export interface PptEditContextSnapshot {
  pageNumber: number
  imageBase64: string
  regionRect?: { x: number; y: number; w: number; h: number }
  pptxPath?: string
  isRegion?: boolean
}

export interface CreatePptDeckParams {
  title: string
  theme: string
  style: string
  outline: string
  outputPath: string
  slideCount?: number
  pendingImages: string[]
  settings: AISettings
}

export interface EditPptDeckParams {
  pptxPath: string
  pageNumber?: number
  pageNumbers?: number[]
  mode: 'regenerate' | 'partial_edit'
  feedback: string
  regionRect?: { x: number; y: number; w: number; h: number }
  regionScreenshot?: string
  settings: AISettings
}

export interface PptDomainAdapterOptions {
  isElectron: boolean
  currentFilePath: string | null
  currentFileName: string | null
  workspacePath: string | null
  refreshFiles: () => Promise<void>
  openFile: (file: FileItem) => Promise<void>
}

function sanitizeFilename(title: string): string {
  const safe = String(title).replace(/[<>:"/\\|?*]/g, '_').slice(0, 60).trim()
  return safe || 'presentation'
}

function getPathSeparator(dir: string): '/' | '\\' {
  return dir.includes('\\') && !dir.includes('/') ? '\\' : '/'
}

function dirnameOf(filePath: string): string | null {
  const normalized = String(filePath || '').trim()
  if (!normalized) return null
  const match = normalized.match(/^(.*)[\\/][^\\/]+$/)
  return match?.[1] || null
}

export class PptDomainAdapter {
  constructor(private readonly options: PptDomainAdapterOptions) {}

  canUsePptTools(): boolean {
    return !!(
      this.options.isElectron &&
      window.electronAPI?.pptGenerateDeck &&
      window.electronAPI?.pptEditSlides
    )
  }

  resolveOutputDir(): string | null {
    if (this.options.currentFilePath) {
      const dir = dirnameOf(this.options.currentFilePath)
      if (dir) return dir
    }
    return this.options.workspacePath || null
  }

  buildOutputPath(title: string): { fileName: string; outputPath: string } | null {
    const dir = this.resolveOutputDir()
    if (!dir) return null
    const safeTitle = sanitizeFilename(title)
    const fileName = safeTitle.toLowerCase().endsWith('.pptx')
      ? safeTitle
      : `${safeTitle}.pptx`
    const separator = getPathSeparator(dir)
    return {
      fileName,
      outputPath: `${dir}${separator}${fileName}`,
    }
  }

  async generatePromptSlides(params: CreatePptDeckParams) {
    if (!window.electronAPI?.openrouterGeminiPptPrompts) {
      return {
        success: false,
        error: 'Missing openrouterGeminiPptPrompts API.',
      }
    }
    const slideCountMatch = params.outline.match(/\u7b2c\s*(\d+)\s*\u9875/g)
    const estimatedSlideCount =
      params.slideCount && params.slideCount > 0
        ? params.slideCount
        : slideCountMatch
          ? slideCountMatch.length
          : 3
    const openRouterApiKey =
      params.settings?.openRouterApiKey || BUILTIN_KEYS.linapiKey
    return window.electronAPI.openrouterGeminiPptPrompts({
      apiKey: openRouterApiKey,
      outline: params.outline,
      slideCount: estimatedSlideCount,
      theme: params.theme,
      style: params.style,
      styleImages: params.pendingImages.length > 0 ? [...params.pendingImages] : undefined,
      mainApiKey: params.settings?.apiKey || '',
      mainBaseUrl: params.settings?.baseUrl || '',
      mainModel: params.settings?.model || '',
    })
  }

  async generateDeck(params: {
    outputPath: string
    slides: Array<{ prompt: string; negativePrompt: string }>
    designConcept?: string
    colorPalette?: string
    outline?: unknown
    settings: AISettings
  }) {
    const openRouterApiKey =
      params.settings?.openRouterApiKey || BUILTIN_KEYS.linapiKey
    const dashscopeApiKey =
      params.settings?.dashscopeApiKey || BUILTIN_KEYS.dashscopeApiKey
    const pptImageModel = params.settings?.pptImageModel || 'z-image-turbo'
    const imageSize =
      pptImageModel === 'z-image-turbo' ? '2048*1152' : '1664*928'
    return window.electronAPI!.pptGenerateDeck({
      outputPath: params.outputPath,
      slides: params.slides,
      mainApiKey: params.settings?.apiKey || '',
      dashscope: {
        apiKey: dashscopeApiKey,
        region: 'cn',
        size: imageSize,
        model: pptImageModel,
        promptExtend: false,
        watermark: false,
        negativePromptDefault:
          'watermark, logo, brand name text, badge, QR code, UI, screenshot, HUD, sci-fi interface, holographic UI, futuristic dashboard, neon cyberpunk, neon cyan, bright cyan, fluorescent cyan, neon teal, cheap turquoise, generic isometric city, isometric cityscape, circuit-board city, lowres, blurry, garbled Chinese, wrong characters, text distortion, misspelling, random letters, gibberish, extra text, english text, ugly typography, amateur layout, noisy background, oversaturated, cheap plastic, toy-like, glossy, harsh specular, overbloom, stock 3d icons, generic template, ai artifacts, uncanny',
      },
      postprocess: { mode: 'letterbox' },
      repair: {
        enabled: !!openRouterApiKey,
        openRouterApiKey,
        model: 'google/gemini-3-pro-preview',
        maxAttempts: 2,
        deckContext: {
          designConcept: params.designConcept || '',
          colorPalette: params.colorPalette || '',
        },
      },
      outline: params.outline,
    })
  }

  async editSlides(params: EditPptDeckParams) {
    const openRouterApiKey =
      params.settings?.openRouterApiKey || BUILTIN_KEYS.linapiKey
    const dashscopeApiKey =
      params.settings?.dashscopeApiKey || BUILTIN_KEYS.dashscopeApiKey
    const pageNumbers =
      params.pageNumbers && params.pageNumbers.length > 0
        ? params.pageNumbers
        : typeof params.pageNumber === 'number'
          ? [params.pageNumber]
          : []
    return window.electronAPI!.pptEditSlides({
      pptxPath: params.pptxPath,
      pageNumbers,
      mode: params.mode,
      feedback: params.feedback,
      regionScreenshot: params.regionScreenshot,
      regionRect: params.regionRect,
      openRouterApiKey,
      dashscopeApiKey,
      mainApiKey: params.settings?.apiKey || '',
      pptImageModel: params.settings?.pptImageModel || 'z-image-turbo',
    })
  }

  async textEditHealth(options?: { bootstrap?: boolean }) {
    if (!window.electronAPI?.pptTextEditHealth) {
      return { success: false, error: 'Missing pptTextEditHealth API.' }
    }
    return window.electronAPI.pptTextEditHealth(options || {})
  }

  async detectTextLayer(params: {
    pptxPath: string
    pageNumber: number
    useCache?: boolean
    cacheOnly?: boolean
  }): Promise<{
    success: boolean
    cached?: boolean
    cacheVersion?: 'v2'
    canvasWidth?: number
    canvasHeight?: number
    boxes?: PptDetectedTextBox[]
    sourceImagePath?: string
    error?: string
  }> {
    if (!window.electronAPI?.pptDetectTextLayer) {
      return { success: false, error: 'Missing pptDetectTextLayer API.' }
    }
    return window.electronAPI.pptDetectTextLayer(params)
  }

  async applyTextEdits(params: {
    pptxPath: string
    pageNumber: number
    edits: PptTextEditOperation[]
  }): Promise<{
    success: boolean
    path?: string
    imageDataUrl?: string
    editedBoxes?: string[]
    appliedCandidateId?: string
    candidateCount?: number
    fontMatchConfidence?: number
    cleanupStrategy?: PptCleanupStrategy
    blendStrategy?: PptBlendStrategy
    candidates?: PptEditCandidate[]
    perBoxCandidates?: Record<string, PptEditCandidate[]>
    logs?: Array<{ boxId?: string; success: boolean; error?: string }>
    fallbackSuggested?: boolean
    error?: string
  }> {
    if (!window.electronAPI?.pptApplyTextEdits) {
      return { success: false, error: 'Missing pptApplyTextEdits API.' }
    }
    return window.electronAPI.pptApplyTextEdits(params)
  }

  async refreshFiles(): Promise<void> {
    await this.options.refreshFiles()
  }

  async openGeneratedDeck(fileName: string, filePath: string): Promise<void> {
    await this.options.openFile({
      name: fileName,
      path: filePath,
      type: 'file',
    })
  }

  async reopenCurrentDeckIfNeeded(pptxPath: string): Promise<void> {
    if (
      this.options.currentFilePath === pptxPath &&
      this.options.currentFileName
    ) {
      await this.options.openFile({
        name: this.options.currentFileName,
        path: pptxPath,
        type: 'file',
      })
    }
  }
}
