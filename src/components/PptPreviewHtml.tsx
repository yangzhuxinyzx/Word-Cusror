import { useEffect, useMemo, useState, useCallback, useRef } from 'react'
import { ChevronLeft, ChevronRight, Loader2, Maximize2, Minimize2, FileWarning, RefreshCw, Paintbrush, CheckSquare, Square, Move } from 'lucide-react'
import JSZip from 'jszip'
import html2canvas from 'html2canvas'
import type {
  PptDetectedTextBox,
  PptEditCandidate,
  PptTextEditOperation,
  PptTextEditStyleOverride,
} from '../types'

type PptPreviewHtmlProps = {
  title: string
  pptxBase64: string
  pptxPath?: string // PPTX 文件路径（用于编辑）
  onEditRequest?: (options: {
    pageNumbers: number[]
    mode: 'regenerate' | 'partial_edit'
  }) => void // 编辑请求回调
}

function base64ToArrayBuffer(base64: string) {
  const binaryString = atob(base64)
  const len = binaryString.length
  const bytes = new Uint8Array(len)
  for (let i = 0; i < len; i++) bytes[i] = binaryString.charCodeAt(i)
  return bytes.buffer
}

export default function PptPreviewHtml({ title, pptxBase64, pptxPath, onEditRequest }: PptPreviewHtmlProps) {
  const [error, setError] = useState<string | null>(null)
  const [loading, setLoading] = useState(true)

  const [activeIndex, setActiveIndex] = useState(0)
  const [scale, setScale] = useState(100)
  const [isFullscreen, setIsFullscreen] = useState(false)
  const [slideCount, setSlideCount] = useState(0)
  const [slideImages, setSlideImages] = useState<(string | null)[]>([])
  const [textEditMode, setTextEditMode] = useState(false)
  const [textEditHealth, setTextEditHealth] = useState<{
    ready: boolean
    message: string
  }>({ ready: false, message: '' })
  const [textEditProgress, setTextEditProgress] = useState<{
    active: boolean
    stage?: string
    progress?: number
    message?: string
    completedCandidates?: number
    totalCandidates?: number
  }>({ active: false })
  const [textLayerLoading, setTextLayerLoading] = useState(false)
  const [textLayerError, setTextLayerError] = useState<string | null>(null)
  const [textLayerSize, setTextLayerSize] = useState<{ width: number; height: number } | null>(null)
  const [textBoxes, setTextBoxes] = useState<PptDetectedTextBox[]>([])
  const [activeTextBoxId, setActiveTextBoxId] = useState<string | null>(null)
  const [textDrafts, setTextDrafts] = useState<Record<string, string>>({})
  const [expertMode, setExpertMode] = useState(false)
  const [boxOverrides, setBoxOverrides] = useState<Record<string, PptTextEditStyleOverride>>({})
  const [lastCandidates, setLastCandidates] = useState<Record<string, PptEditCandidate[]>>({})
  const [applyingTextEdits, setApplyingTextEdits] = useState(false)
  
  // 多选状态
  const [isMultiSelectMode, setIsMultiSelectMode] = useState(false)
  const [selectedPages, setSelectedPages] = useState<Set<number>>(new Set())
  
  // 切换页面选中状态
  const togglePageSelection = useCallback((pageIndex: number) => {
    setSelectedPages((prev) => {
      const next = new Set(prev)
      if (next.has(pageIndex)) {
        next.delete(pageIndex)
      } else {
        next.add(pageIndex)
      }
      return next
    })
  }, [])
  
  // 全选/取消全选
  const toggleSelectAll = useCallback(() => {
    if (selectedPages.size === slideCount) {
      setSelectedPages(new Set())
    } else {
      setSelectedPages(new Set(Array.from({ length: slideCount }, (_, i) => i)))
    }
  }, [selectedPages.size, slideCount])
  
  // 退出多选模式
  const exitMultiSelectMode = useCallback(() => {
    setIsMultiSelectMode(false)
    setSelectedPages(new Set())
  }, [])
  
  // 发起编辑请求
  const handleEditRequest = useCallback((mode: 'regenerate' | 'partial_edit') => {
    const pagesToEdit = isMultiSelectMode && selectedPages.size > 0
      ? Array.from(selectedPages).map((i) => i + 1).sort((a, b) => a - b)
      : [activeIndex + 1]
    
    if (onEditRequest) {
      onEditRequest({ pageNumbers: pagesToEdit, mode })
    } else {
      // 触发自定义事件，由 ChatPanel 捕获处理
      window.dispatchEvent(new CustomEvent('ppt-edit-request', {
        detail: {
          pptxPath,
          pageNumbers: pagesToEdit,
          mode,
        }
      }))
    }
    
    // 退出多选模式
    exitMultiSelectMode()
  }, [isMultiSelectMode, selectedPages, activeIndex, onEditRequest, pptxPath, exitMultiSelectMode])
  
  // 待跳转的页码（编辑完成后跳转）
  const [pendingJumpPage, setPendingJumpPage] = useState<number | null>(null)
  
  // ========== 框选功能状态 ==========
  const [isSelecting, setIsSelecting] = useState(false)
  const [selectionStart, setSelectionStart] = useState<{ x: number; y: number } | null>(null)
  const [selectionRect, setSelectionRect] = useState<{ x: number; y: number; w: number; h: number } | null>(null)
  const mainCanvasRef = useRef<HTMLDivElement>(null)
  const slideContainerRef = useRef<HTMLDivElement>(null)
  
  // 框选开始
  const handleSelectionMouseDown = useCallback((e: React.MouseEvent) => {
    // 只有按住 Ctrl 才启用框选
    if (!e.ctrlKey || !slideContainerRef.current) return
    
    e.preventDefault()
    const rect = slideContainerRef.current.getBoundingClientRect()
    const x = e.clientX - rect.left
    const y = e.clientY - rect.top
    
    setIsSelecting(true)
    setSelectionStart({ x, y })
    setSelectionRect({ x, y, w: 0, h: 0 })
  }, [])
  
  // 框选移动
  const handleSelectionMouseMove = useCallback((e: React.MouseEvent) => {
    if (!isSelecting || !selectionStart || !slideContainerRef.current) return
    
    const rect = slideContainerRef.current.getBoundingClientRect()
    const currentX = Math.max(0, Math.min(e.clientX - rect.left, rect.width))
    const currentY = Math.max(0, Math.min(e.clientY - rect.top, rect.height))
    
    const x = Math.min(selectionStart.x, currentX)
    const y = Math.min(selectionStart.y, currentY)
    const w = Math.abs(currentX - selectionStart.x)
    const h = Math.abs(currentY - selectionStart.y)
    
    setSelectionRect({ x, y, w, h })
  }, [isSelecting, selectionStart])
  
  // 框选结束 - 截图并触发事件
  const handleSelectionMouseUp = useCallback(async () => {
    if (!isSelecting || !selectionRect || !slideContainerRef.current) {
      setIsSelecting(false)
      setSelectionStart(null)
      setSelectionRect(null)
      return
    }
    
    // 如果框选区域太小，忽略
    if (selectionRect.w < 20 || selectionRect.h < 20) {
      setIsSelecting(false)
      setSelectionStart(null)
      setSelectionRect(null)
      return
    }
    
    try {
      // 使用 html2canvas 截取整个 slide 容器
      const canvas = await html2canvas(slideContainerRef.current, {
        useCORS: true,
        allowTaint: true,
        backgroundColor: '#000',
        scale: 1,
      })
      
      // 从 canvas 中裁剪出框选区域
      const croppedCanvas = document.createElement('canvas')
      const ctx = croppedCanvas.getContext('2d')
      if (ctx) {
        // 计算实际裁剪区域（考虑缩放）
        const scaleRatio = canvas.width / slideContainerRef.current.offsetWidth
        const cropX = selectionRect.x * scaleRatio
        const cropY = selectionRect.y * scaleRatio
        const cropW = selectionRect.w * scaleRatio
        const cropH = selectionRect.h * scaleRatio
        
        croppedCanvas.width = cropW
        croppedCanvas.height = cropH
        ctx.drawImage(canvas, cropX, cropY, cropW, cropH, 0, 0, cropW, cropH)
        
        const regionBase64 = croppedCanvas.toDataURL('image/png').split(',')[1]
        
        // 触发自定义事件，通知 ChatPanel
        window.dispatchEvent(new CustomEvent('ppt-region-selected', {
          detail: {
            pageNumber: activeIndex + 1,
            regionBase64,
            regionRect: selectionRect,
            fullPageBase64: slideImages[activeIndex]?.split(',')[1] || '',
            pptxPath,
          }
        }))
      }
    } catch (err) {
      console.error('框选截图失败:', err)
    }
    
    setIsSelecting(false)
    setSelectionStart(null)
    setSelectionRect(null)
  }, [isSelecting, selectionRect, activeIndex, slideImages, pptxPath])
  
  // ========== 缩略图拖拽功能 ==========
  const handleThumbnailDragStart = useCallback((e: React.DragEvent, pageIndex: number) => {
    const img = slideImages[pageIndex]
    if (!img) return
    
    // 设置拖拽数据
    e.dataTransfer.setData('application/ppt-page', JSON.stringify({
      pageNumber: pageIndex + 1,
      imageBase64: img.split(',')[1] || '',
      pptxPath,
    }))
    e.dataTransfer.effectAllowed = 'copy'
    
    // 设置拖拽预览图
    const dragImage = document.createElement('div')
    dragImage.style.cssText = 'position:absolute;top:-9999px;left:-9999px;width:120px;height:75px;background:#333;border-radius:4px;display:flex;align-items:center;justify-content:center;color:#fff;font-size:12px;'
    dragImage.textContent = `第 ${pageIndex + 1} 页`
    document.body.appendChild(dragImage)
    e.dataTransfer.setDragImage(dragImage, 60, 37)
    setTimeout(() => dragImage.remove(), 0)
  }, [slideImages, pptxPath])
  
  // 监听跳转事件
  useEffect(() => {
    const handleJumpToPage = (event: CustomEvent<{ pageNumber: number }>) => {
      const { pageNumber } = event.detail
      // 保存待跳转页码，等 PPTX 加载完成后执行
      setPendingJumpPage(pageNumber)
    }
    
    window.addEventListener('ppt-jump-to-page', handleJumpToPage as EventListener)
    return () => {
      window.removeEventListener('ppt-jump-to-page', handleJumpToPage as EventListener)
    }
  }, [])
  
  // 当 slideImages 加载完成且有待跳转页码时，执行跳转
  useEffect(() => {
    if (pendingJumpPage !== null && slideImages.length > 0 && !loading) {
      const targetIndex = pendingJumpPage - 1
      if (targetIndex >= 0 && targetIndex < slideImages.length) {
        setActiveIndex(targetIndex)
      }
      setPendingJumpPage(null)
    }
  }, [pendingJumpPage, slideImages.length, loading])

  const currentSlideImage = useMemo(() => {
    const idx = Math.min(Math.max(activeIndex, 0), Math.max(slideImages.length - 1, 0))
    return slideImages[idx] ?? null
  }, [activeIndex, slideImages])
  const textOverlayFrame = useMemo(() => {
    if (!textLayerSize) return null
    const containerW = 960
    const containerH = 600
    const scaleRatio = Math.min(containerW / textLayerSize.width, containerH / textLayerSize.height)
    const displayW = textLayerSize.width * scaleRatio
    const displayH = textLayerSize.height * scaleRatio
    return {
      scaleRatio,
      offsetX: (containerW - displayW) / 2,
      offsetY: (containerH - displayH) / 2,
    }
  }, [textLayerSize])

  const safeActiveIndex = Math.min(Math.max(activeIndex, 0), Math.max(slideCount - 1, 0))
  const canPrev = safeActiveIndex > 0
  const canNext = safeActiveIndex < slideCount - 1
  const activePageNumber = safeActiveIndex + 1
  const editedCount = useMemo(
    () => textBoxes.filter((box) => {
      const textChanged = (textDrafts[box.boxId] ?? box.text) !== box.text
      const override = boxOverrides[box.boxId]
      const hasOverride = !!override && Object.values(override).some((value) => value !== undefined && value !== '')
      return textChanged || hasOverride
    }).length,
    [boxOverrides, textBoxes, textDrafts],
  )
  const activeTextBox = useMemo(
    () => textBoxes.find((box) => box.boxId === activeTextBoxId) || null,
    [activeTextBoxId, textBoxes],
  )
  const activeCandidates = useMemo(
    () => (activeTextBoxId ? lastCandidates[activeTextBoxId] || [] : []),
    [activeTextBoxId, lastCandidates],
  )

  const updateBoxOverride = useCallback((boxId: string, patch: Partial<PptTextEditStyleOverride>) => {
    setBoxOverrides((prev) => ({
      ...prev,
      [boxId]: {
        ...(prev[boxId] || {}),
        ...patch,
      },
    }))
  }, [])

  function mimeFromPath(p: string) {
    const lower = p.toLowerCase()
    if (lower.endsWith('.png')) return 'image/png'
    if (lower.endsWith('.jpg') || lower.endsWith('.jpeg')) return 'image/jpeg'
    if (lower.endsWith('.gif')) return 'image/gif'
    if (lower.endsWith('.webp')) return 'image/webp'
    return 'application/octet-stream'
  }

  const loadTextLayer = useCallback(async (options?: { cacheOnly?: boolean }) => {
    if (!pptxPath || !window.electronAPI?.pptDetectTextLayer) return
    setTextLayerLoading(true)
    setTextLayerError(null)
    try {
      const result = await window.electronAPI.pptDetectTextLayer({
        pptxPath,
        pageNumber: activePageNumber,
        useCache: true,
        cacheOnly: options?.cacheOnly === true,
      })
      if (!result.success) {
        if (options?.cacheOnly && result.error) {
          setTextBoxes([])
          setTextLayerSize(null)
          setTextDrafts({})
          setBoxOverrides({})
          setLastCandidates({})
          setActiveTextBoxId(null)
          return
        }
        throw new Error(result.error || '文字识别失败')
      }
      const boxes = Array.isArray(result.boxes) ? result.boxes : []
      setTextBoxes(boxes)
      setBoxOverrides({})
      setLastCandidates({})
      setTextEditHealth({
        ready: true,
        message: 'PPT 文本编辑主链已就绪',
      })
      setTextLayerSize(
        result.canvasWidth && result.canvasHeight
          ? { width: result.canvasWidth, height: result.canvasHeight }
          : null,
      )
      setTextDrafts(Object.fromEntries(boxes.map((box) => [box.boxId, box.text])))
      if (boxes.length > 0) {
        setTextEditMode(true)
        setActiveTextBoxId((prev) => prev && boxes.some((box) => box.boxId === prev) ? prev : boxes[0].boxId)
      } else {
        setActiveTextBoxId(null)
      }
    } catch (error) {
      setTextLayerError(error instanceof Error ? error.message : '文字识别失败')
    } finally {
      setTextLayerLoading(false)
    }
  }, [activePageNumber, pptxPath])

  const applyTextEdits = useCallback(async () => {
    if (!pptxPath || !window.electronAPI?.pptApplyTextEdits) return
    if (window.electronAPI?.pptTextEditHealth && !textEditHealth.ready) {
      setTextLayerError(null)
      setTextEditHealth({ ready: false, message: '正在初始化改字主链...' })
      const warmup = await window.electronAPI.pptTextEditHealth({ bootstrap: true })
      if (!warmup.success || !warmup.ready) {
        setTextEditHealth({ ready: false, message: warmup.error || 'PPT 文本编辑 sidecar 尚未初始化' })
        setTextLayerError(warmup.error || 'PPT 文本编辑 sidecar 尚未初始化')
        return
      }
      setTextEditHealth({ ready: true, message: 'PPT 文本编辑主链已就绪' })
    }
    const edits: PptTextEditOperation[] = textBoxes
      .map((box) => {
        const nextText = (textDrafts[box.boxId] ?? box.text).trim()
        const styleOverride = boxOverrides[box.boxId]
        const hasOverride = !!styleOverride && Object.values(styleOverride).some((value) => value !== undefined && value !== '')
        if (nextText === box.text && !hasOverride) return null
        return {
          boxId: box.boxId,
          fromText: box.text,
          toText: nextText,
          styleMode: 'preserve',
          bounds: box.bounds,
          styleOverride,
        }
      })
      .filter((item): item is PptTextEditOperation => item !== null)
    if (edits.length === 0) return

    setApplyingTextEdits(true)
    setTextLayerError(null)
    try {
      const result = await window.electronAPI.pptApplyTextEdits({
        pptxPath,
        pageNumber: activePageNumber,
        edits,
      })
      if (!result.success) {
        throw new Error(result.error || '应用改字失败')
      }

      if (result.imageDataUrl) {
        setSlideImages((prev) =>
          prev.map((item, idx) => (idx === safeActiveIndex ? result.imageDataUrl || item : item)),
        )
      }
      setLastCandidates(result.perBoxCandidates || {})

      setTextBoxes((prev) =>
        prev.map((box) => {
          const edit = edits.find((item) => item.boxId === box.boxId)
          return edit ? {
            ...box,
            text: edit.toText,
            bounds: edit.bounds || box.bounds,
            styleEstimate: {
              ...(box.styleEstimate || box.styleHint),
              ...(edit.styleOverride || {}),
            },
          } : box
        }),
      )
      setTextDrafts((prev) => {
        const next = { ...prev }
        edits.forEach((edit) => {
          next[edit.boxId] = edit.toText
        })
        return next
      })
      if (edits.length > 0) {
        setActiveTextBoxId(edits[0].boxId)
      }
    } catch (error) {
      setTextLayerError(error instanceof Error ? error.message : '应用改字失败')
    } finally {
      setApplyingTextEdits(false)
    }
  }, [activePageNumber, boxOverrides, pptxPath, safeActiveIndex, textBoxes, textDrafts, textEditHealth.ready])

  useEffect(() => {
    if (!window.electronAPI?.onPptTextEditProgress) return
    return window.electronAPI.onPptTextEditProgress((payload) => {
      setTextEditProgress({
        active: !!payload?.active,
        stage: payload?.stage,
        progress: payload?.progress,
        message: payload?.message,
        completedCandidates: payload?.completedCandidates,
        totalCandidates: payload?.totalCandidates,
      })
    })
  }, [])

  useEffect(() => {
    let cancelled = false
    async function run() {
      if (!pptxPath || !window.electronAPI?.pptTextEditHealth) return
      const result = await window.electronAPI.pptTextEditHealth({ bootstrap: false })
      if (cancelled) return
      if (result.success && result.ready) {
        const adapters = (result.externalAdapters || []).filter((item) => item.available)
        setTextEditHealth({
          ready: true,
          message: result.deterministicRendererAvailable
            ? `无痕改字主链路可用${adapters.length ? ` · 高配适配器 ${adapters.length} 个` : ' · 本地确定性重排' }`
            : '文字改字引擎可用（基础模式）',
        })
      } else {
        setTextEditHealth({ ready: false, message: result.error || '文字改字引擎未就绪' })
      }
    }
    void run()
    return () => {
      cancelled = true
    }
  }, [pptxPath])

  useEffect(() => {
    if (!pptxPath || !window.electronAPI?.pptDetectTextLayer) return
    void loadTextLayer({ cacheOnly: true })
  }, [activePageNumber, loadTextLayer, pptxPath])

  function resolveTargetPath(baseDir: string, target: string) {
    // baseDir like: "ppt/slides/"
    const baseParts = baseDir.split('/').filter(Boolean)
    const targetParts = target.split('/').filter(Boolean)
    const out: string[] = [...baseParts]
    for (const part of targetParts) {
      if (part === '.') continue
      if (part === '..') {
        out.pop()
        continue
      }
      out.push(part)
    }
    return out.join('/')
  }

  useEffect(() => {
    let cancelled = false

    async function run() {
      setLoading(true)
      setError(null)
      setSlideCount(0)
      setActiveIndex(0)
      setSlideImages([])

      try {
        const ab = base64ToArrayBuffer(pptxBase64)

        const zip = await JSZip.loadAsync(ab)
        const slidePaths = Object.keys(zip.files)
          .filter((p) => /^ppt\/slides\/slide\d+\.xml$/i.test(p))
          .sort((a, b) => {
            const na = Number(a.match(/slide(\d+)\.xml/i)?.[1] || 0)
            const nb = Number(b.match(/slide(\d+)\.xml/i)?.[1] || 0)
            return na - nb
          })

        if (slidePaths.length === 0) {
          throw new Error('未找到 slides（ppt/slides/slide*.xml），该文件可能不是有效的 PPTX')
        }

        const imagesPerSlide: (string | null)[] = await Promise.all(
          slidePaths.map(async (slidePath) => {
            const slideIndex = Number(slidePath.match(/slide(\d+)\.xml/i)?.[1] || 0)
            const slideXml = await zip.file(slidePath)?.async('string')
            if (!slideXml) return null

            // 1) 找 rId（图片引用）
            const rIds: string[] = []
            const embedRe = /r:embed="([^"]+)"/g
            let m: RegExpExecArray | null
            while ((m = embedRe.exec(slideXml)) !== null) {
              rIds.push(m[1])
            }

            // 2) 解析 rels：rId -> Target
            const relPath = `ppt/slides/_rels/slide${slideIndex}.xml.rels`
            const relXml = await zip.file(relPath)?.async('string')
            if (!relXml) return null

            const ridToTarget = new Map<string, string>()
            const relRe = /Relationship\b[^>]*\bId="([^"]+)"[^>]*\bType="([^"]+)"[^>]*\bTarget="([^"]+)"/g
            while ((m = relRe.exec(relXml)) !== null) {
              const id = m[1]
              const type = m[2]
              const target = m[3]
              if (type.includes('/image') || /media\//i.test(target)) {
                ridToTarget.set(id, target)
              }
            }

            const baseDir = 'ppt/slides/'
            let pickedTarget: string | undefined
            for (const rid of rIds) {
              const t = ridToTarget.get(rid)
              if (t) {
                pickedTarget = t
                break
              }
            }
            // 如果 slide 里没找到 rId，就从 rels 里兜底取第一个图片关系
            if (!pickedTarget) {
              for (const [, t] of ridToTarget) {
                pickedTarget = t
                break
              }
            }
            if (!pickedTarget) return null

            const imagePath = resolveTargetPath(baseDir, pickedTarget)
            const imgFile = zip.file(imagePath)
            if (!imgFile) return null

            const base64 = await imgFile.async('base64')
            const mime = mimeFromPath(imagePath)
            return `data:${mime};base64,${base64}`
          })
        )

        if (cancelled) return
        setSlideCount(slidePaths.length)
        setSlideImages(imagesPerSlide)
        setActiveIndex(0)
      } catch (e) {
        if (cancelled) return
        setError((e as Error).message || 'PPT 渲染失败')
      } finally {
        if (!cancelled) setLoading(false)
      }
    }

    run()
    return () => {
      cancelled = true
    }
  }, [pptxBase64])

  return (
    <div className={`flex flex-col h-full bg-background ${isFullscreen ? 'fixed inset-0 z-50 bg-background' : ''}`}>
      {/* Ribbon / Toolbar */}
      <div className="glass border-b border-border px-3 py-2 flex items-center gap-3">
        <div className="text-xs text-text font-medium truncate max-w-[30vw]">
          PPT 预览：{title}
        </div>
        
        {/* 编辑操作区 */}
        {pptxPath && (
          <>
            <div className="w-px h-5 bg-black/10 dark:bg-white/10 mx-1" />
            <div className="flex items-center gap-1">
              {/* 多选模式切换 */}
              <button
                onClick={() => {
                  if (isMultiSelectMode) {
                    exitMultiSelectMode()
                  } else {
                    setIsMultiSelectMode(true)
                  }
                }}
                className={`flex items-center gap-1 px-2 py-1 rounded-md text-xs transition-colors ${
                  isMultiSelectMode
                    ? 'bg-accent text-white border border-accent/25'
                    : 'text-text-muted hover:text-text hover:bg-black/5 dark:hover:bg-white/5 border border-transparent'
                }`}
                title={isMultiSelectMode ? '退出多选' : '多选页面'}
              >
                {isMultiSelectMode ? <CheckSquare className="w-3.5 h-3.5" /> : <Square className="w-3.5 h-3.5" />}
                {isMultiSelectMode ? `已选 ${selectedPages.size} 页` : '多选'}
              </button>
              
              {isMultiSelectMode && (
                <button
                  onClick={toggleSelectAll}
                  className="px-2 py-1 rounded-md text-xs text-text-muted hover:text-text hover:bg-black/5 dark:hover:bg-white/5"
                  title={selectedPages.size === slideCount ? '取消全选' : '全选'}
                >
                  {selectedPages.size === slideCount ? '取消全选' : '全选'}
                </button>
              )}
              
              <div className="w-px h-4 bg-black/10 dark:bg-white/10 mx-1" />
              
              {/* 整页重做 */}
              <button
                onClick={() => handleEditRequest('regenerate')}
                disabled={isMultiSelectMode && selectedPages.size === 0}
                className="flex items-center gap-1 px-2 py-1 rounded-md text-xs text-text-muted hover:text-text hover:bg-black/5 dark:hover:bg-white/5 disabled:opacity-40 disabled:cursor-not-allowed"
                title="整页重做：根据反馈重新生成选中页面"
              >
                <RefreshCw className="w-3.5 h-3.5" />
                整页重做
              </button>
              
              {/* 局部编辑 */}
              <button
                onClick={() => handleEditRequest('partial_edit')}
                disabled={isMultiSelectMode && selectedPages.size === 0}
                className="flex items-center gap-1 px-2 py-1 rounded-md text-xs text-text-muted hover:text-text hover:bg-black/5 dark:hover:bg-white/5 disabled:opacity-40 disabled:cursor-not-allowed"
                title="局部编辑：修改背景、文字等局部内容"
              >
                <Paintbrush className="w-3.5 h-3.5" />
                局部编辑
              </button>

              <div className="w-px h-4 bg-black/10 dark:bg-white/10 mx-1" />

              <button
                onClick={() => {
                  if (textEditMode) {
                    setTextEditMode(false)
                    setActiveTextBoxId(null)
                    setTextLayerError(null)
                  } else {
                    void loadTextLayer({ cacheOnly: false })
                  }
                }}
                className={`flex items-center gap-1 px-2 py-1 rounded-md text-xs transition-colors ${
                  textEditMode
                    ? 'bg-accent/14 text-accent border border-accent/25'
                    : 'text-text-muted hover:text-text hover:bg-black/5 dark:hover:bg-white/5'
                }`}
                title={textEditMode ? '退出文字编辑模式' : '识别当前页文字'}
              >
                <Paintbrush className="w-3.5 h-3.5" />
                {textEditMode ? '退出改字' : '识别文字'}
              </button>

              {textEditMode && (
                <button
                  onClick={() => setExpertMode((prev) => !prev)}
                  className={`flex items-center gap-1 px-2 py-1 rounded-md text-xs border transition-colors ${
                    expertMode
                      ? 'border-accent/25 bg-accent/12 text-accent'
                      : 'border-black/10 dark:border-white/10 text-text-muted hover:text-text hover:bg-black/5 dark:hover:bg-white/5'
                  }`}
                  title="切换专家模式"
                >
                  专家模式
                </button>
              )}

              {textEditMode && (
                <button
                  onClick={() => void applyTextEdits()}
                  disabled={editedCount === 0 || applyingTextEdits}
                  className="flex items-center gap-1 px-2 py-1 rounded-md text-xs text-success border border-success/25 bg-success/12 hover:bg-success/18 disabled:opacity-40 disabled:cursor-not-allowed"
                  title="应用文字替换"
                >
                  {applyingTextEdits ? <Loader2 className="w-3.5 h-3.5 animate-spin" /> : <CheckSquare className="w-3.5 h-3.5" />}
                  应用改字
                </button>
              )}
            </div>
          </>
        )}

        <div className="flex items-center gap-1 ml-auto">
          <button
            disabled={!canPrev}
            onClick={() => {
              if (!canPrev) return
              setActiveIndex((i) => Math.max(0, i - 1))
            }}
            className="p-1.5 rounded-md text-text-muted hover:text-text disabled:opacity-40 hover:bg-black/5 dark:hover:bg-white/5"
            title="上一页"
          >
            <ChevronLeft className="w-4 h-4" />
          </button>
          <button
            disabled={!canNext}
            onClick={() => {
              if (!canNext) return
              setActiveIndex((i) => Math.min(Math.max(slideCount - 1, 0), i + 1))
            }}
            className="p-1.5 rounded-md text-text-muted hover:text-text disabled:opacity-40 hover:bg-black/5 dark:hover:bg-white/5"
            title="下一页"
          >
            <ChevronRight className="w-4 h-4" />
          </button>

          <div className="w-px h-5 bg-black/10 dark:bg-white/10 mx-2" />

          <div className="flex items-center bg-black/5 dark:bg-white/5 border border-black/10 dark:border-white/10 rounded-md overflow-hidden">
            <button
              onClick={() => setScale((s) => Math.max(25, s - 10))}
              className="px-2 py-1 text-xs text-text-muted hover:text-text hover:bg-black/10 dark:hover:bg-white/10"
              title="缩小"
            >
              -
            </button>
            <span className="px-2 text-xs text-text-secondary min-w-[52px] text-center">{scale}%</span>
            <button
              onClick={() => setScale((s) => Math.min(200, s + 10))}
              className="px-2 py-1 text-xs text-text-muted hover:text-text hover:bg-black/10 dark:hover:bg-white/10"
              title="放大"
            >
              +
            </button>
          </div>

          <button
            onClick={() => setIsFullscreen((v) => !v)}
            className="p-1.5 rounded-md text-text-muted hover:text-text hover:bg-black/5 dark:hover:bg-white/5 ml-2"
            title={isFullscreen ? '退出全屏' : '全屏'}
          >
            {isFullscreen ? <Minimize2 className="w-4 h-4" /> : <Maximize2 className="w-4 h-4" />}
          </button>
        </div>
      </div>

      {/* Body - 左侧缩略图 + 右侧主画布 */}
      <div className="flex-1 flex overflow-hidden">
        {/* 左侧缩略图导航栏 */}
        <div className="w-[140px] flex-shrink-0 glass-panel overflow-y-auto">
          <div className="p-2 space-y-2">
            {slideImages.map((img, idx) => {
              const isSelected = selectedPages.has(idx)
              const isActive = idx === safeActiveIndex
              
              return (
                <div
                  key={idx}
                  className="relative"
                >
                  {/* 多选复选框 */}
                  {isMultiSelectMode && (
                    <button
                      onClick={(e) => {
                        e.stopPropagation()
                        togglePageSelection(idx)
                      }}
                      className={`absolute top-1 right-1 z-20 w-5 h-5 rounded flex items-center justify-center transition-colors ${
                        isSelected
                          ? 'bg-accent text-white'
                          : 'bg-black/60 text-white/60 hover:bg-black/80'
                      }`}
                    >
                      {isSelected ? (
                        <CheckSquare className="w-3.5 h-3.5" />
                      ) : (
                        <Square className="w-3.5 h-3.5" />
                      )}
                    </button>
                  )}
                  
                  <div
                    draggable={!isMultiSelectMode && !!img}
                    onDragStart={(e) => handleThumbnailDragStart(e, idx)}
                    onClick={() => {
                      if (isMultiSelectMode) {
                        togglePageSelection(idx)
                      } else {
                        setActiveIndex(idx)
                      }
                    }}
                    className={`w-full relative rounded-md overflow-hidden border-2 transition-all cursor-pointer ${
                      isSelected
                        ? 'border-accent ring-2 ring-accent/25'
                        : isActive
                        ? 'border-accent ring-1 ring-accent/20'
                        : 'border-transparent hover:border-black/20 dark:hover:border-white/20'
                    }`}
                  >
                    {/* 页码标签 */}
                    <div className="absolute top-1 left-1 bg-black/70 text-[9px] text-white px-1.5 py-0.5 rounded z-10">
                      {idx + 1}
                    </div>
                    {/* 拖拽提示 */}
                    {!isMultiSelectMode && img && (
                      <div className="absolute bottom-1 right-1 bg-black/70 text-[8px] text-white/60 px-1 py-0.5 rounded z-10 opacity-0 group-hover:opacity-100 transition-opacity">
                        <Move className="w-2.5 h-2.5 inline" />
                      </div>
                    )}
                    {/* 缩略图 */}
                    <div className="aspect-[16/10] bg-black/40 flex items-center justify-center">
                      {img ? (
                        <img
                          src={img}
                          alt={`幻灯片 ${idx + 1}`}
                          className="w-full h-full object-contain"
                          draggable={false}
                        />
                      ) : (
                        <div className="text-[8px] text-text-dim">无图片</div>
                      )}
                    </div>
                  </div>
                </div>
              )
            })}
            {loading && slideImages.length === 0 && (
              <div className="text-[10px] text-text-dim text-center py-4">
                <Loader2 className="w-4 h-4 animate-spin mx-auto mb-2" />
                加载中...
              </div>
            )}
          </div>
        </div>

        {/* 右侧主画布区域 */}
        <div 
          ref={mainCanvasRef}
          className="flex-1 overflow-auto bg-black/5 dark:bg-white/5 relative"
          onMouseUp={handleSelectionMouseUp}
          onMouseLeave={() => {
            if (isSelecting) {
              setIsSelecting(false)
              setSelectionStart(null)
              setSelectionRect(null)
            }
          }}
        >
          <div className="min-h-full p-6 flex flex-col items-center justify-center">
            <div
              ref={slideContainerRef}
              className="bg-black shadow-[0_10px_40px_rgba(0,0,0,0.55)] border border-black/10 dark:border-white/10 origin-center overflow-hidden relative select-none"
              style={{
                width: '960px',
                height: '600px',
                transform: `scale(${scale / 100})`,
                cursor: isSelecting ? 'crosshair' : 'default',
              }}
              onMouseDown={handleSelectionMouseDown}
              onMouseMove={handleSelectionMouseMove}
            >
              {/* image-only preview (pure local JS) */}
              <div className="w-full h-full bg-black flex items-center justify-center pointer-events-none">
                {currentSlideImage ? (
                  <img
                    src={currentSlideImage}
                    alt={`Slide ${safeActiveIndex + 1}`}
                    className="w-full h-full"
                    style={{ objectFit: 'contain' }}
                    draggable={false}
                  />
                ) : (
                  !loading && (
                    <div className="text-xs text-text-dim px-6 text-center">
                      本页未检测到可渲染的图片元素
                    </div>
                  )
                )}
              </div>

              {textEditMode && textLayerSize && (
                <div className="absolute inset-0 z-20">
                  {textBoxes.map((box) => {
                    const scaleRatio = textOverlayFrame?.scaleRatio || 1
                    const offsetX = textOverlayFrame?.offsetX || 0
                    const offsetY = textOverlayFrame?.offsetY || 0
                    const left = offsetX + box.bounds.left * scaleRatio
                    const top = offsetY + box.bounds.top * scaleRatio
                    const width = box.bounds.width * scaleRatio
                    const height = box.bounds.height * scaleRatio
                    const active = box.boxId === activeTextBoxId
                    return (
                      <div
                        key={box.boxId}
                        onClick={(e) => {
                          e.stopPropagation()
                          setActiveTextBoxId(box.boxId)
                        }}
                        className={`absolute rounded-sm border transition-all ${
                          active
                            ? 'border-accent shadow-[0_0_0_1px_rgba(59,130,246,0.45)]'
                            : 'border-white/45 hover:border-accent/70'
                        }`}
                        style={{
                          left,
                          top,
                          width,
                          height,
                          minHeight: 18,
                          backgroundColor: active ? `${box.styleHint.backgroundColor || '#ffffff'}ee` : 'transparent',
                          opacity: 1,
                        }}
                      >
                        <div
                          className={`absolute -top-5 left-0 px-1.5 py-0.5 rounded text-[10px] leading-none ${
                            active
                              ? 'bg-accent text-white'
                              : 'bg-black/70 text-white/80'
                          }`}
                        >
                          {box.readingOrder}
                        </div>
                        {active ? (
                          <textarea
                            value={textDrafts[box.boxId] ?? box.text}
                            onChange={(e) =>
                              setTextDrafts((prev) => ({ ...prev, [box.boxId]: e.target.value }))
                            }
                            className="w-full h-full bg-transparent text-white text-[12px] leading-snug px-1 py-0.5 resize-none focus:outline-none"
                            style={{
                              color: box.styleHint.textColor,
                              fontSize: `${Math.max(10, box.styleHint.fontSize * scaleRatio)}px`,
                              textAlign: box.styleHint.align,
                            }}
                          />
                        ) : (
                          <div className="w-full h-full" />
                        )}
                      </div>
                    )
                  })}
                </div>
              )}
              
              {/* 框选区域可视化 */}
              {isSelecting && selectionRect && selectionRect.w > 0 && selectionRect.h > 0 && (
                <div
                  className="absolute border-2 border-dashed border-accent bg-accent/20 pointer-events-none z-30"
                  style={{
                    left: selectionRect.x,
                    top: selectionRect.y,
                    width: selectionRect.w,
                    height: selectionRect.h,
                  }}
                />
              )}

              {loading && (
                <div className="absolute inset-0 flex items-center justify-center text-sm text-text-dim gap-2 bg-black/40">
                  <Loader2 className="w-4 h-4 animate-spin" />
                  正在渲染 PPT…
                </div>
              )}

              {!loading && error && (
                <div className="absolute inset-0 flex flex-col items-center justify-center text-sm px-6 text-center bg-black/60">
                  <FileWarning className="w-12 h-12 text-amber-500 mb-4" />
                  <div className="text-error mb-2">预览加载失败</div>
                  <div className="text-xs text-text-muted mb-4 max-w-[300px]">{error}</div>
                </div>
              )}
            </div>

            {textEditMode && activeTextBox && (
              <div className="mt-4 w-[960px] max-w-full rounded-xl border border-black/10 dark:border-white/10 bg-black/5 dark:bg-white/5 p-4">
                <div className="flex items-center justify-between gap-4 mb-3">
                  <div>
                    <div className="text-sm text-text font-medium">
                      文字框 {activeTextBox.readingOrder} · {activeTextBox.text}
                    </div>
                    <div className="text-[11px] text-text-dim mt-1">
                      背景复杂度：{activeTextBox.backgroundComplexity || 'medium'} ·
                      风格复杂度：{activeTextBox.styleComplexity || 'plain'} ·
                      候选字体：{activeTextBox.fontCandidates?.length || 0}
                    </div>
                  </div>
                  <div className="text-[11px] text-text-dim">
                    {expertMode ? '专家模式已开启' : '默认自动选优'}
                  </div>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div className="space-y-3">
                    <label className="block">
                      <div className="text-[11px] text-text-dim mb-1">字体候选</div>
                      <select
                        value={boxOverrides[activeTextBox.boxId]?.fontCandidateId || ''}
                        onChange={(e) => updateBoxOverride(activeTextBox.boxId, { fontCandidateId: e.target.value || undefined })}
                        className="w-full rounded-md border border-black/10 dark:border-white/10 bg-background px-2 py-1.5 text-sm"
                      >
                        <option value="">自动选择</option>
                        {(activeTextBox.fontCandidates || []).map((candidate) => (
                          <option key={candidate.candidateId} value={candidate.candidateId}>
                            {candidate.family} · {Math.round(candidate.confidence * 100)}%
                          </option>
                        ))}
                      </select>
                    </label>

                    {expertMode && (
                      <>
                        <label className="block">
                          <div className="flex items-center justify-between text-[11px] text-text-dim mb-1">
                            <span>字距</span>
                            <span>{boxOverrides[activeTextBox.boxId]?.letterSpacing ?? activeTextBox.styleEstimate?.letterSpacing ?? 0}px</span>
                          </div>
                          <input
                            type="range"
                            min={-2}
                            max={12}
                            step={0.5}
                            value={boxOverrides[activeTextBox.boxId]?.letterSpacing ?? activeTextBox.styleEstimate?.letterSpacing ?? 0}
                            onChange={(e) => updateBoxOverride(activeTextBox.boxId, { letterSpacing: Number(e.target.value) })}
                            className="w-full"
                          />
                        </label>

                        <label className="block">
                          <div className="flex items-center justify-between text-[11px] text-text-dim mb-1">
                            <span>行高</span>
                            <span>{(boxOverrides[activeTextBox.boxId]?.lineHeight ?? activeTextBox.styleEstimate?.lineHeight ?? 1).toFixed(2)}</span>
                          </div>
                          <input
                            type="range"
                            min={0.8}
                            max={2}
                            step={0.05}
                            value={boxOverrides[activeTextBox.boxId]?.lineHeight ?? activeTextBox.styleEstimate?.lineHeight ?? 1}
                            onChange={(e) => updateBoxOverride(activeTextBox.boxId, { lineHeight: Number(e.target.value) })}
                            className="w-full"
                          />
                        </label>
                      </>
                    )}
                  </div>

                  <div className="space-y-3">
                    {expertMode && (
                      <>
                        <label className="block">
                          <div className="flex items-center justify-between text-[11px] text-text-dim mb-1">
                            <span>描边</span>
                            <span>{(boxOverrides[activeTextBox.boxId]?.strokeWidth ?? activeTextBox.styleEstimate?.strokeWidth ?? 0).toFixed(1)}px</span>
                          </div>
                          <input
                            type="range"
                            min={0}
                            max={6}
                            step={0.25}
                            value={boxOverrides[activeTextBox.boxId]?.strokeWidth ?? activeTextBox.styleEstimate?.strokeWidth ?? 0}
                            onChange={(e) => updateBoxOverride(activeTextBox.boxId, { strokeWidth: Number(e.target.value) })}
                            className="w-full"
                          />
                        </label>

                        <label className="block">
                          <div className="flex items-center justify-between text-[11px] text-text-dim mb-1">
                            <span>阴影模糊</span>
                            <span>{(boxOverrides[activeTextBox.boxId]?.shadowBlur ?? activeTextBox.styleEstimate?.shadowBlur ?? 0).toFixed(1)}px</span>
                          </div>
                          <input
                            type="range"
                            min={0}
                            max={24}
                            step={0.5}
                            value={boxOverrides[activeTextBox.boxId]?.shadowBlur ?? activeTextBox.styleEstimate?.shadowBlur ?? 0}
                            onChange={(e) => updateBoxOverride(activeTextBox.boxId, { shadowBlur: Number(e.target.value) })}
                            className="w-full"
                          />
                        </label>

                        <label className="block">
                          <div className="text-[11px] text-text-dim mb-1">清底策略</div>
                          <select
                            value={boxOverrides[activeTextBox.boxId]?.cleanupStrategy || ''}
                            onChange={(e) => updateBoxOverride(activeTextBox.boxId, { cleanupStrategy: (e.target.value || undefined) as PptTextEditStyleOverride['cleanupStrategy'] })}
                            className="w-full rounded-md border border-black/10 dark:border-white/10 bg-background px-2 py-1.5 text-sm"
                          >
                            <option value="">自动</option>
                            <option value="analytic_fill">analytic_fill</option>
                            <option value="local_inpaint">local_inpaint</option>
                          </select>
                        </label>
                      </>
                    )}
                  </div>
                </div>

                {activeCandidates.length > 0 && (
                  <div className="mt-4">
                    <div className="text-[11px] text-text-dim mb-2">最近一次生成的候选</div>
                    <div className="flex gap-3 overflow-x-auto pb-1">
                      {activeCandidates.map((candidate) => (
                        <button
                          key={candidate.candidateId}
                          onClick={() => updateBoxOverride(activeTextBox.boxId, { fontCandidateId: candidate.fontCandidateId })}
                          className={`min-w-[160px] rounded-lg border text-left overflow-hidden ${
                            candidate.applied
                              ? 'border-accent ring-1 ring-accent/20'
                              : 'border-black/10 dark:border-white/10'
                          }`}
                        >
                          {candidate.previewDataUrl && (
                            <img
                              src={candidate.previewDataUrl}
                              alt={candidate.label}
                              className="w-full h-[72px] object-cover bg-black/5"
                            />
                          )}
                          <div className="p-2">
                            <div className="text-[11px] text-text font-medium truncate">{candidate.label}</div>
                            <div className="text-[10px] text-text-dim mt-1">
                              score {candidate.score.total.toFixed(2)} · OCR {candidate.score.ocrExactness.toFixed(2)}
                            </div>
                          </div>
                        </button>
                      ))}
                    </div>
                  </div>
                )}
              </div>
            )}
            
            {/* Ctrl 框选提示 */}
            {!loading && !error && currentSlideImage && (
              <div className="mt-2 flex flex-col items-center gap-1 text-[10px] text-text-dim">
                <div className="flex items-center gap-1">
                  <span className="px-1.5 py-0.5 bg-black/10 dark:bg-white/10 border border-black/10 dark:border-white/10 rounded text-text-dim">Ctrl</span>
                  + 拖拽框选区域进行局部编辑
                </div>
                {textEditProgress.active && (
                  <div className="mt-2 w-[360px] max-w-full rounded-lg border border-accent/20 bg-accent/5 px-3 py-2">
                    <div className="flex items-center justify-between gap-3">
                      <div className="text-[11px] text-text">
                        {textEditProgress.message || '正在处理改字...'}
                      </div>
                      <div className="text-[10px] text-text-dim">
                        {typeof textEditProgress.progress === 'number'
                          ? `${Math.round(textEditProgress.progress * 100)}%`
                          : ''}
                      </div>
                    </div>
                    <div className="mt-2 h-1.5 rounded-full bg-black/10 dark:bg-white/10 overflow-hidden">
                      <div
                        className="h-full bg-accent transition-all duration-300"
                        style={{ width: `${Math.max(6, Math.round((textEditProgress.progress || 0) * 100))}%` }}
                      />
                    </div>
                    {typeof textEditProgress.completedCandidates === 'number' && typeof textEditProgress.totalCandidates === 'number' && textEditProgress.totalCandidates > 0 && (
                      <div className="mt-1 text-[10px] text-text-dim">
                        候选进度：{Math.min(textEditProgress.completedCandidates, textEditProgress.totalCandidates)} / {textEditProgress.totalCandidates}
                      </div>
                    )}
                  </div>
                )}
                {textEditMode && (
                  <div className="text-center">
                    {textLayerLoading
                      ? '正在识别当前页文字...'
                      : textLayerError
                      ? `文字编辑不可用：${textLayerError}`
                      : `已识别 ${textBoxes.length} 个文字框，可直接点击框并修改文字`}
                  </div>
                )}
                {textLayerError && (
                  <button
                    onClick={() => handleEditRequest('partial_edit')}
                    className="mt-1 px-2 py-1 rounded-md border border-orange-400/30 text-orange-400 hover:bg-orange-400/10 transition-colors"
                  >
                    改用局部编辑
                  </button>
                )}
                {!textEditMode && textEditHealth.message && (
                  <div className={textEditHealth.ready ? 'text-success' : 'text-warning'}>
                    {textEditHealth.message}
                  </div>
                )}
              </div>
            )}

            {/* 底部翻页控制 */}
            <div className="mt-4 flex items-center gap-4">
              <button
                disabled={!canPrev}
                onClick={() => setActiveIndex((i) => Math.max(0, i - 1))}
                className="flex items-center gap-1 px-3 py-1.5 rounded-md bg-black/10 dark:bg-white/10 border border-black/10 dark:border-white/10 text-text-secondary text-xs disabled:opacity-40 hover:bg-black/15 dark:hover:bg-white/15 transition-colors"
              >
                <ChevronLeft className="w-4 h-4" />
                上一页
              </button>
              <div className="text-sm text-text font-medium">
                {safeActiveIndex + 1} / {slideCount}
              </div>
              <button
                disabled={!canNext}
                onClick={() => setActiveIndex((i) => Math.min(Math.max(slideCount - 1, 0), i + 1))}
                className="flex items-center gap-1 px-3 py-1.5 rounded-md bg-black/10 dark:bg-white/10 border border-black/10 dark:border-white/10 text-text-secondary text-xs disabled:opacity-40 hover:bg-black/15 dark:hover:bg-white/15 transition-colors"
              >
                下一页
                <ChevronRight className="w-4 h-4" />
              </button>
            </div>
          </div>
        </div>
      </div>
    </div>
  )
}


