import { useEffect, useMemo, useState, useCallback, useRef } from 'react'
import { ChevronLeft, ChevronRight, Loader2, Maximize2, Minimize2, FileWarning, RefreshCw, Paintbrush, CheckSquare, Square, Move } from 'lucide-react'
import JSZip from 'jszip'
import html2canvas from 'html2canvas'

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

  const safeActiveIndex = Math.min(Math.max(activeIndex, 0), Math.max(slideCount - 1, 0))
  const canPrev = safeActiveIndex > 0
  const canNext = safeActiveIndex < slideCount - 1

  function mimeFromPath(p: string) {
    const lower = p.toLowerCase()
    if (lower.endsWith('.png')) return 'image/png'
    if (lower.endsWith('.jpg') || lower.endsWith('.jpeg')) return 'image/jpeg'
    if (lower.endsWith('.gif')) return 'image/gif'
    if (lower.endsWith('.webp')) return 'image/webp'
    return 'application/octet-stream'
  }

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
    <div className={`flex flex-col h-full bg-[#1e1e1e] ${isFullscreen ? 'fixed inset-0 z-50' : ''}`}>
      {/* Ribbon / Toolbar */}
      <div className="border-b border-[#2d2d2d] bg-[#252526] px-3 py-2 flex items-center gap-3">
        <div className="text-xs text-[#cccccc] font-medium truncate max-w-[30vw]">
          PPT 预览：{title}
        </div>
        
        {/* 编辑操作区 */}
        {pptxPath && (
          <>
            <div className="w-px h-5 bg-[#3a3a3a] mx-1" />
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
                    ? 'bg-[#0e639c] text-white'
                    : 'text-[#cfcfcf] hover:bg-[#2d2d2d]'
                }`}
                title={isMultiSelectMode ? '退出多选' : '多选页面'}
              >
                {isMultiSelectMode ? <CheckSquare className="w-3.5 h-3.5" /> : <Square className="w-3.5 h-3.5" />}
                {isMultiSelectMode ? `已选 ${selectedPages.size} 页` : '多选'}
              </button>
              
              {isMultiSelectMode && (
                <button
                  onClick={toggleSelectAll}
                  className="px-2 py-1 rounded-md text-xs text-[#cfcfcf] hover:bg-[#2d2d2d]"
                  title={selectedPages.size === slideCount ? '取消全选' : '全选'}
                >
                  {selectedPages.size === slideCount ? '取消全选' : '全选'}
                </button>
              )}
              
              <div className="w-px h-4 bg-[#3a3a3a] mx-1" />
              
              {/* 整页重做 */}
              <button
                onClick={() => handleEditRequest('regenerate')}
                disabled={isMultiSelectMode && selectedPages.size === 0}
                className="flex items-center gap-1 px-2 py-1 rounded-md text-xs text-[#cfcfcf] hover:bg-[#2d2d2d] disabled:opacity-40 disabled:cursor-not-allowed"
                title="整页重做：根据反馈重新生成选中页面"
              >
                <RefreshCw className="w-3.5 h-3.5" />
                整页重做
              </button>
              
              {/* 局部编辑 */}
              <button
                onClick={() => handleEditRequest('partial_edit')}
                disabled={isMultiSelectMode && selectedPages.size === 0}
                className="flex items-center gap-1 px-2 py-1 rounded-md text-xs text-[#cfcfcf] hover:bg-[#2d2d2d] disabled:opacity-40 disabled:cursor-not-allowed"
                title="局部编辑：修改背景、文字等局部内容"
              >
                <Paintbrush className="w-3.5 h-3.5" />
                局部编辑
              </button>
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
            className="p-1.5 rounded-md text-[#cfcfcf] disabled:opacity-40 hover:bg-[#2d2d2d]"
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
            className="p-1.5 rounded-md text-[#cfcfcf] disabled:opacity-40 hover:bg-[#2d2d2d]"
            title="下一页"
          >
            <ChevronRight className="w-4 h-4" />
          </button>

          <div className="w-px h-5 bg-[#3a3a3a] mx-2" />

          <div className="flex items-center bg-[#1e1e1e] border border-[#2d2d2d] rounded-md overflow-hidden">
            <button
              onClick={() => setScale((s) => Math.max(25, s - 10))}
              className="px-2 py-1 text-xs text-[#cfcfcf] hover:bg-[#2d2d2d]"
              title="缩小"
            >
              -
            </button>
            <span className="px-2 text-xs text-[#cfcfcf] min-w-[52px] text-center">{scale}%</span>
            <button
              onClick={() => setScale((s) => Math.min(200, s + 10))}
              className="px-2 py-1 text-xs text-[#cfcfcf] hover:bg-[#2d2d2d]"
              title="放大"
            >
              +
            </button>
          </div>

          <button
            onClick={() => setIsFullscreen((v) => !v)}
            className="p-1.5 rounded-md text-[#cfcfcf] hover:bg-[#2d2d2d] ml-2"
            title={isFullscreen ? '退出全屏' : '全屏'}
          >
            {isFullscreen ? <Minimize2 className="w-4 h-4" /> : <Maximize2 className="w-4 h-4" />}
          </button>
        </div>
      </div>

      {/* Body - 左侧缩略图 + 右侧主画布 */}
      <div className="flex-1 flex overflow-hidden">
        {/* 左侧缩略图导航栏 */}
        <div className="w-[140px] flex-shrink-0 bg-[#1a1a1a] border-r border-[#2d2d2d] overflow-y-auto">
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
                          ? 'bg-[#0e639c] text-white'
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
                        ? 'border-[#0e639c] ring-2 ring-[#0e639c]/50'
                        : isActive
                        ? 'border-[#0e639c] ring-1 ring-[#0e639c]/50'
                        : 'border-transparent hover:border-[#3c3c3c]'
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
                    <div className="aspect-[16/10] bg-[#111] flex items-center justify-center">
                      {img ? (
                        <img
                          src={img}
                          alt={`幻灯片 ${idx + 1}`}
                          className="w-full h-full object-contain"
                          draggable={false}
                        />
                      ) : (
                        <div className="text-[8px] text-[#666]">无图片</div>
                      )}
                    </div>
                  </div>
                </div>
              )
            })}
            {loading && slideImages.length === 0 && (
              <div className="text-[10px] text-[#666] text-center py-4">
                <Loader2 className="w-4 h-4 animate-spin mx-auto mb-2" />
                加载中...
              </div>
            )}
          </div>
        </div>

        {/* 右侧主画布区域 */}
        <div 
          ref={mainCanvasRef}
          className="flex-1 overflow-auto bg-[#111] relative"
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
              className="bg-black shadow-[0_10px_40px_rgba(0,0,0,0.55)] border border-[#2d2d2d] origin-center overflow-hidden relative select-none"
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
                    <div className="text-xs text-[#9aa0a6] px-6 text-center">
                      本页未检测到可渲染的图片元素
                    </div>
                  )
                )}
              </div>
              
              {/* 框选区域可视化 */}
              {isSelecting && selectionRect && selectionRect.w > 0 && selectionRect.h > 0 && (
                <div
                  className="absolute border-2 border-dashed border-[#0e639c] bg-[#0e639c]/20 pointer-events-none z-30"
                  style={{
                    left: selectionRect.x,
                    top: selectionRect.y,
                    width: selectionRect.w,
                    height: selectionRect.h,
                  }}
                />
              )}

              {loading && (
                <div className="absolute inset-0 flex items-center justify-center text-sm text-[#9aa0a6] gap-2 bg-black/40">
                  <Loader2 className="w-4 h-4 animate-spin" />
                  正在渲染 PPT…
                </div>
              )}

              {!loading && error && (
                <div className="absolute inset-0 flex flex-col items-center justify-center text-sm px-6 text-center bg-[#1a1a1a]">
                  <FileWarning className="w-12 h-12 text-amber-500 mb-4" />
                  <div className="text-[#ffb4b4] mb-2">预览加载失败</div>
                  <div className="text-xs text-[#888] mb-4 max-w-[300px]">{error}</div>
                </div>
              )}
            </div>
            
            {/* Ctrl 框选提示 */}
            {!loading && !error && currentSlideImage && (
              <div className="mt-2 text-[10px] text-[#666] flex items-center gap-1">
                <span className="px-1.5 py-0.5 bg-[#2d2d2d] rounded text-[#999]">Ctrl</span>
                + 拖拽框选区域进行局部编辑
              </div>
            )}

            {/* 底部翻页控制 */}
            <div className="mt-4 flex items-center gap-4">
              <button
                disabled={!canPrev}
                onClick={() => setActiveIndex((i) => Math.max(0, i - 1))}
                className="flex items-center gap-1 px-3 py-1.5 rounded-md bg-[#2d2d2d] text-[#cfcfcf] text-xs disabled:opacity-40 hover:bg-[#3c3c3c] transition-colors"
              >
                <ChevronLeft className="w-4 h-4" />
                上一页
              </button>
              <div className="text-sm text-[#cfcfcf] font-medium">
                {safeActiveIndex + 1} / {slideCount}
              </div>
              <button
                disabled={!canNext}
                onClick={() => setActiveIndex((i) => Math.min(Math.max(slideCount - 1, 0), i + 1))}
                className="flex items-center gap-1 px-3 py-1.5 rounded-md bg-[#2d2d2d] text-[#cfcfcf] text-xs disabled:opacity-40 hover:bg-[#3c3c3c] transition-colors"
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


