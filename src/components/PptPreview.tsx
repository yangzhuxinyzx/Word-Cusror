import { useMemo, useState } from 'react'
import { ChevronLeft, ChevronRight, Maximize2, Minimize2 } from 'lucide-react'

type PptPreviewProps = {
  title: string
  slideUrls: string[]
}

export default function PptPreview({ title, slideUrls }: PptPreviewProps) {
  const [activeIndex, setActiveIndex] = useState(0)
  const [scale, setScale] = useState(100)
  const [isFullscreen, setIsFullscreen] = useState(false)

  const safeActiveIndex = Math.min(Math.max(activeIndex, 0), Math.max(slideUrls.length - 1, 0))
  const activeUrl = slideUrls[safeActiveIndex] || ''

  const canPrev = safeActiveIndex > 0
  const canNext = safeActiveIndex < slideUrls.length - 1

  const thumbItems = useMemo(() => {
    return slideUrls.map((url, idx) => ({
      url,
      idx,
      label: String(idx + 1),
    }))
  }, [slideUrls])

  return (
    <div className={`flex flex-col h-full bg-[#1e1e1e] ${isFullscreen ? 'fixed inset-0 z-50' : ''}`}>
      {/* Ribbon / Toolbar */}
      <div className="border-b border-[#2d2d2d] bg-[#252526] px-3 py-2 flex items-center gap-3">
        <div className="text-xs text-[#cccccc] font-medium truncate max-w-[40vw]">
          PPT 预览：{title}
        </div>

        <div className="flex items-center gap-1 ml-auto">
          <button
            disabled={!canPrev}
            onClick={() => canPrev && setActiveIndex(safeActiveIndex - 1)}
            className="p-1.5 rounded-md text-[#cfcfcf] disabled:opacity-40 hover:bg-[#2d2d2d]"
            title="上一页"
          >
            <ChevronLeft className="w-4 h-4" />
          </button>
          <button
            disabled={!canNext}
            onClick={() => canNext && setActiveIndex(safeActiveIndex + 1)}
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

      {/* Body */}
      <div className="flex-1 flex overflow-hidden">
        {/* Thumbnails */}
        <div className="w-[220px] bg-[#1b1b1b] border-r border-[#2d2d2d] overflow-y-auto py-3">
          <div className="px-3 pb-2 text-[10px] text-[#9aa0a6] uppercase tracking-wider">幻灯片</div>
          <div className="space-y-2 px-2">
            {thumbItems.map((item) => {
              const selected = item.idx === safeActiveIndex
              return (
                <button
                  key={item.idx}
                  onClick={() => setActiveIndex(item.idx)}
                  className={`w-full flex items-start gap-2 p-2 rounded-md border transition-colors ${
                    selected
                      ? 'border-violet-500/50 bg-violet-500/10'
                      : 'border-[#2d2d2d] bg-[#1e1e1e] hover:bg-[#252526]'
                  }`}
                >
                  <div className="text-[10px] text-[#9aa0a6] w-5 shrink-0 pt-0.5">{item.label}</div>
                  <div className="flex-1 overflow-hidden">
                    <div className="aspect-[16/10] bg-black/40 rounded border border-[#2d2d2d] overflow-hidden">
                      {/* eslint-disable-next-line jsx-a11y/img-redundant-alt */}
                      <img src={item.url} alt={`slide-${item.idx + 1}`} className="w-full h-full object-contain" />
                    </div>
                  </div>
                </button>
              )
            })}
          </div>
        </div>

        {/* Stage */}
        <div className="flex-1 overflow-auto bg-[#111] relative">
          <div className="min-h-full p-8 flex justify-center">
            <div
              className="bg-black shadow-[0_10px_40px_rgba(0,0,0,0.55)] border border-[#2d2d2d] origin-top"
              style={{
                width: '960px',
                height: '600px',
                transform: `scale(${scale / 100})`,
              }}
            >
              {activeUrl ? (
                // eslint-disable-next-line jsx-a11y/img-redundant-alt
                <img src={activeUrl} alt="active-slide" className="w-full h-full object-contain" />
              ) : (
                <div className="w-full h-full flex items-center justify-center text-sm text-[#9aa0a6]">
                  暂无可预览的幻灯片
                </div>
              )}
            </div>
          </div>

          {/* Bottom status */}
          <div className="sticky bottom-0 border-t border-[#2d2d2d] bg-[#1b1b1b] px-3 py-1.5 text-[10px] text-[#9aa0a6] flex items-center justify-between">
            <div>
              第 {safeActiveIndex + 1} / {slideUrls.length} 页
            </div>
            <div>只读预览（Phase 0）</div>
          </div>
        </div>
      </div>
    </div>
  )
}



