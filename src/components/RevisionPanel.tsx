import { useMemo } from 'react'
import { X, Check, Ban, CornerDownRight } from 'lucide-react'
import { useDocument } from '../context/DocumentContext'

type RevisionPanelProps = {
  open: boolean
  onClose: () => void
}

export default function RevisionPanel({ open, onClose }: RevisionPanelProps) {
  const {
    pendingChanges,
    pendingChangesTotal,
    acceptChange,
    rejectChange,
    acceptAllChanges,
    rejectAllChanges,
    scrollToDiffId,
  } = useDocument()

  const items = useMemo(() => {
    return [...pendingChanges].sort((a, b) => a.timestamp - b.timestamp)
  }, [pendingChanges])

  if (!open) return null

  return (
    <>
      {/* overlay */}
      <div
        className="fixed inset-0 z-[9998] bg-black/40"
        onClick={onClose}
      />

      {/* panel */}
      <div className="fixed right-0 top-0 bottom-0 z-[9999] w-[420px] max-w-[95vw] bg-zinc-950 border-l border-zinc-800 shadow-2xl flex flex-col">
        <div className="flex items-center justify-between px-4 py-3 border-b border-zinc-800">
          <div className="flex flex-col">
            <div className="text-sm text-zinc-100 font-medium">修订面板</div>
            <div className="text-xs text-zinc-400">
              {items.length} 条修改 · 命中总计 {pendingChangesTotal}
            </div>
          </div>
          <button
            className="p-1.5 rounded hover:bg-zinc-800 transition-colors"
            onClick={onClose}
            title="关闭"
          >
            <X className="w-4 h-4 text-zinc-300" />
          </button>
        </div>

        <div className="px-4 py-3 border-b border-zinc-800 flex items-center gap-2">
          <button
            className="flex items-center gap-1.5 px-3 py-1.5 bg-green-600 hover:bg-green-500 text-white text-xs rounded-lg transition-colors"
            onClick={acceptAllChanges}
            disabled={items.length === 0}
            title="全部接受"
          >
            <Check className="w-3.5 h-3.5" />
            <span>全部接受</span>
          </button>
          <button
            className="flex items-center gap-1.5 px-3 py-1.5 bg-zinc-800 hover:bg-red-600/20 text-zinc-300 hover:text-red-300 text-xs rounded-lg transition-colors border border-zinc-700 hover:border-red-500/30"
            onClick={rejectAllChanges}
            disabled={items.length === 0}
            title="全部拒绝"
          >
            <Ban className="w-3.5 h-3.5" />
            <span>全部拒绝</span>
          </button>
        </div>

        <div className="flex-1 overflow-auto">
          {items.length === 0 ? (
            <div className="px-4 py-8 text-center text-sm text-zinc-400">
              暂无待审阅修改
            </div>
          ) : (
            <div className="divide-y divide-zinc-800">
              {items.map((c) => (
                <div key={c.id} className="px-4 py-3">
                  <div className="flex items-start justify-between gap-3">
                    <div className="min-w-0">
                      <div className="text-xs text-zinc-400">
                        {c.kind === 'replace_text' ? '文字替换' : c.kind}
                        {c.stats?.matches ? ` · 命中 ${c.stats.matches} 处` : ''}
                      </div>
                      <div className="text-sm text-zinc-100 font-medium mt-1">
                        {c.summary || '修改'}
                      </div>
                    </div>
                    <button
                      className="shrink-0 p-1.5 rounded hover:bg-zinc-800 transition-colors"
                      onClick={() => scrollToDiffId(c.id)}
                      title="定位到文档"
                    >
                      <CornerDownRight className="w-4 h-4 text-zinc-300" />
                    </button>
                  </div>

                  <div className="mt-2 grid grid-cols-2 gap-2 text-xs">
                    <div className="bg-zinc-900/50 border border-zinc-800 rounded-lg p-2">
                      <div className="text-[10px] text-zinc-500 mb-1">原文</div>
                      <div className="text-red-300/90 line-through break-words">
                        {c.beforePreview ? (c.beforePreview.length > 140 ? c.beforePreview.slice(0, 140) + '…' : c.beforePreview) : '—'}
                      </div>
                    </div>
                    <div className="bg-zinc-900/50 border border-zinc-800 rounded-lg p-2">
                      <div className="text-[10px] text-zinc-500 mb-1">修改后</div>
                      <div className="text-green-300/90 break-words">
                        {c.afterPreview ? (c.afterPreview.length > 140 ? c.afterPreview.slice(0, 140) + '…' : c.afterPreview) : '—'}
                      </div>
                    </div>
                  </div>

                  <div className="mt-2 flex items-center gap-2">
                    <button
                      className="flex items-center gap-1.5 px-3 py-1.5 bg-green-600 hover:bg-green-500 text-white text-xs rounded-lg transition-colors"
                      onClick={() => acceptChange(c.id)}
                      title="接受此修改"
                    >
                      <Check className="w-3.5 h-3.5" />
                      <span>接受</span>
                    </button>
                    <button
                      className="flex items-center gap-1.5 px-3 py-1.5 bg-zinc-800 hover:bg-red-600/20 text-zinc-300 hover:text-red-300 text-xs rounded-lg transition-colors border border-zinc-700 hover:border-red-500/30"
                      onClick={() => rejectChange(c.id)}
                      title="拒绝此修改"
                    >
                      <Ban className="w-3.5 h-3.5" />
                      <span>拒绝</span>
                    </button>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
    </>
  )
}



