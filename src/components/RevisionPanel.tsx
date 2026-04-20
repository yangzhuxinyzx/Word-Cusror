import { useEffect, useMemo, useState } from 'react'
import { X, Check, Ban, CornerDownRight } from 'lucide-react'
import { useDocument } from '../context/useDocument'

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
    document,
  } = useDocument()

  const items = useMemo(() => {
    return [...pendingChanges].sort((a, b) => a.timestamp - b.timestamp)
  }, [pendingChanges])

  const [trackAuthorFilter, setTrackAuthorFilter] = useState<string>('all')
  const [hadPendingSinceOpen, setHadPendingSinceOpen] = useState(false)

  useEffect(() => {
    if (!open) {
      setHadPendingSinceOpen(false)
      return
    }
    if (items.length > 0) {
      setHadPendingSinceOpen(true)
    }
  }, [items.length, open])

  const trackChanges = useMemo(() => {
    const html = document?.content || ''
    if (!html) return { items: [] as Array<{ key: string; type: 'insert' | 'delete'; author?: string; date?: string; text: string }>, authors: [] as string[] }
    try {
      const parser = new DOMParser()
      const docx = parser.parseFromString(html, 'text/html')
      const nodes = Array.from(docx.querySelectorAll('span.docx-track')) as HTMLElement[]
      const grouped = new Map<string, { key: string; type: 'insert' | 'delete'; author?: string; date?: string; text: string }>()
      const authors = new Set<string>()

      for (const n of nodes) {
        const typeRaw = (n.getAttribute('data-track-type') || 'insert').toLowerCase()
        const type = (typeRaw === 'delete' ? 'delete' : 'insert') as 'insert' | 'delete'
        const id = (n.getAttribute('data-track-id') || '').trim()
        const author = (n.getAttribute('data-track-author') || '').trim() || undefined
        const date = (n.getAttribute('data-track-date') || '').trim() || undefined
        const text = (n.textContent || '').trim()

        if (author) authors.add(author)
        const key = `${type}:${id || `${author || 'unknown'}:${date || 'nodate'}:${text.slice(0, 24)}`}`
        const prev = grouped.get(key)
        if (prev) {
          if (text && !prev.text.includes(text)) prev.text = (prev.text ? prev.text + ' ' : '') + text
        } else {
          grouped.set(key, { key, type, author, date, text })
        }
      }

      return { items: Array.from(grouped.values()), authors: Array.from(authors).sort() }
    } catch {
      return { items: [], authors: [] }
    }
  }, [document?.content])

  const filteredTrackItems = useMemo(() => {
    if (trackAuthorFilter === 'all') return trackChanges.items
    return trackChanges.items.filter(i => i.author === trackAuthorFilter)
  }, [trackChanges.items, trackAuthorFilter])

  if (!open) return null

  return (
    <>
      {/* overlay */}
      <div
        className="fixed inset-0 z-[9998] bg-black/40"
        onClick={onClose}
      />

      {/* panel */}
      <div className="fixed right-0 top-0 bottom-0 z-[9999] w-[420px] max-w-[95vw] bg-surface backdrop-blur-xl border-l border-border shadow-2xl flex flex-col">
        <div className="flex items-center justify-between px-4 py-3 border-b border-border">
          <div className="flex flex-col">
            <div className="text-sm text-text font-medium">修订面板</div>
            <div className="text-xs text-text-muted">
              {items.length} 条修改 · 命中总计 {pendingChangesTotal}
            </div>
          </div>
          <button
            className="p-1.5 rounded hover:bg-black/5 dark:hover:bg-white/5 transition-colors"
            onClick={onClose}
            title="关闭"
          >
            <X className="w-4 h-4 text-text-secondary" />
          </button>
        </div>

        <div className="px-4 py-3 border-b border-border flex items-center gap-2">
          <button
            className="flex items-center gap-1.5 px-3 py-1.5 bg-success/16 text-success border border-success/25 hover:bg-success/22 text-xs rounded-lg transition-colors"
            onClick={acceptAllChanges}
            disabled={items.length === 0}
            title="全部接受"
          >
            <Check className="w-3.5 h-3.5" />
            <span>全部接受</span>
          </button>
          <button
            className="flex items-center gap-1.5 px-3 py-1.5 bg-black/10 dark:bg-white/10 hover:bg-error/12 text-text-secondary hover:text-error text-xs rounded-lg transition-colors border border-black/10 dark:border-white/10 hover:border-error/30"
            onClick={rejectAllChanges}
            disabled={items.length === 0}
            title="全部拒绝"
          >
            <Ban className="w-3.5 h-3.5" />
            <span>全部拒绝</span>
          </button>
        </div>

        <div className="flex-1 overflow-auto">
          <div className="divide-y divide-border">
            {/* AI 修订（现有） */}
            <div className="px-4 py-3">
              <div className="text-xs text-text-muted mb-2">AI 待审阅修订</div>
              {items.length === 0 ? (
                <div className="py-6 text-center">
                  <div className="text-sm text-text font-medium">
                    {hadPendingSinceOpen ? '所有修订已处理' : '暂无待审阅修改'}
                  </div>
                  <div className="text-xs text-text-muted mt-1">
                    {hadPendingSinceOpen ? '当前文档里的待处理修订已经清空。' : '当前还没有需要你确认的 AI 修订。'}
                  </div>
                  <button
                    className="mt-3 inline-flex items-center gap-1.5 px-3 py-1.5 bg-black/10 dark:bg-white/10 hover:bg-black/15 dark:hover:bg-white/15 text-text-secondary text-xs rounded-lg transition-colors border border-black/10 dark:border-white/10"
                    onClick={onClose}
                  >
                    <X className="w-3.5 h-3.5" />
                    <span>关闭面板</span>
                  </button>
                </div>
              ) : (
                <div className="space-y-3">
                  {items.map((c) => {
                    const reviewTypeBadge: Record<string, { label: string; cls: string }> = {
                      grammar: { label: '语法', cls: 'bg-red-500/20 text-red-400 border-red-500/30' },
                      logic: { label: '逻辑', cls: 'bg-orange-500/20 text-orange-400 border-orange-500/30' },
                      style: { label: '措辞', cls: 'bg-blue-500/20 text-blue-400 border-blue-500/30' },
                      typo: { label: '错别字', cls: 'bg-purple-500/20 text-purple-400 border-purple-500/30' },
                      format: { label: '格式', cls: 'bg-gray-500/20 text-gray-400 border-gray-500/30' },
                    }
                    const badge = c.reviewType ? reviewTypeBadge[c.reviewType] : undefined
                    const isReview = !!(c.reviewReason || c.reviewType)

                    return (
                    <div key={c.id} className="border border-border rounded-lg p-3 bg-black/5 dark:bg-white/5">
                      <div className="flex items-start justify-between gap-3">
                        <div className="min-w-0">
                          <div className="flex items-center gap-1.5 flex-wrap">
                            {badge && (
                              <span className={`inline-flex items-center px-1.5 py-0.5 text-[10px] font-medium rounded border ${badge.cls}`}>
                                {badge.label}
                              </span>
                            )}
                            <span className="text-xs text-text-muted">
                              {isReview ? '审查建议' : (c.kind === 'replace_text' ? '文字替换' : c.kind)}
                              {c.stats?.matches ? ` · ${c.stats.matches} 处` : ''}
                            </span>
                          </div>
                          <div className="text-sm text-text font-medium mt-1">{c.summary || '修改'}</div>
                          {c.reviewReason && (
                            <div className="text-xs text-text-muted mt-0.5 italic">
                              {c.reviewReason}
                            </div>
                          )}
                        </div>
                        <button
                          className="shrink-0 p-1.5 rounded hover:bg-black/5 dark:hover:bg-white/5 transition-colors"
                          onClick={() => scrollToDiffId(c.id)}
                          title="定位到文档"
                        >
                          <CornerDownRight className="w-4 h-4 text-text-secondary" />
                        </button>
                      </div>

                      <div className="mt-2 grid grid-cols-2 gap-2 text-xs">
                        <div className="bg-black/5 dark:bg-white/5 border border-black/10 dark:border-white/10 rounded-lg p-2">
                          <div className="text-[10px] text-text-dim mb-1">原文</div>
                          <div className="text-red-300/90 line-through break-words">
                            {c.beforePreview ? (c.beforePreview.length > 140 ? c.beforePreview.slice(0, 140) + '…' : c.beforePreview) : '—'}
                          </div>
                        </div>
                        <div className="bg-black/5 dark:bg-white/5 border border-black/10 dark:border-white/10 rounded-lg p-2">
                          <div className="text-[10px] text-text-dim mb-1">修改后</div>
                          <div className="text-success/90 break-words">
                            {c.afterPreview ? (c.afterPreview.length > 140 ? c.afterPreview.slice(0, 140) + '…' : c.afterPreview) : '—'}
                          </div>
                        </div>
                      </div>

                      <div className="mt-2 flex items-center gap-2">
                        <button
                          className="flex items-center gap-1.5 px-3 py-1.5 bg-success/16 text-success border border-success/25 hover:bg-success/22 text-xs rounded-lg transition-colors"
                          onClick={() => acceptChange(c.id)}
                          title="接受此修改"
                        >
                          <Check className="w-3.5 h-3.5" />
                          <span>接受</span>
                        </button>
                        <button
                          className="flex items-center gap-1.5 px-3 py-1.5 bg-black/10 dark:bg-white/10 hover:bg-error/12 text-text-secondary hover:text-error text-xs rounded-lg transition-colors border border-black/10 dark:border-white/10 hover:border-error/30"
                          onClick={() => rejectChange(c.id)}
                          title="拒绝此修改"
                        >
                          <Ban className="w-3.5 h-3.5" />
                          <span>拒绝</span>
                        </button>
                      </div>
                    </div>
                    )
                  })}
                </div>
              )}
            </div>

            {/* Word 原生修订（docx w:ins/w:del） */}
            <div className="px-4 py-3">
              <div className="flex items-center justify-between gap-2 mb-2">
                <div className="text-xs text-text-muted">文档原生修订（Track Changes）</div>
                {trackChanges.authors.length > 0 && (
                  <select
                    className="text-xs px-2 py-1 rounded bg-black/5 dark:bg-white/5 border border-black/10 dark:border-white/10 outline-none"
                    value={trackAuthorFilter}
                    onChange={(e) => setTrackAuthorFilter(e.target.value)}
                    title="按作者筛选"
                  >
                    <option value="all">全部作者</option>
                    {trackChanges.authors.map(a => (
                      <option key={a} value={a}>{a}</option>
                    ))}
                  </select>
                )}
              </div>

              {filteredTrackItems.length === 0 ? (
                <div className="py-4 text-center text-sm text-text-muted">当前文档未检测到原生修订</div>
              ) : (
                <div className="space-y-2">
                  {filteredTrackItems.slice(0, 200).map((t) => (
                    <div key={t.key} className="border border-border rounded-lg p-3 bg-black/5 dark:bg-white/5">
                      <div className="text-[10px] text-text-dim">
                        {t.type === 'insert' ? '插入' : '删除'}
                        {t.author ? ` · ${t.author}` : ''}
                        {t.date ? ` · ${t.date}` : ''}
                      </div>
                      <div className="mt-1 text-xs break-words">
                        {t.type === 'delete' ? (
                          <span className="text-red-300/90 line-through">{t.text || '—'}</span>
                        ) : (
                          <span className="text-success/90">{t.text || '—'}</span>
                        )}
                      </div>
                      <div className="mt-2 text-[10px] text-text-dim">
                        key: {t.key}
                      </div>
                    </div>
                  ))}
                  {filteredTrackItems.length > 200 && (
                    <div className="text-[10px] text-text-muted text-center py-2">
                      已展示前 200 条（共 {filteredTrackItems.length} 条）
                    </div>
                  )}
                </div>
              )}
            </div>
          </div>
        </div>
      </div>
    </>
  )
}



