import { useMemo, useState } from 'react'
import { X, Check, Trash2, CornerDownRight, MessageSquarePlus } from 'lucide-react'
import { useComments } from '../context/CommentContext'

type CommentPanelProps = {
  open: boolean
  onClose: () => void
  onAddComment?: (text: string) => void
  onReply?: (parentId: string, text: string) => void
  onLocate?: (commentId: string) => void
}

export default function CommentPanel({ open, onClose, onAddComment, onReply, onLocate }: CommentPanelProps) {
  const { comments, selectedCommentId, setSelectedCommentId, resolveComment, deleteComment } = useComments()
  const [newCommentText, setNewCommentText] = useState('')
  const [replyDraft, setReplyDraft] = useState<Record<string, string>>({})

  const tree = useMemo(() => {
    const byParent: Record<string, typeof comments> = {}
    const roots: typeof comments = []
    for (const c of comments) {
      if (c.parentId) {
        if (!byParent[c.parentId]) byParent[c.parentId] = []
        byParent[c.parentId].push(c)
      } else {
        roots.push(c)
      }
    }
    const sortByDate = (a: any, b: any) => String(a.date || '').localeCompare(String(b.date || ''))
    roots.sort(sortByDate)
    Object.values(byParent).forEach(list => list.sort(sortByDate))
    return { roots, byParent }
  }, [comments])

  if (!open) return null

  return (
    <>
      <div className="fixed inset-0 z-[9998] bg-black/40" onClick={onClose} />

      <div className="fixed right-0 top-0 bottom-0 z-[9999] w-[420px] max-w-[95vw] bg-surface backdrop-blur-xl border-l border-border shadow-2xl flex flex-col">
        <div className="flex items-center justify-between px-4 py-3 border-b border-border">
          <div className="flex flex-col">
            <div className="text-sm text-text font-medium">批注面板</div>
            <div className="text-xs text-text-muted">{comments.length} 条批注</div>
          </div>
          <button
            className="p-1.5 rounded hover:bg-black/5 dark:hover:bg-white/5 transition-colors"
            onClick={onClose}
            title="关闭"
          >
            <X className="w-4 h-4 text-text-secondary" />
          </button>
        </div>

        {/* quick add (selection-based add is handled by editor; this is just text entry) */}
        <div className="px-4 py-3 border-b border-border">
          <div className="text-[10px] text-text-dim mb-1">新增批注（会应用到当前选区）</div>
          <div className="flex gap-2">
            <textarea
              value={newCommentText}
              onChange={(e) => setNewCommentText(e.target.value)}
              placeholder="输入批注内容…"
              className="flex-1 h-[56px] resize-none text-xs px-3 py-2 rounded-lg bg-black/5 dark:bg-white/5 border border-black/10 dark:border-white/10 outline-none focus:ring-2 focus:ring-sage-500/40"
            />
            <button
              className="shrink-0 h-[56px] px-3 rounded-lg bg-sage-500/20 hover:bg-sage-500/28 border border-sage-500/30 text-sage-200 text-xs flex items-center gap-1.5"
              onClick={() => {
                const t = newCommentText.trim()
                if (!t) return
                onAddComment?.(t)
                setNewCommentText('')
              }}
              title="添加批注"
            >
              <MessageSquarePlus className="w-4 h-4" />
              添加
            </button>
          </div>
        </div>

        <div className="flex-1 overflow-auto">
          {comments.length === 0 ? (
            <div className="px-4 py-8 text-center text-sm text-text-muted">暂无批注</div>
          ) : (
            <div className="divide-y divide-border">
              {tree.roots.map((c) => {
                const replies = tree.byParent[c.id] || []
                const active = selectedCommentId === c.id
                return (
                  <div key={c.id} className="px-4 py-3">
                    <div
                      className={
                        'rounded-lg border px-3 py-2 transition-colors ' +
                        (active
                          ? 'border-sage-500/40 bg-sage-500/10'
                          : 'border-black/10 dark:border-white/10 bg-black/5 dark:bg-white/5')
                      }
                      onClick={() => setSelectedCommentId(c.id)}
                      role="button"
                      tabIndex={0}
                    >
                      <div className="flex items-start justify-between gap-2">
                        <div className="min-w-0">
                          <div className="text-[10px] text-text-dim">
                            {c.author || 'Unknown'}
                            {c.date ? ` · ${c.date}` : ''}
                            {c.resolved ? ' · 已解决' : ''}
                          </div>
                          <div className="text-xs text-text mt-1 whitespace-pre-wrap break-words">
                            {c.text || '—'}
                          </div>
                        </div>
                        <div className="shrink-0 flex items-center gap-1">
                          <button
                            className="p-1.5 rounded hover:bg-black/10 dark:hover:bg-white/10 transition-colors"
                            onClick={(e) => {
                              e.stopPropagation()
                              onLocate?.(c.id)
                            }}
                            title="定位到文档"
                          >
                            <CornerDownRight className="w-4 h-4 text-text-secondary" />
                          </button>
                        </div>
                      </div>

                      <div className="mt-2 flex items-center gap-2">
                        <button
                          className="flex items-center gap-1.5 px-3 py-1.5 bg-success/16 text-success border border-success/25 hover:bg-success/22 text-xs rounded-lg transition-colors"
                          onClick={(e) => {
                            e.stopPropagation()
                            resolveComment(c.id)
                          }}
                          title="标记为已解决"
                        >
                          <Check className="w-3.5 h-3.5" />
                          <span>解决</span>
                        </button>
                        <button
                          className="flex items-center gap-1.5 px-3 py-1.5 bg-black/10 dark:bg-white/10 hover:bg-error/12 text-text-secondary hover:text-error text-xs rounded-lg transition-colors border border-black/10 dark:border-white/10 hover:border-error/30"
                          onClick={(e) => {
                            e.stopPropagation()
                            deleteComment(c.id)
                          }}
                          title="删除"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                          <span>删除</span>
                        </button>
                      </div>

                      <div className="mt-2">
                        <div className="text-[10px] text-text-dim mb-1">回复</div>
                        <div className="flex gap-2">
                          <input
                            value={replyDraft[c.id] || ''}
                            onChange={(e) => setReplyDraft(prev => ({ ...prev, [c.id]: e.target.value }))}
                            placeholder="输入回复…"
                            className="flex-1 text-xs px-3 py-2 rounded-lg bg-black/5 dark:bg-white/5 border border-black/10 dark:border-white/10 outline-none focus:ring-2 focus:ring-sage-500/40"
                          />
                          <button
                            className="shrink-0 px-3 rounded-lg bg-sage-500/20 hover:bg-sage-500/28 border border-sage-500/30 text-sage-200 text-xs"
                            onClick={(e) => {
                              e.stopPropagation()
                              const t = (replyDraft[c.id] || '').trim()
                              if (!t) return
                              onReply?.(c.id, t)
                              setReplyDraft(prev => ({ ...prev, [c.id]: '' }))
                            }}
                          >
                            发送
                          </button>
                        </div>
                      </div>
                    </div>

                    {replies.length > 0 && (
                      <div className="mt-2 pl-4 space-y-2">
                        {replies.map(r => (
                          <div
                            key={r.id}
                            className="rounded-lg border border-black/10 dark:border-white/10 bg-black/5 dark:bg-white/5 px-3 py-2"
                            onClick={() => setSelectedCommentId(r.id)}
                            role="button"
                            tabIndex={0}
                          >
                            <div className="text-[10px] text-text-dim">
                              {r.author || 'Unknown'}
                              {r.date ? ` · ${r.date}` : ''}
                              {r.resolved ? ' · 已解决' : ''}
                            </div>
                            <div className="text-xs text-text mt-1 whitespace-pre-wrap break-words">{r.text || '—'}</div>
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                )
              })}
            </div>
          )}
        </div>
      </div>
    </>
  )
}











