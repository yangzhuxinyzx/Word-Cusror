import { createContext, useCallback, useContext, useMemo, useRef, useState, type ReactNode } from 'react'

export type CommentItem = {
  id: string
  author?: string
  date?: string
  text: string
  parentId?: string
  resolved?: boolean
}

type CommentContextType = {
  comments: CommentItem[]
  selectedCommentId: string | null
  setSelectedCommentId: (id: string | null) => void
  setComments: (items: CommentItem[]) => void
  addComment: (text: string, meta?: { author?: string; date?: string; parentId?: string }) => string
  replyToComment: (parentId: string, text: string, meta?: { author?: string; date?: string }) => string
  resolveComment: (id: string) => void
  deleteComment: (id: string) => void
}

const CommentContext = createContext<CommentContextType | null>(null)

function safeToInt(v: string | undefined | null): number | null {
  if (!v) return null
  const n = Number(String(v).trim())
  return Number.isFinite(n) && n >= 0 ? Math.floor(n) : null
}

export function CommentProvider({ children }: { children: ReactNode }) {
  const [comments, setCommentsState] = useState<CommentItem[]>([])
  const [selectedCommentId, setSelectedCommentId] = useState<string | null>(null)
  const nextIdRef = useRef<number>(1)

  const setComments = useCallback((items: CommentItem[]) => {
    const list = Array.isArray(items) ? items : []
    setCommentsState(list)
    // update next numeric id (OOXML w:id requires decimal number)
    let max = 0
    for (const c of list) {
      const n = safeToInt(c.id)
      if (n !== null && n > max) max = n
    }
    nextIdRef.current = Math.max(1, max + 1)
  }, [])

  const newId = useCallback(() => {
    const id = String(nextIdRef.current)
    nextIdRef.current += 1
    return id
  }, [])

  const addComment = useCallback((text: string, meta?: { author?: string; date?: string; parentId?: string }) => {
    const id = newId()
    setCommentsState(prev => [
      ...prev,
      {
        id,
        text,
        author: meta?.author,
        date: meta?.date || new Date().toISOString(),
        parentId: meta?.parentId,
        resolved: false,
      },
    ])
    setSelectedCommentId(id)
    return id
  }, [newId])

  const replyToComment = useCallback((parentId: string, text: string, meta?: { author?: string; date?: string }) => {
    const id = newId()
    setCommentsState(prev => [
      ...prev,
      {
        id,
        text,
        author: meta?.author,
        date: meta?.date || new Date().toISOString(),
        parentId,
        resolved: false,
      },
    ])
    setSelectedCommentId(id)
    return id
  }, [newId])

  const resolveComment = useCallback((id: string) => {
    setCommentsState(prev => prev.map(c => (c.id === id ? { ...c, resolved: true } : c)))
  }, [])

  const deleteComment = useCallback((id: string) => {
    setCommentsState(prev => prev.filter(c => c.id !== id))
    setSelectedCommentId(prev => (prev === id ? null : prev))
  }, [])

  const value = useMemo<CommentContextType>(
    () => ({
      comments,
      selectedCommentId,
      setSelectedCommentId,
      setComments,
      addComment,
      replyToComment,
      resolveComment,
      deleteComment,
    }),
    [comments, selectedCommentId, setComments, addComment, replyToComment, resolveComment, deleteComment]
  )

  return <CommentContext.Provider value={value}>{children}</CommentContext.Provider>
}

export function useComments() {
  const ctx = useContext(CommentContext)
  if (!ctx) throw new Error('useComments must be used within CommentProvider')
  return ctx
}


