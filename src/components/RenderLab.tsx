import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import WordEditor from './WordEditor'
import { useDocument } from '../context/useDocument'

type RenderLabStatus = 'idle' | 'loading-fixture' | 'rendering' | 'ready' | 'error'

declare global {
  interface Window {
    __wordCursorLayoutDebug?: {
      status?: 'loading' | 'ready'
      pageCount?: number
      layoutResult?: { pages?: unknown[] }
      scale?: number
    }
    __wordCursorRenderLab?: {
      status: RenderLabStatus
      fixture: string
      error: string | null
      currentFile: string | null
      pageCount: number
      waitUntilReady: (timeoutMs?: number) => Promise<{
        fixture: string
        currentFile: string | null
        pageCount: number
      }>
      reloadFixture: (fixture?: string) => Promise<void>
      getSnapshot: () => {
        status: RenderLabStatus
        fixture: string
        error: string | null
        currentFile: string | null
        pageCount: number
        renderMode: 'canvas' | 'tiptap' | 'unknown'
        layoutStatus: string
      }
    }
  }
}

function getSearchParam(name: string): string {
  if (typeof window === 'undefined') return ''
  return new URLSearchParams(window.location.search).get(name)?.trim() || ''
}

function normalizeFixturePath(value: string): string {
  const raw = (value || '').trim()
  if (!raw) return '/render-lab/fixtures/render-lab-sample-cn.docx'
  if (/^https?:\/\//i.test(raw)) return raw
  if (raw.startsWith('/')) return raw
  return `/render-lab/fixtures/${raw}`
}

function resolveFileNameFromPath(path: string): string {
  const clean = path.split('?')[0].split('#')[0]
  const fileName = clean.split('/').pop() || 'render-lab.docx'
  return fileName.toLowerCase().endsWith('.docx') ? fileName : `${fileName}.docx`
}

async function fetchFixtureFile(fixturePath: string): Promise<File> {
  const response = await fetch(fixturePath, { cache: 'no-store' })
  if (!response.ok) {
    throw new Error(`获取 fixture 失败（HTTP ${response.status}）`)
  }
  const blob = await response.blob()
  return new File([blob], resolveFileNameFromPath(fixturePath), {
    type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  })
}

function getLayoutPageCount(): number {
  const directCount = window.__wordCursorLayoutDebug?.pageCount
  if (typeof directCount === 'number' && directCount > 0) return directCount
  const pages = window.__wordCursorLayoutDebug?.layoutResult?.pages
  return Array.isArray(pages) ? pages.length : 0
}

function getEditorPageCount(): number {
  const total = window.__wordCursorWordEditorDebug?.pageStats?.total
  return typeof total === 'number' && total > 0 ? total : 0
}

function getDomPageCount(): number {
  const breaks = document.querySelectorAll('.word-editor-content .pm-page-break').length
  return breaks > 0 ? breaks + 1 : 0
}

function getRenderMode(): 'canvas' | 'tiptap' | 'unknown' {
  const mode = window.__wordCursorWordEditorDebug?.printRenderMode
  return mode === 'canvas' || mode === 'tiptap' ? mode : 'unknown'
}

function getResolvedPageCount(): number {
  return getLayoutPageCount() || getDomPageCount() || getEditorPageCount()
}

function parseCssDimension(value: string): number {
  const raw = String(value || '').trim()
  if (!raw) return 0
  const numeric = Number.parseFloat(raw)
  if (!Number.isFinite(numeric)) return 0
  if (raw.endsWith('mm')) return numeric * 96 / 25.4
  if (raw.endsWith('cm')) return numeric * 96 / 2.54
  if (raw.endsWith('in')) return numeric * 96
  if (raw.endsWith('pt')) return numeric * 96 / 72
  return numeric
}

function isRenderReady(currentFileName: string | null): boolean {
  if (!currentFileName) return false

  const renderMode = getRenderMode()
  const pageCount = getResolvedPageCount()
  if (pageCount <= 0) return false

  if (renderMode === 'canvas') {
    return window.__wordCursorLayoutDebug?.status === 'ready'
  }

  if (renderMode === 'tiptap') {
    const editorDebug = window.__wordCursorWordEditorDebug
    const pageEl = document.querySelector('[data-testid="word-render-tiptap-page"]') as HTMLElement | null
    const contentEl = pageEl?.querySelector('.word-editor-content') as HTMLElement | null
    const pageHeight = pageEl ? parseCssDimension(getComputedStyle(pageEl).getPropertyValue('--page-height')) : 0
    const contentHeight = contentEl?.scrollHeight || contentEl?.getBoundingClientRect().height || 0
    const needsPagination = pageHeight > 0 && contentHeight > pageHeight * 1.15
    const hasExplicitPagination = getDomPageCount() > 1 || getEditorPageCount() > 1

    return !!(
      editorDebug &&
      editorDebug.isLoading === false &&
      editorDebug.hasDocxData &&
      contentEl &&
      (!needsPagination || hasExplicitPagination)
    )
  }

  return false
}

export default function RenderLab() {
  const { uploadDocxFile, currentFile, setEditorMode } = useDocument()
  const initialFixture = useMemo(() => normalizeFixturePath(getSearchParam('fixture')), [])
  const [fixture, setFixture] = useState(initialFixture)
  const [status, setStatus] = useState<RenderLabStatus>('idle')
  const [error, setError] = useState<string | null>(null)
  const [pageCount, setPageCount] = useState(0)
  const stablePageStateRef = useRef<{ count: number; since: number }>({ count: 0, since: 0 })

  const loadFixture = useCallback(async (nextFixture?: string) => {
    const resolvedFixture = normalizeFixturePath(nextFixture || initialFixture)
    setFixture(resolvedFixture)
    setStatus('loading-fixture')
    setError(null)

    try {
      setEditorMode('tiptap')
      const file = await fetchFixtureFile(resolvedFixture)
      await uploadDocxFile(file)
      setStatus('rendering')
    } catch (fixtureError) {
      console.error('[RenderLab] fixture 加载失败:', fixtureError)
      setError((fixtureError as Error).message || '未知错误')
      setStatus('error')
    }
  }, [initialFixture, setEditorMode, uploadDocxFile])

  useEffect(() => {
    void loadFixture(initialFixture)
  }, [initialFixture, loadFixture])

  useEffect(() => {
    const timer = window.setInterval(() => {
      const nextPageCount = getResolvedPageCount()
      setPageCount((prev) => (prev === nextPageCount ? prev : nextPageCount))
    }, 120)

    return () => window.clearInterval(timer)
  }, [])

  useEffect(() => {
    if (status !== 'loading-fixture' && status !== 'rendering') return

    const timer = window.setInterval(() => {
      const nextPageCount = getResolvedPageCount()
      const ready = isRenderReady(currentFile?.name || null)

      if (!ready || nextPageCount <= 0) {
        stablePageStateRef.current = { count: 0, since: 0 }
        return
      }

      const now = Date.now()
      if (stablePageStateRef.current.count !== nextPageCount) {
        stablePageStateRef.current = { count: nextPageCount, since: now }
        return
      }

      if (now - stablePageStateRef.current.since >= 480) {
        setStatus('ready')
      }
    }, 120)

    return () => window.clearInterval(timer)
  }, [currentFile, status])

  const waitUntilReady = useCallback(async (timeoutMs = 30000) => {
    const deadline = Date.now() + timeoutMs
    let stableCount = 0
    let stableSince = 0

    while (Date.now() < deadline) {
      const nextPageCount = getResolvedPageCount()
      const ready = status === 'ready' || isRenderReady(currentFile?.name || null)
      if (ready && currentFile && nextPageCount > 0) {
        const now = Date.now()
        if (stableCount !== nextPageCount) {
          stableCount = nextPageCount
          stableSince = now
        } else if (now - stableSince >= 480) {
          return {
            fixture,
            currentFile: currentFile.name,
            pageCount: nextPageCount,
          }
        }
      } else {
        stableCount = 0
        stableSince = 0
      }
      await new Promise((resolve) => window.setTimeout(resolve, 120))
    }

    throw new Error(`等待渲染完成超时：${fixture}`)
  }, [currentFile, fixture, status])

  useEffect(() => {
    const api = {
      status,
      fixture,
      error,
      currentFile: currentFile?.name || null,
      pageCount,
      waitUntilReady,
      reloadFixture: (nextFixture?: string) => loadFixture(nextFixture || fixture),
      getSnapshot: () => ({
        status,
        fixture,
        error,
        currentFile: currentFile?.name || null,
        pageCount: getResolvedPageCount(),
        renderMode: getRenderMode(),
        layoutStatus: window.__wordCursorLayoutDebug?.status || 'unknown',
      }),
    }

    window.__wordCursorRenderLab = api
    return () => {
      if (window.__wordCursorRenderLab === api) {
        delete window.__wordCursorRenderLab
      }
    }
  }, [currentFile?.name, error, fixture, loadFixture, pageCount, status, waitUntilReady])

  return (
    <div
      className="h-screen w-screen overflow-hidden bg-[var(--word-canvas-bg)]"
      data-testid="render-lab-root"
      data-render-lab-status={status}
    >
      <div className="flex items-center justify-between gap-4 px-4 py-2 border-b border-black/10 bg-white/70 backdrop-blur-sm">
        <div className="min-w-0">
          <div className="text-sm font-medium text-text">Word Render Lab</div>
          <div className="text-xs text-text-muted truncate">
            fixture: {fixture}
          </div>
        </div>
        <div className="flex items-center gap-3 text-xs text-text-secondary">
          <span data-testid="render-lab-status">状态：{status}</span>
          <span data-testid="render-lab-current-file">文件：{currentFile?.name || '未加载'}</span>
          <span data-testid="render-lab-page-count">页数：{pageCount}</span>
          {error ? <span className="text-red-600">错误：{error}</span> : null}
        </div>
      </div>

      <div className="h-[calc(100vh-45px)] overflow-hidden" data-testid="render-lab-editor-region">
        <WordEditor />
      </div>
    </div>
  )
}
