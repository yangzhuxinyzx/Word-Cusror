import React, { useMemo, useState, useEffect } from 'react'
import { 
  Plus,
  FileText,
  FolderOpen,
  Sparkles,
  Presentation,
  ArrowUpRight
} from 'lucide-react'
import { useDocument } from '../context/useDocument'
import { FileItem } from '../types'

export default function Dashboard() {
  const { 
    files, 
    openFile, 
    createNewDocument, 
    openFolder,
    isElectron,
    effectiveWorkspacePath,
    sessionMode,
  } = useDocument()

  const [greeting, setGreeting] = useState('')

  useEffect(() => {
    const hour = new Date().getHours()
    if (hour < 12) setGreeting('早上好')
    else if (hour < 18) setGreeting('下午好')
    else setGreeting('晚上好')
  }, [])

  const handleNewDoc = () => {
    const title = `新文档_${Date.now()}`
    createNewDocument(title, '')
  }

  const handleNewPPT = () => {
    if (isElectron) {
      const topic = window.prompt('请输入 PPT 主题:')
      if (topic) {
        window.dispatchEvent(new CustomEvent('ppt-create-request', { detail: { topic, slideCount: 12 } }))
      }
    } else {
      alert('PPT 生成仅支持桌面端')
    }
  }

  // Flatten files for display
  const displayFiles = useMemo(() => {
    const flatten = (items: FileItem[]): FileItem[] => {
      return items.reduce((acc, item) => {
        if (item.type === 'file') acc.push(item)
        if (item.children) acc.push(...flatten(item.children))
        return acc
      }, [] as FileItem[])
    }
    return flatten(files).slice(0, 6)
  }, [files])

  return (
    <div className="flex-1 h-full overflow-y-auto bg-transparent relative">
      {/* Main Content Container (keep center mostly empty for preview) */}
      <div className="p-8 max-w-5xl mx-auto h-full flex flex-col">
        {/* Top: greeting + compact actions */}
        <div className="pt-2">
          <h1 className="text-3xl font-bold text-text tracking-tight">{greeting}</h1>
          <p className="text-text-secondary text-sm mt-2">
            选择一个文件开始编辑，或创建新文档。中间区域保持留白，用作预览/编辑承载区。
          </p>
          {effectiveWorkspacePath && (
            <p className="text-text-muted text-xs mt-2">
              {sessionMode === 'single-file'
                ? `当前处于轻工作区：${effectiveWorkspacePath}`
                : `当前工作区：${effectiveWorkspacePath}`}
            </p>
          )}

          {/* Compact pill actions (macOS-like) */}
          <div className="mt-4 flex flex-wrap gap-2">
            <button
              onClick={handleNewDoc}
              className="inline-flex items-center gap-2 px-3.5 py-2 rounded-full bg-black/5 dark:bg-white/5 hover:bg-black/10 dark:hover:bg-white/10 border border-black/10 dark:border-white/10 text-xs font-medium text-text-secondary hover:text-text transition-colors"
            >
              <span className="w-6 h-6 rounded-full bg-accent/15 text-accent inline-flex items-center justify-center">
                <Plus className="w-3.5 h-3.5" />
              </span>
              新建文档
            </button>

            <button
              onClick={handleNewPPT}
              disabled={!isElectron}
              className="inline-flex items-center gap-2 px-3.5 py-2 rounded-full bg-black/5 dark:bg-white/5 hover:bg-black/10 dark:hover:bg-white/10 border border-black/10 dark:border-white/10 text-xs font-medium text-text-secondary hover:text-text transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
            >
              <span className="w-6 h-6 rounded-full bg-system-purple/15 text-system-purple inline-flex items-center justify-center">
                <Presentation className="w-3.5 h-3.5" />
              </span>
              新建 PPT
              <ArrowUpRight className="w-3.5 h-3.5 opacity-50" />
            </button>

            <button
              onClick={() => isElectron && openFolder()}
              disabled={!isElectron}
              className="inline-flex items-center gap-2 px-3.5 py-2 rounded-full bg-black/5 dark:bg-white/5 hover:bg-black/10 dark:hover:bg-white/10 border border-black/10 dark:border-white/10 text-xs font-medium text-text-secondary hover:text-text transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
            >
              <span className="w-6 h-6 rounded-full bg-system-teal/15 text-system-teal inline-flex items-center justify-center">
                <FolderOpen className="w-3.5 h-3.5" />
              </span>
              打开文件夹
            </button>
          </div>
        </div>

        {/* Center: preview placeholder (keep mostly empty) */}
        <div className="flex-1 flex items-center justify-center py-10">
          <div className="w-full max-w-3xl h-[56vh] rounded-3xl border border-border bg-black/5 dark:bg-white/5 backdrop-blur-xl shadow-glass-lg relative overflow-hidden">
            <div className="absolute inset-0 pointer-events-none opacity-60">
              <div className="absolute inset-x-10 top-10 h-px bg-black/10 dark:bg-white/10" />
              <div className="absolute inset-x-10 top-16 h-px bg-black/8 dark:bg-white/8" />
              <div className="absolute inset-x-10 top-22 h-px bg-black/6 dark:bg-white/6" />
            </div>
            <div className="h-full w-full flex flex-col items-center justify-center text-center px-10">
              <div className="w-14 h-14 rounded-2xl bg-black/5 dark:bg-white/5 flex items-center justify-center mb-4">
                <Sparkles className="w-7 h-7 text-text-muted" />
              </div>
              <div className="text-sm font-semibold text-text">预览区域</div>
              <div className="text-xs text-text-muted mt-1 max-w-sm">
                打开文件后，这里会显示编辑器/预览。当前保持留白，减少干扰。
              </div>
            </div>
          </div>
        </div>

        {/* Bottom: recent files (compact, optional) */}
        {displayFiles.length > 0 && (
          <div className="pb-2">
            <div className="text-[11px] font-semibold text-text-muted uppercase tracking-wider mb-2">最近文件</div>
            <div className="flex flex-wrap gap-2">
              {displayFiles.slice(0, 5).map((file) => (
                <button
                  key={file.path}
                  onClick={() => openFile(file)}
                  className="inline-flex items-center gap-2 px-3 py-2 rounded-full bg-black/5 dark:bg-white/5 hover:bg-black/10 dark:hover:bg-white/10 border border-black/10 dark:border-white/10 text-xs text-text-secondary hover:text-text transition-colors max-w-[260px]"
                  title={file.name}
                >
                  <FileText className="w-3.5 h-3.5 text-text-muted" />
                  <span className="truncate">{file.name}</span>
                </button>
              ))}
            </div>
          </div>
        )}
      </div>
    </div>
  )
}

