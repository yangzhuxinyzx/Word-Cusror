import { 
  FileText, 
  MessageSquare, 
  Settings,
  Download,
  Save,
  FolderOpen
} from 'lucide-react'
import { useDocument } from '../context/DocumentContext'
import { useCallback, useState } from 'react'

interface HeaderProps {
  showChat: boolean
  showPreview: boolean
  activeView: 'editor' | 'preview' | 'split'
  onToggleChat: () => void
  onTogglePreview: () => void
  onViewChange: (view: 'editor' | 'preview' | 'split') => void
  onOpenSettings: () => void
}

export default function Header({
  showChat,
  onToggleChat,
  onOpenSettings,
}: HeaderProps) {
  const { document, saveDocument, hasUnsavedChanges, isElectron, openFolder } = useDocument()
  const [isSaving, setIsSaving] = useState(false)

  const handleSave = useCallback(async () => {
    setIsSaving(true)
    try {
      await saveDocument()
    } catch (error) {
      console.error('Save failed:', error)
    } finally {
      setIsSaving(false)
    }
  }, [saveDocument])

  return (
    <header className="h-12 bg-background border-b border-border flex items-center justify-between px-4 select-none app-drag-region z-20">
      {/* 左侧区域：Logo 和 文件名 */}
      <div className="flex items-center gap-4 no-drag">
        <div className="flex items-center gap-2 text-text hover:text-primary transition-colors cursor-pointer group">
          <div className="w-6 h-6 rounded bg-primary/10 flex items-center justify-center group-hover:bg-primary/20 transition-colors">
            <FileText className="w-3.5 h-3.5 text-primary" />
          </div>
          <span className="text-sm font-semibold tracking-tight">Word-Cursor</span>
          {isElectron && (
            <span className="text-[10px] px-1.5 py-0.5 rounded bg-primary/10 text-primary border border-primary/20">
              Desktop
            </span>
          )}
        </div>

        <div className="h-4 w-px bg-border/50" />

        {/* 打开文件夹按钮 */}
        {isElectron && (
          <button
            onClick={openFolder}
            className="flex items-center gap-1.5 px-2 py-1 text-text-muted hover:text-text hover:bg-surface-hover rounded-md transition-all"
            title="打开文件夹"
          >
            <FolderOpen className="w-4 h-4" />
            <span className="text-xs">打开</span>
          </button>
        )}

        <div className="flex items-center gap-2 group cursor-pointer">
          <span className="text-sm text-text-muted group-hover:text-text transition-colors truncate max-w-[300px]">
            {document.title}
          </span>
          {hasUnsavedChanges ? (
            <span className="text-[10px] px-1.5 py-0.5 rounded-full bg-amber-500/10 border border-amber-500/20 text-amber-400">
              未保存
            </span>
          ) : (
            <span className="text-[10px] px-1.5 py-0.5 rounded-full bg-surface border border-border text-text-dim">
              已保存
            </span>
          )}
        </div>
      </div>

      {/* 右侧区域：工具栏 */}
      <div className="flex items-center gap-2 no-drag">
        <button
          onClick={onToggleChat}
          className={`p-2 rounded-md transition-all ${
            showChat 
              ? 'text-primary bg-primary/10' 
              : 'text-text-muted hover:text-text hover:bg-surface-hover'
          }`}
          title={showChat ? '隐藏AI对话' : '显示AI对话'}
        >
          <MessageSquare className="w-4 h-4" />
        </button>

        <div className="h-4 w-px bg-border/50" />

        {/* 保存按钮 */}
        <button
          onClick={handleSave}
          disabled={isSaving || !hasUnsavedChanges}
          className={`flex items-center gap-2 px-3 py-1.5 text-xs font-medium rounded-md transition-all ${
            hasUnsavedChanges
              ? 'bg-primary hover:bg-primary-hover text-white shadow-glow'
              : 'bg-surface text-text-muted border border-border'
          } disabled:opacity-50 disabled:cursor-not-allowed`}
        >
          {isSaving ? (
            <div className="w-3.5 h-3.5 border-2 border-white/30 border-t-white rounded-full animate-spin" />
          ) : (
            <Save className="w-3.5 h-3.5" />
          )}
          <span>{isSaving ? '保存中' : '保存'}</span>
        </button>

        {/* 导出按钮 */}
        <button
          onClick={handleSave}
          className="flex items-center gap-2 px-3 py-1.5 bg-surface hover:bg-surface-hover border border-border text-text text-xs font-medium rounded-md transition-all"
        >
          <Download className="w-3.5 h-3.5" />
          <span>导出</span>
        </button>

        <button
          onClick={onOpenSettings}
          className="p-2 rounded-md text-text-muted hover:text-text hover:bg-surface-hover transition-all"
          title="设置"
        >
          <Settings className="w-4 h-4" />
        </button>
      </div>
    </header>
  )
}
