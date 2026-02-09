import { 
  MessageSquare, 
  Settings,
  Download,
  Save,
  FolderOpen,
  Bold,
  Italic,
  Link,
  Share2,
  Sparkles
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
    <header className="h-14 glass border-b border-border flex items-center justify-between px-5 select-none app-drag-region z-30 relative">
      {/* 左侧区域：打开和文件名 */}
      <div className="flex items-center gap-4 no-drag">
        {/* 打开文件夹按钮 */}
        {isElectron && (
          <button
            onClick={openFolder}
            className="flex items-center gap-2 px-3 py-1.5 rounded-xl text-text-muted hover:text-text hover:bg-black/5 dark:hover:bg-white/5 transition-all"
            title="打开文件夹"
          >
            <FolderOpen className="w-4 h-4" />
            <span className="text-xs font-medium">打开</span>
          </button>
        )}

        <div className="h-5 w-px bg-black/10 dark:bg-white/10" />

        {/* 文件名和状态 */}
        <div className="flex items-center gap-3">
          <span className="text-sm font-medium text-text truncate max-w-[200px]">
            {document.title || '新建文档'}
          </span>
          {hasUnsavedChanges ? (
            <span className="flex items-center gap-1.5 text-[10px] px-2.5 py-1 rounded-full bg-warning/12 border border-warning/25 text-warning font-medium">
              <span className="w-1.5 h-1.5 rounded-full bg-warning animate-pulse" />
              未保存
            </span>
          ) : (
            <span className="flex items-center gap-1.5 text-[10px] px-2.5 py-1 rounded-full bg-black/5 dark:bg-white/5 border border-black/10 dark:border-white/10 text-text-dim font-medium">
              已保存
            </span>
          )}
        </div>
      </div>

      {/* 中间工具栏 - 参考图一的居中设计 */}
      <div className="absolute left-1/2 -translate-x-1/2 flex items-center gap-1 no-drag">
        <div className="flex items-center gap-0.5 p-1.5 rounded-2xl glass-card-soft">
          <button
            className="p-2 rounded-xl text-text-muted hover:text-text hover:bg-black/5 dark:hover:bg-white/8 transition-all"
            title="粗体"
          >
            <Bold className="w-4 h-4" />
          </button>
          <button
            className="p-2 rounded-xl text-text-muted hover:text-text hover:bg-black/5 dark:hover:bg-white/8 transition-all"
            title="斜体"
          >
            <Italic className="w-4 h-4" />
          </button>
          <button
            className="p-2 rounded-xl text-text-muted hover:text-text hover:bg-black/5 dark:hover:bg-white/8 transition-all"
            title="链接"
          >
            <Link className="w-4 h-4" />
          </button>
          
          <div className="w-px h-5 bg-black/10 dark:bg-white/10 mx-1.5" />
          
          {/* AI Rewrite 按钮 - 参考图一的设计 */}
          <button
            className="flex items-center gap-2 px-3 py-1.5 rounded-xl bg-black/5 dark:bg-white/5 hover:bg-black/10 dark:hover:bg-white/10 transition-all group border border-black/10 dark:border-white/10 hover:border-accent/30"
            title="AI 重写"
          >
            <div className="w-5 h-5 rounded-lg bg-gradient-to-br from-accent to-accent-hover flex items-center justify-center shadow-sm shadow-accent/20">
              <Sparkles className="w-3 h-3 text-white" />
            </div>
            <span className="text-xs text-text-muted group-hover:text-text font-medium">AI Rewrite</span>
          </button>
        </div>
      </div>

      {/* 右侧区域：工具栏 */}
      <div className="flex items-center gap-2 no-drag">
        {/* 状态指示器 - 柔和风格 */}
        <div className="status-indicator online mr-2 hidden sm:flex">
          ONLINE
        </div>

        <div className="h-5 w-px bg-black/10 dark:bg-white/10" />

        {/* AI 聊天切换 */}
        <button
          onClick={onToggleChat}
          className={`p-2.5 rounded-xl transition-all ${
            showChat 
              ? 'text-accent bg-accent/12 border border-accent/25' 
              : 'text-text-muted hover:text-text hover:bg-black/5 dark:hover:bg-white/5 border border-transparent'
          }`}
          title={showChat ? '隐藏AI助手' : '显示AI助手'}
        >
          <MessageSquare className="w-4 h-4" />
        </button>

        <div className="h-5 w-px bg-black/10 dark:bg-white/10" />

        {/* 保存按钮 */}
        <button
          onClick={handleSave}
          disabled={isSaving || !hasUnsavedChanges}
          className={`flex items-center gap-2 px-4 py-2 text-xs font-semibold rounded-xl transition-all ${
            hasUnsavedChanges
              ? 'btn-sage'
              : 'bg-black/5 dark:bg-white/5 text-text-dim border border-black/10 dark:border-white/10'
          } disabled:opacity-50 disabled:cursor-not-allowed disabled:transform-none`}
        >
          {isSaving ? (
            <div className="w-4 h-4 border-2 border-white/30 border-t-white rounded-full animate-spin" />
          ) : (
            <Save className="w-4 h-4" />
          )}
          <span>{isSaving ? '保存中' : '保存'}</span>
        </button>

        {/* 导出按钮 */}
        <button
          onClick={handleSave}
          className="flex items-center gap-2 px-4 py-2 glass-button text-text text-xs font-semibold rounded-xl"
        >
          <Download className="w-4 h-4" />
          <span>导出</span>
        </button>

        {/* 分享按钮 */}
        <button
          className="p-2.5 rounded-xl text-text-muted hover:text-text hover:bg-white/5 transition-all hidden md:inline-flex"
          title="分享"
        >
          <Share2 className="w-4 h-4" />
        </button>

        {/* 设置按钮 */}
        <button
          onClick={onOpenSettings}
          className="p-2.5 rounded-xl text-text-muted hover:text-text hover:bg-black/5 dark:hover:bg-white/5 transition-all"
          title="设置"
        >
          <Settings className="w-4 h-4" />
        </button>
      </div>
    </header>
  )
}
