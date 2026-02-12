import { useState, useRef } from 'react'
import { 
  ChevronRight, 
  ChevronDown, 
  FileText, 
  FileCode,
  FileImage,
  FileSpreadsheet,
  Folder, 
  FolderOpen,
  Plus,
  Search,
  Upload,
  FilePlus,
  FolderPlus,
  RefreshCw,
  Settings,
  Image,
  Clock,
  Pencil
} from 'lucide-react'
import { useDocument } from '../context/DocumentContext'
import { FileItem } from '../types'

interface FileTreeItemProps {
  item: FileItem
  level: number
  onSelect: (item: FileItem) => void
  onDragStart: (item: FileItem) => void
  selectedPath: string | null
}

function FileTreeItem({ item, level, onSelect, onDragStart, selectedPath }: FileTreeItemProps) {
  const [isExpanded, setIsExpanded] = useState(level < 2)
  const isFolder = item.type === 'folder'
  const isSelected = item.path === selectedPath

  const handleDragStart = (e: React.DragEvent) => {
    e.dataTransfer.setData('application/json', JSON.stringify(item))
    e.dataTransfer.effectAllowed = 'copy'
    onDragStart(item)
  }

  // Finder 风格：按文件类型给 icon 轻量点缀色（避免整体太单一）
  const getFileVisual = (name: string): { Icon: React.ElementType; fg: string; bg: string } => {
    const lower = name.toLowerCase()

    // Images
    if (/\.(png|jpg|jpeg|gif|webp|svg)$/.test(lower)) {
      return { Icon: FileImage, fg: 'text-system-purple', bg: 'bg-system-purple/12' }
    }
    // PDF
    if (lower.endsWith('.pdf')) {
      return { Icon: FileText, fg: 'text-error', bg: 'bg-error/12' }
    }
    // PPT
    if (/\.(pptx|ppt)$/.test(lower)) {
      return { Icon: Image, fg: 'text-warning', bg: 'bg-warning/12' }
    }
    // Excel
    if (/\.(xlsx|xls|csv)$/.test(lower)) {
      return { Icon: FileSpreadsheet, fg: 'text-success', bg: 'bg-success/12' }
    }
    // Code-like
    if (/\.(ts|tsx|js|jsx|json|yml|yaml|py|go|rs|java|md|txt)$/.test(lower)) {
      return { Icon: FileCode, fg: 'text-system-gray', bg: 'bg-system-gray/12' }
    }
    // Default docs
    return { Icon: FileText, fg: 'text-accent', bg: 'bg-accent/12' }
  }

  return (
    <div>
      <div
        draggable
        onDragStart={handleDragStart}
        className={`group flex items-center gap-2 px-3 py-2 mx-2 cursor-pointer rounded-lg transition-all duration-200 select-none ${
          isSelected
            ? 'bg-accent/18 text-text border border-accent/25 shadow-sm'
            : 'text-text-secondary hover:bg-black/5 dark:hover:bg-white/10 hover:text-text'
        } cursor-grab active:cursor-grabbing`}
        style={{ paddingLeft: `${level * 12 + 12}px` }}
        onClick={() => {
          if (isFolder) {
            setIsExpanded(!isExpanded)
          } else {
            onSelect(item)
          }
        }}
      >
        <div className={`flex items-center justify-center w-4 h-4 shrink-0 transition-colors ${isSelected ? 'text-white' : 'text-text-muted group-hover:text-text-secondary'}`}>
          {isFolder && (
            <ChevronRight className={`w-3.5 h-3.5 transition-transform duration-200 ${isExpanded ? 'rotate-90' : ''}`} />
          )}
        </div>
        
        {isFolder ? (
          <div
            className={`w-6 h-6 rounded-lg flex items-center justify-center ${
              isSelected ? 'bg-white/18 text-white' : 'bg-accent/12 text-accent'
            }`}
          >
            {isExpanded ? <FolderOpen className="w-4 h-4" /> : <Folder className="w-4 h-4" />}
          </div>
        ) : (
          (() => {
            const { Icon, fg, bg } = getFileVisual(item.name)
            return (
              <div
                className={`w-6 h-6 rounded-lg flex items-center justify-center ${
                  isSelected ? 'bg-white/18 text-white' : `${bg} ${fg}`
                }`}
              >
                <Icon className="w-4 h-4" />
              </div>
            )
          })()
        )}
        
        <span className="text-[13px] truncate flex-1 font-medium">{item.name}</span>
      </div>
      
      {isFolder && isExpanded && item.children && (
        <div className="relative">
          {/* 连接线 - 更淡更细 */}
          <div 
            className="absolute left-[19px] top-0 bottom-0 w-px bg-border-light" 
            style={{ left: `${level * 12 + 19}px` }}
          />
          {item.children.map((child, index) => (
            <FileTreeItem
              key={`${child.path}-${index}`}
              item={child}
              level={level + 1}
              onSelect={onSelect}
              onDragStart={onDragStart}
              selectedPath={selectedPath}
            />
          ))}
        </div>
      )}
    </div>
  )
}

export default function Sidebar() {
  const { 
    files, 
    currentFile, 
    workspacePath,
    isElectron,
    openFolder, 
    openFile,
    uploadDocxFile, 
    createNewDocument,
    refreshFiles
  } = useDocument()
  
  const [searchQuery, setSearchQuery] = useState('')
  const [isCollapsed, setIsCollapsed] = useState(false)
  const [isDragging, setIsDragging] = useState(false)
  const [draggedFile, setDraggedFile] = useState<FileItem | null>(null)
  const [activeSection, setActiveSection] = useState<'workspace' | 'recent'>('workspace')
  const fileInputRef = useRef<HTMLInputElement>(null)

  // 打开设置（由 App.tsx 监听 open-settings 事件来弹出 SettingsModal）
  const openSettings = () => {
    window.dispatchEvent(new CustomEvent('open-settings'))
  }

  const handleSelectFile = async (item: FileItem) => {
    if (item.type === 'file') {
      await openFile(item)
    }
  }

  const handleDragStart = (item: FileItem) => {
    setDraggedFile(item)
  }

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]
    if (file && (file.name.endsWith('.docx') || file.name.endsWith('.doc'))) {
      try {
        await uploadDocxFile(file)
      } catch (error) {
        console.error('Upload failed:', error)
        alert('文件上传失败，请重试')
      }
    }
    e.target.value = ''
  }

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault()
    if (!isElectron) {
      setIsDragging(true)
    }
  }

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault()
    setIsDragging(false)
  }

  const handleDrop = async (e: React.DragEvent) => {
    e.preventDefault()
    setIsDragging(false)
    
    if (!isElectron) {
      const file = e.dataTransfer.files?.[0]
      if (file && (file.name.endsWith('.docx') || file.name.endsWith('.doc'))) {
        try {
          await uploadDocxFile(file)
        } catch (error) {
          console.error('Upload failed:', error)
          alert('文件上传失败，请重试')
        }
      }
    }
  }

  const handleNewDocument = () => {
    const title = `新文档_${Date.now()}`
    createNewDocument(title, `# ${title}\n\n在这里开始编写你的文档...`)
  }

  const handleNewPpt = () => {
    if (!isElectron) {
      alert('新建 PPT 仅支持桌面版（Electron）')
      return
    }
    const topic = window.prompt('请输入 PPT 主题/需求（将自动生成整套 PPT 海报页，含文字排版）：')
    if (!topic) return
    const countStr = window.prompt('请输入页数（默认 12，建议 10-15）：', '12') || '12'
    const slideCount = Math.max(1, Math.min(30, parseInt(countStr, 10) || 12))
    window.dispatchEvent(new CustomEvent('ppt-create-request', { detail: { topic, slideCount } }))
  }

  if (isCollapsed) {
    return (
      <div className="w-16 glass border-r border-border flex flex-col items-center py-4 gap-2">
        <button
          onClick={() => setIsCollapsed(false)}
          className="p-2.5 rounded-xl text-text-muted hover:text-text hover:bg-white/5 transition-all"
        >
          <ChevronRight className="w-4 h-4" />
        </button>
        <div className="w-8 h-px bg-white/10 my-2" />
        {isElectron && (
          <button 
            onClick={openFolder}
            className="p-2.5 rounded-xl text-text-muted hover:text-accent hover:bg-accent/10 transition-all"
            title="打开文件夹"
          >
            <FolderPlus className="w-4 h-4" />
          </button>
        )}
        {isElectron && (
          <button 
            onClick={handleNewPpt}
            className="p-2.5 rounded-xl text-text-muted hover:text-accent hover:bg-accent/10 transition-all"
            title="新建 PPT"
          >
            <Image className="w-4 h-4" />
          </button>
        )}
        <button 
          onClick={handleNewDocument}
          className="p-2.5 rounded-xl text-text-muted hover:text-accent hover:bg-accent/10 transition-all"
          title="新建文档"
        >
          <FilePlus className="w-4 h-4" />
        </button>

        <div className="flex-1" />

        <button 
          onClick={openSettings}
          className="p-2.5 rounded-xl text-text-muted hover:text-text hover:bg-white/5 transition-all"
          title="设置"
        >
          <Settings className="w-4 h-4" />
        </button>
      </div>
    )
  }

  return (
    <div 
      className={`w-64 glass-panel flex flex-col transition-all duration-300 ease-out z-20 ${
        isDragging ? 'ring-2 ring-accent/40 ring-inset bg-accent/5' : ''
      }`}
      onDragOver={handleDragOver}
      onDragLeave={handleDragLeave}
      onDrop={handleDrop}
    >
      <input
        ref={fileInputRef}
        type="file"
        accept=".docx,.doc"
        className="hidden"
        onChange={handleFileUpload}
      />

      {/* Logo 和品牌 - 极简风格 */}
      <div className="px-5 pt-5 pb-3">
        <div className="flex items-center gap-3">
          <div className="w-8 h-8 rounded-lg bg-accent flex items-center justify-center shadow-md shadow-accent/20">
            <Pencil className="w-4 h-4 text-white" strokeWidth={2.5} />
          </div>
          <div>
            <h1 className="text-sm font-bold text-text tracking-wide">Word Cursor</h1>
          </div>
          <button
            onClick={() => setIsCollapsed(true)}
            className="ml-auto p-1.5 rounded-md text-text-muted hover:text-text hover:bg-black/5 dark:hover:bg-white/10 transition-all"
          >
            <ChevronRight className="w-3.5 h-3.5 rotate-180" />
          </button>
        </div>
      </div>

      {/* 搜索框 */}
      <div className="px-3 pb-2">
        <div className="relative group">
          <Search className="absolute left-2.5 top-1/2 -translate-y-1/2 w-3.5 h-3.5 text-text-muted group-focus-within:text-accent transition-colors" />
          <input
            type="text"
            placeholder="Search..."
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            className="w-full bg-black/5 dark:bg-white/10 hover:bg-black/10 dark:hover:bg-white/15 focus:bg-white dark:focus:bg-white/10 rounded-lg pl-8 pr-3 py-1.5 text-xs text-text placeholder-text-dim border border-transparent focus:border-accent/30 focus:ring-2 focus:ring-accent/10 transition-all outline-none"
          />
        </div>
      </div>

      {/* 分区标题 - WORKSPACE */}
      <div className="px-4 py-2 flex items-center justify-between group">
        <span className="text-[10px] font-bold text-text-muted tracking-wider uppercase">
          Workspace
        </span>
        <div className="flex opacity-0 group-hover:opacity-100 transition-opacity gap-0.5">
          {isElectron && (
             <button 
               onClick={refreshFiles}
               className="p-1 rounded hover:bg-black/5 dark:hover:bg-white/10 text-text-muted hover:text-text transition-colors"
               title="Refresh"
             >
               <RefreshCw className="w-3 h-3" />
             </button>
          )}
          <button 
             onClick={handleNewDocument}
             className="p-1 rounded hover:bg-black/5 dark:hover:bg-white/10 text-text-muted hover:text-text transition-colors"
             title="New Document"
           >
             <Plus className="w-3 h-3" />
           </button>
        </div>
      </div>


      {/* 拖放提示 */}
      {isDragging && !isElectron && (
        <div className="mx-3 mb-3 p-6 border-2 border-dashed border-accent/40 rounded-xl bg-accent/5 flex flex-col items-center justify-center gap-2">
          <Upload className="w-8 h-8 text-accent" />
          <span className="text-sm text-accent font-medium">释放以上传文档</span>
        </div>
      )}

      {/* 文件列表 */}
      <div className="flex-1 overflow-y-auto py-1 chat-scrollbar">
        {files.length > 0 ? (
          files.map((item, index) => (
            <FileTreeItem
              key={`${item.path}-${index}`}
              item={item}
              level={0}
              onSelect={handleSelectFile}
              onDragStart={handleDragStart}
              selectedPath={currentFile?.path || null}
            />
          ))
        ) : (
          <div className="flex flex-col items-center justify-center h-full text-center px-6 py-8">
            <div className="w-12 h-12 rounded-xl bg-black/5 dark:bg-white/5 flex items-center justify-center mb-3">
              <FolderOpen className="w-6 h-6 text-text-dim" />
            </div>
            <p className="text-sm text-text-secondary mb-1 font-medium">
              {isElectron
                ? (workspacePath ? '文件夹为空' : '没有打开的文件夹')
                : '没有文档'}
            </p>
            <p className="text-xs text-text-muted mb-4 leading-relaxed max-w-[180px]">
              {isElectron 
                ? (workspacePath
                    ? '该文件夹内暂无可用文件'
                    : '点击上方按钮打开一个本地文件夹')
                : '上传一个 .docx 文件开始编辑'}
            </p>
            {isElectron ? (
              <button
                onClick={openFolder}
                className="flex items-center justify-center gap-2 w-full px-3 py-2 bg-accent text-white text-xs font-medium rounded-lg hover:bg-accent-hover transition-colors shadow-sm"
                title="打开文件夹"
              >
                <FolderPlus className="w-3.5 h-3.5" />
                打开文件夹
              </button>
            ) : (
              <button
                onClick={() => fileInputRef.current?.click()}
                className="flex items-center justify-center gap-2 w-full px-3 py-2 bg-accent text-white text-xs font-medium rounded-lg hover:bg-accent-hover transition-colors shadow-sm"
              >
                <Upload className="w-3.5 h-3.5" />
                上传文档
              </button>
            )}
          </div>
        )}
      </div>

      {/* 分区标题 - RECENT */}
      <div className="px-4 py-2 border-t border-border-light">
        <span className="text-[10px] font-bold text-text-muted tracking-wider uppercase">
          Recent
        </span>
      </div>

      {/* 最近文件列表 - 简化展示 */}
      <div className="px-2 pb-2">
        {currentFile && (
          <div className="flex items-center gap-3 px-3 py-2 mx-1 rounded-lg bg-accent/10 border border-accent/20">
            <FileText className="w-4 h-4 text-accent" />
            <div className="flex-1 min-w-0">
              <p className="text-[13px] font-medium text-text truncate">{currentFile.name}</p>
              <p className="text-[10px] text-text-muted">刚刚编辑</p>
            </div>
            <div className="w-1.5 h-1.5 rounded-full bg-accent animate-pulse" />
          </div>
        )}
      </div>

      {/* 底部用户信息 - 参考图一风格 */}
      <div className="p-2 border-t border-border-light">
        <div
          className="flex items-center gap-3 p-2 rounded-lg hover:bg-black/5 dark:hover:bg-white/5 transition-all cursor-pointer group"
          onClick={openSettings}
          title="打开设置"
        >
          <div className="w-8 h-8 rounded-full bg-accent flex items-center justify-center shadow-sm text-white font-medium text-xs">
            AC
          </div>
          <div className="flex-1 min-w-0">
            <p className="text-xs font-medium text-text truncate">Alex Chen</p>
            <p className="text-[10px] text-text-muted">Pro Workspace</p>
          </div>
          <Settings className="w-3.5 h-3.5 text-text-muted group-hover:text-text transition-colors" />
        </div>
      </div>

    </div>
  )
}
