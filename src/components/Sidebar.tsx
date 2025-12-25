import { useState, useRef } from 'react'
import { 
  ChevronRight, 
  ChevronDown, 
  FileText, 
  Folder, 
  FolderOpen,
  Plus,
  Search,
  MoreHorizontal,
  Upload,
  FilePlus,
  FolderPlus,
  RefreshCw,
  ExternalLink
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
    if (item.type === 'file') {
      e.dataTransfer.setData('application/json', JSON.stringify(item))
      e.dataTransfer.effectAllowed = 'copy'
      onDragStart(item)
    }
  }

  return (
    <div>
      <div
        draggable={item.type === 'file'}
        onDragStart={handleDragStart}
        className={`group flex items-center gap-1.5 px-2 py-1.5 mx-2 cursor-pointer rounded-md transition-all select-none ${
          isSelected 
            ? 'bg-primary/15 text-primary border border-primary/20' 
            : 'text-text-muted hover:bg-surface-hover hover:text-text border border-transparent'
        } ${item.type === 'file' ? 'cursor-grab active:cursor-grabbing' : ''}`}
        style={{ paddingLeft: `${level * 12 + 8}px` }}
        onClick={() => {
          if (isFolder) {
            setIsExpanded(!isExpanded)
          } else {
            onSelect(item)
          }
        }}
      >
        <div className="flex items-center justify-center w-4 h-4 shrink-0 text-text-dim group-hover:text-text-muted transition-colors">
          {isFolder && (
            isExpanded ? <ChevronDown className="w-3 h-3" /> : <ChevronRight className="w-3 h-3" />
          )}
        </div>
        
        {isFolder ? (
          isExpanded ? 
            <FolderOpen className={`w-4 h-4 ${isSelected ? 'text-primary' : 'text-amber-500/70'}`} /> : 
            <Folder className={`w-4 h-4 ${isSelected ? 'text-primary' : 'text-amber-500/70'}`} />
        ) : (
          <FileText className={`w-4 h-4 ${isSelected ? 'text-primary' : 'text-blue-400/70'}`} />
        )}
        
        <span className="text-xs truncate flex-1 font-medium">{item.name}</span>
        
        {item.type === 'file' && (
          <span className="text-[10px] text-text-dim opacity-0 group-hover:opacity-100 transition-opacity">
            拖拽
          </span>
        )}
      </div>
      
      {isFolder && isExpanded && item.children && (
        <div className="relative">
          <div 
            className="absolute left-[15px] top-0 bottom-0 w-px bg-border/30" 
            style={{ left: `${level * 12 + 15}px` }}
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
  const fileInputRef = useRef<HTMLInputElement>(null)

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
      <div className="w-12 bg-background border-r border-border flex flex-col items-center py-4 gap-4">
        <button
          onClick={() => setIsCollapsed(false)}
          className="p-2 rounded-md text-text-muted hover:text-text hover:bg-surface-hover transition-all"
        >
          <ChevronRight className="w-4 h-4" />
        </button>
        <div className="w-6 h-px bg-border" />
        {isElectron && (
          <button 
            onClick={openFolder}
            className="p-2 rounded-md text-text-muted hover:text-primary hover:bg-primary/10 transition-all"
            title="打开文件夹"
          >
            <FolderPlus className="w-4 h-4" />
          </button>
        )}
        {isElectron && (
          <button 
            onClick={handleNewPpt}
            className="p-2 rounded-md text-text-muted hover:text-text hover:bg-surface-hover transition-all"
            title="新建 PPT（AI 海报式生成）"
          >
            <FileText className="w-4 h-4" />
          </button>
        )}
        <button 
          onClick={handleNewDocument}
          className="p-2 rounded-md text-text-muted hover:text-text hover:bg-surface-hover transition-all"
          title="新建文档"
        >
          <FilePlus className="w-4 h-4" />
        </button>
      </div>
    )
  }

  return (
    <div 
      className={`w-64 bg-background border-r border-border flex flex-col transition-all duration-300 ease-in-out ${
        isDragging ? 'ring-2 ring-primary ring-inset' : ''
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

      {/* 头部 */}
      <div className="p-3 space-y-3">
        <div className="flex items-center justify-between px-2">
          <span className="text-xs font-semibold text-text-muted uppercase tracking-wider">
            {isElectron ? '工作区' : '文档'}
          </span>
          <div className="flex items-center gap-1">
            {isElectron && (
              <>
                <button 
                  onClick={openFolder}
                  className="p-1.5 rounded text-text-muted hover:text-primary hover:bg-primary/10 transition-all"
                  title="打开文件夹"
                >
                  <FolderPlus className="w-3.5 h-3.5" />
                </button>
                <button 
                  onClick={refreshFiles}
                  className="p-1.5 rounded text-text-muted hover:text-text hover:bg-surface-hover transition-all"
                  title="刷新"
                >
                  <RefreshCw className="w-3.5 h-3.5" />
                </button>
              </>
            )}
            {!isElectron && (
              <button 
                onClick={() => fileInputRef.current?.click()}
                className="p-1.5 rounded text-text-muted hover:text-primary hover:bg-primary/10 transition-all"
                title="上传 Word 文档"
              >
                <Upload className="w-3.5 h-3.5" />
              </button>
            )}
            <button 
              onClick={handleNewDocument}
              className="p-1.5 rounded text-text-muted hover:text-primary hover:bg-primary/10 transition-all"
              title="新建文档"
            >
              <Plus className="w-3.5 h-3.5" />
            </button>
            {isElectron && (
              <button 
                onClick={handleNewPpt}
                className="p-1.5 rounded text-text-muted hover:text-primary hover:bg-primary/10 transition-all"
                title="新建 PPT（AI 海报式生成）"
              >
                <FileText className="w-3.5 h-3.5" />
              </button>
            )}
            <button
              onClick={() => setIsCollapsed(true)}
              className="p-1.5 rounded text-text-muted hover:text-text hover:bg-surface-hover transition-all"
              title="折叠"
            >
              <ChevronRight className="w-3.5 h-3.5 rotate-180" />
            </button>
          </div>
        </div>

        {/* 工作区路径 */}
        {isElectron && workspacePath && (
          <div className="px-2 py-1.5 bg-surface rounded-md border border-border">
            <p className="text-[10px] text-text-dim truncate" title={workspacePath}>
              📁 {workspacePath.split(/[/\\]/).pop()}
            </p>
          </div>
        )}
        
        {/* 搜索框 */}
        <div className="relative px-2">
          <Search className="absolute left-4 top-1/2 -translate-y-1/2 w-3.5 h-3.5 text-text-muted" />
          <input
            type="text"
            placeholder="搜索文档..."
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            className="w-full bg-surface border border-border rounded-md pl-8 pr-3 py-1.5 text-xs text-text placeholder-text-dim focus:outline-none focus:border-primary/50 focus:bg-surface-hover transition-all"
          />
        </div>
      </div>

      {/* 拖放提示 */}
      {isDragging && !isElectron && (
        <div className="mx-3 mb-3 p-4 border-2 border-dashed border-primary/50 rounded-lg bg-primary/5 flex flex-col items-center justify-center gap-2">
          <Upload className="w-6 h-6 text-primary" />
          <span className="text-xs text-primary font-medium">释放以上传文档</span>
        </div>
      )}

      {/* 文件列表 */}
      <div className="flex-1 overflow-y-auto py-2 scrollbar-thin">
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
            <FolderOpen className="w-10 h-10 text-text-dim mb-3" />
            <p className="text-sm text-text-muted mb-2">
              {isElectron
                ? (workspacePath ? '文件夹为空' : '没有打开的文件夹')
                : '没有文档'}
            </p>
            <p className="text-xs text-text-dim mb-4">
              {isElectron 
                ? (workspacePath
                    ? '该文件夹内暂无可用文件。你可以新建文档，或点击上方按钮切换到其它文件夹。'
                    : '点击上方按钮打开一个本地文件夹')
                : '上传一个 .docx 文件开始编辑'}
            </p>
            {isElectron ? (
              <div className="flex items-center gap-2">
                {workspacePath && (
                  <button
                    onClick={handleNewDocument}
                    className="flex items-center gap-2 px-4 py-2 bg-primary text-white text-xs rounded-md hover:bg-primary-hover transition-all"
                    title="在当前文件夹中新建文档"
                  >
                    <FilePlus className="w-4 h-4" />
                    新建文档
                  </button>
                )}
              <button
                onClick={openFolder}
                  className="flex items-center gap-2 px-4 py-2 bg-surface border border-border text-text text-xs rounded-md hover:bg-surface-hover transition-all"
                  title="打开/切换文件夹"
              >
                <FolderPlus className="w-4 h-4" />
                打开文件夹
              </button>
              </div>
            ) : (
              <button
                onClick={() => fileInputRef.current?.click()}
                className="flex items-center gap-2 px-4 py-2 bg-primary text-white text-xs rounded-md hover:bg-primary-hover transition-all"
              >
                <Upload className="w-4 h-4" />
                上传文档
              </button>
            )}
          </div>
        )}
      </div>

      {/* 底部提示 */}
      <div className="p-3 border-t border-border">
        <p className="text-[10px] text-text-dim text-center">
          {isElectron 
            ? '拖拽文件到 AI 对话框进行分析' 
            : '桌面版支持打开本地文件夹'}
        </p>
      </div>
    </div>
  )
}
