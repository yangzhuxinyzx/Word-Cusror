import { useCallback, useState, useRef, useEffect } from 'react'
import Sidebar from './components/Sidebar'
import WordEditor from './components/WordEditor'
import OnlyOfficeEditor from './components/OnlyOfficeEditor'
import ChatPanel from './components/ChatPanel'
import Header from './components/Header'
import SettingsModal from './components/SettingsModal'
import { DocumentProvider, useDocument } from './context/DocumentContext'
import { AIProvider } from './context/AIContext'

// 内部组件，可以访问 DocumentContext
function AppContent() {
  const { editorMode, setEditorMode } = useDocument()
  const [showChat, setShowChat] = useState(true)
  const [showSettings, setShowSettings] = useState(false)
  const [activeView, setActiveView] = useState<'editor' | 'preview' | 'split'>('editor')
  
  // 可拖拽调节对话框宽度
  const [chatWidth, setChatWidth] = useState(380)
  const isResizing = useRef(false)
  const startX = useRef(0)
  const startWidth = useRef(0)

  const toggleChat = useCallback(() => setShowChat(prev => !prev), [])
  const toggleSettings = useCallback(() => setShowSettings(prev => !prev), [])

  // 拖拽调节宽度
  const handleMouseDown = useCallback((e: React.MouseEvent) => {
    isResizing.current = true
    startX.current = e.clientX
    startWidth.current = chatWidth
    document.body.style.cursor = 'col-resize'
    document.body.style.userSelect = 'none'
  }, [chatWidth])

  useEffect(() => {
    const handleMouseMove = (e: MouseEvent) => {
      if (!isResizing.current) return
      const delta = startX.current - e.clientX
      const newWidth = Math.min(Math.max(startWidth.current + delta, 280), 600)
      setChatWidth(newWidth)
    }

    const handleMouseUp = () => {
      isResizing.current = false
      document.body.style.cursor = ''
      document.body.style.userSelect = ''
    }

    document.addEventListener('mousemove', handleMouseMove)
    document.addEventListener('mouseup', handleMouseUp)
    return () => {
      document.removeEventListener('mousemove', handleMouseMove)
      document.removeEventListener('mouseup', handleMouseUp)
    }
  }, [])

  return (
    <div className="h-screen w-screen flex flex-col bg-background text-text overflow-hidden">
      {/* 顶部导航栏 */}
      <Header 
        showChat={showChat}
        showPreview={false}
        activeView={activeView}
        onToggleChat={toggleChat}
        onTogglePreview={() => {}}
        onViewChange={setActiveView}
        onOpenSettings={toggleSettings}
      />
      
      {/* 主内容区 */}
      <div className="flex-1 flex overflow-hidden">
        {/* 侧边栏 - 文件浏览器 */}
        <Sidebar />
        
        {/* 主编辑区域 */}
        <div className="flex-1 flex flex-col overflow-hidden relative">
          {/* 编辑器切换按钮 */}
          <div className="flex items-center gap-2 px-4 py-2 bg-surface border-b border-border">
            <span className="text-xs text-text-muted">编辑器:</span>
            <button
              onClick={() => setEditorMode('tiptap')}
              className={`px-3 py-1 text-xs rounded-md transition-colors ${
                editorMode === 'tiptap' 
                  ? 'bg-primary text-white' 
                  : 'bg-surface-hover text-text-muted hover:text-text'
              }`}
            >
              内置编辑器
            </button>
            <button
              onClick={() => setEditorMode('onlyoffice')}
              className={`px-3 py-1 text-xs rounded-md transition-colors ${
                editorMode === 'onlyoffice' 
                  ? 'bg-primary text-white' 
                  : 'bg-surface-hover text-text-muted hover:text-text'
              }`}
            >
              ONLYOFFICE
            </button>
            {editorMode === 'tiptap' && (
              <span className="text-[10px] text-green-400 ml-2">✓ AI 编辑已启用</span>
            )}
            {editorMode === 'onlyoffice' && (
              <span className="text-[10px] text-blue-400 ml-2">✓ 完美兼容 Office</span>
            )}
          </div>
          
          {/* 编辑器内容 */}
          <div className="flex-1 overflow-hidden">
            {editorMode === 'tiptap' ? <WordEditor /> : <OnlyOfficeEditor />}
          </div>
        </div>
        
        {/* AI对话面板 - 可拖拽调节宽度 */}
        {showChat && (
          <div 
            className="flex flex-col overflow-hidden bg-surface transition-colors duration-200"
            style={{ width: chatWidth }}
          >
            {/* 拖拽调节条 */}
            <div
              onMouseDown={handleMouseDown}
              className="absolute left-0 top-0 bottom-0 w-1 cursor-col-resize hover:bg-primary/50 active:bg-primary transition-colors z-10 group"
              style={{ marginLeft: -2 }}
            >
              <div className="absolute left-0 top-1/2 -translate-y-1/2 w-1 h-12 bg-border group-hover:bg-primary/70 rounded-full transition-colors" />
            </div>
            <ChatPanel />
          </div>
        )}
      </div>
      
      {/* 设置弹窗 */}
      {showSettings && (
        <SettingsModal onClose={toggleSettings} />
      )}
    </div>
  )
}

function App() {
  return (
    <AIProvider>
      <DocumentProvider>
        <AppContent />
      </DocumentProvider>
    </AIProvider>
  )
}

export default App
