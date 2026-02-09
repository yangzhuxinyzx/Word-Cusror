import { useCallback, useState, useRef, useEffect } from 'react'
import Sidebar from './components/Sidebar'
import WordEditor from './components/WordEditor'
import OnlyOfficeEditor from './components/OnlyOfficeEditor'
import ChatPanel from './components/ChatPanel'
import Header from './components/Header'
import SettingsModal from './components/SettingsModal'
import { DocumentProvider, useDocument } from './context/DocumentContext'
import { AIProvider } from './context/AIContext'
import { CommentProvider } from './context/CommentContext'

import Dashboard from './components/Dashboard'

// 内部组件，可以访问 DocumentContext
function AppContent() {
  const { editorMode, currentFile } = useDocument()
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

  // 统一的“打开设置”入口：支持窗口事件 + Ctrl/Cmd + ,
  useEffect(() => {
    const handleOpenSettings = () => setShowSettings(true)

    const handleKeyDown = (e: KeyboardEvent) => {
      const isShortcut = (e.ctrlKey || e.metaKey) && !e.shiftKey && !e.altKey && e.key === ','
      if (!isShortcut) return
      e.preventDefault()
      e.stopPropagation()
      setShowSettings(true)
    }

    window.addEventListener('open-settings', handleOpenSettings as EventListener)
    window.addEventListener('keydown', handleKeyDown, true)
    return () => {
      window.removeEventListener('open-settings', handleOpenSettings as EventListener)
      window.removeEventListener('keydown', handleKeyDown, true)
    }
  }, [])

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
      const newWidth = Math.min(Math.max(startWidth.current + delta, 320), 600)
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
    <div className="h-screen w-screen flex flex-col overflow-hidden relative bg-background transition-colors duration-300">
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
        
        {/* 中间区域：如果没打开文件显示 Dashboard，否则显示编辑器 */}
        <div className="flex-1 flex flex-col overflow-hidden relative">
          {!currentFile ? (
            <Dashboard />
          ) : (
            <>
              {/* 编辑器内容 */}
              <div className="flex-1 overflow-hidden">
                {editorMode === 'tiptap' ? <WordEditor /> : <OnlyOfficeEditor />}
              </div>
            </>
          )}
        </div>
        
        {/* AI对话面板 - 可拖拽调节宽度 */}
        {showChat && (
          <div 
            className="flex flex-col overflow-hidden relative glass-card rounded-none border-l border-white/5 transition-all duration-200"
            style={{ width: chatWidth }}
          >
            {/* 拖拽调节条 */}
            <div
              onMouseDown={handleMouseDown}
              className="absolute left-0 top-0 bottom-0 w-1.5 cursor-col-resize z-10 group hover:bg-sage-500/20 transition-colors"
            >
              <div className="absolute left-0.5 top-1/2 -translate-y-1/2 w-0.5 h-16 bg-white/8 group-hover:bg-sage-400 rounded-full transition-colors" />
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
      <CommentProvider>
        <DocumentProvider>
          <AppContent />
        </DocumentProvider>
      </CommentProvider>
    </AIProvider>
  )
}

export default App
