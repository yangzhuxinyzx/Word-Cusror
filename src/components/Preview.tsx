import { useMemo, useState } from 'react'
import { FileText, Maximize2, Minimize2, Printer } from 'lucide-react'
import { useDocument } from '../context/DocumentContext'
import ReactMarkdown from 'react-markdown'

export default function Preview() {
  const { document } = useDocument()
  const [isFullscreen, setIsFullscreen] = useState(false)
  const [scale, setScale] = useState(100)

  const previewContent = useMemo(() => {
    return document.content
  }, [document.content])

  return (
    <div className={`flex flex-col h-full bg-surface/50 ${isFullscreen ? 'fixed inset-0 z-50 bg-background' : ''} transition-colors duration-300`}>
      {/* 预览头部 */}
      <div className="flex items-center justify-between px-4 py-2 border-b border-border bg-background/50 backdrop-blur-sm select-none">
        <div className="flex items-center gap-2">
          <div className="px-2 py-1 rounded bg-primary/10 text-primary text-xs font-medium border border-primary/20">
            Preview
          </div>
          <span className="text-xs text-text-muted">Word Document (A4)</span>
        </div>
        
        <div className="flex items-center gap-2">
           <div className="flex items-center bg-surface rounded-md border border-border/50 overflow-hidden">
              <button onClick={() => setScale(s => Math.max(50, s - 10))} className="px-2 py-1 text-xs text-text-muted hover:bg-surface-hover hover:text-text transition-colors">-</button>
              <span className="px-2 text-xs text-text min-w-[3ch] text-center">{scale}%</span>
              <button onClick={() => setScale(s => Math.min(200, s + 10))} className="px-2 py-1 text-xs text-text-muted hover:bg-surface-hover hover:text-text transition-colors">+</button>
           </div>
           
           <div className="w-px h-4 bg-border/50" />
           
           <button
            className="p-1.5 rounded-md text-text-muted hover:text-text hover:bg-surface-hover transition-all"
            title="Print"
          >
            <Printer className="w-3.5 h-3.5" />
          </button>
           
          <button
            onClick={() => setIsFullscreen(!isFullscreen)}
            className="p-1.5 rounded-md text-text-muted hover:text-text hover:bg-surface-hover transition-all"
            title={isFullscreen ? '退出全屏' : '全屏预览'}
          >
            {isFullscreen ? (
              <Minimize2 className="w-3.5 h-3.5" />
            ) : (
              <Maximize2 className="w-3.5 h-3.5" />
            )}
          </button>
        </div>
      </div>

      {/* 预览内容区 */}
      <div className="flex-1 overflow-auto p-8 bg-surface/30 relative scrollbar-thin">
        {/* 背景网格装饰 */}
        <div className="absolute inset-0 opacity-[0.03] pointer-events-none" 
             style={{ backgroundImage: 'radial-gradient(#fff 1px, transparent 1px)', backgroundSize: '20px 20px' }}>
        </div>

        {/* A4纸张效果 */}
        <div 
            className="mx-auto bg-white shadow-paper transition-transform duration-200 origin-top"
            style={{ 
                width: '210mm', 
                minHeight: '297mm',
                transform: `scale(${scale / 100})`,
                marginBottom: '2rem'
            }}
        >
          <div className="word-preview p-[2.54cm]"> {/* 标准 Word 页边距 */}
            {previewContent ? (
              <ReactMarkdown>
                {previewContent}
              </ReactMarkdown>
            ) : (
              <div className="flex flex-col items-center justify-center h-[600px] text-gray-300 select-none">
                <FileText className="w-16 h-16 mb-4 opacity-50" />
                <p className="text-lg font-medium">No Content</p>
                <p className="text-sm mt-2">Start typing in the editor...</p>
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  )
}
