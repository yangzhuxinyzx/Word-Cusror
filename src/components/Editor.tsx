import { useCallback, useState, useRef, useEffect } from 'react'
import { 
  Bold, 
  Italic, 
  Underline, 
  AlignLeft, 
  AlignCenter, 
  AlignRight,
  AlignJustify,
  List,
  ListOrdered,
  Heading1,
  Heading2,
  Heading3,
  Undo,
  Redo,
  Sparkles,
  Wand2,
  Save,
  FileUp,
  Copy,
  Scissors,
  ClipboardPaste
} from 'lucide-react'
import { useDocument } from '../context/DocumentContext'
import { useAI } from '../context/AIContext'

export default function Editor() {
  const { document, updateContent, updateStyles, saveDocument, hasUnsavedChanges, applyAIEdit } = useDocument()
  const { sendMessage, addMessage, isLoading } = useAI()
  const [showAIHelper, setShowAIHelper] = useState(false)
  const [aiPrompt, setAiPrompt] = useState('')
  const [isSaving, setIsSaving] = useState(false)
  const [selectedText, setSelectedText] = useState('')
  const textareaRef = useRef<HTMLTextAreaElement>(null)
  const [history, setHistory] = useState<string[]>([])
  const [historyIndex, setHistoryIndex] = useState(-1)

  // 记录编辑历史（用于撤销/重做）
  useEffect(() => {
    if (document.content && (history.length === 0 || history[history.length - 1] !== document.content)) {
      const newHistory = [...history.slice(0, historyIndex + 1), document.content]
      if (newHistory.length > 50) newHistory.shift() // 限制历史记录数量
      setHistory(newHistory)
      setHistoryIndex(newHistory.length - 1)
    }
  }, [document.content])

  const handleContentChange = useCallback((e: React.ChangeEvent<HTMLTextAreaElement>) => {
    updateContent(e.target.value)
  }, [updateContent])

  // 获取选中的文本
  const handleSelect = useCallback(() => {
    if (textareaRef.current) {
      const start = textareaRef.current.selectionStart
      const end = textareaRef.current.selectionEnd
      if (start !== end) {
        setSelectedText(document.content.substring(start, end))
      } else {
        setSelectedText('')
      }
    }
  }, [document.content])

  // 撤销
  const handleUndo = useCallback(() => {
    if (historyIndex > 0) {
      setHistoryIndex(historyIndex - 1)
      updateContent(history[historyIndex - 1])
    }
  }, [history, historyIndex, updateContent])

  // 重做
  const handleRedo = useCallback(() => {
    if (historyIndex < history.length - 1) {
      setHistoryIndex(historyIndex + 1)
      updateContent(history[historyIndex + 1])
    }
  }, [history, historyIndex, updateContent])

  // 保存文档
  const handleSave = useCallback(async () => {
    setIsSaving(true)
    try {
      await saveDocument()
    } catch (error) {
      console.error('Save failed:', error)
      alert('保存失败，请重试')
    } finally {
      setIsSaving(false)
    }
  }, [saveDocument])

  // AI 辅助编辑
  const handleAIAssist = useCallback(async () => {
    if (!aiPrompt.trim()) return

    // 构建带有选中文本上下文的提示
    let fullPrompt = aiPrompt
    if (selectedText) {
      fullPrompt = `请对以下选中的文本进行操作：\n\n"${selectedText}"\n\n操作要求：${aiPrompt}\n\n请直接返回修改后的完整文档内容（Markdown格式）。`
    } else {
      fullPrompt = `${aiPrompt}\n\n请基于当前文档内容进行修改，直接返回修改后的完整文档内容（Markdown格式）。`
    }

    addMessage({ role: 'user', content: aiPrompt })
    const response = await sendMessage(fullPrompt, document.content)
    addMessage({ role: 'assistant', content: response })

    // 如果返回内容看起来像文档，应用到编辑器
    if (response.includes('#') || response.includes('-') || response.length > 50) {
      // 检测是否是完整文档内容
      const isDocumentContent = response.startsWith('#') || 
                                response.includes('\n#') || 
                                response.includes('\n-') ||
                                response.includes('\n1.')
      
      if (isDocumentContent) {
        applyAIEdit(response)
      }
    }

    setAiPrompt('')
    setShowAIHelper(false)
  }, [aiPrompt, selectedText, document.content, addMessage, sendMessage, applyAIEdit])

  // 插入 Markdown 格式
  const insertMarkdown = useCallback((prefix: string, suffix: string = '') => {
    if (!textareaRef.current) return

    const start = textareaRef.current.selectionStart
    const end = textareaRef.current.selectionEnd
    const selectedText = document.content.substring(start, end)
    
    const newText = 
      document.content.substring(0, start) + 
      prefix + selectedText + suffix + 
      document.content.substring(end)
    
    updateContent(newText)
    
    // 恢复光标位置
    setTimeout(() => {
      if (textareaRef.current) {
        const newPos = start + prefix.length + selectedText.length + suffix.length
        textareaRef.current.selectionStart = newPos
        textareaRef.current.selectionEnd = newPos
        textareaRef.current.focus()
      }
    }, 0)
  }, [document.content, updateContent])

  // 快捷键处理
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.ctrlKey || e.metaKey) {
        switch (e.key.toLowerCase()) {
          case 's':
            e.preventDefault()
            handleSave()
            break
          case 'z':
            if (e.shiftKey) {
              e.preventDefault()
              handleRedo()
            } else {
              e.preventDefault()
              handleUndo()
            }
            break
          case 'y':
            e.preventDefault()
            handleRedo()
            break
          case 'b':
            e.preventDefault()
            insertMarkdown('**', '**')
            break
          case 'i':
            e.preventDefault()
            insertMarkdown('*', '*')
            break
        }
      }
    }

    window.addEventListener('keydown', handleKeyDown)
    return () => window.removeEventListener('keydown', handleKeyDown)
  }, [handleSave, handleUndo, handleRedo, insertMarkdown])

  const ToolbarButton = ({ 
    icon: Icon, 
    onClick, 
    title,
    active = false,
    disabled = false
  }: { 
    icon: React.ElementType
    onClick: () => void
    title: string
    active?: boolean
    disabled?: boolean
  }) => (
    <button
      onClick={onClick}
      title={title}
      disabled={disabled}
      className={`p-1.5 rounded-md transition-all ${
        disabled 
          ? 'text-text-dim cursor-not-allowed'
          : active 
            ? 'bg-primary/10 text-primary' 
            : 'text-text-muted hover:text-text hover:bg-surface-hover'
      }`}
    >
      <Icon className="w-4 h-4" />
    </button>
  )

  return (
    <div className="flex flex-col h-full bg-surface/30 relative group">
      {/* 悬浮式 AI 助手入口 */}
      <div className="absolute bottom-6 right-6 z-10">
        <button
          onClick={() => setShowAIHelper(!showAIHelper)}
          className={`flex items-center justify-center w-12 h-12 rounded-full shadow-glow transition-all duration-300 hover:scale-105 ${
            showAIHelper 
              ? 'bg-surface text-text rotate-45 border border-border' 
              : 'bg-primary text-white hover:bg-primary-hover'
          }`}
          title="AI 编辑助手 (Ctrl+K)"
        >
          {showAIHelper ? <Sparkles className="w-5 h-5" /> : <Wand2 className="w-5 h-5" />}
        </button>
      </div>

      {/* AI 快捷编辑框 */}
      {showAIHelper && (
        <div className="absolute bottom-20 right-6 left-6 z-20 animate-enter">
          <div className="bg-surface/95 backdrop-blur-xl border border-primary/20 rounded-xl shadow-2xl p-4 ring-1 ring-primary/10">
            <div className="flex gap-3 mb-3">
              <div className="w-8 h-8 rounded-lg bg-gradient-to-br from-indigo-500 to-purple-600 flex items-center justify-center shrink-0">
                <Sparkles className="w-4 h-4 text-white" />
              </div>
              <div className="flex-1">
                <h3 className="text-sm font-semibold text-text">AI 编辑</h3>
                <p className="text-xs text-text-muted">
                  {selectedText ? `已选中 ${selectedText.length} 个字符` : '描述你想对文档做的修改...'}
                </p>
              </div>
            </div>
            
            <div className="relative">
              <input
                type="text"
                value={aiPrompt}
                onChange={(e) => setAiPrompt(e.target.value)}
                onKeyDown={(e) => e.key === 'Enter' && handleAIAssist()}
                placeholder="例如：把这段改得更正式、添加一个总结部分..."
                className="w-full bg-background border border-border rounded-lg px-4 py-3 text-sm text-text placeholder-text-dim focus:outline-none focus:border-primary/50 transition-all shadow-inner"
                autoFocus
              />
              <button
                onClick={handleAIAssist}
                disabled={isLoading || !aiPrompt.trim()}
                className="absolute right-2 top-1/2 -translate-y-1/2 px-3 py-1.5 bg-primary text-white rounded-md text-xs font-medium hover:bg-primary-hover transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
              >
                {isLoading ? '...' : '执行'}
              </button>
            </div>

            <div className="flex gap-2 mt-3 overflow-x-auto pb-1 scrollbar-none">
              <button onClick={() => setAiPrompt('润色这段文字，使其更专业')} className="whitespace-nowrap px-3 py-1.5 bg-surface-hover hover:bg-primary/10 text-xs text-text-muted hover:text-primary rounded-full transition-colors border border-border/50">
                ✨ 润色
              </button>
              <button onClick={() => setAiPrompt('扩展这部分内容，增加更多细节')} className="whitespace-nowrap px-3 py-1.5 bg-surface-hover hover:bg-primary/10 text-xs text-text-muted hover:text-primary rounded-full transition-colors border border-border/50">
                📝 扩展
              </button>
              <button onClick={() => setAiPrompt('精简这段文字，保留核心内容')} className="whitespace-nowrap px-3 py-1.5 bg-surface-hover hover:bg-primary/10 text-xs text-text-muted hover:text-primary rounded-full transition-colors border border-border/50">
                📉 精简
              </button>
              <button onClick={() => setAiPrompt('修正语法和拼写错误')} className="whitespace-nowrap px-3 py-1.5 bg-surface-hover hover:bg-primary/10 text-xs text-text-muted hover:text-primary rounded-full transition-colors border border-border/50">
                🔍 纠错
              </button>
              <button onClick={() => setAiPrompt('翻译成英文')} className="whitespace-nowrap px-3 py-1.5 bg-surface-hover hover:bg-primary/10 text-xs text-text-muted hover:text-primary rounded-full transition-colors border border-border/50">
                🌐 翻译
              </button>
            </div>
          </div>
        </div>
      )}

      {/* 工具栏 */}
      <div className="flex items-center gap-1 px-4 py-2 border-b border-border bg-background/50 backdrop-blur-sm select-none">
        {/* 撤销/重做 */}
        <div className="flex items-center bg-surface rounded-lg p-0.5 border border-border/50">
          <ToolbarButton icon={Undo} onClick={handleUndo} title="撤销 (Ctrl+Z)" disabled={historyIndex <= 0} />
          <ToolbarButton icon={Redo} onClick={handleRedo} title="重做 (Ctrl+Y)" disabled={historyIndex >= history.length - 1} />
        </div>
        
        <div className="w-px h-5 bg-border/50 mx-2" />
        
        {/* 标题 */}
        <div className="flex items-center gap-0.5">
          <ToolbarButton icon={Heading1} onClick={() => insertMarkdown('# ')} title="标题1" />
          <ToolbarButton icon={Heading2} onClick={() => insertMarkdown('## ')} title="标题2" />
          <ToolbarButton icon={Heading3} onClick={() => insertMarkdown('### ')} title="标题3" />
        </div>
        
        <div className="w-px h-5 bg-border/50 mx-2" />
        
        {/* 文本格式 */}
        <div className="flex items-center gap-0.5">
          <ToolbarButton icon={Bold} onClick={() => insertMarkdown('**', '**')} title="粗体 (Ctrl+B)" />
          <ToolbarButton icon={Italic} onClick={() => insertMarkdown('*', '*')} title="斜体 (Ctrl+I)" />
          <ToolbarButton icon={Underline} onClick={() => insertMarkdown('<u>', '</u>')} title="下划线" />
        </div>
        
        <div className="w-px h-5 bg-border/50 mx-2" />
        
        {/* 对齐 */}
        <div className="flex items-center gap-0.5">
          <ToolbarButton icon={AlignLeft} onClick={() => updateStyles({ textAlign: 'left' })} title="左对齐" active={document.styles.textAlign === 'left'} />
          <ToolbarButton icon={AlignCenter} onClick={() => updateStyles({ textAlign: 'center' })} title="居中" active={document.styles.textAlign === 'center'} />
          <ToolbarButton icon={AlignRight} onClick={() => updateStyles({ textAlign: 'right' })} title="右对齐" active={document.styles.textAlign === 'right'} />
          <ToolbarButton icon={AlignJustify} onClick={() => updateStyles({ textAlign: 'justify' })} title="两端对齐" active={document.styles.textAlign === 'justify'} />
        </div>
        
        <div className="w-px h-5 bg-border/50 mx-2" />
        
        {/* 列表 */}
        <div className="flex items-center gap-0.5">
          <ToolbarButton icon={List} onClick={() => insertMarkdown('- ')} title="无序列表" />
          <ToolbarButton icon={ListOrdered} onClick={() => insertMarkdown('1. ')} title="有序列表" />
        </div>

        <div className="flex-1" />

        {/* 保存按钮 */}
        <button
          onClick={handleSave}
          disabled={isSaving || !hasUnsavedChanges}
          className={`flex items-center gap-1.5 px-3 py-1.5 rounded-md text-xs font-medium transition-all ${
            hasUnsavedChanges
              ? 'bg-primary text-white hover:bg-primary-hover shadow-glow'
              : 'bg-surface text-text-muted border border-border'
          } disabled:opacity-50 disabled:cursor-not-allowed`}
        >
          {isSaving ? (
            <div className="w-3.5 h-3.5 border-2 border-white/30 border-t-white rounded-full animate-spin" />
          ) : (
            <Save className="w-3.5 h-3.5" />
          )}
          <span>{isSaving ? '保存中' : hasUnsavedChanges ? '保存' : '已保存'}</span>
        </button>
      </div>

      {/* 编辑器核心区域 */}
      <div className="flex-1 overflow-hidden relative">
        <textarea
          ref={textareaRef}
          value={document.content}
          onChange={handleContentChange}
          onSelect={handleSelect}
          placeholder="开始编辑你的文档...

支持 Markdown 语法：
# 标题1
## 标题2
### 标题3

**粗体** *斜体*

- 无序列表
1. 有序列表

点击右下角的 AI 按钮，让 AI 帮你编辑文档！"
          className="w-full h-full resize-none bg-transparent text-text p-8 focus:outline-none font-mono text-sm leading-relaxed selection:bg-primary/30 scrollbar-thin"
          style={{
            fontSize: `${document.styles.fontSize}px`,
            lineHeight: document.styles.lineHeight,
            textAlign: document.styles.textAlign,
          }}
          spellCheck={false}
        />
      </div>

      {/* 底部状态栏 */}
      <div className="flex items-center justify-between px-4 py-1.5 border-t border-border bg-surface/30 text-[10px] text-text-dim select-none">
        <div className="flex items-center gap-4">
          <span className="hover:text-text transition-colors">{document.content.length} 字符</span>
          <span className="hover:text-text transition-colors">{document.content.split(/\s+/).filter(Boolean).length} 词</span>
          <span className="hover:text-text transition-colors">{document.content.split('\n').length} 行</span>
          {selectedText && (
            <span className="text-primary">已选中 {selectedText.length} 字符</span>
          )}
        </div>
        <div className="flex items-center gap-4">
          {hasUnsavedChanges && (
            <span className="text-amber-400">● 未保存</span>
          )}
          <span className="uppercase">Markdown</span>
          <span>Ctrl+S 保存</span>
        </div>
      </div>
    </div>
  )
}
