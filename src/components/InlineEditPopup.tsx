import { useState, useRef, useEffect } from 'react'
import { Sparkles, Loader2, X, Check, Wand2 } from 'lucide-react'
import { useAI } from '../context/AIContext'

interface InlineEditPopupProps {
  selectedText: string
  selectedHtml?: string  // 选区的 HTML 格式（包含样式）
  position: { x: number; y: number }
  onClose: () => void
  onApply: (newText: string, isHtml?: boolean) => void  // isHtml 表示返回的是 HTML 格式
}

export default function InlineEditPopup({ 
  selectedText, 
  selectedHtml,
  position, 
  onClose, 
  onApply 
}: InlineEditPopupProps) {
  const [instruction, setInstruction] = useState('')
  const [isProcessing, setIsProcessing] = useState(false)
  const [result, setResult] = useState<string | null>(null)
  const [error, setError] = useState<string | null>(null)
  const inputRef = useRef<HTMLInputElement>(null)
  const { settings } = useAI()
  const apiUrl = settings.apiUrl || `${settings.baseUrl}/chat/completions`

  // 自动聚焦输入框
  useEffect(() => {
    inputRef.current?.focus()
  }, [])

  // 是否有 HTML 格式可用
  const hasHtmlFormat = selectedHtml && selectedHtml !== selectedText && selectedHtml.includes('<')

  // 生成 AI 的 system prompt
  const getSystemPrompt = () => {
    if (hasHtmlFormat) {
      return `你是一位专业的文档编辑助手。请严格按照用户的修改指令处理文字。

【输入格式】
用户会提供原文的 HTML 格式，其中包含格式标签（如 <strong>粗体</strong>、<em>斜体</em>、<u>下划线</u>、<span style="color:xxx">颜色</span> 等）。

【输出规则】
1. 输出格式为 HTML，保留并合理应用原文的格式标签
2. 只输出修改后的 HTML 内容，不要有任何解释、引号、前缀或后缀
3. 不要输出完整的 HTML 文档结构（如 <html>、<body> 等），只输出内容片段
4. 如果原文有粗体/斜体/颜色等格式，修改后的对应内容也应保持相同格式
5. 你可以根据内容语义决定是否调整格式（如重点内容可以加粗）
6. 支持的格式标签：<strong>/<b>、<em>/<i>、<u>、<span style="...">

【示例】
原文 HTML: <strong>重要通知</strong>：会议时间为 <em>下午3点</em>
修改指令: 把时间改成4点
输出: <strong>重要通知</strong>：会议时间为 <em>下午4点</em>`
    } else {
      return `你是一位专业的文档编辑助手。请严格按照用户的修改指令处理文字。

【输出规则】
- 只输出修改后的文字，不要有任何解释、引号、前缀或后缀
- 保持原文的格式风格（标点、换行、缩进等）
- 如果是翻译任务，保持原文的语义和风格`
    }
  }

  // 处理 AI 请求
  const handleSubmit = async () => {
    if (!instruction.trim()) return
    
    setIsProcessing(true)
    setError(null)
    setResult(null)

    try {
      // 使用 HTML 格式（如果有）
      const contentToProcess = hasHtmlFormat ? selectedHtml : selectedText
      const formatHint = hasHtmlFormat ? '（HTML格式，请保留格式标签）' : ''

      const response = await fetch(apiUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${settings.apiKey}`,
        },
        body: JSON.stringify({
          model: settings.model,
          messages: [
            {
              role: 'system',
              content: getSystemPrompt()
            },
            {
              role: 'user',
              content: `【原文${formatHint}】
${contentToProcess}

【修改指令】${instruction}

请直接输出修改后的${hasHtmlFormat ? 'HTML' : '文字'}：`
            }
          ],
          temperature: 0.7,
          max_tokens: 4000,
        }),
      })

      if (!response.ok) {
        const errorText = await response.text().catch(() => '')
        throw new Error(`AI 请求失败 (${response.status})${errorText ? ': ' + errorText.slice(0, 100) : ''}`)
      }

      const data = await response.json()
      let newText = data.choices?.[0]?.message?.content?.trim()
      
      if (newText) {
        // 清理可能的 markdown 代码块包裹
        if (newText.startsWith('```html')) {
          newText = newText.slice(7)
        } else if (newText.startsWith('```')) {
          newText = newText.slice(3)
        }
        if (newText.endsWith('```')) {
          newText = newText.slice(0, -3)
        }
        newText = newText.trim()
        
        setResult(newText)
      } else {
        throw new Error('AI 返回内容为空')
      }
    } catch (err) {
      setError((err as Error).message || '处理失败')
    } finally {
      setIsProcessing(false)
    }
  }

  // 快捷操作（优化后的 prompt 指令）
  const quickActions = [
    { label: '润色', icon: '✨', instruction: '优化这段文字的表达，使其更加流畅、专业，但保持原意和风格不变' },
    { label: '精简', icon: '✂️', instruction: '精简这段文字，删除冗余内容，保留核心信息，字数减少30%-50%' },
    { label: '扩写', icon: '📝', instruction: '扩展这段文字，补充更多细节和论据，字数增加50%-100%' },
    { label: '正式', icon: '👔', instruction: '将这段文字改成正式的书面语风格，适合公文或商务场合' },
    { label: '翻译', icon: '🌐', instruction: '翻译成英文，保持原文语义和风格' },
  ]

  const handleQuickAction = (actionInstruction: string) => {
    setInstruction(actionInstruction)
    // 自动提交
    setTimeout(() => {
      handleSubmitWithInstruction(actionInstruction)
    }, 0)
  }

  const handleSubmitWithInstruction = async (inst: string) => {
    setIsProcessing(true)
    setError(null)
    setResult(null)

    try {
      // 使用 HTML 格式（如果有）
      const contentToProcess = hasHtmlFormat ? selectedHtml : selectedText
      const formatHint = hasHtmlFormat ? '（HTML格式，请保留格式标签）' : ''

      const response = await fetch(apiUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${settings.apiKey}`,
        },
        body: JSON.stringify({
          model: settings.model,
          messages: [
            {
              role: 'system',
              content: getSystemPrompt()
            },
            {
              role: 'user',
              content: `【原文${formatHint}】
${contentToProcess}

【修改指令】${inst}

请直接输出修改后的${hasHtmlFormat ? 'HTML' : '文字'}：`
            }
          ],
          temperature: 0.7,
          max_tokens: 4000,
        }),
      })

      if (!response.ok) {
        const errorText = await response.text().catch(() => '')
        throw new Error(`AI 请求失败 (${response.status})${errorText ? ': ' + errorText.slice(0, 100) : ''}`)
      }

      const data = await response.json()
      let newText = data.choices?.[0]?.message?.content?.trim()
      
      if (newText) {
        // 清理可能的 markdown 代码块包裹
        if (newText.startsWith('```html')) {
          newText = newText.slice(7)
        } else if (newText.startsWith('```')) {
          newText = newText.slice(3)
        }
        if (newText.endsWith('```')) {
          newText = newText.slice(0, -3)
        }
        newText = newText.trim()
        
        setResult(newText)
      } else {
        throw new Error('AI 返回内容为空')
      }
    } catch (err) {
      setError((err as Error).message || '处理失败')
    } finally {
      setIsProcessing(false)
    }
  }

  // 计算弹窗位置
  const popupStyle: React.CSSProperties = {
    position: 'fixed',
    left: Math.min(position.x, window.innerWidth - 420),
    top: Math.min(position.y + 10, window.innerHeight - 400),
    zIndex: 9999,
  }

  return (
    <>
      {/* 背景遮罩 */}
      <div 
        className="fixed inset-0 z-[9998]" 
        onClick={onClose}
      />
      
      {/* 弹窗 */}
      <div style={popupStyle} className="w-[400px]">
        <div className="glass-card border border-border rounded-xl shadow-2xl overflow-hidden">
          {/* 标题栏 */}
          <div className="flex items-center justify-between px-3 py-2 bg-black/10 dark:bg-white/10 backdrop-blur-md border-b border-border">
            <div className="flex items-center gap-2">
              <Wand2 className="w-4 h-4 text-accent" />
              <span className="text-sm text-text">AI 编辑</span>
            </div>
            <button 
              onClick={onClose}
              className="p-1 hover:bg-black/10 dark:hover:bg-white/10 rounded transition-colors"
            >
              <X className="w-4 h-4 text-text-muted" />
            </button>
          </div>

          {/* 选中的文本预览 */}
          <div className="px-3 py-2 bg-black/5 dark:bg-white/5 border-b border-border">
            <p className="text-xs text-text-dim mb-1">选中的文本：</p>
            <p className="text-sm text-text-secondary line-clamp-2">{selectedText}</p>
          </div>

          {/* 快捷操作按钮 */}
          <div className="flex flex-wrap gap-1.5 px-3 py-2 border-b border-border">
            {quickActions.map((action) => (
              <button
                key={action.label}
                onClick={() => handleQuickAction(action.instruction)}
                disabled={isProcessing}
                className="flex items-center gap-1 px-2.5 py-1 bg-black/5 dark:bg-white/5 hover:bg-black/10 dark:hover:bg-white/10 border border-black/10 dark:border-white/10 rounded-lg text-xs text-text-secondary transition-colors disabled:opacity-50"
              >
                <span>{action.icon}</span>
                <span>{action.label}</span>
              </button>
            ))}
          </div>

          {/* 自定义指令输入 */}
          <div className="px-3 py-2">
            <div className="flex items-center gap-2">
              <input
                ref={inputRef}
                type="text"
                value={instruction}
                onChange={(e) => setInstruction(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter' && !e.shiftKey) {
                    e.preventDefault()
                    handleSubmit()
                  }
                  if (e.key === 'Escape') {
                    onClose()
                  }
                }}
                placeholder="输入修改指令，如：改成正式语气..."
                className="flex-1 px-3 py-2 bg-black/10 dark:bg-white/10 border border-black/10 dark:border-white/10 rounded-lg text-sm text-text placeholder-text-dim focus:outline-none focus:border-accent focus:ring-2 focus:ring-accent/10"
                disabled={isProcessing}
              />
              <button
                onClick={handleSubmit}
                disabled={isProcessing || !instruction.trim()}
                className="px-3 py-2 btn-sage rounded-lg transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
              >
                {isProcessing ? (
                  <Loader2 className="w-4 h-4 animate-spin" />
                ) : (
                  <Sparkles className="w-4 h-4" />
                )}
              </button>
            </div>
          </div>

          {/* 错误提示 */}
          {error && (
            <div className="px-3 py-2 bg-red-900/30 border-t border-red-800">
              <p className="text-sm text-red-400">{error}</p>
            </div>
          )}

          {/* 结果预览 */}
          {result && (
            <div className="border-t border-border">
              <div className="px-3 py-2 bg-black/5 dark:bg-white/5">
                <p className="text-xs text-text-dim mb-1">修改结果：</p>
                <p className="text-sm text-success whitespace-pre-wrap leading-relaxed">
                  {result}
                </p>
              </div>
              
              {/* 操作按钮 */}
              <div className="flex items-center gap-2 px-3 py-2 bg-black/10 dark:bg-white/10 border-t border-border">
                <button
                  onClick={() => {
                    onApply(result, hasHtmlFormat ? true : undefined)
                    onClose()
                  }}
                  className="flex items-center gap-1.5 px-3 py-1.5 bg-success/16 text-success border border-success/25 hover:bg-success/22 text-sm rounded-lg transition-colors"
                >
                  <Check className="w-4 h-4" />
                  <span>应用修改</span>
                </button>
                <button
                  onClick={() => setResult(null)}
                  className="flex items-center gap-1.5 px-3 py-1.5 bg-black/10 dark:bg-white/10 hover:bg-black/15 dark:hover:bg-white/15 text-text-secondary text-sm rounded-lg transition-colors border border-black/10 dark:border-white/10"
                >
                  <X className="w-4 h-4" />
                  <span>取消</span>
                </button>
              </div>
            </div>
          )}

          {/* 快捷键提示 */}
          <div className="px-3 py-1.5 bg-black/10 dark:bg-white/10 border-t border-border">
            <p className="text-[10px] text-text-dim text-center">
              <kbd className="px-1 py-0.5 bg-black/10 dark:bg-white/10 border border-black/10 dark:border-white/10 rounded">Enter</kbd> 提交 · 
              <kbd className="px-1 py-0.5 bg-black/10 dark:bg-white/10 border border-black/10 dark:border-white/10 rounded ml-1">Esc</kbd> 关闭
            </p>
          </div>
        </div>
      </div>
    </>
  )
}

