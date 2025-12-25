import { useEffect, useRef } from 'react'
import { 
  Sparkles, 
  Scissors, 
  FileText, 
  Languages, 
  Wand2,
  MessageSquare,
  Copy,
  Trash2
} from 'lucide-react'

interface ContextMenuProps {
  position: { x: number; y: number }
  selectedText: string
  onClose: () => void
  onAction: (action: string, instruction?: string) => void
}

export default function ContextMenu({ 
  position, 
  selectedText,
  onClose, 
  onAction 
}: ContextMenuProps) {
  const menuRef = useRef<HTMLDivElement>(null)

  // 点击外部关闭
  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (menuRef.current && !menuRef.current.contains(e.target as Node)) {
        onClose()
      }
    }
    
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        onClose()
      }
    }

    document.addEventListener('mousedown', handleClickOutside)
    document.addEventListener('keydown', handleKeyDown)
    
    return () => {
      document.removeEventListener('mousedown', handleClickOutside)
      document.removeEventListener('keydown', handleKeyDown)
    }
  }, [onClose])

  // AI 操作选项（优化后的 prompt 指令）
  const aiActions = [
    { 
      id: 'polish', 
      label: '润色', 
      icon: Sparkles, 
      instruction: '优化这段文字的表达，使其更加流畅、专业，但保持原意和风格不变。注意：不要改变原文的核心观点和信息',
      color: 'text-violet-400'
    },
    { 
      id: 'simplify', 
      label: '精简', 
      icon: Scissors, 
      instruction: '精简这段文字，删除冗余和重复的内容，保留核心信息。目标：字数减少30%-50%，但不丢失关键信息',
      color: 'text-blue-400'
    },
    { 
      id: 'expand', 
      label: '扩写', 
      icon: FileText, 
      instruction: '扩展这段文字，补充更多细节、论据和例子，使内容更加丰富完整。目标：字数增加50%-100%',
      color: 'text-green-400'
    },
    { 
      id: 'formal', 
      label: '正式化', 
      icon: MessageSquare, 
      instruction: '将这段文字改成正式的书面语风格，适合公文、商务报告或学术场合。去除口语化表达，使用规范用语',
      color: 'text-amber-400'
    },
    { 
      id: 'translate', 
      label: '翻译成英文', 
      icon: Languages, 
      instruction: '将这段中文准确翻译成英文，保持原文的语义、语气和风格。使用地道的英语表达',
      color: 'text-cyan-400'
    },
    { 
      id: 'translate_cn', 
      label: '翻译成中文', 
      icon: Languages, 
      instruction: '将这段英文准确翻译成中文，保持原文的语义、语气和风格。使用地道的中文表达',
      color: 'text-cyan-400'
    },
  ]

  // 基础操作
  const basicActions = [
    { id: 'copy', label: '复制', icon: Copy },
    { id: 'delete', label: '删除', icon: Trash2 },
  ]

  // 计算菜单位置，确保不超出屏幕
  const menuStyle: React.CSSProperties = {
    position: 'fixed',
    left: Math.min(position.x, window.innerWidth - 220),
    top: Math.min(position.y, window.innerHeight - 400),
    zIndex: 9999,
  }

  return (
    <div ref={menuRef} style={menuStyle}>
      <div className="w-52 bg-zinc-900 border border-zinc-700 rounded-xl shadow-2xl overflow-hidden">
        {/* AI 编辑标题 */}
        <div className="px-3 py-2 bg-zinc-800 border-b border-zinc-700">
          <div className="flex items-center gap-2">
            <Wand2 className="w-4 h-4 text-violet-400" />
            <span className="text-xs font-medium text-zinc-300">AI 编辑</span>
          </div>
        </div>

        {/* AI 操作列表 */}
        <div className="py-1">
          {aiActions.map((action) => (
            <button
              key={action.id}
              onClick={() => onAction('ai', action.instruction)}
              className="w-full flex items-center gap-3 px-3 py-2 hover:bg-zinc-800 transition-colors text-left"
            >
              <action.icon className={`w-4 h-4 ${action.color}`} />
              <span className="text-sm text-zinc-300">{action.label}</span>
            </button>
          ))}
        </div>

        {/* 分隔线 */}
        <div className="border-t border-zinc-700" />

        {/* 自定义编辑 */}
        <div className="py-1">
          <button
            onClick={() => onAction('custom')}
            className="w-full flex items-center gap-3 px-3 py-2 hover:bg-zinc-800 transition-colors text-left"
          >
            <Sparkles className="w-4 h-4 text-violet-400" />
            <span className="text-sm text-zinc-300">自定义编辑...</span>
            <span className="ml-auto text-[10px] text-zinc-600 bg-zinc-800 px-1.5 py-0.5 rounded">Ctrl+K</span>
          </button>
        </div>

        {/* 分隔线 */}
        <div className="border-t border-zinc-700" />

        {/* 基础操作 */}
        <div className="py-1">
          {basicActions.map((action) => (
            <button
              key={action.id}
              onClick={() => onAction(action.id)}
              className="w-full flex items-center gap-3 px-3 py-2 hover:bg-zinc-800 transition-colors text-left"
            >
              <action.icon className="w-4 h-4 text-zinc-500" />
              <span className="text-sm text-zinc-400">{action.label}</span>
            </button>
          ))}
        </div>
      </div>
    </div>
  )
}

