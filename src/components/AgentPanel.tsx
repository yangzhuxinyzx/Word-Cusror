import { useState, useEffect } from 'react'
import { 
  CheckCircle2, 
  Circle, 
  Loader2, 
  FileText, 
  ChevronDown,
  ChevronRight,
  Clock,
  Sparkles,
  Pencil,
  Plus,
  Trash2,
  Search,
  Eye
} from 'lucide-react'

export interface AgentStep {
  id: string
  type: 'thinking' | 'reading' | 'searching' | 'editing' | 'creating' | 'deleting' | 'completed'
  description: string
  status: 'pending' | 'running' | 'completed' | 'error'
  details?: string
  timestamp?: Date
}

export interface FileChange {
  name: string
  path?: string
  additions: number
  deletions: number
  status: 'pending' | 'writing' | 'done'
  operations?: string[]  // 具体的操作描述
}

interface AgentPanelProps {
  isVisible: boolean
  steps: AgentStep[]
  fileChanges: FileChange[]
  thinkingTime: number
  currentAction?: string
  onAccept?: () => void
  onReject?: () => void
}

// 获取步骤图标
function getStepIcon(type: AgentStep['type'], status: AgentStep['status']) {
  if (status === 'running') {
    return <Loader2 className="w-4 h-4 text-primary animate-spin" />
  }
  if (status === 'completed') {
    return <CheckCircle2 className="w-4 h-4 text-green-400" />
  }
  if (status === 'error') {
    return <Circle className="w-4 h-4 text-red-400" />
  }
  
  switch (type) {
    case 'thinking':
      return <Sparkles className="w-4 h-4 text-text-dim" />
    case 'reading':
      return <Eye className="w-4 h-4 text-text-dim" />
    case 'searching':
      return <Search className="w-4 h-4 text-text-dim" />
    case 'editing':
      return <Pencil className="w-4 h-4 text-text-dim" />
    case 'creating':
      return <Plus className="w-4 h-4 text-text-dim" />
    case 'deleting':
      return <Trash2 className="w-4 h-4 text-text-dim" />
    default:
      return <Circle className="w-4 h-4 text-text-dim" />
  }
}

export default function AgentPanel({
  isVisible,
  steps,
  fileChanges,
  thinkingTime,
  currentAction,
}: AgentPanelProps) {
  const [isStepsExpanded, setIsStepsExpanded] = useState(true)
  const [isFilesExpanded, setIsFilesExpanded] = useState(true)

  if (!isVisible) return null

  const completedSteps = steps.filter(s => s.status === 'completed').length
  const runningStep = steps.find(s => s.status === 'running')

  return (
    <div className="border-t border-border bg-gradient-to-b from-surface/80 to-surface/50 animate-enter">
      {/* 当前操作状态栏 */}
      {(thinkingTime > 0 || currentAction) && (
        <div className="px-4 py-2.5 border-b border-border/50 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <div className="relative">
              <Sparkles className="w-4 h-4 text-violet-400" />
              <span className="absolute -top-0.5 -right-0.5 w-2 h-2 bg-violet-400 rounded-full animate-pulse" />
            </div>
            <span className="text-sm font-medium text-text">
              {currentAction || 'AI 正在处理...'}
            </span>
          </div>
          {thinkingTime > 0 && (
            <div className="flex items-center gap-1.5 text-xs text-text-muted bg-background/50 px-2 py-1 rounded-md">
              <Clock className="w-3 h-3" />
              <span>{thinkingTime}s</span>
            </div>
          )}
        </div>
      )}

      {/* 步骤列表 */}
      {steps.length > 0 && (
        <div className="border-b border-border/50">
          <button
            onClick={() => setIsStepsExpanded(!isStepsExpanded)}
            className="w-full px-4 py-2.5 flex items-center justify-between hover:bg-surface-hover/50 transition-colors"
          >
            <div className="flex items-center gap-2">
              {isStepsExpanded ? (
                <ChevronDown className="w-3.5 h-3.5 text-text-muted" />
              ) : (
                <ChevronRight className="w-3.5 h-3.5 text-text-muted" />
              )}
              <span className="text-xs font-medium text-text">执行步骤</span>
              <span className="text-[10px] text-text-dim bg-background/50 px-1.5 py-0.5 rounded">
                {completedSteps}/{steps.length}
              </span>
            </div>
            {runningStep && (
              <div className="flex items-center gap-1.5 text-xs text-violet-400">
                <Loader2 className="w-3 h-3 animate-spin" />
                <span className="truncate max-w-[180px]">{runningStep.description}</span>
              </div>
            )}
          </button>
          
          {isStepsExpanded && (
            <div className="px-4 pb-3 space-y-1">
              {steps.map((step, index) => (
                <div 
                  key={step.id} 
                  className={`flex items-start gap-2.5 py-1.5 px-2 rounded-md transition-colors ${
                    step.status === 'running' ? 'bg-violet-500/10' : 'hover:bg-surface-hover/30'
                  }`}
                >
                  <div className="mt-0.5">
                    {getStepIcon(step.type, step.status)}
                  </div>
                  <div className="flex-1 min-w-0">
                    <div className={`text-xs ${
                      step.status === 'completed' ? 'text-text-muted' :
                      step.status === 'running' ? 'text-text font-medium' : 
                      step.status === 'error' ? 'text-red-400' : 'text-text-muted'
                    }`}>
                      {step.description}
                    </div>
                    {step.details && step.status === 'running' && (
                      <div className="text-[10px] text-text-dim mt-0.5 truncate">
                        {step.details}
                      </div>
                    )}
                  </div>
                  {step.status === 'completed' && step.timestamp && (
                    <span className="text-[10px] text-text-dim">
                      {step.timestamp.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' })}
                    </span>
                  )}
                </div>
              ))}
            </div>
          )}
        </div>
      )}

      {/* 文件变更列表 */}
      {fileChanges.length > 0 && (
        <div>
          <button
            onClick={() => setIsFilesExpanded(!isFilesExpanded)}
            className="w-full px-4 py-2.5 flex items-center gap-2 hover:bg-surface-hover/50 transition-colors"
          >
            {isFilesExpanded ? (
              <ChevronDown className="w-3.5 h-3.5 text-text-muted" />
            ) : (
              <ChevronRight className="w-3.5 h-3.5 text-text-muted" />
            )}
            <span className="text-xs font-medium text-text">文件变更</span>
            <span className="text-[10px] text-text-dim bg-background/50 px-1.5 py-0.5 rounded">
              {fileChanges.length} 个文件
            </span>
          </button>
          
          {isFilesExpanded && (
            <div className="px-4 pb-3 space-y-1">
              {fileChanges.map((file, index) => (
                <div 
                  key={index} 
                  className={`py-2 px-2.5 rounded-lg transition-colors ${
                    file.status === 'writing' ? 'bg-violet-500/10 border border-violet-500/20' : 'hover:bg-surface-hover/30'
                  }`}
                >
                  <div className="flex items-center justify-between">
                    <div className="flex items-center gap-2">
                      {file.status === 'writing' ? (
                        <Loader2 className="w-4 h-4 text-violet-400 animate-spin" />
                      ) : file.status === 'done' ? (
                        <CheckCircle2 className="w-4 h-4 text-green-400" />
                      ) : (
                        <FileText className="w-4 h-4 text-text-dim" />
                      )}
                      <div>
                        <span className="text-xs text-text font-medium">{file.name}</span>
                        {file.path && (
                          <span className="text-[10px] text-text-dim ml-1.5">{file.path}</span>
                        )}
                      </div>
                    </div>
                    <div className="flex items-center gap-2 text-xs font-mono">
                      {file.additions > 0 && (
                        <span className="text-green-400 bg-green-500/10 px-1.5 py-0.5 rounded">
                          +{file.additions}
                        </span>
                      )}
                      {file.deletions > 0 && (
                        <span className="text-red-400 bg-red-500/10 px-1.5 py-0.5 rounded">
                          -{file.deletions}
                        </span>
                      )}
                    </div>
                  </div>
                  
                  {/* 操作详情 */}
                  {file.operations && file.operations.length > 0 && (
                    <div className="mt-2 pl-6 space-y-0.5">
                      {file.operations.map((op, i) => (
                        <div key={i} className="text-[10px] text-text-dim flex items-center gap-1.5">
                          <span className="w-1 h-1 rounded-full bg-text-dim" />
                          <span>{op}</span>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              ))}
            </div>
          )}
        </div>
      )}
    </div>
  )
}

// Hook to manage agent state - 重新设计
export function useAgentState() {
  const [isVisible, setIsVisible] = useState(false)
  const [steps, setSteps] = useState<AgentStep[]>([])
  const [fileChanges, setFileChanges] = useState<FileChange[]>([])
  const [thinkingTime, setThinkingTime] = useState(0)
  const [currentAction, setCurrentAction] = useState<string>('')
  const [thinkingStartTime, setThinkingStartTime] = useState<number | null>(null)

  // 更新思考时间
  useEffect(() => {
    let interval: NodeJS.Timeout
    if (thinkingStartTime) {
      interval = setInterval(() => {
        setThinkingTime(Math.floor((Date.now() - thinkingStartTime) / 1000))
      }, 1000)
    }
    return () => clearInterval(interval)
  }, [thinkingStartTime])

  // 开始操作
  const startOperation = (operationType: 'create' | 'edit' | 'analyze') => {
    setIsVisible(true)
    setThinkingStartTime(Date.now())
    setThinkingTime(0)
    
    if (operationType === 'create') {
      setCurrentAction('正在创建文档...')
      setSteps([
        { id: '1', type: 'thinking', description: '分析用户需求', status: 'running' },
        { id: '2', type: 'creating', description: '生成文档内容', status: 'pending' },
        { id: '3', type: 'editing', description: '写入文件', status: 'pending' },
      ])
      setFileChanges([
        { name: '新文档', additions: 0, deletions: 0, status: 'pending' }
      ])
    } else if (operationType === 'edit') {
      setCurrentAction('正在修改文档...')
      setSteps([
        { id: '1', type: 'reading', description: '读取当前文档', status: 'running' },
        { id: '2', type: 'thinking', description: '分析修改需求', status: 'pending' },
        { id: '3', type: 'editing', description: '执行修改', status: 'pending' },
      ])
      setFileChanges([
        { name: '当前文档', additions: 0, deletions: 0, status: 'pending' }
      ])
    } else {
      setCurrentAction('正在分析文档...')
      setSteps([
        { id: '1', type: 'reading', description: '读取文档内容', status: 'running' },
        { id: '2', type: 'thinking', description: '分析处理', status: 'pending' },
      ])
    }
  }

  // 添加步骤
  const addStep = (step: Omit<AgentStep, 'id'>) => {
    const newStep: AgentStep = {
      ...step,
      id: Date.now().toString(),
    }
    setSteps(prev => [...prev, newStep])
    return newStep.id
  }

  // 更新步骤状态
  const updateStep = (stepId: string, updates: Partial<AgentStep>) => {
    setSteps(prev => prev.map(s => 
      s.id === stepId ? { ...s, ...updates } : s
    ))
  }

  // 完成当前步骤并开始下一个
  const completeCurrentStep = () => {
    setSteps(prev => {
      const runningIndex = prev.findIndex(s => s.status === 'running')
      if (runningIndex === -1) return prev
      
      const newSteps = [...prev]
      newSteps[runningIndex] = { 
        ...newSteps[runningIndex], 
        status: 'completed',
        timestamp: new Date()
      }
      
      // 开始下一个步骤
      if (runningIndex + 1 < newSteps.length) {
        newSteps[runningIndex + 1] = { 
          ...newSteps[runningIndex + 1], 
          status: 'running' 
        }
      }
      
      return newSteps
    })
  }

  // 设置当前操作描述
  const setAction = (action: string) => {
    setCurrentAction(action)
  }

  // 更新文件变更
  const updateFileChange = (index: number, updates: Partial<FileChange>) => {
    setFileChanges(prev => prev.map((f, i) => 
      i === index ? { ...f, ...updates } : f
    ))
  }

  // 添加文件操作记录
  const addFileOperation = (index: number, operation: string) => {
    setFileChanges(prev => prev.map((f, i) => 
      i === index ? { 
        ...f, 
        operations: [...(f.operations || []), operation] 
      } : f
    ))
  }

  // 完成操作
  const completeOperation = (fileName?: string, additions?: number, deletions?: number) => {
    // 完成所有步骤
    setSteps(prev => prev.map(s => ({ 
      ...s, 
      status: 'completed' as const,
      timestamp: s.timestamp || new Date()
    })))
    
    // 更新文件变更
    if (fileName) {
      setFileChanges([{
        name: fileName,
        additions: additions || 0,
        deletions: deletions || 0,
        status: 'done'
      }])
    } else {
      setFileChanges(prev => prev.map(f => ({ ...f, status: 'done' as const })))
    }
    
    setThinkingStartTime(null)
    setCurrentAction('操作完成')
    
    // 3秒后隐藏面板
    setTimeout(() => {
      setIsVisible(false)
      setSteps([])
      setFileChanges([])
      setThinkingTime(0)
      setCurrentAction('')
    }, 3000)
  }

  // 重置
  const reset = () => {
    setIsVisible(false)
    setSteps([])
    setFileChanges([])
    setThinkingTime(0)
    setThinkingStartTime(null)
    setCurrentAction('')
  }

  return {
    isVisible,
    steps,
    fileChanges,
    thinkingTime,
    currentAction,
    startOperation,
    addStep,
    updateStep,
    completeCurrentStep,
    setAction,
    updateFileChange,
    addFileOperation,
    completeOperation,
    reset,
    // 兼容旧接口
    tasks: steps.map(s => ({ id: s.id, text: s.description, status: s.status === 'running' ? 'running' : s.status === 'completed' ? 'completed' : 'pending' })),
    streamingPreview: '',
    updateTaskStatus: (taskId: string, status: 'pending' | 'running' | 'completed') => updateStep(taskId, { status }),
    completeTask: completeCurrentStep,
    updateStreamingPreview: () => {},
  }
}
