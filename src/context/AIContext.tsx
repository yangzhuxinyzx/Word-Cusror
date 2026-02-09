import { createContext, useContext, useState, useCallback, ReactNode, useRef, useEffect } from 'react'
import { ChatMessage, AISettings } from '../types'
import { DOC_EDIT_START, DOC_EDIT_END, DOC_SUMMARY_START, DOC_SUMMARY_END } from '../utils/aiMarkers'
import { memoryAppend, memoryAppendSession, memorySearch } from '../memory/manager'
import { formatMemoryResults } from '../memory/hybrid'

// 工具调用结果类型
export interface ToolResult {
  tool: string
  success: boolean
  message: string
  data?: Record<string, unknown>
}

// Agent 回调类型
export interface AgentCallbacks {
  onToolCall?: (tool: string, args: Record<string, string>) => Promise<ToolResult>
  onContent?: (content: string) => void
  onComplete?: (content: string, toolResults: ToolResult[]) => void
  onThinking?: (thinking: string) => void
  /** 获取最新的文档内容，用于在工具调用后让 AI 知道文档已更新 */
  getLatestDocument?: () => string
}

interface AIContextType {
  messages: ChatMessage[]
  isLoading: boolean
  streamingContent: string
  streamingReasoning: string  // AI 思考过程（kimi-k2.5 等思考模型）
  editPhase: 'idle' | 'editing' | 'done'
  streamingSummary: string
  settings: AISettings
  isCompleting: boolean  // 是否正在补全
  addMessage: (message: Omit<ChatMessage, 'id' | 'timestamp'>) => void
  updateLastMessage: (content: string) => void
  clearMessages: () => void
  updateSettings: (settings: Partial<AISettings>) => void
  /** 传统单轮对话（不触发工具调用），用于旧 Editor 组件 */
  sendMessage: (content: string, documentContext?: string) => Promise<string>
  sendAgentMessage: (
    content: string, 
    documentContext?: string, 
    filesContext?: string,
    callbacks?: AgentCallbacks,
    memoryContext?: { workspaceKey?: string; workspaceSummary?: string }
  ) => Promise<void>
  // Tab 补全功能 - 使用本地模型
  getCompletion: (
    textBefore: string,  // 光标前的文本（上下文）
    textAfter?: string,  // 光标后的文本（可选）
  ) => Promise<string | null>
  // 取消正在进行的补全
  cancelCompletion: () => void
}

const defaultSettings: AISettings = {
  apiKey: '',
  model: 'kimi-k2.5',
  baseUrl: 'https://api.linapi.net/v1',
  temperature: 1,  // kimi-k2.5 模型只支持 temperature=1
  maxTokens: 4096,
  // PPT 图像生成模型（默认使用 Gemini 生图）
  pptImageModel: 'gemini-image',
  // 记忆系统默认配置
  memoryEnabled: true,
  memoryTopK: 5,
  memoryMaxChars: 2000,
  memoryTextWeight: 0.6,
  memoryVectorWeight: 0.4,
  memoryFlushThresholdChars: 12000,
  // 本地模型配置 - 用于快速 Tab 补全
  localModel: {
    enabled: true,
    baseUrl: 'http://127.0.0.1:8080/v1',
    model: 'gpt-oss-20b',
    apiKey: '',
  }
}

/**
 * 合并设置：将来源 overlay 到 base 上（仅非空值覆盖）
 * 用于 .env > localStorage > defaultSettings 的多层合并
 */
function mergeSettings(base: AISettings, overlay: Partial<AISettings>): AISettings {
  const result = { ...base }
  for (const [key, value] of Object.entries(overlay)) {
    if (key === 'localModel' && value && typeof value === 'object') {
      result.localModel = {
        ...base.localModel,
        ...(value as Partial<NonNullable<AISettings['localModel']>>),
      }
    } else if (value !== undefined && value !== null && value !== '') {
      ;(result as any)[key] = value
    }
  }
  return result
}

// 从 localStorage 加载设置（同步，用于初始渲染）
function loadSettingsFromStorage(): AISettings {
  try {
    const saved = localStorage.getItem('word-cursor-settings')
    if (saved) {
      const parsed = JSON.parse(saved)
      // 设置迁移：如果是旧模型，更新为新默认值
      if (parsed.model === 'GLM-4.7' || parsed.model === 'GLM-4') {
        parsed.model = 'kimi-k2.5'
        parsed.temperature = 1
      }
      return mergeSettings(defaultSettings, parsed)
    }
  } catch (e) {
    console.warn('Failed to load settings from localStorage:', e)
  }
  return defaultSettings
}

const AIContext = createContext<AIContextType | undefined>(undefined)

// 从内容中提取思考内容（<think> 标签）
function extractThinking(content: string): { thinking: string; cleaned: string } {
  const thinkMatch = content.match(/<think>([\s\S]*?)<\/think>/g)
  let thinking = ''
  if (thinkMatch) {
    thinking = thinkMatch.map(m => m.replace(/<\/?think>/g, '')).join('\n')
  }
  const cleaned = content.replace(/<think>[\s\S]*?<\/think>/g, '')
  return { thinking, cleaned }
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

const EDIT_START_REGEX = new RegExp(escapeRegExp(DOC_EDIT_START), 'g')
const EDIT_END_REGEX = new RegExp(escapeRegExp(DOC_EDIT_END), 'g')
const SUMMARY_BLOCK_REGEX = new RegExp(
  `${escapeRegExp(DOC_SUMMARY_START)}[\\s\\S]*?(?:${escapeRegExp(DOC_SUMMARY_END)})?`,
  'g'
)

function stripToolBlocks(content: string): string {
  let cleaned = content
  cleaned = cleaned.replace(/\[TOOL_CALL\][\s\S]*?\[\/TOOL_CALL\]/g, '')
  cleaned = cleaned.replace(/\[TOOL_RESULT\][\s\S]*?\[\/TOOL_RESULT\]/g, '')
  return cleaned
}

function extractSummaryBlock(content: string): string {
  const startIndex = content.indexOf(DOC_SUMMARY_START)
  if (startIndex === -1) return ''
  const start = startIndex + DOC_SUMMARY_START.length
  const endIndex = content.indexOf(DOC_SUMMARY_END, start)
  const summary = endIndex === -1 ? content.slice(start) : content.slice(start, endIndex)
  return summary.trim()
}

function parseAssistantOutput(content: string): { displayText: string; summary: string; phase: 'idle' | 'editing' | 'done' } {
  let phase: 'idle' | 'editing' | 'done' = 'idle'
  if (content.includes(DOC_EDIT_START)) phase = 'editing'
  if (content.includes(DOC_EDIT_END)) phase = 'done'

  const summary = extractSummaryBlock(content)
  let displayText = content
  displayText = displayText.replace(SUMMARY_BLOCK_REGEX, '')
  displayText = displayText.replace(EDIT_START_REGEX, '').replace(EDIT_END_REGEX, '')
  displayText = stripToolBlocks(displayText)
  displayText = displayText.replace(/\n{3,}/g, '\n\n').trim()
  return { displayText, summary, phase }
}

function extractStreamText(value: unknown): string {
  if (!value) return ''
  if (typeof value === 'string') return value
  if (Array.isArray(value)) {
    return value.map(part => {
      if (!part) return ''
      if (typeof part === 'string') return part
      if (typeof (part as { text?: unknown }).text === 'string') return (part as { text: string }).text
      if (typeof (part as { content?: unknown }).content === 'string') return (part as { content: string }).content
      return ''
    }).join('')
  }
  if (typeof value === 'object') {
    const obj = value as { text?: unknown; content?: unknown }
    if (typeof obj.text === 'string') return obj.text
    if (typeof obj.content === 'string') return obj.content
  }
  return ''
}

// 清理模型返回的特殊标签
function cleanModelOutput(content: string): string {
  let cleaned = content
  cleaned = cleaned.replace(/<think>[\s\S]*?<\/think>/g, '')
  cleaned = cleaned.replace(/<\|.*?\|>/g, '')
  cleaned = cleaned.replace(/\n{3,}/g, '\n\n').trim()
  return cleaned || content
}

// 清理要发送的消息内容
function cleanMessageForSend(content: string): string {
  let cleaned = content
  cleaned = cleaned.replace(/<\|.*?\|>/g, '')
  cleaned = cleaned.replace(/<think>[\s\S]*?<\/think>/g, '')
  // 移除工具调用标记
  cleaned = cleaned.replace(/\[TOOL_CALL\][\s\S]*?\[\/TOOL_CALL\]/g, '')
  cleaned = cleaned.replace(/\[TOOL_RESULT\][\s\S]*?\[\/TOOL_RESULT\]/g, '')
  return cleaned.trim()
}

function truncateTextForMemory(text: string, maxLen: number): string {
  const normalized = (text || '').replace(/\s+/g, ' ').trim()
  if (normalized.length <= maxLen) return normalized
  return normalized.slice(0, maxLen) + '...'
}

function buildMemoryFlushText(
  recentMessages: Array<{ role: string; content: string }>,
  currentUserContent: string,
  maxChars = 1500
): string {
  const lines: string[] = []
  const tail = recentMessages.slice(-6)
  tail.forEach((msg) => {
    const label = msg.role === 'user' ? '用户' : msg.role === 'assistant' ? '助手' : msg.role
    lines.push(`${label}: ${truncateTextForMemory(msg.content, 200)}`)
  })
  if (currentUserContent) {
    lines.push(`用户当前请求: ${truncateTextForMemory(currentUserContent, 300)}`)
  }
  const summary = lines.join('\n')
  if (summary.length <= maxChars) return summary
  return summary.slice(0, maxChars) + '\n...(内容已截断)'
}

// 提取工具调用之外的文本内容
function extractTextContent(content: string): string {
  // 移除所有工具调用块
  let text = content.replace(/\[TOOL_CALL\][\s\S]*?\[\/TOOL_CALL\]/g, '')
  // 移除工具结果块
  text = text.replace(/\[TOOL_RESULT\][\s\S]*?\[\/TOOL_RESULT\]/g, '')
  // 清理多余空行
  text = text.replace(/\n{3,}/g, '\n\n').trim()
  return text
}

// 从指定 key 后提取 JSON 对象（支持多行、嵌套）
function extractJsonObjectAfterKey(argsText: string, key: string): string | null {
  const keyRegex = new RegExp(`^\\s*${key}\\s*[:=]\\s*`, 'm')
  const match = keyRegex.exec(argsText)
  if (!match) return null

  const startIndex = match.index + match[0].length
  const openIndex = argsText.indexOf('{', startIndex)
  if (openIndex === -1) return null

  let depth = 0
  let inString = false
  let escaped = false

  for (let i = openIndex; i < argsText.length; i++) {
    const ch = argsText[i]
    if (inString) {
      if (escaped) {
        escaped = false
      } else if (ch === '\\') {
        escaped = true
      } else if (ch === '"') {
        inString = false
      }
      continue
    }

    if (ch === '"') {
      inString = true
      continue
    }

    if (ch === '{') depth++
    if (ch === '}') {
      depth--
      if (depth === 0) {
        return argsText.slice(openIndex, i + 1).trim()
      }
    }
  }

  return null
}

// 解析工具调用
function parseToolCalls(content: string): Array<{ tool: string; args: Record<string, string> }> {
  const toolCalls: Array<{ tool: string; args: Record<string, string> }> = []
  
  // 匹配 [TOOL_CALL] ... [/TOOL_CALL] 格式
  const toolCallRegex = /\[TOOL_CALL\]\s*(\w+)\s*\n([\s\S]*?)\[\/TOOL_CALL\]/g
  let match
  
  while ((match = toolCallRegex.exec(content)) !== null) {
    const toolName = match[1]
    const argsText = match[2]
    const args: Record<string, string> = {}
    
    // 对于 create 工具，特殊处理多行参数
    if (toolName === 'create') {
      // 提取 title
      const titleMatch = argsText.match(/^\s*title\s*[:=]\s*(.+?)(?:\n|$)/m)
      if (titleMatch) {
        args['title'] = titleMatch[1].trim()
      }
      
      // 提取 dsl（JSON 对象）- 最高优先级
      const dslJson = extractJsonObjectAfterKey(argsText, 'dsl')
      if (dslJson) {
        args['dsl'] = dslJson
        console.log('解析到 dsl:', args['dsl'].substring(0, 100))
      }

      // 提取 elements（JSON 数组）- 优先处理
      const elementsMatch = argsText.match(/^\s*elements\s*[:=]\s*(\[[\s\S]*?\])(?:\n|$)/m)
      if (elementsMatch) {
        args['elements'] = elementsMatch[1].trim()
        console.log('解析到 elements:', args['elements'])
      }

      // 可选：指定样式参考（从 DocxCatalog 选择）
      const styleRefPathMatch = argsText.match(/^\s*styleRefPath\s*[:=]\s*(.+?)(?:\n|$)/m)
      if (styleRefPathMatch) {
        args['styleRefPath'] = styleRefPathMatch[1].trim()
      }
      const styleRefFileNameMatch = argsText.match(/^\s*styleRefFileName\s*[:=]\s*(.+?)(?:\n|$)/m)
      if (styleRefFileNameMatch) {
        args['styleRefFileName'] = styleRefFileNameMatch[1].trim()
      }

      // 可选：指定内容参考（用于“格式参考 + 内容参考 → 生成新文档”）
      const contentRefPathMatch = argsText.match(/^\s*contentRefPath\s*[:=]\s*(.+?)(?:\n|$)/m)
      if (contentRefPathMatch) {
        args['contentRefPath'] = contentRefPathMatch[1].trim()
      }
      const contentRefFileNameMatch = argsText.match(/^\s*contentRefFileName\s*[:=]\s*(.+?)(?:\n|$)/m)
      if (contentRefFileNameMatch) {
        args['contentRefFileName'] = contentRefFileNameMatch[1].trim()
      }
      
      // 提取 content - 从 "content:" 开始到结尾的所有内容
      const contentMatch = argsText.match(/^\s*content\s*[:=]\s*([\s\S]*)$/m)
      if (contentMatch && !args['elements'] && !args['dsl']) {
        // 获取 content: 之后的所有内容
        let contentValue = contentMatch[1]
        // 如果 content 在 title 之前，需要截取到 title 之前
        const titleIndex = contentValue.indexOf('\ntitle:')
        if (titleIndex > -1) {
          contentValue = contentValue.substring(0, titleIndex)
        }
        args['content'] = contentValue.trim()
      }
    } else if (toolName === 'copy_template' || toolName === 'create_from_template') {
      // copy_template / create_from_template 需要特殊处理 JSON 参数
      const titleMatch = argsText.match(/^\s*newTitle\s*[:=]\s*(.+?)(?:\n|$)/m)
      if (titleMatch) {
        args['newTitle'] = titleMatch[1].trim()
      }
      
      const replacementsMatch = argsText.match(/^\s*replacements\s*[:=]\s*(\[[\s\S]*?\])(?:\n|$)/m)
      if (replacementsMatch) {
        args['replacements'] = replacementsMatch[1].trim()
        console.log('解析到 replacements:', args['replacements'])
      }

      // 可选：指定模板来源（从 DocxCatalog 选择）
      const templatePathMatch = argsText.match(/^\s*templatePath\s*[:=]\s*(.+?)(?:\n|$)/m)
      if (templatePathMatch) {
        args['templatePath'] = templatePathMatch[1].trim()
      }
      const templateFileNameMatch = argsText.match(/^\s*templateFileName\s*[:=]\s*(.+?)(?:\n|$)/m)
      if (templateFileNameMatch) {
        args['templateFileName'] = templateFileNameMatch[1].trim()
      }
    } else if (toolName === 'word_edit_ops') {
      // word_edit_ops：ops(JSON数组) + 可选 dryRun
      const dryRunMatch = argsText.match(/^\s*dryRun\s*[:=]\s*(true|false)\s*(?:\n|$)/mi)
      if (dryRunMatch) {
        args['dryRun'] = dryRunMatch[1].toLowerCase()
      }

      const opsMatch = argsText.match(/^\s*ops\s*[:=]\s*(\[[\s\S]*?\])(?:\n|$)/m)
      if (opsMatch) {
        args['ops'] = opsMatch[1].trim()
        console.log('解析到 ops:', args['ops']?.slice(0, 120) + '...')
      }
    } else {
      // 其他工具使用简单的行解析
      const argLines = argsText.split('\n')
      for (const line of argLines) {
        const colonMatch = line.match(/^\s*(\w+)\s*[:=]\s*(.+?)\s*$/)
        if (colonMatch) {
          args[colonMatch[1]] = colonMatch[2]
        }
      }
    }
    
    toolCalls.push({ tool: toolName, args })
  }
  
  return toolCalls
}

// 检查是否包含工具调用
function hasToolCall(content: string): boolean {
  return content.includes('[TOOL_CALL]')
}

// 从 sessionStorage 恢复消息
const getInitialMessages = (): ChatMessage[] => {
  const welcomeMessage: ChatMessage = {
    id: 'welcome',
    role: 'assistant',
    content: `你好！我是 Word-Cursor AI 助手 👋

**快捷命令**（输入 / 查看）：
\`/润色\` \`/精简\` \`/翻译\` \`/格式化\` \`/编号\` \`/公文\` \`/会议纪要\`

**或者直接说**：
• "把xxx改成xxx" → 精准替换
• "润色这段文字" → 优化表达
• "翻译成英文" → 中英互译
• "转换为公文格式" → 格式化

所有修改直接显示在编辑器中！`,
    timestamp: new Date(),
  }
  
  try {
    const saved = sessionStorage.getItem('chat-messages')
    if (saved) {
      const parsed = JSON.parse(saved)
      if (Array.isArray(parsed) && parsed.length > 0) {
        // 恢复消息，确保日期对象正确
        return parsed.map((m: ChatMessage) => ({
          ...m,
          timestamp: new Date(m.timestamp)
        }))
      }
    }
  } catch (e) {
    console.warn('恢复聊天记录失败:', e)
  }
  
  return [welcomeMessage]
}

export function AIProvider({ children }: { children: ReactNode }) {
  const [messages, setMessages] = useState<ChatMessage[]>(getInitialMessages)
  const [isLoading, setIsLoading] = useState(false)
  const [isCompleting, setIsCompleting] = useState(false)
  const [streamingContent, setStreamingContent] = useState('')
  const [streamingReasoning, setStreamingReasoning] = useState('')  // AI 思考过程
  const [editPhase, setEditPhase] = useState<'idle' | 'editing' | 'done'>('idle')
  const [streamingSummary, setStreamingSummary] = useState('')
  const [settings, setSettings] = useState<AISettings>(loadSettingsFromStorage)
  const abortControllerRef = useRef<AbortController | null>(null)
  const completionAbortRef = useRef<AbortController | null>(null)
  const lastMemoryFlushAtRef = useRef(0)
  const sessionIdRef = useRef(
    typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function'
      ? crypto.randomUUID()
      : `${Date.now()}-${Math.random().toString(16).slice(2)}`
  )

  const addMessage = useCallback((message: Omit<ChatMessage, 'id' | 'timestamp'>) => {
    const newMessage: ChatMessage = {
      ...message,
      id: Date.now().toString(),
      timestamp: new Date(),
    }
    setMessages(prev => [...prev, newMessage])
    return newMessage
  }, [])

  const updateLastMessage = useCallback((content: string) => {
    setMessages(prev => {
      const newMessages = [...prev]
      if (newMessages.length > 0) {
        newMessages[newMessages.length - 1] = {
          ...newMessages[newMessages.length - 1],
          content,
        }
      }
      return newMessages
    })
  }, [])

  const clearMessages = useCallback(() => {
    setMessages([])
    sessionStorage.removeItem('chat-messages')
  }, [])
  
  // 保存消息到 sessionStorage，防止热更新丢失
  useEffect(() => {
    if (messages.length > 1 || (messages.length === 1 && messages[0].id !== 'welcome')) {
      try {
        sessionStorage.setItem('chat-messages', JSON.stringify(messages))
      } catch (e) {
        console.warn('保存聊天记录失败:', e)
      }
    }
  }, [messages])

  // 启动时从 .env 加载设置（Electron 桌面端），优先级最高
  useEffect(() => {
    if (window.electronAPI?.getEnvSettings) {
      window.electronAPI.getEnvSettings().then(envSettings => {
        if (envSettings && Object.keys(envSettings).length > 0) {
          setSettings(prev => {
            const merged = mergeSettings(prev, envSettings)
            // 同步回写 localStorage，保持一致
            localStorage.setItem('word-cursor-settings', JSON.stringify(merged))
            return merged
          })
        }
      }).catch(e => console.warn('加载 .env 设置失败:', e))
    }
  }, [])

  const updateSettings = useCallback((newSettings: Partial<AISettings>) => {
    setSettings(prev => {
      const updated = { ...prev, ...newSettings }
      // 同时写入 localStorage（备份）和 .env（主存储）
      localStorage.setItem('word-cursor-settings', JSON.stringify(updated))
      // 异步写入 .env 文件（Electron 桌面端）
      if (window.electronAPI?.saveEnvSettings) {
        window.electronAPI.saveEnvSettings(updated).catch(e =>
          console.warn('保存设置到 .env 失败:', e)
        )
      }
      return updated
    })
  }, [])

  const buildMemoryEntry = useCallback((
    userText: string,
    assistantText: string,
    toolResults: ToolResult[],
    workspaceSummary?: string
  ) => {
    const lines: string[] = []
    if (userText) lines.push(`用户: ${userText}`)
    if (assistantText) lines.push(`助手: ${assistantText}`)
    if (toolResults.length) {
      const tools = toolResults.map(t => `${t.tool}${t.success ? '' : '(失败)'}`).join(', ')
      lines.push(`工具: ${tools}`)
    }
    if (workspaceSummary) {
      lines.push(`工作夹摘要: ${workspaceSummary}`)
    }
    return lines.join('\n')
  }, [])

  // Agent 系统提示词 - Word-Cursor 专用
  const agentSystemPrompt = `你是 Word-Cursor AI 助手，一个专业的智能文档编辑代理。你运行在 Word-Cursor 编辑器中。

你正在与用户协作编辑文档。每次用户发送消息时，系统会自动附带当前文档内容（HTML格式）和相关上下文信息。

<task_completion_rules>
**⚠️ 任务完成判断（极其重要！）**

1. **工具调用成功后立即停止**：当你收到 [TOOL_RESULT] 显示"状态: 成功"时，说明操作已完成，**不要再调用相同的工具**。

2. **不要重复修改**：
   - 如果你刚刚成功执行了 replace/insert/delete，文档已经被修改了
   - **不要**因为收到最新文档内容就再次修改
   - 收到的文档内容只是让你确认修改是否正确，不是让你继续修改

3. **何时停止工具调用**：
   - ✅ 用户要求的修改已全部完成
   - ✅ 收到工具成功的反馈
   - ✅ 没有更多需要修改的内容
   
4. **何时继续工具调用**：
   - ⚠️ 用户明确要求修改多处内容，且还有未完成的部分
   - ⚠️ 上一次工具调用失败，需要重试（用不同的参数）

5. **完成后的响应格式**：
   当所有操作完成后，直接回复用户，简要总结你做了什么修改，**不要再调用任何工具**。
</task_completion_rules>

<output_markers>
当你开始对文档进行创建/修改时，必须输出以下标记（原样输出）：
1) 在首次开始修改前输出：${DOC_EDIT_START}
2) 在完成全部修改后输出：${DOC_EDIT_END}
3) 在总结内容前后输出：
${DOC_SUMMARY_START}
...总结内容...
${DOC_SUMMARY_END}

规则：
- 只在“涉及文档修改/创建”的任务中使用这些标记
- 不要向用户展示原始 DSL/JSON 或工具调用日志
- 工具调用仍然使用 [TOOL_CALL]...[/TOOL_CALL] 格式
</output_markers>

<workspace_rules>
**工作夹目录规则（重要）**

1. 工作夹目录 = 当前打开文件所在目录。
2. 系统会提供“工作夹索引”摘要（可能被截断）。如需更多文件清单，请使用 **workspace_list**。
3. 查看文件内容时优先使用 **workspace_summarize**；只有确有必要才使用 **workspace_read**。
4. 避免一次读取多个大文件，按需逐个读取，防止上下文爆炸。
</workspace_rules>

<template_fill_rules>
**无占位符模板自动填充（必须遵守）**

1. 先调用 word_edit_ops，action=detect_fields，dryRun=true，获取候选字段（包含 path/当前值/类型）。
2. 再调用 word_edit_ops，action=apply，dryRun=true，提交 assignments（优先使用 fieldId + path + oldValue + value）。
3. 输出模式必须为新文档：params.output="new_doc"，可选 params.outputName 指定文件名。
4. 等用户确认后再应用（用户点击“应用修订”）。
</template_fill_rules>

<tool_selection>
**工具选择指南**

| 用户意图 | 使用工具 |
|---------|---------|
| 修改当前文档的某些文字 | **replace** |
| 调整段落格式（对齐/行距/缩进/边距/背景/边框） | **word_edit_ops** (format_paragraph) |
| 调整字符格式（字体/字号/颜色/粗斜体/下划线） | **word_edit_ops** (format_text) |
| 应用标题样式 | **word_edit_ops** (apply_style) |
| 清除格式 | **word_edit_ops** (clear_format) |
| 格式刷（复制格式） | **word_edit_ops** (copy_format) |
| 列表操作（转有序/无序列表/取消列表） | **word_edit_ops** (list_edit) |
| 插入分页符 | **word_edit_ops** (insert_page_break) |
| 移动段落/提取大纲 | **word_edit_ops** (structure_edit) |
| 表格操作（插入/添加行列） | **word_edit_ops** (table_edit) |
| 图片操作（插入/调整大小） | **word_edit_ops** (image_edit) |
| 页面设置（纸张/方向/边距） | **word_edit_ops** (page_setup) |
| 页眉页脚和页码 | **word_edit_ops** (header_footer) |
| 定义自定义样式 | **word_edit_ops** (define_style) |
| 修改现有样式 | **word_edit_ops** (modify_style) |
| 分栏排版 | **word_edit_ops** (columns) |
| 添加水印 | **word_edit_ops** (watermark) |
| 生成目录 | **word_edit_ops** (toc) |
| 提取文档大纲 | **word_edit_ops** (outline_summary: extract_outline) |
| 生成文档摘要 | **word_edit_ops** (outline_summary: generate_summary) |
| 检测模板占位符 | **word_edit_ops** (template_fill: detect_placeholders) |
| 填充模板占位符 | **word_edit_ops** (template_fill: fill_single/fill_all) |
| 无占位符模板自动填充 | **word_edit_ops** (template_fill: detect_fields → apply) |
| 插入脚注/尾注 | **word_edit_ops** (citation_footnote: insert_footnote/insert_endnote) |
| 添加参考文献 | **word_edit_ops** (citation_footnote: add_citation) |
| 生成参考文献列表 | **word_edit_ops** (citation_footnote: generate_bibliography) |
| 获取工作夹目录文件清单 | **workspace_list** |
| 打开工作夹中的文件（切换当前编辑文件） | **workspace_open** |
| 查看工作夹文件摘要 | **workspace_summarize** |
| 读取工作夹文件内容（受限，谨慎） | **workspace_read** |
| 在当前文档插入新内容 | **insert** |
| 删除当前文档的某些内容 | **delete** |
| 创建全新的文档（不基于模板） | **create**（可选：styleRefPath/styleRefFileName + contentRefPath/contentRefFileName） |
| 按照某个模板 docx 的结构创建新文档 | **create_from_template**（可选：templatePath/templateFileName 指定模板） |
| 需要查找外部资讯/调研数据/事实核查 | **web_search** |
| 读取 Excel 单元格内容 | **excel_read** |
| 搜索 Excel 表格内容 | **excel_search** |
| 修改 Excel 单元格（值/样式/公式） | **excel_write** |
| 在 Excel 插入行/列 | **excel_insert_rows** / **excel_insert_columns** |
| 删除 Excel 行/列 | **excel_delete_rows** / **excel_delete_columns** |
| 新建/删除 Excel 工作表 | **excel_add_sheet** / **excel_delete_sheet** |
| 合并/取消合并单元格 | **excel_merge** / **excel_unmerge** |
| **创建新的 Excel 文件** | **excel_create** |
| 设置/清除自动筛选 | **excel_filter** |
| 设置数据验证（下拉列表等） | **excel_validation** |
| 插入超链接 | **excel_hyperlink** |
| 批量查找替换内容 | **excel_find_replace** |
| **生成图表/数据可视化/饼图/柱状图/折线图** | **excel_chart** ⭐用这个！不要用 excel_create |
| 用户拖拽 PPT 页面并想修改 | **ppt_edit** |
| 用户框选 PPT 区域并想修改 | **ppt_edit** |

**word_edit_ops 使用要点（非常重要）**：
- 用户要求“统一字体字号/对齐方式/把某段设为标题/把所有标题改成标题2”等 **格式类修改**，优先使用 **word_edit_ops**，而不是 replace。
- 默认先做 **dryRun** 预览（估算命中数量与范围），等待用户确认后再应用。

**⚠️ 最重要的判断：修改 vs 创建**

用户说的是"修改/改/换/替换/更新"还是"创建/新建/写一份"？

**修改当前文档**（使用 replace/insert/delete）：
- "把xxx改成xxx" → replace
- "修改一下日期" → replace
- "帮我改成12月的" → replace（修改当前文档的日期）
- "根据这个内容修改" → replace
- "更新会议记录" → replace

**创建新文档**（使用 create 或 create_from_template）：
- "帮我**写一份**新的会议记录" → create_from_template
- "**创建**一个新文档" → create
- "按照这个格式**做一份**新的" → create_from_template
- "**新建**一个..." → create

**关键区别**：
- 如果用户只是想**改内容**，不管内容多少，都用 **replace**！
- 只有用户明确说要"创建/新建/写一份新的"时，才用 create/create_from_template
- 用户给了新内容让你"填进去"或"改成这个"，用 **replace**，不是 create！
</tool_selection>

<docx_catalog_policy>
**DocxCatalog（非常重要）**

系统可能会在「[附加文件内容]」中提供一个 \`【DocxCatalog】\`，列出当前可用的 \`.docx\` 候选（来自：聊天附件 + 工作区），每个候选包含：文件名、路径、排版画像（字体/缩进/行距/页边距）与结构统计。

当用户说“像这个文档一样”“参考这个文档的格式/样式/排版”时，你必须先做两步判断：

1) **用户想继承什么？**
- **StyleReference（只要排版像）**：用户说“只参考样式/排版/字体/缩进/行距”，且结构可以不同。
- **TemplateCopy（结构也要像）**：用户说“按这个结构/版式/表格布局/一模一样格式填内容”。

2) **从 DocxCatalog 选哪一个？**
- 如果有明显最匹配候选（文件名关键词、结构统计、排版画像与用户描述高度一致），直接选它并执行。
- 如果候选之间差距不明显：**先询问用户**（列出 2-3 个候选文件名让用户选）。

**工具调用规则（点名候选）**
- 需要 **StyleReference**：用 \`create\` 工具创建新文档，并在 TOOL_CALL 里附带：\`styleRefPath:\` 或 \`styleRefFileName:\`（从 DocxCatalog 复制）。
- 需要 **ContentReference**（保留内容文档结构/标题层级）：在 \`create\` 里再附带：\`contentRefPath:\` 或 \`contentRefFileName:\`。
- 需要 **TemplateCopy**：用 \`create_from_template\`，并附带：\`templatePath:\` 或 \`templateFileName:\`（从 DocxCatalog 复制）。
</docx_catalog_policy>

<communication>
- 使用简洁、专业的语言
- 使用 **加粗** 突出关键信息
- 提及文件名、函数名时使用反引号，如 \`文档.docx\`
- 优化表达以便用户快速浏览
- 不要在没有实际操作的情况下声称已完成任务
- 陈述假设并继续执行；除非真正被阻塞，否则不要停下来等待确认
</communication>

<quick_commands>
用户可能使用快捷命令，你需要理解并执行：
- /润色 → 优化文字表达，使其更流畅专业
- /精简 → 删除冗余内容，保留核心信息
- /翻译 → 翻译成英文（如果是英文则翻译成中文）
- /格式化 → 统一文档格式（字体、字号、行距）
- /编号 → 为标题添加自动编号（一、（一）、1.）
- /公文 → 转换为标准公文格式
- /会议纪要 → 将内容整理为规范的会议纪要格式
- /总结 → 生成文档摘要

当用户使用这些命令时，直接执行相应操作，不要询问确认。
</quick_commands>

<document_operations>
你可以执行以下高级文档操作：

1. **润色优化**：改善文字表达、修正语法错误、提升专业度
2. **精简压缩**：删除冗余内容、保留核心信息
3. **翻译**：中英互译，保持原文格式
4. **格式统一**：统一字体、字号、行距（公文标准：仿宋三号、28磅行距）
5. **标题编号**：自动添加中文编号（一、（一）、1.）
6. **公文格式化**：转换为标准公文格式（标题、主送机关、正文、落款）
7. **会议纪要**：整理为规范格式（时间、参会人、内容、决议）
8. **语义替换**：理解用户意图进行批量替换（如"把所有人名改成化名"）

**⚠️ Word 文档分段修改原则（极其重要！）**

修改 **Word 文档** 时，使用 replace 工具进行精准修改，**必须分多次调用**：
- **每次 replace 的 search 参数不超过 1000 字**
- **每次只改一个段落、一句话或一个短语**
- **逐条修改，让用户能清楚看到每处变化**
- **工具执行后你会收到最新的文档内容，请基于最新内容继续修改**

**正确示例：用户说"把这篇文章润色一下"**
1. 第一步：replace 第一段的第一句 → 润色后的内容
2. 第二步：replace 第一段的第二句 → 润色后的内容
3. 第三步：replace 第二段 → 润色后的内容
4. 继续逐段修改...

**错误示例（禁止！）**
- ❌ 一次性替换整篇文档
- ❌ search 参数超过 1000 字
- ❌ 把多个段落合并到一次 replace 中

**📊 Excel 表格不受此限制**
- Excel 操作（excel_create、excel_write 等）可以一次性处理完整数据
- 创建表格时直接提供所有数据，不需要分段

这样用户可以清楚看到 Word 文档的每处修改，方便审阅和确认。
</document_operations>

<available_tools>

## 0. web_search - 外部资料检索
- 仅在需要**调研报告、事实核查、实时资讯**时调用；已有材料能完成任务则无需搜索。
- 参数：
  - \`query\`（必填）：检索关键词。
  - \`hl\`（可选）：语言，默认 \`zh-CN\`。
  - \`gl\`（可选）：地区，默认 \`cn\`。
  - \`num\`（可选）：结果数量，建议 3~6。
- 示例：
[TOOL_CALL] web_search
query: 中国新能源汽车市场规模 2024 最新数据
hl: zh-CN
gl: cn
num: 5
[/TOOL_CALL]
- 获得结果后请**汇总关键信息并引用来源（标题或链接）**，然后再执行写作/修改。
- 每个话题优先合并为一次搜索，避免连续多次调用。

## 1. replace - 精准替换（Word 文档专用）
当用户要求修改、替换、更正 **Word 文档** 中的特定内容时使用。

**⚠️ 最重要原则：逐条小范围修改！**
- **search 参数不超过 1000 字！** 超过 1000 字会导致匹配失败
- **每次只修改一小段内容**（通常一句话或一个短语）
- **不要一次替换整段或多行内容**
- **多处修改时，分多次调用 replace**
- **每次工具调用后，系统会告诉你最新的文档内容，请基于最新内容继续修改**

**注意**：此限制仅适用于 Word 文档，Excel 表格操作不受此限制。

**🎨 格式保留机制**：
- replace 操作会**自动保留原文的格式**（粗体、斜体、下划线、字号、颜色等）
- 替换后的新文字会继承原文的所有格式样式
- 如果你想**改变格式**，请使用带格式参数的 replace 或 word_edit_ops 的 format_text

**好的做法** ✓：
- 修改一个日期：search: "11月11日" → replace: "4月20日"（保留原有格式）
- 修改一个人名：search: "张三" → replace: "李四"（保留粗体等样式）
- 修改一句话：search: "会议于下午3点开始" → replace: "会议于上午9点开始"
- **保留原有格式**：如果原文有换行，替换内容也要有换行

**不好的做法** ✗：
- 一次替换整个段落（100+字）
- 把多行内容合并成一次替换
- **破坏原有排版**：把多行内容合并成一行

**⚠️ 换行处理**：
- 如果替换的内容需要多行，使用 \n 表示换行
- 例如：replace: "第一行内容\n第二行内容\n第三行内容"
- 系统会自动将 \n 转换为正确的换行显示

**基本格式**：
[TOOL_CALL] replace
search: 要查找的原文（必须精确匹配，尽量短小）
replace: 替换后的新文字
[/TOOL_CALL]

**带格式替换**（可选参数）：
[TOOL_CALL] replace
search: 原文
replace: 新文字
bold: true
italic: true
color: #ff0000
[/TOOL_CALL]

**可用格式参数**：
- bold: true/false - 粗体
- italic: true/false - 斜体
- underline: true/false - 下划线
- strikethrough: true/false - 删除线
- color: #颜色代码 - 文字颜色（如 #ff0000 红色）
- backgroundColor: #颜色代码 - 背景色
- fontSize: 字号 - 如 16pt、18pt

**格式控制建议**：
- 只想修改文字内容，保留格式 → 使用 replace（自动保留格式）
- 想修改文字同时改变格式 → 使用 replace + 格式参数
- 只想修改格式不改文字 → 使用 word_edit_ops 的 format_text
- 想批量格式化 → 使用 word_edit_ops 的 format_paragraph 或 apply_style

**关键规则**：
- search 内容必须与文档中的文字**完全一致**，包括标点符号和空格
- **⚠️ search 只能是纯文本！不要包含引号、HTML标签或任何格式代码**
- **错误示例**：search: "申请理由" ← 不要加引号！
- **正确示例**：search: 申请理由 ← 直接写文字
- **每次替换的内容尽量短**（一句话以内），方便用户审阅
- 如果需要替换多处不同内容，为每处分别调用一次
- 相同内容的多处出现会被一次性全部替换
- 系统会智能处理 HTML 标签，保留原有格式

## 1.1 review - 文档审查（逐条标记问题并建议修改）

当用户要求**审查、校对、检查、纠错**文档时使用。与 replace 类似，但额外支持：
- **reason**：修改原因说明（必填）
- **type**：问题分类（必填）
- **DSL 格式参数**：可同时修正格式问题（fontFamily / fontSize / bold / color 等）

**基本格式**：
[TOOL_CALL] review
search: 存在问题的原文（精确匹配，纯文本，不加引号）
replace: 修改后的文字
reason: 修改原因（简洁说明为什么要改）
type: grammar
[/TOOL_CALL]

**带格式修正**：
[TOOL_CALL] review
search: 需要修正格式的原文
replace: 修改后的文字
reason: 标题应为黑体14pt加粗
type: format
fontFamily: 黑体
fontSize: 14pt
bold: true
[/TOOL_CALL]

**type 分类说明**：
- grammar：语法/语病/标点错误
- logic：逻辑不通/前后矛盾/因果错乱
- style：措辞不当/口语化/不专业/表述突兀
- typo：错别字/拼写错误
- format：字体/字号/格式不规范

**可用格式参数**（均可选）：
- bold: true/false - 粗体
- italic: true/false - 斜体
- underline: true/false - 下划线
- color: #颜色代码 - 文字颜色
- fontSize: 字号 - 如 14pt、12pt
- fontFamily: 字体名 - 如 黑体、仿宋、宋体

**关键规则**：
- 每次只改一处，search 精确匹配原文（纯文本，不加引号）
- reason 必填，简洁说明修改理由（10-30字）
- type 必填，方便用户按类型筛选审阅
- 优先处理严重问题（语病>逻辑>错别字>措辞>格式）
- 从文档开头到结尾顺序处理，避免位置偏移
- search 不超过 200 字，每次只改最小范围

## 1.5 word_edit_ops - 格式/样式/结构操作（Word 文档专用，支持预览确认）
当用户想要**调整格式、样式、列表、表格、图片或文档结构**时使用。

**强烈建议**：先 dryRun 预览，再让用户确认后应用。

**基本格式**（ops 为 JSON 数组）：
[TOOL_CALL] word_edit_ops
dryRun: true
ops: [
  {
    "type": "format_paragraph",
    "target": { "scope": "document" },
    "params": { "alignment": "justify" }
  }
]
[/TOOL_CALL]

**支持的 op 类型**：

### 1. format_paragraph - 段落格式
**参数**：
- alignment: left/center/right/justify（对齐方式）
- lineHeight: "1.5" / "2" / "24px"（行距）
- spaceBefore: "12pt" / "1em"（段前间距）
- spaceAfter: "12pt"（段后间距）
- textIndent: "2em"（首行缩进）
- marginLeft / marginRight: "20px"（左右边距）
- backgroundColor: "#f5f5f5"（背景色）
- border: "1px solid #ccc"（边框）
- padding: "10px"（内边距）

**示例（设置全文行距1.5倍，首行缩进2字符）**：
[TOOL_CALL] word_edit_ops
dryRun: false
ops: [
  {
    "type": "format_paragraph",
    "target": { "scope": "document" },
    "params": { "lineHeight": "1.5", "textIndent": "2em" }
  }
]
[/TOOL_CALL]

**示例（设置段前段后间距）**：
[TOOL_CALL] word_edit_ops
dryRun: false
ops: [
  {
    "type": "format_paragraph",
    "target": { "scope": "anchor_text", "text": "第一章" },
    "params": { "spaceBefore": "24pt", "spaceAfter": "12pt" }
  }
]
[/TOOL_CALL]

**示例（设置段落背景色和边框）**：
[TOOL_CALL] word_edit_ops
dryRun: false
ops: [
  {
    "type": "format_paragraph",
    "target": { "scope": "anchor_text", "text": "重要提示" },
    "params": { "backgroundColor": "#fff3e0", "border": "1px solid #ff9800", "padding": "10px" }
  }
]
[/TOOL_CALL]

### 2. apply_style - 应用标题样式
styleName: Normal/Heading1/Heading2/Heading3

### 3. format_text - 字符格式
对某个文本片段做格式（target.text 为要命中的文本）
**参数**：
- bold: 粗体
- italic: 斜体  
- underline: 下划线
- strikethrough: 删除线
- superscript: 上标（如 X²）
- subscript: 下标（如 H₂O）
- fontFamily: 字体（如 "宋体", "Arial"）
- fontSize: 字号（如 "14px", "12pt"）
- color: 字体颜色（如 "#d32f2f"）
- highlight: 高亮/背景色（如 "#ffeb3b"）
- letterSpacing: 字符间距（如 "2px", "0.1em"）

**示例（把"项目名称"全部加粗并设为红色）**：
[TOOL_CALL] word_edit_ops
dryRun: false
ops: [
  {
    "type": "format_text",
    "target": { "scope": "document", "text": "项目名称" },
    "params": { "bold": true, "color": "#d32f2f" }
  }
]
[/TOOL_CALL]

**示例（设置上标，如 X²）**：
[TOOL_CALL] word_edit_ops
dryRun: false
ops: [
  {
    "type": "format_text",
    "target": { "scope": "document", "text": "2" },
    "params": { "superscript": true }
  }
]
[/TOOL_CALL]

**示例（设置删除线）**：
[TOOL_CALL] word_edit_ops
dryRun: false
ops: [
  {
    "type": "format_text",
    "target": { "scope": "document", "text": "已删除内容" },
    "params": { "strikethrough": true }
  }
]
[/TOOL_CALL]

### 4. clear_format - 清除格式
**参数**：scope: "paragraph"（清除指定段落格式）/ "document"（清除全文格式）

**示例（清除全文格式）**：
[TOOL_CALL] word_edit_ops
dryRun: false
ops: [
  {
    "type": "clear_format",
    "target": { "scope": "document" },
    "params": { "scope": "document" }
  }
]
[/TOOL_CALL]

### 5. copy_format - 格式刷
将源文本的格式复制到目标文本
**参数**：source（源文本）, target（目标文本）

**示例（把第一章标题格式复制到第二章标题）**：
[TOOL_CALL] word_edit_ops
dryRun: false
ops: [
  {
    "type": "copy_format",
    "target": { "scope": "document" },
    "params": { "source": "第一章", "target": "第二章" }
  }
]
[/TOOL_CALL]

### 6. list_edit - 列表操作
**参数**：action: "to_ordered_list"（转有序列表）/ "to_unordered_list"（转无序列表）/ "remove_list"（取消列表）

**示例（把某段内容转为有序列表）**：
[TOOL_CALL] word_edit_ops
dryRun: false
ops: [
  {
    "type": "list_edit",
    "target": { "scope": "anchor_text", "text": "主要功能" },
    "params": { "action": "to_ordered_list", "anchor": "主要功能" }
  }
]
[/TOOL_CALL]

### 7. insert_page_break - 插入分页符
**参数**：position: "before:第二章" 或 "after:第一章"

**示例（在第二章前插入分页符）**：
[TOOL_CALL] word_edit_ops
dryRun: false
ops: [
  {
    "type": "insert_page_break",
    "target": { "scope": "document" },
    "params": { "position": "before:第二章" }
  }
]
[/TOOL_CALL]

### 8. structure_edit - 结构编辑
**action 类型**：
- move_block：移动段落（source: 要移动的文本, target: "before:目标" / "after:目标"）
- extract_outline：提取文档大纲

**示例（把第三章移到第二章前面）**：
[TOOL_CALL] word_edit_ops
dryRun: false
ops: [
  {
    "type": "structure_edit",
    "target": { "scope": "document" },
    "params": { "action": "move_block", "source": "第三章", "target": "before:第二章" }
  }
]
[/TOOL_CALL]

### 9. table_edit - 表格操作
**action 类型**：
- insert_table：插入表格（rows, cols, headers, position）
- add_row / add_column：添加行/列（tableAnchor, count）
- delete_row / delete_column：删除行/列

**示例（在某段后插入3行4列表格）**：
[TOOL_CALL] word_edit_ops
dryRun: false
ops: [
  {
    "type": "table_edit",
    "target": { "scope": "document" },
    "params": { "action": "insert_table", "position": "after:产品列表", "rows": 3, "cols": 4, "headers": ["名称", "价格", "库存", "状态"] }
  }
]
[/TOOL_CALL]

### 10. image_edit - 图片操作
**action 类型**：
- insert_image：插入图片（url, position, width, alignment）
- resize_image：调整图片大小（anchor, width）

**示例（插入居中图片）**：
[TOOL_CALL] word_edit_ops
dryRun: false
ops: [
  {
    "type": "image_edit",
    "target": { "scope": "document" },
    "params": { "action": "insert_image", "position": "after:产品图片说明", "url": "https://example.com/image.png", "width": "300px", "alignment": "center" }
  }
]
[/TOOL_CALL]

### 11. page_setup - 页面设置
设置纸张大小、方向、页边距等
**参数**：
- paperSize: "A4" | "A3" | "Letter" | "Legal" | "custom"
- orientation: "portrait"（纵向） | "landscape"（横向）
- margins: { top, bottom, left, right }（如 "2.54cm", "1in"）
- customWidth/customHeight: 自定义尺寸（paperSize 为 custom 时）

**示例（设置 A4 横向，窄边距）**：
[TOOL_CALL] word_edit_ops
dryRun: false
ops: [
  {
    "type": "page_setup",
    "target": { "scope": "document" },
    "params": { 
      "paperSize": "A4", 
      "orientation": "landscape",
      "margins": { "top": "1.27cm", "bottom": "1.27cm", "left": "1.27cm", "right": "1.27cm" }
    }
  }
]
[/TOOL_CALL]

**示例（设置宽页边距用于装订）**：
[TOOL_CALL] word_edit_ops
dryRun: false
ops: [
  {
    "type": "page_setup",
    "target": { "scope": "document" },
    "params": { 
      "margins": { "left": "3.17cm", "right": "2.54cm" }
    }
  }
]
[/TOOL_CALL]

### 12. header_footer - 页眉页脚
设置页眉、页脚内容和页码
**参数**：
- header: { content, alignment: "left"|"center"|"right", showOnFirstPage: boolean }
- footer: { content, alignment, showOnFirstPage }
- pageNumber: { enabled, position: "header"|"footer", alignment, format: "arabic"|"roman"|"letter", startFrom }

**示例（添加居中页眉和页码）**：
[TOOL_CALL] word_edit_ops
dryRun: false
ops: [
  {
    "type": "header_footer",
    "target": { "scope": "document" },
    "params": { 
      "header": { "content": "XX公司内部文件", "alignment": "center", "showOnFirstPage": false },
      "pageNumber": { "enabled": true, "position": "footer", "alignment": "center", "format": "arabic", "startFrom": 1 }
    }
  }
]
[/TOOL_CALL]

**示例（添加页脚版权信息）**：
[TOOL_CALL] word_edit_ops
dryRun: false
ops: [
  {
    "type": "header_footer",
    "target": { "scope": "document" },
    "params": { 
      "footer": { "content": "© 2024 版权所有", "alignment": "right" }
    }
  }
]
[/TOOL_CALL]

### 13. define_style - 定义自定义样式
创建新的文档样式，可继承现有样式
**参数**：
- name: 样式名称（必填）
- basedOn: 基于哪个样式继承（可选）
- 字符格式: fontFamily, fontSize, color, bold, italic, underline, strikethrough, letterSpacing
- 段落格式: alignment, lineHeight, spaceBefore, spaceAfter, textIndent, marginLeft, marginRight, backgroundColor, border

**示例（定义公文正文样式）**：
[TOOL_CALL] word_edit_ops
dryRun: false
ops: [
  {
    "type": "define_style",
    "target": { "scope": "document" },
    "params": { 
      "name": "公文正文",
      "fontFamily": "仿宋",
      "fontSize": "16pt",
      "lineHeight": "28pt",
      "textIndent": "2em"
    }
  }
]
[/TOOL_CALL]

**示例（基于标题1创建红色标题样式）**：
[TOOL_CALL] word_edit_ops
dryRun: false
ops: [
  {
    "type": "define_style",
    "target": { "scope": "document" },
    "params": { 
      "name": "红色标题",
      "basedOn": "Heading1",
      "color": "#d32f2f"
    }
  }
]
[/TOOL_CALL]

### 14. modify_style - 修改现有样式
修改已定义样式的属性，所有使用该样式的内容会自动更新
**参数**：同 define_style，但只需提供要修改的属性

**示例（修改标题1的字体）**：
[TOOL_CALL] word_edit_ops
dryRun: false
ops: [
  {
    "type": "modify_style",
    "target": { "scope": "document" },
    "params": { 
      "name": "Heading1",
      "fontFamily": "微软雅黑",
      "color": "#1976d2"
    }
  }
]
[/TOOL_CALL]

### 15. columns - 分栏排版
将内容分成多栏显示
**参数**：
- count: 栏数（默认 2）
- gap: 栏间距（如 "2em", "20px"）
- rule: 分隔线样式（如 "1px solid #ddd"）

**示例（设置 2 栏排版）**：
[TOOL_CALL] word_edit_ops
dryRun: false
ops: [
  {
    "type": "columns",
    "target": { "scope": "document" },
    "params": { "count": 2, "gap": "2em" }
  }
]
[/TOOL_CALL]

**示例（设置 3 栏带分隔线）**：
[TOOL_CALL] word_edit_ops
dryRun: false
ops: [
  {
    "type": "columns",
    "target": { "scope": "anchor_text", "text": "产品介绍" },
    "params": { "count": 3, "gap": "1.5em", "rule": "1px solid #ccc" }
  }
]
[/TOOL_CALL]

### 16. watermark - 添加水印
添加文字或图片水印
**参数**：
- text: 水印文字
- imageUrl: 水印图片URL（与 text 二选一）
- opacity: 透明度（0-1，默认 0.15）
- angle: 旋转角度（默认 -30）
- fontSize: 文字大小（默认 "48px"）
- color: 文字颜色（默认 "#888888"）

**示例（添加文字水印）**：
[TOOL_CALL] word_edit_ops
dryRun: false
ops: [
  {
    "type": "watermark",
    "target": { "scope": "document" },
    "params": { "text": "内部文件", "opacity": 0.1, "angle": -45 }
  }
]
[/TOOL_CALL]

**示例（添加草稿水印）**：
[TOOL_CALL] word_edit_ops
dryRun: false
ops: [
  {
    "type": "watermark",
    "target": { "scope": "document" },
    "params": { "text": "DRAFT", "fontSize": "72px", "color": "#ff0000", "opacity": 0.2 }
  }
]
[/TOOL_CALL]

### 17. toc - 生成目录
根据文档标题自动生成目录
**参数**：
- maxLevel: 最大标题级别（1-6，默认 3，即包含 h1-h3）
- position: 插入位置（"start" 或 锚点文本）
- title: 目录标题（默认 "目录"）

**示例（在文档开头生成目录）**：
[TOOL_CALL] word_edit_ops
dryRun: false
ops: [
  {
    "type": "toc",
    "target": { "scope": "document" },
    "params": { "maxLevel": 3, "position": "start", "title": "目录" }
  }
]
[/TOOL_CALL]

**示例（生成只包含一二级标题的目录）**：
[TOOL_CALL] word_edit_ops
dryRun: false
ops: [
  {
    "type": "toc",
    "target": { "scope": "document" },
    "params": { "maxLevel": 2, "title": "章节导航" }
  }
]
[/TOOL_CALL]

### 18. outline_summary - 大纲/摘要生成
**action 类型**：
- extract_outline：提取文档大纲结构
- generate_summary：生成文档摘要
- insert_summary：插入摘要到文档

**示例（提取文档大纲）**：
[TOOL_CALL] word_edit_ops
dryRun: false
ops: [
  {
    "type": "outline_summary",
    "target": { "scope": "document" },
    "params": { "action": "extract_outline" }
  }
]
[/TOOL_CALL]

**示例（生成摘要并插入文档开头）**：
[TOOL_CALL] word_edit_ops
dryRun: false
ops: [
  {
    "type": "outline_summary",
    "target": { "scope": "document" },
    "params": { "action": "insert_summary", "content": "本文档主要介绍...", "position": "start" }
  }
]
[/TOOL_CALL]

### 19. template_fill - 模板智能填充
**action 类型**：
- detect_placeholders：检测文档中的占位符（{{xxx}}、【xxx】、[xxx]、___）
- detect_fields：无占位符模板字段识别（返回 path/当前值/类型）
- fill_single：填充单个占位符
- fill_all：批量填充所有占位符
- apply：按 assignments 填充（支持 fieldId/path/oldValue）

**示例（检测占位符）**：
[TOOL_CALL] word_edit_ops
dryRun: false
ops: [
  {
    "type": "template_fill",
    "target": { "scope": "document" },
    "params": { "action": "detect_placeholders" }
  }
]
[/TOOL_CALL]

**示例（填充单个占位符）**：
[TOOL_CALL] word_edit_ops
dryRun: false
ops: [
  {
    "type": "template_fill",
    "target": { "scope": "document" },
    "params": { "action": "fill_single", "placeholder": "{{公司名称}}", "value": "北京科技有限公司" }
  }
]
[/TOOL_CALL]

**示例（无占位符：先识别，再按 path 精准填充）**：
[TOOL_CALL] word_edit_ops
dryRun: true
ops: [
  {
    "type": "template_fill",
    "target": { "scope": "document" },
    "params": { "action": "detect_fields" }
  }
]
[/TOOL_CALL]

[TOOL_CALL] word_edit_ops
dryRun: true
ops: [
  {
    "type": "template_fill",
    "target": { "scope": "document" },
    "params": {
      "action": "apply",
      "assignments": [
        {
          "fieldId": "p:12:colon",
          "path": "h1[1]/p[4]",
          "oldValue": "",
          "value": "张三",
          "fieldType": "person"
        }
      ],
      "output": "new_doc",
      "outputName": "已填充-报名表.docx"
    }
  }
]
[/TOOL_CALL]

### 20. citation_footnote - 引用/脚注管理
**action 类型**：
- insert_footnote：插入脚注（页面底部）
- insert_endnote：插入尾注（文档末尾）
- add_citation：添加参考文献引用
- generate_bibliography：整理参考文献列表

**示例（插入脚注）**：
[TOOL_CALL] word_edit_ops
dryRun: false
ops: [
  {
    "type": "citation_footnote",
    "target": { "scope": "document", "text": "关键数据" },
    "params": { "action": "insert_footnote", "anchorText": "关键数据", "content": "数据来源：2024年行业报告" }
  }
]
[/TOOL_CALL]

**示例（添加参考文献，GB/T 7714格式）**：
[TOOL_CALL] word_edit_ops
dryRun: false
ops: [
  {
    "type": "citation_footnote",
    "target": { "scope": "document" },
    "params": { 
      "action": "add_citation", 
      "format": "gb7714",
      "source": {
        "type": "book",
        "title": "人工智能导论",
        "author": "张三",
        "year": "2024"
      }
    }
  }
]
[/TOOL_CALL]

## 2. create - 从零创建新文档
**⚠️ 注意**：如果用户要求"按照当前文档格式"创建新文档，请使用 **create_from_template** 工具！

create 工具只适用于：
- 用户没有打开任何文档
- 用户明确要求从零开始创建
- 创建简单文档

**⚠️ 长文档 DSL 链式创建策略（极其重要！）**

当文档内容预计超过 2000 字时，**必须用 DSL 分段创建**，禁止一次性生成全部内容。
全程使用 DSL 格式，每一步都可以自主指定字体、字号、颜色、行距等排版参数。

**第 1 步：用 create + DSL 创建文档的第一部分**
- 用 create 工具，传入 dsl 参数，包含文档标题和第一部分内容（1~2 个章节，不超过 1500 字）
- DSL 中为每个标题、段落设置合适的字体/字号/格式

**第 2~N 步：用 insert + DSL 追加后续部分**
- 用 **insert** 工具，position="end"，传入 dsl 参数，追加下一部分的内容
- 每步只写 1~2 个章节，不超过 1500 字
- DSL 中同样指定完整的排版格式

**最后一步：整体审查**
- 全部写完后，系统会附带最新完整文档内容
- 检查有无遗漏/错误，用 replace 修正细节；无问题则回复总结

**示例流程**：
1. create dsl:{"blocks":[标题+一、项目概述（带字体字号格式）]}
2. insert position:end dsl:{"blocks":[二、业务需求（带格式）]}
3. insert position:end dsl:{"blocks":[三、技术方案（带格式）]}
4. insert position:end dsl:{"blocks":[四、实施计划（带格式）]}  ← 最后一部分
5. 审查完整文档 → 回复总结

**DSL 格式排版建议（中文商务/公文文档）**：
- 一级标题：黑体，16pt，居中，段前12pt段后6pt
- 二级标题：黑体，14pt，左对齐，加粗
- 三级标题：黑体，12pt，左对齐，加粗
- 正文段落：仿宋，12pt（小四），首行缩进2em，行距28pt
- 表格内容：宋体，10.5pt（五号）

**核心规则**：
- 每步都用 DSL，AI 自主控制每个段落的字体/字号/颜色/行距
- 每步不超过 1500 字，分多步完成长文档
- 每步完成后系统会返回最新文档，据此继续下一步
- 全部写完后必须审查一遍

**方式一：DSL 结构化格式（推荐，可控性最强）**
[TOOL_CALL] create
title: 文档标题
dsl: {"blocks":[{"type":"heading","level":1,"content":"标题"},{"type":"paragraph","content":[{"text":"正文内容","bold":true}]}]}
[/TOOL_CALL]

**DSL 块类型**：
- heading: {"type":"heading","level":1,"content":"标题","format":{"alignment":"center"}}
- paragraph: {"type":"paragraph","content":"正文","format":{"firstLineIndent":"2em"}}
- paragraph（带格式）: {"type":"paragraph","content":[{"text":"普通"},{"text":"粗体","bold":true}]}
- list: {"type":"list","listType":"bullet","items":[{"content":"项目1"}]}
- table: {"type":"table","rows":[{"cells":[{"content":"表头"}],"isHeader":true}]}
- table（合并单元格）: {"type":"table","rows":[{"cells":[{"content":"合并","colSpan":2}]}]}
- image: {"type":"image","src":"https://...","width":"300px","alignment":"center"}

**DSL 行内格式（Run）**: text, bold, italic, underline, strikethrough, fontFamily, fontSize, color, highlight
**DSL 段落格式（format）**: alignment, firstLineIndent, leftIndent, spaceBefore, spaceAfter, lineHeight

**方式二：参考另一个 DOCX 的排版**
[TOOL_CALL] create
title: 合并后的新文档
styleRefFileName: 格式参考.docx
contentRefFileName: 内容参考.docx
content: （可留空或简短说明，系统会优先用 contentRefFileName 解析正文结构）
[/TOOL_CALL]

**方式三：HTML 内容（简单场景）**
[TOOL_CALL] create
title: 文档标题
content: <h1 style="text-align: center">标题</h1><p>正文内容...</p>
[/TOOL_CALL]

**HTML 支持标签**：
- 标题: <h1>/<h2>/<h3> - 可加 style="text-align: center" 居中
- 段落: <p> - 默认首行缩进
- 粗体: <strong> 或 <b>
- 斜体: <em> 或 <i>
- 下划线: <u>
- 颜色: <span style="color: #ff0000">红色文字</span>
- 表格: <table><tr><td>单元格</td></tr></table>
- 列表: <ul><li>项目</li></ul> 或 <ol><li>项目</li></ol>

**方式四：elements 数组（兼容旧格式）**
[TOOL_CALL] create
title: 文档标题
elements: [{"type":"heading","content":"标题","level":1},{"type":"paragraph","content":"正文","bold":true}]
[/TOOL_CALL]

**elements 格式**（JSON数组）：
- **标题**: {"type":"heading","content":"标题文字","level":1}
- **段落**: {"type":"paragraph","content":"段落内容","bold":true,"fontSize":14}
- **表格**: {"type":"table","rows":3,"cols":2,"data":[["表头1","表头2"],["数据1","数据2"]]}

**格式选择建议**：
- 需要精确控制格式（字体/颜色/表格合并等）→ 使用 DSL
- 简单文本创建 → 使用 HTML 或 elements
- 需要保留原文档复杂格式 → 使用 create_from_template

**⚠️ 文档审查策略（用户请求审查/校对/检查/纠错时使用）**

当用户要求审查、校对、检查文档时，**必须使用 review 工具**（不要用 replace），按以下流程执行：

**第 1 步：通读全文**，从以下维度识别问题：
  1. 语言通顺性（语病、错别字、标点）→ type: grammar / typo
  2. 逻辑连贯性（段落衔接、前后一致）→ type: logic
  3. 表述专业性（用词准确、避免口语化）→ type: style
  4. 格式规范性（标题层级、字体字号一致）→ type: format
  5. 内容完整性（遗漏、重复）→ type: logic

**第 2 步：概述文档整体质量**
  - 用一段文字总结文档的优缺点
  - 列出即将修改的问题清单（编号+简述）

**第 3 步：逐条使用 review 工具标记修改**
  - 从文档开头到结尾依次处理，避免位置偏移
  - 每条**必须填写 reason 和 type**
  - 内容/措辞问题 → review（纯文字替换）
  - 格式问题 → review + 格式参数（fontFamily/fontSize/bold 等）
  - 问题较多时**先处理最重要的 10 处**，然后告知用户可继续审查

**第 4 步：给出审查总结**
  - 共发现 N 个问题，已提交 N 处修改建议
  - 按类别统计：语法 X 处、逻辑 X 处、措辞 X 处...
  - 提醒用户去修订面板逐条接受/拒绝

## 3. insert - 插入内容
在文档的指定位置插入新内容。支持**纯文本**和**DSL 结构化格式**两种方式。

**方式一：纯文本插入**
[TOOL_CALL] insert
position: start | end | after:锚点文字
content: 要插入的文本内容
[/TOOL_CALL]

**方式二：DSL 格式插入（推荐用于长文档续写，可精确控制排版）**
[TOOL_CALL] insert
position: end
dsl: {"blocks":[...]}
[/TOOL_CALL]

DSL 格式与 create 工具的 dsl 参数完全一致，支持 heading/paragraph/list/table 等所有块类型，以及 fontFamily/fontSize/color/bold/lineHeight 等全部排版属性。

**DSL 插入示例（追加一个章节到末尾）**：
[TOOL_CALL] insert
position: end
dsl: {"blocks":[{"type":"heading","level":2,"content":[{"text":"二、业务需求分析","fontFamily":"黑体","fontSize":14,"bold":true}]},{"type":"paragraph","content":[{"text":"    本项目旨在解决企业日常运营中的核心痛点...","fontFamily":"仿宋","fontSize":12}],"format":{"lineHeight":"28pt","firstLineIndent":"2em"}}]}
[/TOOL_CALL]

**position 参数说明**：
- \`start\`：在文档开头插入
- \`end\`：在文档末尾插入（长文档续写时必须用 end）
- \`after:某段文字\`：在指定文字后面插入

**⚠️ 长文档创建时，第 2 步起必须用 insert + dsl + position:end 来续写，不要用 replace。**

## 4. delete - 删除内容
删除文档中的指定内容。

调用格式：
[TOOL_CALL] delete
target: 要删除的文字（精确匹配）
[/TOOL_CALL]

## 5. create_from_template - 基于当前文档创建新文档
**当用户要求"按照这个格式"、"照着这个模板"、"用同样的格式"创建新文档时使用。**

这个工具会：
1. 复制当前打开的文档（100%保留所有格式）
2. 在新文档中自动替换你指定的内容

调用格式：
[TOOL_CALL] create_from_template
newTitle: 新文档的标题
replacements: [{"search":"原文字","replace":"新文字"}]
[/TOOL_CALL]

**⚠️ 关键注意事项**：
- **search 必须完全精确匹配**文档中的文字，一个字都不能差！
- 查看系统提供的"文档结构"信息，从中**直接复制**要替换的文字
- 不要猜测或编造 search 内容
- 如果不确定原文是什么，先用简短的、确定存在的文字

**示例**：用户打开了会议记录模板，说"帮我按这个格式写12月5日的会议记录"

假设系统提供的文档结构显示：
- 表格第1行: "会议时间" | "2024年11月11日21时10分至21时30分"
- 表格第2行: "会议地点" | "精工园3-102"

那么调用：
[TOOL_CALL] create_from_template
newTitle: 2024年12月5日会议记录
replacements: [{"search":"精工园3-102","replace":"行政楼201"}]
[/TOOL_CALL]

**如果替换失败**，说明 search 文字不匹配，请检查文档结构中的原文。

## 6. Excel 表格操作工具（当用户打开 .xlsx 文件时可用）

### 6.1 excel_read - 读取单元格
读取指定单元格或区域的内容。

[TOOL_CALL] excel_read
sheet: Sheet1
range: A1
[/TOOL_CALL]

或读取区域：
[TOOL_CALL] excel_read
sheet: Sheet1
range: A1:C10
[/TOOL_CALL]

### 6.2 excel_search - 搜索内容
在工作表中搜索包含指定文字的单元格。

[TOOL_CALL] excel_search
sheet: Sheet1
text: 要搜索的文字
[/TOOL_CALL]

### 6.3 excel_write - 写入/修改单元格
修改一个或多个单元格的值和样式。

**单个单元格：**
[TOOL_CALL] excel_write
sheet: Sheet1
updates: [{"address":"A1","value":"新内容"}]
[/TOOL_CALL]

**多个单元格：**
[TOOL_CALL] excel_write
sheet: Sheet1
updates: [{"address":"A1","value":"姓名"},{"address":"B1","value":"年龄"},{"address":"A2","value":"张三"},{"address":"B2","value":25}]
[/TOOL_CALL]

**带样式（完整格式）：**
[TOOL_CALL] excel_write
sheet: Sheet1
updates: [{"address":"A1","value":"标题","style":{"font":{"bold":true,"size":14,"color":{"argb":"FFFF0000"}},"alignment":{"horizontal":"center"},"fill":{"type":"pattern","pattern":"solid","fgColor":{"argb":"FFFFFF00"}}}}]
[/TOOL_CALL]

**样式说明：**
- font: {bold, italic, underline, size, name, color:{argb}}
- alignment: {horizontal: left/center/right, vertical: top/middle/bottom, wrapText}
- fill: {type:"pattern", pattern:"solid", fgColor:{argb:"FFRRGGBB"}}
- border: {top/bottom/left/right: {style:"thin", color:{argb}}}

**写入公式：**
[TOOL_CALL] excel_write
sheet: Sheet1
updates: [{"address":"C1","value":"=SUM(A1:B1)"}]
[/TOOL_CALL]

### 6.4 excel_insert_rows - 插入行
在指定位置插入新行。

[TOOL_CALL] excel_insert_rows
sheet: Sheet1
startRow: 5
count: 3
data: [["数据1","数据2"],["数据3","数据4"],["数据5","数据6"]]
[/TOOL_CALL]

### 6.5 excel_insert_columns - 插入列
在指定位置插入新列。

[TOOL_CALL] excel_insert_columns
sheet: Sheet1
startCol: 3
count: 2
[/TOOL_CALL]

### 6.6 excel_delete_rows - 删除行
删除指定行。

[TOOL_CALL] excel_delete_rows
sheet: Sheet1
startRow: 5
count: 2
[/TOOL_CALL]

### 6.7 excel_delete_columns - 删除列
删除指定列。

[TOOL_CALL] excel_delete_columns
sheet: Sheet1
startCol: 3
count: 1
[/TOOL_CALL]

### 6.8 excel_add_sheet - 新建工作表
创建新的工作表。

[TOOL_CALL] excel_add_sheet
name: 新工作表
[/TOOL_CALL]

### 6.9 excel_delete_sheet - 删除工作表
删除指定的工作表。

[TOOL_CALL] excel_delete_sheet
name: Sheet2
[/TOOL_CALL]

### 6.10 excel_merge - 合并单元格
合并指定区域的单元格。

[TOOL_CALL] excel_merge
sheet: Sheet1
range: A1:C1
[/TOOL_CALL]

### 6.11 excel_unmerge - 取消合并
取消合并指定区域的单元格。

[TOOL_CALL] excel_unmerge
sheet: Sheet1
range: A1:C1
[/TOOL_CALL]

### 6.12 excel_create - 创建新的 Excel 文件 ⭐重要
创建一个全新的 Excel 文件，自动带有专业格式（表头样式、边框、自动列宽等）。

**⚠️ 重要：多工作表必须在一次调用中创建！**
- 如果需要多个工作表（如"员工信息"和"统计分析"），必须用 sheets 参数一次性创建
- **错误做法**：分多次调用创建多个文件 ❌
- **正确做法**：一次调用，sheets 数组包含所有工作表 ✅

**参数说明：**
- filename: 文件名（如 "调研报告.xlsx"）
- data: 二维数组（简单用法，只创建一个工作表）
- sheets: JSON 格式的工作表配置数组（**多工作表必须用这个**）

**sheets 参数格式（JSON 数组）：**
\`[{"name":"工作表名","data":[[数据行1],[数据行2]...]},{"name":"工作表2","data":[[...]]}]\`

**数据格式支持：**
1. **简单值**：直接写值，如 "张三", 100, "25%"
2. **公式**：以=开头，如 "=SUM(A1:A10)", "=VLOOKUP(A2,Sheet1!A:B,2,FALSE)"
3. **跨工作表公式**：用 '工作表名'! 格式，如 "=SUM('员工信息'!E:E)"
4. **带样式的值**：{"v": "内容", "s": "样式字符串"}

**样式字符串格式**（逗号分隔）：
- 字体：bold, italic, underline
- 对齐：center, left, right
- 字号：数字如 14, 16, 18
- 字体颜色：#FF0000（红色）, #00FF00（绿色）
- 背景色：bg#FFFF00（黄色背景）

**示例1：简单单工作表**
[TOOL_CALL] excel_create
filename: 员工名单.xlsx
data: [["姓名","年龄","部门","薪资"],["张三",28,"技术部",15000],["李四",32,"市场部",12000]]
[/TOOL_CALL]

**示例2：⭐多工作表（一次创建，跨表公式）**
这是创建包含多个关联工作表的正确方式！
[TOOL_CALL] excel_create
filename: 员工管理.xlsx
sheets: [{"name":"员工信息","data":[["姓名","部门","薪资"],["张三","技术部",15000],["李四","销售部",12000],["王芳","财务部",10000]]},{"name":"统计分析","data":[["统计项","数值"],["总人数","=COUNTA('员工信息'!A2:A100)"],["总薪资","=SUM('员工信息'!C:C)"],["平均薪资","=AVERAGE('员工信息'!C:C)"],["技术部人数","=COUNTIF('员工信息'!B:B,\\"技术部\\")"]]}]
[/TOOL_CALL]

**示例3：带公式计算的表格**
[TOOL_CALL] excel_create
filename: 销售报表.xlsx
sheets: [{"name":"销售数据","data":[["产品","数量","单价","金额"],["iPhone",100,5000,"=B2*C2"],["iPad",50,3000,"=B3*C3"],["总计","","","=SUM(D2:D3)"]]}]
[/TOOL_CALL]

**工作表配置项：**
- name: 工作表名称（必填）
- data: 二维数组数据
- columnWidths: 列宽数组 [15, 10, 20]
- rowHeight: 数据行高（默认20）
- headerHeight: 表头行高（默认25）
- firstRowIsHeader: 第一行是否为表头（默认true）
- freezeHeader: 是否冻结表头（默认true）

### 6.13 excel_formula - 设置公式 ⭐常用
批量设置单元格公式，支持所有 Excel 公式。

**支持的常用公式：**
- SUM(A1:A10) - 求和
- AVERAGE(A1:A10) - 平均值
- COUNT(A1:A10) - 计数
- MAX(A1:A10) - 最大值
- MIN(A1:A10) - 最小值
- IF(条件, 真值, 假值) - 条件判断
- VLOOKUP(查找值, 范围, 列号, 模式) - 垂直查找
- SUMIF(范围, 条件, 求和范围) - 条件求和
- COUNTIF(范围, 条件) - 条件计数
- CONCATENATE(A1, B1) 或 A1&B1 - 文本连接
- ROUND(数值, 小数位数) - 四舍五入
- TODAY() / NOW() - 日期时间

**单个公式：**
[TOOL_CALL] excel_formula
sheet: Sheet1
address: B10
formula: =SUM(B2:B9)
[/TOOL_CALL]

**批量公式（JSON 格式）：**
[TOOL_CALL] excel_formula
sheet: Sheet1
formulas: [{"address":"B10","formula":"=SUM(B2:B9)"},{"address":"C10","formula":"=AVERAGE(C2:C9)"}]
[/TOOL_CALL]

### 6.14 excel_sort - 排序数据
按指定列对数据进行排序。

[TOOL_CALL] excel_sort
sheet: Sheet1
range: A1:D10
column: B
ascending: true
hasHeader: true
[/TOOL_CALL]

参数说明：
- range: 要排序的范围（如 A1:D10）
- column: 排序依据的列（如 B）
- ascending: true=升序, false=降序
- hasHeader: true=第一行是表头不参与排序

### 6.15 excel_autofill - 自动填充/序列填充
从源范围自动填充到目标范围。

**复制填充：**
[TOOL_CALL] excel_autofill
sheet: Sheet1
sourceRange: A1
targetRange: A2:A10
fillType: copy
[/TOOL_CALL]

**序列填充（数字递增）：**
[TOOL_CALL] excel_autofill
sheet: Sheet1
sourceRange: A1
targetRange: A2:A10
fillType: series
[/TOOL_CALL]

**公式填充：**
[TOOL_CALL] excel_autofill
sheet: Sheet1
sourceRange: C2
targetRange: C3:C10
fillType: formula
[/TOOL_CALL]

### 6.16 excel_dimensions - 设置列宽行高
调整列宽和行高。

[TOOL_CALL] excel_dimensions
sheet: Sheet1
columns: [{"column":"A","width":20},{"column":"B","width":15},{"column":"C","width":30}]
rows: [{"row":1,"height":25},{"row":2,"height":20}]
[/TOOL_CALL]

### 6.17 excel_conditional_format - 条件格式
根据条件设置单元格格式（如高亮显示）。

**数值大于条件：**
[TOOL_CALL] excel_conditional_format
sheet: Sheet1
range: B2:B10
type: cellIs
operator: greaterThan
value: 100
fill: FF00FF00
[/TOOL_CALL]

**色阶（从红到绿）：**
[TOOL_CALL] excel_conditional_format
sheet: Sheet1
range: C2:C10
rules: [{"type":"colorScale","minColor":"FFF8696B","maxColor":"FF63BE7B"}]
[/TOOL_CALL]

**数据条：**
[TOOL_CALL] excel_conditional_format
sheet: Sheet1
range: D2:D10
rules: [{"type":"dataBar","color":"FF638EC6"}]
[/TOOL_CALL]

### 6.18 excel_calculate - 获取计算结果
获取单元格的值或公式计算结果。

[TOOL_CALL] excel_calculate
sheet: Sheet1
addresses: ["B10","C10","D10"]
[/TOOL_CALL]

### 6.19 excel_filter - 自动筛选 ⭐新增
设置或清除工作表的自动筛选（AutoFilter）。

**设置筛选：**
[TOOL_CALL] excel_filter
sheet: Sheet1
range: A1:D100
action: set
[/TOOL_CALL]

**清除筛选：**
[TOOL_CALL] excel_filter
sheet: Sheet1
action: remove
[/TOOL_CALL]

### 6.20 excel_validation - 数据验证 ⭐新增
设置单元格的数据验证规则（下拉列表、数值限制等）。

**下拉列表：**
[TOOL_CALL] excel_validation
sheet: Sheet1
range: B2:B100
type: list
values: ["是", "否", "待定"]
[/TOOL_CALL]

**数值范围限制：**
[TOOL_CALL] excel_validation
sheet: Sheet1
range: C2:C100
type: whole
min: 1
max: 100
[/TOOL_CALL]

**参数说明：**
- type: list（下拉列表）、whole（整数）、decimal（小数）、textLength（文本长度）
- values: 下拉选项数组（仅 list 类型）
- min/max: 数值范围（仅数值类型）

### 6.21 excel_hyperlink - 超链接 ⭐新增
在单元格中插入超链接。

[TOOL_CALL] excel_hyperlink
sheet: Sheet1
cell: A1
url: https://www.baidu.com
text: 点击访问百度
[/TOOL_CALL]

### 6.22 excel_find_replace - 查找替换 ⭐新增
批量查找并替换工作表中的内容。

[TOOL_CALL] excel_find_replace
sheet: Sheet1
find: 北京
replace: 上海
matchCase: false
[/TOOL_CALL]

**参数说明：**
- find: 要查找的文本
- replace: 替换为的文本
- matchCase: 是否区分大小写（true/false）
- matchWholeCell: 是否匹配整个单元格（true/false）
- allSheets: 是否搜索所有工作表（true/false）

### 6.23 excel_chart - 图表/数据可视化 ⭐重要
**当用户说"做图表"、"可视化"、"饼图"、"柱状图"、"折线图"等需求时，必须使用此工具！**
不要用 excel_create 创建新表格，而是用 excel_chart 在现有数据旁边插入图表图片。

[TOOL_CALL] excel_chart
sheet: 饼图数据
type: pie
dataRange: A1:B6
title: 基层就业项目分布
position: D1
width: 450
height: 350
[/TOOL_CALL]

**参数说明：**
- sheet: **必须使用当前打开的工作表名称**（从上下文中获取）
- type: 图表类型
  - pie（饼图）⭐用于占比分析
  - column（柱状图）⭐用于对比分析
  - bar（横向条形图）
  - line（折线图）⭐用于趋势分析
  - doughnut（环形图）
  - area（面积图）
- dataRange: 数据所在范围，格式如 A1:B6
  - **第一行**是标题行（如"项目名称"、"数量"）
  - **第一列**是分类标签
  - **其他列**是数值数据
- title: 图表标题
- position: 图表插入位置（建议放在数据右侧，如 D1、E1）
- width/height: 图表尺寸（像素），默认 500x300

**典型用例：**
用户数据：A1:B6（A列是名称，B列是数值）
→ 使用 excel_chart, dataRange: A1:B6, type: pie

**⚠️ 注意：这个工具会在 Excel 中插入真实的图表图片！**

**⚠️ Excel 操作注意事项：**
- 只有打开 .xlsx 文件时这些工具才可用
- sheet 参数必须是实际存在的工作表名称
- 行号从 1 开始，列号也从 1 开始（或用字母 A, B, C...）
- 修改会自动保存到文件

## 7. ppt_create - 生成 PPTX（海报式 image-only，每页一张成片）⭐重要
用于生成并导出 ".pptx" 演示文稿。**每一页都是一张完整海报图**，图里必须包含中文文案与排版（不是只做背景）。

### 两阶段工作流（必须遵守）
1) **先做大纲（不调用工具）**：当用户提出“做 PPT/生成 PPT”时，先输出一个结构化大纲（建议 JSON），让用户确认。
2) **确认后再生成**：只有当用户明确回复“开始生成/确认生成”时，才调用 ppt_create。

### 阶段1（大纲）输出格式要求（强制）
- **只输出一个 JSON 大纲**（可在前后加 1~2 句解释，但必须包含一个完整 JSON 对象，且可直接复制解析）
- **页数规则（重要）**：
  - 用户指定页数 N → 必须输出 **正好 N 页**
  - 用户未指定页数 → **默认推荐 10~15 页**（内容充实、结构完整）；如果主题特别复杂/涉及多个章节，可以推荐 15~20 页
  - 除非用户明确要求"精简/简短/3页就够"，否则不要少于 10 页
- **字段必须齐全且稳定**：请使用如下结构（字段名不要随意改动）

\`\`\`json
{
  "title": "PPT 标题（中文）",
  "theme": "主题/用途（中文）",
  "slideCount": 12,
  "styleHint": "给 Gemini 的风格倾向（可空；例如：'玻璃质感高级商务 / 手绘插画高级 / 极简瑞士排版'）",
  "slides": [
    {
      "pageNumber": 1,
      "pageType": "cover|agenda|section|content|timeline|chart|ending",
      "headline": "主标题（中文）",
      "subheadline": "副标题（可空）",
      "bullets": ["要点1","要点2","要点3"],
      "footerNote": "页脚短句（可空）",
      "layoutIntent": "版式意图（如：左文右图/上标题下三栏/大标题居左+右侧主视觉等）",
      "visualElements": "可选：主视觉意象/图标/装饰元素建议（给 Gemini 用）"
    }
  ]
}
\`\`\`

### 调用参数（必须）
[TOOL_CALL] ppt_create
title: PPT 文件名（不含扩展名）
theme: 主题/用途
style: 风格倾向（可空；如“Fluent+柔光+抽象3D，商务高级”）
outline: 阶段1输出的大纲原文（建议 JSON 原样粘贴）
[/TOOL_CALL]

### 质量要求（功能优先）
- outline 必须包含每页的**完整中文文案**（headline/subheadline/bullets/footerNote），避免临场瞎编
- **页数强约束**：\`slideCount\` 与 \`slides.length\` 必须一致；\`pageNumber\` 必须从 1 连续递增到 \`slideCount\`，不允许缺页/多页
- **信息更强**：每页 bullets 建议 3~6 条，表达具体、可落地；避免"待补充/XXX/自行发挥"等占位词
- 禁止：水印/徽章/二维码/乱码/错别字
- 排版描述必须明确：层级、对齐方式、留白、网格、阅读动线

## 8. ppt_edit - 编辑已生成的 PPT 页面（拖拽/框选触发）

当用户**拖拽 PPT 页面到对话框**或**Ctrl+框选区域**后发送修改要求时，使用此工具。

### 触发条件
上下文中包含 "=== PPT 编辑请求 ===" 标记时，说明用户正在请求编辑 PPT 页面。

### 强制约束（非常重要）
当出现 "=== PPT 编辑请求 ===" 时：
- **只能**调用 \`ppt_edit\`
- **禁止**调用任何 Word/Excel 工具：\`replace\` / \`insert\` / \`delete\` / \`create\` / \`create_from_template\` / \`excel_*\`
如果不调用 \`ppt_edit\`，会导致修改对象错误（用户编辑的是 PPT，不是 Word 文档）。

### 判断编辑模式（重要！）
根据用户的措辞判断使用哪种模式：

**mode="regenerate"（整页重做）**：
- 用户对整页不满意：太丑、不好看、换个风格、重新生成、重做、再来一个
- 用户想要完全不同的设计

**mode="partial_edit"（局部调整）**：
- 用户只想改局部：把XX改成YY、调整颜色、换个背景、修改文字、移动位置
- 用户提到具体细节的修改

### 调用参数
[TOOL_CALL] ppt_edit
pageNumber: 页码（从1开始）
mode: regenerate 或 partial_edit
feedback: 用户的修改要求（原文）
pptxPath: PPTX 文件路径（从上下文获取）
[/TOOL_CALL]

### 示例
用户拖拽第3页并说"这页太丑了，换个古风风格"
→ mode="regenerate", feedback="这页太丑了，换个古风风格"

用户框选某区域并说"把这里的颜色改成蓝色"
→ mode="partial_edit", feedback="把这里的颜色改成蓝色"

</available_tools>

<workflow>
1. **分析阶段**：理解用户需求，确定需要使用哪个工具
2. **执行阶段**：调用相应工具执行操作
3. **验证阶段**：根据工具返回结果确认是否成功
4. **迭代阶段**：如果需要多次操作，继续调用工具直到完成
5. **总结阶段**：用简短的话告诉用户完成了什么

**重要**：如果你说要做某事，必须在同一回合内实际执行（调用工具）。
</workflow>

<tool_usage_examples>

### 示例1：简单替换
用户：把"小明"改成"小红"

[TOOL_CALL] replace
search: 小明
replace: 小红
[/TOOL_CALL]

⚠️ 注意：search 和 replace 的值直接写文字，**不要加引号**！
- 错误：search: "小明"  ← 会搜索包含引号的字符串
- 正确：search: 小明    ← 直接搜索"小明"两个字

### 示例2：多处不同修改
用户：把标题改成"工作报告"，把日期改成"2024年1月"

[TOOL_CALL] replace
search: 原标题内容
replace: 工作报告
[/TOOL_CALL]

[TOOL_CALL] replace
search: 原日期内容
replace: 2024年1月
[/TOOL_CALL]

### 示例3：⭐ 基于模板创建新文档（最常见场景！）
用户打开了"2024年11月会议记录.docx"，说：帮我写一份12月的会议记录，时间是12月5日下午2点，地点行政楼201

**使用 create_from_template 保留表格和格式！**

[TOOL_CALL] create_from_template
newTitle: 2024年12月5日会议记录
replacements: [{"search":"2024年11月11日","replace":"2024年12月5日"},{"search":"21时10分至21时30分","replace":"14时00分至15时00分"},{"search":"精工园3-102","replace":"行政楼201"}]
[/TOOL_CALL]

### 示例4：只修改当前文档（不创建新文档）
用户：把日期改成12月

[TOOL_CALL] replace
search: 11月
replace: 12月
[/TOOL_CALL]

### 示例5：从零创建（没有打开任何文档时）
用户：帮我写一份简单的通知

[TOOL_CALL] create
title: 通知
content: <h1>通知</h1><p>内容...</p>
[/TOOL_CALL]

</tool_usage_examples>

<constraints>
- **不要**在没有使用工具的情况下声称已修改文档
- **不要**输出完整的文档内容来"展示"修改，使用 replace 工具进行精准修改
- **不要**猜测文档内容，根据系统提供的 [当前文档内容] 进行操作
- **不要**输出冗长的解释，保持简洁
- **不要**在工具调用前后添加不必要的确认语句
- 如果 search 内容在文档中找不到，系统会返回失败，此时应该检查是否有拼写差异并重试
</constraints>

<response_style>
完成操作后的回复示例：
- ✅ 已将"小明"替换为"小红"，共 3 处
- ✅ 已创建文档 \`会议纪要.docx\`
- ⚠️ 未找到"xxx"，请确认文档中是否存在该内容

保持回复简短、信息密度高。用户可以在编辑器中看到实际的修改效果。
</response_style>`

  // 轻量编辑器提示词：禁止工具调用，仅返回内容
  const editorSystemPrompt = `你是一个写作与改写助手。
规则：
- 不要输出任何 [TOOL_CALL]/[/TOOL_CALL]、[TOOL_RESULT]/[/TOOL_RESULT] 等标记
- 不要提出要调用工具或“已修改文档”的说法
- 用户要求“返回修改后的完整文档内容”时：直接返回最终内容（Markdown）
- 其它情况：给出简洁、可直接复制使用的答案`

  const streamFallbackContent = async (content: string, signal: AbortSignal) => {
    if (!content) return
    const total = content.length
    const chunkSize = total > 4000 ? 120 : total > 1500 ? 60 : total > 400 ? 30 : 12
    const delay = 24
    for (let i = 0; i < total; i += chunkSize) {
      if (signal.aborted) break
      setStreamingContent(content.slice(0, i + chunkSize))
      await new Promise(resolve => setTimeout(resolve, delay))
    }
  }

  // 单次 API 调用
  const callAPI = async (
    allMessages: Array<{ role: string; content: string }>,
    signal: AbortSignal,
    options?: { returnRaw?: boolean }
  ): Promise<string> => {
    const headers: Record<string, string> = {
      'Content-Type': 'application/json',
    }
    if (settings.apiKey) {
      headers['Authorization'] = `Bearer ${settings.apiKey}`
    }

    // Electron 环境：优先走主进程代理（Node fetch/HTTP1.1），避免渲染进程 HTTP/2 的 ERR_HTTP2_PING_FAILED
    if (window.electronAPI?.isElectron && window.electronAPI.aiChatCompletions && window.electronAPI.onAIStreamDelta) {
      const requestId = `${Date.now()}-${Math.random().toString(16).slice(2)}`
      let fullContent = ''
      let fullReasoning = ''  // 思考内容（kimi-k2.5 等思考模型）
      let contentDeltaCount = 0
      
      // 清空之前的思考内容
      setStreamingReasoning('')

      const unsubscribe = window.electronAPI.onAIStreamDelta((payload) => {
        if (!payload || payload.requestId !== requestId) return
        
        // 处理内容增量
        const delta = payload.delta || ''
        if (delta) {
          contentDeltaCount += 1
          fullContent += delta
          // 从内容中提取 <think> 标签内的思考内容
          const { thinking, cleaned } = extractThinking(fullContent)
          if (thinking && thinking !== fullReasoning) {
            fullReasoning = thinking
            setStreamingReasoning(fullReasoning)
          }
          const parsed = parseAssistantOutput(cleaned)
          if (parsed.summary) setStreamingSummary(parsed.summary)
          if (parsed.phase === 'editing') setEditPhase('editing')
          if (parsed.phase === 'done') setEditPhase('done')
          setStreamingContent(parsed.displayText)
        }
        
        // 处理思考增量（kimi-k2.5 等思考模型 - 通过单独字段返回）
        const reasoningDelta = (payload as { reasoningDelta?: string }).reasoningDelta || ''
        if (reasoningDelta) {
          fullReasoning += reasoningDelta
          setStreamingReasoning(fullReasoning)
        }

      })

      const abortHandler = () => {
        try {
          window.electronAPI?.aiCancel?.(requestId)
        } catch {
          // ignore
        }
      }

      signal.addEventListener('abort', abortHandler, { once: true })

      try {
        // kimi-k2.5 模型只支持 temperature=1，自动修正
        const effectiveTemperature = settings.model.toLowerCase().includes('kimi-k2') ? 1 : settings.temperature
        
        const result = await window.electronAPI.aiChatCompletions({
          requestId,
          baseUrl: settings.baseUrl,
          apiKey: settings.apiKey,
          model: settings.model,
          messages: allMessages,
          temperature: effectiveTemperature,
          maxTokens: settings.maxTokens,
        })

        if (!result?.success) {
          throw new Error(result?.error || '请求失败')
        }

        const content = result.content || fullContent
        const cleaned = cleanModelOutput(content)
        const parsed = parseAssistantOutput(cleaned)
        const displayText = parsed.displayText || parsed.summary || cleaned
        if (parsed.summary) setStreamingSummary(parsed.summary)
        if (parsed.phase === 'editing') setEditPhase('editing')
        if (parsed.phase === 'done') setEditPhase('done')
        if (contentDeltaCount === 0 && displayText) {
          await streamFallbackContent(displayText, signal)
        }
        return options?.returnRaw ? cleaned : displayText
      } finally {
        try {
          unsubscribe?.()
        } catch {
          // ignore
        }
      }
    }

    // kimi-k2.5 模型只支持 temperature=1，自动修正
    const effectiveTemperature = settings.model.toLowerCase().includes('kimi-k2') ? 1 : settings.temperature

    const response = await fetch(`${settings.baseUrl}/chat/completions`, {
      method: 'POST',
      headers,
      signal,
      body: JSON.stringify({
        model: settings.model,
        messages: allMessages,
        temperature: effectiveTemperature,
        max_tokens: settings.maxTokens,
        stream: true,
      }),
    })

    if (!response.ok) {
      const errorText = await response.text()
      throw new Error(errorText || '请求失败')
    }

    const reader = response.body?.getReader()
    if (!reader) throw new Error('无法读取响应')

    const decoder = new TextDecoder()
    let fullContent = ''
    let fullReasoning = ''  // 存储完整的思考过程
    let buffer = ''
    
    // 清空之前的思考内容
    setStreamingReasoning('')
    
    // 读取超时包装函数
    const readWithTimeout = async (timeoutMs: number) => {
      const timeoutPromise = new Promise<{ done: true; value: undefined }>((_, reject) => {
        setTimeout(() => reject(new Error('读取超时')), timeoutMs)
      })
      return Promise.race([reader.read(), timeoutPromise])
    }

    const READ_TIMEOUT = 60000 // 60秒读取超时

    while (true) {
      let result
      try {
        result = await readWithTimeout(READ_TIMEOUT)
      } catch (e) {
        console.warn('[API] 流响应读取超时，返回已有内容')
        break
      }
      
      const { done, value } = result
      if (done) break

      buffer += decoder.decode(value, { stream: true })
      const lines = buffer.split('\n')
      buffer = lines.pop() || ''

      for (const line of lines) {
        if (line.startsWith('data: ')) {
          const data = line.slice(6).trim()
          if (data === '[DONE]') continue

          try {
            const json = JSON.parse(data)
            const choice = json.choices?.[0] || {}
            const delta = choice?.delta || {}
            const message = choice?.message || {}
            
            // 调试：首次收到 delta 时打印结构，便于确认 kimi-k2.5 等模型的字段名
            if (delta && !fullContent && !fullReasoning) {
              console.log('[AI stream] delta keys:', Object.keys(delta || {}), 'sample:', JSON.stringify(delta).slice(0, 300))
            }
            
            // 处理思考内容（kimi-k2.5 等思考模型）- 尝试多种字段名
            const reasoningDelta = delta?.reasoning_content || delta?.thinking_content || delta?.reasoning || delta?.thinking || delta?.thought || message?.reasoning_content || json?.reasoning_content || ''
            if (reasoningDelta) {
              fullReasoning += reasoningDelta
              setStreamingReasoning(fullReasoning)
            }
            
            // 处理正常内容
            let contentDelta = delta?.content || delta?.text || ''
            if (!contentDelta) {
              const messageContent = extractStreamText(message?.content ?? message?.text)
              if (messageContent) {
                contentDelta = messageContent
              } else if (choice?.text) {
                contentDelta = choice.text
              }
            }
            if (contentDelta && fullContent && contentDelta.startsWith(fullContent)) {
              contentDelta = contentDelta.slice(fullContent.length)
            }
            if (contentDelta) {
              fullContent += contentDelta
              // 从内容中提取 <think> 标签内的思考内容
              const { thinking, cleaned } = extractThinking(fullContent)
              if (thinking && thinking !== fullReasoning) {
                fullReasoning = thinking
                setStreamingReasoning(fullReasoning)
              }
              const parsed = parseAssistantOutput(cleaned)
              if (parsed.summary) setStreamingSummary(parsed.summary)
              if (parsed.phase === 'editing') setEditPhase('editing')
              if (parsed.phase === 'done') setEditPhase('done')
              setStreamingContent(parsed.displayText)
            }
          } catch {
            // 忽略解析错误
          }
        }
      }
    }

    const cleaned = cleanModelOutput(fullContent)
    const parsed = parseAssistantOutput(cleaned)
    if (parsed.summary) setStreamingSummary(parsed.summary)
    if (parsed.phase === 'editing') setEditPhase('editing')
    if (parsed.phase === 'done') setEditPhase('done')
    const displayText = parsed.displayText || parsed.summary || cleaned
    return options?.returnRaw ? cleaned : displayText
  }

  // 传统单轮消息（不走 Agent 工具循环）
  const sendMessage = useCallback(async (
    content: string,
    documentContext?: string
  ): Promise<string> => {
    setIsLoading(true)
    setStreamingContent('')
    setStreamingReasoning('')  // 清空思考内容
    setEditPhase('idle')
    setStreamingSummary('')

    if (abortControllerRef.current) {
      abortControllerRef.current.abort()
    }
    abortControllerRef.current = new AbortController()

    try {
      let userContent = content
      if (documentContext) {
        userContent += `\n\n[当前文档内容]\n${documentContext}`
      }
      const resp = await callAPI(
        [
          { role: 'system', content: editorSystemPrompt },
          { role: 'user', content: userContent },
        ],
        abortControllerRef.current.signal
      )
      return resp
    } finally {
      setIsLoading(false)
    }
  }, [callAPI, editorSystemPrompt])

  // Agent 消息发送 - 支持工具调用循环
  const sendAgentMessage = useCallback(async (
    content: string,
    documentContext?: string,
    filesContext?: string,
    callbacks?: AgentCallbacks,
    memoryContext?: { workspaceKey?: string; workspaceSummary?: string }
  ): Promise<void> => {
    setIsLoading(true)
    setStreamingContent('')
    setStreamingReasoning('')  // 清空思考内容
    setEditPhase('idle')
    setStreamingSummary('')

    if (abortControllerRef.current) {
      abortControllerRef.current.abort()
    }
    abortControllerRef.current = new AbortController()

    const allToolResults: ToolResult[] = []
    const conversationMessages: Array<{ role: string; content: string }> = []

    try {
      // 构建初始用户消息
      let userContent = content
      if (documentContext) {
        userContent += `\n\n[当前文档内容]\n${documentContext}`
      }
      if (filesContext) {
        userContent += `\n\n[附加文件内容]\n${filesContext}`
      }

      const memoryEnabled = settings.memoryEnabled !== false
      const memoryTopK = Math.max(1, settings.memoryTopK || 5)
      const memoryMaxChars = Math.max(500, settings.memoryMaxChars || 2000)
      const memoryTextWeight = typeof settings.memoryTextWeight === 'number' ? settings.memoryTextWeight : 0.6
      const memoryVectorWeight = typeof settings.memoryVectorWeight === 'number' ? settings.memoryVectorWeight : 0.4
      const workspaceKey = memoryContext?.workspaceKey || ''

      if (memoryEnabled && window.electronAPI?.memorySearch && content.trim().length >= 6) {
        try {
              const memoryResp = await memorySearch({
            query: content,
            topK: memoryTopK,
            textWeight: memoryTextWeight,
            vectorWeight: memoryVectorWeight,
            workspaceKey,
          })
          if (memoryResp.success && memoryResp.results.length > 0) {
            const memoryBlock = formatMemoryResults(
              memoryResp.results,
              memoryMaxChars
            )
            if (memoryBlock) {
              userContent += `\n\n[记忆检索]\n${memoryBlock}`
            }
          }
        } catch {
          // ignore memory failures
        }
      }

      // 获取历史消息 - 保留完整上下文，让 AI 能处理长对话和复杂任务
      // 保留 200 条消息，充分利用模型的上下文窗口
      const recentMessages = messages
        .filter(m => m.id !== 'welcome')
        .slice(-200)
        .map(m => ({
          role: m.role as string,
          content: cleanMessageForSend(m.content),
        }))
        .filter(m => m.content.length > 0)

      if (memoryEnabled && window.electronAPI?.memoryAppend) {
        const flushThreshold = Math.max(2000, settings.memoryFlushThresholdChars || 12000)
        const now = Date.now()
        if (now - lastMemoryFlushAtRef.current > 60_000) {
          const approxChars = recentMessages.reduce((sum, m) => sum + m.content.length, 0) + userContent.length
          if (approxChars > flushThreshold) {
            const summary = buildMemoryFlushText(recentMessages, content)
            if (summary) {
              try {
                await memoryAppend({
                  text: `【自动记忆刷新】\n${summary}`,
                  source: 'auto_flush',
                  tags: ['L2'],
                })
                await memoryAppendSession({
                  sessionId: sessionIdRef.current,
                  text: summary,
                  meta: { workspaceKey, type: 'flush' },
                })
                lastMemoryFlushAtRef.current = now
              } catch {
                // ignore memory failures
              }
            }
          }
        }
      }

      // 初始化对话
      conversationMessages.push(
        { role: 'system', content: agentSystemPrompt },
        ...recentMessages,
        { role: 'user', content: userContent }
      )

      let maxIterations = 20 // 防止无限循环，增加到20次以支持复杂任务
      let iteration = 0
      let accumulatedContent = '' // 累积所有响应中的文本内容
      let latestSummary = '' // 解析到的最新总结
      let lastResponse = ''
      let finalContentForMemory = ''

      // 【防重复修改】追踪已修改的内容
      const modifiedSearchTexts = new Set<string>() // 已被替换的原文
      const modifiedReplaceTexts = new Set<string>() // 替换后的新文本
      let totalReplaceCount = 0 // 总 replace 次数
      let consecutiveReplaceCount = 0 // 连续 replace 次数
      const MAX_CONSECUTIVE_REPLACE = 10 // 连续 replace 上限
      let shouldForceStop = false // 是否强制停止

      while (iteration < maxIterations && !shouldForceStop) {
        iteration++
        
        // 安全过滤：确保没有空 content 的消息（API 会拒绝空消息）
        for (let i = conversationMessages.length - 1; i >= 0; i--) {
          if (!conversationMessages[i].content || !conversationMessages[i].content.trim()) {
            conversationMessages[i].content = '[空]'
          }
        }

        // 调用 API
        const response = await callAPI(
          conversationMessages,
          abortControllerRef.current.signal,
          { returnRaw: true }
        )
        lastResponse = response

        // 检查是否有工具调用
        if (hasToolCall(response)) {
          const toolCalls = parseToolCalls(response)
          
          // 提取工具调用之外的文本内容并累积
          const parsedOutput = parseAssistantOutput(cleanModelOutput(response))
          console.log('[Agent] 提取的文本内容:', parsedOutput.displayText?.substring(0, 200))
          if (parsedOutput.summary) {
            latestSummary = parsedOutput.summary
            setStreamingSummary(parsedOutput.summary)
          }
          if (parsedOutput.displayText) {
            accumulatedContent = parsedOutput.displayText // 用最新的内容替换，因为 AI 会在最后给出完整总结
            console.log('[Agent] 累积内容已更新:', accumulatedContent.substring(0, 200))
          }
          
          // 将 AI 响应添加到对话
          conversationMessages.push({ role: 'assistant', content: response })

          // 执行每个工具调用
          const results: string[] = []
          let allSuccessful = true
          let hasReplaceInThisBatch = false
          let skippedCount = 0
          
          for (const call of toolCalls) {
            // 【防重复修改】检测 replace 工具的重复调用
            if (call.tool === 'replace') {
              hasReplaceInThisBatch = true
              const searchText = call.args.search || ''
              const replaceText = call.args.replace || ''
              
              // 检查是否正在修改之前已经修改过的内容
              if (modifiedReplaceTexts.has(searchText)) {
                console.warn(`[Agent] 跳过重复修改: 该内容是之前修改的结果`)
                results.push(`[TOOL_RESULT]\n工具: replace\n状态: 跳过 - 该内容已被修改过，无需再次修改\n[/TOOL_RESULT]`)
                skippedCount++
                continue
              }
              
              // 检查是否修改相同的原文
              if (modifiedSearchTexts.has(searchText)) {
                console.warn(`[Agent] 跳过重复修改: 相同原文已被修改`)
                results.push(`[TOOL_RESULT]\n工具: replace\n状态: 跳过 - 相同内容已被修改过\n[/TOOL_RESULT]`)
                skippedCount++
                continue
              }
            }
            
            if (callbacks?.onToolCall) {
              const result = await callbacks.onToolCall(call.tool, call.args)
              allToolResults.push(result)
              if (!result.success) allSuccessful = false
              
              // 【追踪修改】记录成功的 replace 操作
              if (call.tool === 'replace' && result.success) {
                const searchText = call.args.search || ''
                const replaceText = call.args.replace || ''
                modifiedSearchTexts.add(searchText)
                modifiedReplaceTexts.add(replaceText)
                totalReplaceCount++
                console.log(`[Agent] 记录修改 #${totalReplaceCount}: "${searchText.substring(0, 30)}..." → "${replaceText.substring(0, 30)}..."`)
              }
              
              // 更明确的结果反馈，包含进度信息
              const statusText = result.success 
                ? '成功 ✓'
                : `失败: ${result.message}`
              
              const progressInfo = call.tool === 'replace' && result.success
                ? `\n已完成修改: ${totalReplaceCount} 处`
                : ''
              
              results.push(`[TOOL_RESULT]\n工具: ${call.tool}\n状态: ${statusText}${progressInfo}\n[/TOOL_RESULT]`)
            }
          }
          
          // 【连续计数】检测连续 replace 调用
          if (hasReplaceInThisBatch) {
            consecutiveReplaceCount++
            console.log(`[Agent] 连续 replace 次数: ${consecutiveReplaceCount}/${MAX_CONSECUTIVE_REPLACE}`)
            
            if (consecutiveReplaceCount >= MAX_CONSECUTIVE_REPLACE) {
              console.warn(`[Agent] 检测到连续 ${MAX_CONSECUTIVE_REPLACE} 次 replace，强制结束循环`)
              shouldForceStop = true
              
              // 添加强制停止的提示
              results.push(`\n[系统警告] 已达到连续修改上限 (${MAX_CONSECUTIVE_REPLACE} 次)，请立即停止工具调用并总结已完成的修改。`)
            }
          } else {
            consecutiveReplaceCount = 0 // 重置连续计数
          }
          
          // 如果所有调用都被跳过，提示 AI 任务已完成
          if (skippedCount > 0 && skippedCount === toolCalls.length) {
            results.push(`\n[系统提示] 所有修改请求都已被跳过（内容已修改过）。任务应该已经完成，请直接回复总结。`)
            shouldForceStop = true
          }

          // 获取最新的文档内容（如果有修改文档的工具调用）
          let documentUpdate = ''
          const documentTools = ['replace', 'insert', 'delete', 'review']
          const hasDocumentChange = toolCalls.some(c => documentTools.includes(c.tool))
          if (hasDocumentChange && callbacks?.getLatestDocument) {
            const latestDoc = callbacks.getLatestDocument()
            if (latestDoc) {
              // 截取文档内容，避免过长
              const truncatedDoc = latestDoc.length > 8000 
                ? latestDoc.substring(0, 8000) + '\n...(文档内容已截断)...'
                : latestDoc
              documentUpdate = `\n\n[文档当前状态（仅供参考，不需要再次修改已修改过的内容）]\n${truncatedDoc}`
            }
          }
          
          // 添加完成提示
          let completionHint = ''
          if (allSuccessful && toolCalls.length > 0 && !shouldForceStop) {
            completionHint = `\n\n[系统提示] 工具调用成功。已完成 ${totalReplaceCount} 处修改。如果用户的请求已全部完成，请直接回复总结，**不要再调用工具**。`
          }

          // 将工具结果添加到对话（附带最新文档内容）
          const toolResultContent = (results.join('\n\n') + documentUpdate + completionHint).trim()
          conversationMessages.push({
            role: 'user',
            content: toolResultContent || '[工具调用已完成，无额外输出]'
          })

          // 如果强制停止，跳出循环
          if (shouldForceStop) {
            console.log('[Agent] 强制停止，准备输出总结')
            // 再调用一次 API 让 AI 输出总结
            const summaryResponse = await callAPI(
              conversationMessages,
              abortControllerRef.current.signal
            )
            const parsedSummary = parseAssistantOutput(cleanModelOutput(summaryResponse))
            if (parsedSummary.summary) {
              latestSummary = parsedSummary.summary
              setStreamingSummary(parsedSummary.summary)
            }
            accumulatedContent = parsedSummary.summary || parsedSummary.displayText || summaryResponse
            finalContentForMemory = accumulatedContent
            break
          }

          // 继续循环，让 AI 处理工具结果
          continue
        }

        // 没有工具调用，AI 完成了任务
        // 优先使用当前响应，如果为空则使用累积的内容
        console.log('[Agent] 最终响应:', response?.substring(0, 200))
        console.log('[Agent] 累积内容:', accumulatedContent?.substring(0, 200))
        const parsedFinal = parseAssistantOutput(cleanModelOutput(response))
        if (parsedFinal.summary) {
          latestSummary = parsedFinal.summary
          setStreamingSummary(parsedFinal.summary)
        }
        const finalContent =
          parsedFinal.summary ||
          parsedFinal.displayText ||
          latestSummary ||
          response.trim() ||
          accumulatedContent
        console.log('[Agent] 最终内容:', finalContent?.substring(0, 200))
        finalContentForMemory = finalContent
        callbacks?.onContent?.(finalContent)
        callbacks?.onComplete?.(finalContent, allToolResults)
        break
      }

      // 如果达到最大迭代次数，也要调用 onComplete
      if (iteration >= maxIterations) {
        console.warn(`[Agent] 达到最大迭代次数 ${maxIterations}，强制结束`)
        console.log('[Agent] 累积内容:', accumulatedContent?.substring(0, 200))
        console.log('[Agent] 最后响应:', lastResponse?.substring(0, 200))
        // 使用累积的内容或最后的响应
        const finalContent = latestSummary || accumulatedContent || lastResponse || '任务已完成（达到最大步骤数）'
        console.log('[Agent] 最终内容:', finalContent?.substring(0, 200))
        finalContentForMemory = finalContent
        callbacks?.onComplete?.(finalContent, allToolResults)
      }

      if (finalContentForMemory && memoryEnabled && window.electronAPI?.memoryAppend) {
        try {
          const entry = buildMemoryEntry(
            cleanMessageForSend(content),
            cleanMessageForSend(finalContentForMemory),
            allToolResults,
            memoryContext?.workspaceSummary
          )
          const maxChars = settings.memoryMaxChars || 2000
          const trimmed = entry.length > maxChars ? entry.slice(0, maxChars) + '... (已截断)' : entry
          await memoryAppend({
            text: trimmed,
            source: 'chat',
            tags: ['L1'],
          })
          await memoryAppendSession({
            sessionId: sessionIdRef.current,
            text: trimmed,
            meta: { workspaceKey, type: 'turn' },
          })
        } catch {
          // ignore memory failures
        }
      }

    } catch (error) {
      if ((error as Error).name === 'AbortError') {
        console.log('请求已取消')
      } else {
        console.error('AI request failed:', error)
        callbacks?.onComplete?.(`请求失败：${(error as Error).message}`, allToolResults)
      }
    } finally {
      setIsLoading(false)
      setStreamingContent('')
    }
  }, [settings, messages])

  // Tab 补全功能 - 仅使用本地模型
  const getCompletion = useCallback(async (
    textBefore: string,
    _textAfter?: string
  ): Promise<string | null> => {
    const localConfig = settings.localModel
    if (!localConfig?.enabled || !localConfig.baseUrl) {
      console.log('本地模型未配置，Tab 补全不可用')
      return null
    }

    // 取消之前的补全请求
    if (completionAbortRef.current) {
      completionAbortRef.current.abort()
    }
    completionAbortRef.current = new AbortController()

    setIsCompleting(true)

    try {
      // 只取光标前最近的文本作为上下文（减少延迟）
      const contextLength = 500  // 最多500字符的上下文
      const recentText = textBefore.slice(-contextLength)
      
      // 补全专用提示词 - 简洁高效
      const completionPrompt = `你是一个文档写作助手。请根据上文内容，直接续写下一句话。

要求：
- 只输出续写的内容，不要任何解释或开场白
- 续写1-2句话即可，不要太长
- 保持与上文风格一致
- 如果上文是列表，继续列表格式

上文内容：
${recentText}

请直接续写：`

      console.log('使用本地模型补全:', localConfig.baseUrl)
      
      const headers: Record<string, string> = {
        'Content-Type': 'application/json',
      }
      if (localConfig.apiKey) {
        headers['Authorization'] = `Bearer ${localConfig.apiKey}`
      }

      const response = await fetch(`${localConfig.baseUrl}/chat/completions`, {
        method: 'POST',
        headers,
        signal: completionAbortRef.current.signal,
        body: JSON.stringify({
          model: localConfig.model,
          messages: [
            { role: 'user', content: completionPrompt }
          ],
          temperature: 0.3,
          max_tokens: 100,
          stream: false,
        }),
      })

      if (!response.ok) {
        console.error('本地模型补全请求失败:', response.status)
        return null
      }

      const data = await response.json()
      let completion = data.choices?.[0]?.message?.content || ''
      completion = cleanModelOutput(completion)
      completion = completion.replace(/^["']|["']$/g, '').trim()
      
      return completion || null

    } catch (error) {
      if ((error as Error).name === 'AbortError') {
        console.log('补全请求已取消')
      } else {
        console.error('本地模型补全失败:', error)
      }
      return null
    } finally {
      setIsCompleting(false)
    }
  }, [settings])

  // 取消补全
  const cancelCompletion = useCallback(() => {
    if (completionAbortRef.current) {
      completionAbortRef.current.abort()
      completionAbortRef.current = null
    }
    setIsCompleting(false)
  }, [])

  return (
    <AIContext.Provider
      value={{
        messages,
        isLoading,
        isCompleting,
        streamingContent,
        streamingReasoning,
        editPhase,
        streamingSummary,
        settings,
        addMessage,
        updateLastMessage,
        clearMessages,
        updateSettings,
        sendMessage,
        sendAgentMessage,
        getCompletion,
        cancelCompletion,
      }}
    >
      {children}
    </AIContext.Provider>
  )
}

export function useAI() {
  const context = useContext(AIContext)
  if (!context) {
    throw new Error('useAI must be used within an AIProvider')
  }
  return context
}
