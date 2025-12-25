export interface DocumentContent {
  title: string
  content: string
  styles: DocumentStyles
  lastModified: Date
}

export interface DocumentStyles {
  fontSize: number
  fontFamily: string
  lineHeight: number
  textAlign: 'left' | 'center' | 'right' | 'justify'
}

// 页面设置
export interface PageSetup {
  // 纸张大小
  paperSize: 'A4' | 'A3' | 'Letter' | 'Legal' | 'custom'
  customWidth?: string  // 自定义宽度，如 "210mm"
  customHeight?: string // 自定义高度，如 "297mm"
  // 页面方向
  orientation: 'portrait' | 'landscape'
  // 页边距
  margins: {
    top: string     // 如 "2.54cm", "1in"
    bottom: string
    left: string
    right: string
  }
}

// 页眉页脚设置
export interface HeaderFooterSetup {
  header?: {
    content: string
    alignment: 'left' | 'center' | 'right'
    showOnFirstPage: boolean
  }
  footer?: {
    content: string
    alignment: 'left' | 'center' | 'right'
    showOnFirstPage: boolean
  }
  pageNumber?: {
    enabled: boolean
    position: 'header' | 'footer'
    alignment: 'left' | 'center' | 'right'
    format: 'arabic' | 'roman' | 'letter' // 1,2,3 / I,II,III / A,B,C
    startFrom: number
  }
}

// 自定义样式定义
export interface CustomStyle {
  name: string
  basedOn?: string // 基于哪个样式继承
  // 字符格式
  fontFamily?: string
  fontSize?: string
  color?: string
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strikethrough?: boolean
  letterSpacing?: string
  // 段落格式
  alignment?: 'left' | 'center' | 'right' | 'justify'
  lineHeight?: string
  spaceBefore?: string
  spaceAfter?: string
  textIndent?: string
  marginLeft?: string
  marginRight?: string
  backgroundColor?: string
  border?: string
}

export interface FileItem {
  name: string
  path: string
  type: 'file' | 'folder'
  children?: FileItem[]
  content?: string
}

export interface DiffChange {
  searchText: string
  replaceText: string
  count: number
}

// Agent 执行步骤
export interface AgentStep {
  id: string
  type: 'thinking' | 'reading' | 'searching' | 'editing' | 'creating' | 'deleting' | 'completed'
  description: string
  status: 'pending' | 'running' | 'completed' | 'error'
  details?: string
  timestamp?: Date
}

// Agent 文件变更
export interface AgentFileChange {
  name: string
  additions: number
  deletions: number
  status: 'pending' | 'writing' | 'done'
  operations?: string[]
}

export interface ChatMessage {
  id: string
  role: 'user' | 'assistant' | 'system'
  content: string
  timestamp: Date
  isStreaming?: boolean
  operationType?: 'create' | 'edit' | 'analyze' | 'chat'
  diffChanges?: DiffChange[]  // 修改记录，用于显示 Diff 和跳转
  fileName?: string           // 相关文件名
  // Agent 进度信息（用于在聊天中显示进度）
  agentStatus?: {
    isActive: boolean
    currentAction?: string
    steps?: AgentStep[]
    fileChanges?: AgentFileChange[]
    thinkingTime?: number
  }
}

export interface AISettings {
  apiKey: string
  model: string
  baseUrl: string
  /** 兼容旧字段：部分组件仍读取 apiUrl */
  apiUrl?: string
  temperature: number
  maxTokens: number
  // 本地模型配置（用于快速补全）
  localModel?: {
    enabled: boolean
    baseUrl: string
    model: string
    apiKey?: string
  }
  // OpenRouter API Key（用于调用 Gemini 生成 PPT 视觉设计）
  openRouterApiKey?: string
  // DashScope API Key（阿里云百炼，用于 PPT 图像生成）
  dashscopeApiKey?: string
  // PPT 图像生成模型
  pptImageModel?: 'z-image-turbo' | 'qwen-image-plus' | 'gemini-image'
  // Brave Search API Key（用于联网搜索）
  braveApiKey?: string
}

export interface EditorCommand {
  type: 'insert' | 'replace' | 'delete' | 'format' | 'create'
  target?: string
  content?: string
  position?: 'start' | 'end' | 'cursor'
}

