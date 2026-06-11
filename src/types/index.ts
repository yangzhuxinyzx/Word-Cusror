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
  relativePath?: string
  extension?: string
  children?: FileItem[]
  content?: string
}

export interface DiffChange {
  diffId?: string
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
  reasoning?: string          // AI 思考过程（kimi-k2.5 等思考模型）
  images?: string[]           // 用户发送的图片 base64 URL（用于回显）
  // 工具调用卡片快照（onComplete 时从 streamItems 中提取，用于历史消息内联展示）
  toolCards?: {
    id: string
    tool: string
    label: string
    status: 'running' | 'success' | 'error' | 'skipped'
    detail?: string
    searchText?: string
    replaceText?: string
  }[]
  knowledgeHits?: {
    sourceScope: KnowledgeSourceScope
    sourcePath: string
    relativePath?: string
    fileType?: string
    title?: string
    score: number
    snippet: string
    category?: string
    statement?: string
  }[]
  // 流式交替快照：保留文字与工具卡片的交替顺序，用于历史消息交替渲染
  streamSnapshot?: (
    | { type: 'text'; id: string; content: string }
    | { type: 'tool'; id: string; toolCard: NonNullable<ChatMessage['toolCards']>[number] }
  )[]
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
  // Adobe Firefly（用于高保真清底/融合）
  adobeFireflyClientId?: string
  adobeFireflyClientSecret?: string
  // Black Forest Labs（用于后续 FLUX 高配图像编辑）
  bflApiKey?: string
  // PPT 图像生成模型
  pptImageModel?: 'gpt-image-2' | 'gemini-image' | 'z-image-turbo' | 'qwen-image-plus' | 'qwen-image-max'
  // Brave Search API Key（用于联网搜索）
  braveApiKey?: string
  // 记忆系统
  memoryEnabled?: boolean
  memoryTopK?: number
  memoryMaxChars?: number
  memoryTextWeight?: number
  memoryVectorWeight?: number
  memoryFlushThresholdChars?: number
  knowledgeEnabled?: boolean
  workspaceKnowledgeEnabled?: boolean
  globalKnowledgePath?: string
  profileMemoryEnabled?: boolean
  embeddingBaseUrl?: string
  embeddingApiKey?: string
  embeddingModel?: string
  knowledgeTopK?: number
}

export type KnowledgeSourceScope =
  | 'workspace'
  | 'global'
  | 'profile'
  | 'daily'
  | 'sessions'
  | string

export interface KnowledgeSearchResult {
  sourceScope: KnowledgeSourceScope
  sourcePath: string
  relativePath?: string
  fileType?: string
  title?: string
  score: number
  snippet: string
  metadata?: Record<string, unknown>
  category?: string
  statement?: string
}

export interface KnowledgePendingProfileItem {
  id: string
  category: string
  statement: string
  evidenceHash: string
  evidenceText: string
  sourceScope?: KnowledgeSourceScope
  sourcePath?: string
  metadata?: Record<string, unknown>
  createdAt: string
  updatedAt: string
}

export interface KnowledgeProfileFact extends KnowledgePendingProfileItem {}

export interface KnowledgeStatusResponse {
  success: boolean
  configured?: {
    knowledgeEnabled: boolean
    workspaceKnowledgeEnabled: boolean
    profileMemoryEnabled: boolean
    globalKnowledgePath: string
    embeddingBaseUrl: string
    embeddingModel: string
    embeddingConfigured: boolean
    knowledgeTopK: number
  }
  workspace?: {
    rootPath: string
    status: string
    fileCount: number
    indexedFileCount: number
    chunkCount: number
    lastIndexedAt?: string | null
    lastError?: string
  } | null
  global?: {
    rootPath: string
    status: string
    fileCount: number
    indexedFileCount: number
    chunkCount: number
    lastIndexedAt?: string | null
    lastError?: string
  } | null
  profile?: {
    pendingCount: number
    factCount: number
  }
  error?: string
}

export interface PptTextStyleHint {
  fontSize: number
  textColor: string
  backgroundColor: string
  rotation: number
  lineCount: number
  align: 'left' | 'center' | 'right'
  familyHint?: string
  fontWeight?: number
  shadowColor?: string
  shadowOpacity?: number
  shadowOffsetX?: number
  shadowOffsetY?: number
  shadowBlur?: number
  strokeColor?: string
  strokeWidth?: number
  letterSpacing?: number
  lineHeight?: number
  opacity?: number
  blendMode?: 'normal' | 'multiply'
  textBounds?: {
    left: number
    top: number
    width: number
    height: number
  }
}

export interface PptCharBox {
  char: string
  index: number
  bounds: {
    left: number
    top: number
    width: number
    height: number
  }
}

export interface PptFontCandidate {
  candidateId: string
  family: string
  confidence: number
  source: 'workspace' | 'bundled' | 'system' | 'remote'
  fontPath?: string
  previewText?: string
}

export interface PptStyleEstimate extends PptTextStyleHint {
  textDirection: 'ltr' | 'rtl' | 'ttb'
  skewX?: number
  skewY?: number
}

export type PptCleanupStrategy =
  | 'analytic_fill'
  | 'local_inpaint'
  | 'adobe_firefly_fill'
  | 'none'

export type PptBlendStrategy =
  | 'deterministic'
  | 'adobe_composite'
  | 'flux_refine'
  | 'none'

export interface PptEditScore {
  total: number
  ocrExactness: number
  fontStyleSimilarity: number
  backgroundPreservation: number
  edgeArtifactScore: number
  overflowPenalty: number
}

export interface PptEditCandidate {
  candidateId: string
  boxId: string
  label: string
  previewDataUrl?: string
  fontCandidateId?: string
  cleanupStrategy: PptCleanupStrategy
  blendStrategy: PptBlendStrategy
  score: PptEditScore
  applied?: boolean
  metrics?: Record<string, number | string | boolean | undefined>
}

export interface PptDetectedTextBoxV2 {
  boxId: string
  text: string
  confidence: number
  polygon: Array<[number, number]>
  bounds: {
    left: number
    top: number
    width: number
    height: number
  }
  readingOrder: number
  styleHint: PptTextStyleHint
  charBoxes?: PptCharBox[]
  rotation?: number
  skew?: number
  textDirection?: 'ltr' | 'rtl' | 'ttb'
  backgroundComplexity?: 'simple' | 'medium' | 'complex'
  styleComplexity?: 'plain' | 'styled' | 'textured'
  fontCandidates?: PptFontCandidate[]
  styleEstimate?: PptStyleEstimate
}

export type PptDetectedTextBox = PptDetectedTextBoxV2

export interface PptTextEditStyleOverride {
  fontCandidateId?: string
  fontFamily?: string
  fontSize?: number
  textColor?: string
  strokeColor?: string
  strokeWidth?: number
  shadowBlur?: number
  shadowColor?: string
  shadowOffsetX?: number
  shadowOffsetY?: number
  letterSpacing?: number
  lineHeight?: number
  opacity?: number
  cleanupStrategy?: PptCleanupStrategy
  blendStrategy?: PptBlendStrategy
}

export interface PptTextEditOperation {
  boxId: string
  fromText: string
  toText: string
  styleMode?: 'preserve'
  bounds?: {
    left: number
    top: number
    width: number
    height: number
  }
  styleOverride?: PptTextEditStyleOverride
}

export interface DocxCompatSettings {
  compatibilityMode?: string
  characterSpacingControl?: string
  noPunctuationKerning?: boolean
  useFELayout?: boolean
  doNotUseEastAsianBreakRules?: boolean
  compressPunctuation?: boolean
}

export interface DocxFontTableEntry {
  name: string
  altName?: string
  family?: string
  charset?: string
  pitch?: string
  panose1?: string
}

export interface DocxStyleGraphNode {
  styleId: string
  type?: string
  name?: string
  basedOn?: string
  next?: string
  link?: string
  isDefault?: boolean
}

export interface DocxRelationshipSummary {
  id: string
  type?: string
  target?: string
}

export interface DocxImageAssetSummary {
  relId: string
  target: string
  size: number
}

export interface DocxTableSummary {
  index: number
  rows: number
  columns: number
  widthTwips?: number
  layout?: string
  floating?: boolean
}

export interface DocxReferencedFont {
  name: string
  alternates: string[]
}

export interface DocxInspectionReport {
  sourcePath: string
  extractedDir: string
  extractedFiles: string[]
  xmlPaths: {
    document?: string
    styles?: string
    settings?: string
    fontTable?: string
    numbering?: string
    theme?: string
    footers?: string[]
    rels?: string[]
  }
  createdAt: string
  summary: {
    pageSettings?: {
      widthTwips?: number
      heightTwips?: number
      marginTopTwips?: number
      marginRightTwips?: number
      marginBottomTwips?: number
      marginLeftTwips?: number
      headerTwips?: number
      footerTwips?: number
      columns?: number
      columnSpacingTwips?: number
      docGridLinePitch?: number
      docGridCharSpace?: number
    }
    compat: DocxCompatSettings
    fontTable: DocxFontTableEntry[]
    referencedFonts: DocxReferencedFont[]
    styleGraph: DocxStyleGraphNode[]
    relationships: DocxRelationshipSummary[]
    images: DocxImageAssetSummary[]
    tables: DocxTableSummary[]
    tocFields: string[]
    footerTargets: string[]
    referencedStyleIds: string[]
  }
}

export interface WordOracleMissingFont {
  name: string
  alternates?: string[]
  resolvedName?: string
}

export interface WordOraclePageImage {
  pageIndex: number
  path: string
  width: number
  height: number
}

export interface WordOracleArtifact {
  exportId: string
  sourcePath: string
  pdfPath: string
  imageDir: string
  pageCount: number
  pages: WordOraclePageImage[]
  exportedAt: string
  wordAppPath?: string
  inspectorExtractedDir?: string
  missingFonts?: WordOracleMissingFont[]
}

export interface RenderMismatch {
  pageIndex: number
  oracleImagePath: string
  diffImagePath: string
  mismatchRatio: number
  mismatchPixels: number
  thresholdRatio: number
  thresholdExceeded: boolean
  oracleSize: {
    width: number
    height: number
  }
  currentSize: {
    width: number
    height: number
  }
  bbox?: {
    x: number
    y: number
    width: number
    height: number
  } | null
}

export interface RenderAlignmentReport {
  artifactId: string
  sourcePath: string
  createdAt: string
  expectedPageCount: number
  actualPageCount: number
  pageCountMatches: boolean
  thresholdRatio: number
  mismatchCount: number
  pages: RenderMismatch[]
  currentPageIndicesOverThreshold: number[]
  status: 'aligned' | 'misaligned' | 'unavailable'
}

export interface NativeTextMeasureEntry {
  id: string
  text: string
  fontFamily: string
  fontSize: number
  fontWeight?: 'normal' | 'bold' | number
  fontStyle?: 'normal' | 'italic'
  letterSpacing?: number
  lineHeight?: number
}

export interface NativeTextMeasureResult {
  id: string
  width: number
  ascent: number
  descent: number
  lineHeight: number
  baseline: number
  resolvedFontFamily?: string
  usedFallback?: boolean
}

export interface NativeTextFontAvailability {
  name: string
  available: boolean
  resolvedName?: string
}

export interface NativeTextMeasureResponse {
  success: boolean
  measurements?: NativeTextMeasureResult[]
  fonts?: NativeTextFontAvailability[]
  error?: string
}

export interface EditorCommand {
  type: 'insert' | 'replace' | 'delete' | 'format' | 'create'
  target?: string
  content?: string
  position?: 'start' | 'end' | 'cursor'
}

