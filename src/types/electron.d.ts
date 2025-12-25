// Electron API 类型定义
export interface FileItem {
  name: string
  path: string
  relativePath: string
  type: 'file' | 'folder'
  extension?: string
  children?: FileItem[]
}

export interface FileReadResult {
  success: boolean
  data?: string
  type?: 'text' | 'docx' | 'doc-html' | 'pptx'
  error?: string
}

export interface FileWriteResult {
  success: boolean
  error?: string
}

export interface FolderReadResult {
  success: boolean
  data?: FileItem[]
  error?: string
}

export interface FileInfo {
  size: number
  created: Date
  modified: Date
  isFile: boolean
  isDirectory: boolean
}

export interface WebSearchResultItem {
  title: string
  link: string
  snippet: string
  extraSnippets?: string[]
}

export interface WebSearchFaqItem {
  question: string
  answer: string
  title?: string
  link?: string
}

export interface WebSearchNewsItem {
  title: string
  link: string
  source?: string
  description?: string
  breaking?: boolean
  isLive?: boolean
  age?: string
}

export interface WebSearchVideoItem {
  title: string
  link: string
  description?: string
  duration?: string
  thumbnail?: string
  viewCount?: number | string
  creator?: string
  publisher?: string
}

export interface WebSearchDiscussionItem {
  link: string
  forumName?: string
  question?: string
  topComment?: string
}

export interface WebSearchSections {
  web: WebSearchResultItem[]
  faq: WebSearchFaqItem[]
  news: WebSearchNewsItem[]
  videos: WebSearchVideoItem[]
  discussions: WebSearchDiscussionItem[]
}

export interface WebSearchResponse {
  success: boolean
  results?: WebSearchResultItem[]
  sections?: WebSearchSections
  summarizerKey?: string
  message?: string
  raw?: unknown
}

export interface ExcelCell {
  r: number
  c: number
  v: unknown
  t?: string
  w?: string
  f?: string
  l?: { Target?: string; tooltip?: string; [key: string]: any }
  z?: string
  cmt?: string
  display?: string
  s?: Record<string, any>
}

export interface ExcelSheet {
  name: string
  range: { s: { r: number; c: number }; e: { r: number; c: number } }
  merges: Array<{ s: { r: number; c: number }; e: { r: number; c: number } }>
  colWidths?: Array<number | undefined>
  rowHeights?: Array<number | undefined>
  autoFilter?: any
  printArea?: any
  margins?: any
  dataValidations?: any
  cells: ExcelCell[]
}

export interface ExcelOpenResponse {
  success: boolean
  sheets?: ExcelSheet[]
  names?: any[]
  error?: string
  warning?: string  // xls 格式样式支持有限时的警告
  isXls?: boolean   // 是否为 xls 文件
  originalPath?: string  // 原始文件路径
}

export interface ExcelConvertResponse {
  success: boolean
  xlsxPath?: string
  message?: string
  error?: string
}

// Excel 操作返回类型
export interface ExcelCellData {
  address: string
  r: number
  c: number
  value: any
  text?: string
  formula?: string
  type?: number
}

export interface ExcelReadResult {
  success: boolean
  cells?: ExcelCellData[]
  range?: string
  error?: string
}

export interface ExcelSearchOptions {
  caseSensitive?: boolean
  matchWholeCell?: boolean
}

export interface ExcelSearchResult {
  success: boolean
  results?: ExcelCellData[]
  count?: number
  error?: string
}

export interface ExcelCellUpdate {
  address: string
  value?: any
  style?: {
    font?: {
      name?: string
      size?: number
      bold?: boolean
      italic?: boolean
      underline?: boolean | string
      strike?: boolean
      color?: { argb?: string }
    }
    fill?: {
      fgColor?: { argb?: string }
    }
    alignment?: {
      horizontal?: 'left' | 'center' | 'right' | 'justify'
      vertical?: 'top' | 'middle' | 'bottom'
      wrapText?: boolean
    }
    border?: {
      top?: { style?: string; color?: { argb?: string } }
      bottom?: { style?: string; color?: { argb?: string } }
      left?: { style?: string; color?: { argb?: string } }
      right?: { style?: string; color?: { argb?: string } }
    }
    numFmt?: string
  }
}

export interface ExcelWriteResult {
  success: boolean
  updatedCells?: string[]
  count?: number
  error?: string
}

export interface ExcelInsertResult {
  success: boolean
  insertedAt?: number
  count?: number
  error?: string
}

export interface ExcelDeleteResult {
  success: boolean
  deletedFrom?: number
  count?: number
  error?: string
}

export interface ExcelSheetResult {
  success: boolean
  sheetName?: string
  deletedSheet?: string
  error?: string
}

export interface ExcelListSheetsResult {
  success: boolean
  sheets?: Array<{
    name: string
    rowCount: number
    columnCount: number
  }>
  error?: string
}

export interface ExcelMergeResult {
  success: boolean
  mergedRange?: string
  unmergedRange?: string
  error?: string
}

// Excel 创建选项
export interface ExcelCreateSheetConfig {
  name?: string
  data?: (string | number | boolean | null | { value: any; style?: any })[][]
  columnWidths?: number[]
  merges?: string[]
}

export interface ExcelCreateOptions {
  sheets?: ExcelCreateSheetConfig[]
  openAfterCreate?: boolean
}

export interface ExcelCreateResult {
  success: boolean
  filePath?: string
  sheetsCreated?: string[]
  openAfterCreate?: boolean
  error?: string
}

export interface ElectronAPI {
  // 文件夹操作
  selectFolder: () => Promise<string | null>
  readFolder: (folderPath: string) => Promise<FolderReadResult>
  
  // 文件操作
  readFile: (filePath: string) => Promise<FileReadResult>
  writeFile: (filePath: string, content: string) => Promise<FileWriteResult>
  writeBinaryFile: (filePath: string, base64Data: string) => Promise<FileWriteResult>
  
  // 对话框
  saveFileDialog: (defaultName: string) => Promise<string | null>
  
  // 文件管理
  createFile: (folderPath: string, fileName: string, content?: string) => Promise<{ success: boolean; path?: string; error?: string }>
  deleteFile: (filePath: string) => Promise<FileWriteResult>
  renameFile: (oldPath: string, newPath: string) => Promise<FileWriteResult>
  showInFolder: (filePath: string) => Promise<{ success: boolean }>
  getFileInfo: (filePath: string) => Promise<{ success: boolean; data?: FileInfo; error?: string }>
  
  // ONLYOFFICE 文件服务
  getFileUrl: (filePath: string) => Promise<string>
  // 渲染进程直接使用的本地 URL（用于预览图片等）
  getLocalFileUrl: (filePath: string) => Promise<string>
  
  // ONLYOFFICE Document Builder - 创建带格式的文档
  createFormattedDocument: (options: {
    filePath: string
    elements: Array<{
      type: 'heading' | 'paragraph' | 'table'
      content?: string
      level?: number
      bold?: boolean
      fontSize?: number
      fontFamily?: string
      alignment?: 'left' | 'center' | 'right' | 'justify'
      color?: string
      rows?: number
      cols?: number
      data?: string[][]
    }>
    title: string
  }) => Promise<{ success: boolean; path?: string; error?: string; fallback?: boolean }>
  
  // 模板文档替换（使用 docxtemplater，需要 {{占位符}} 格式）
  fillTemplate: (options: {
    templatePath: string
    outputPath: string
    replacements: Record<string, string>
  }) => Promise<{ success: boolean; path?: string; error?: string }>
  
  // DOCX 直接搜索替换（不需要占位符，完美保留格式）
  docxSearchReplace: (options: {
    sourcePath: string
    outputPath: string
    replacements: Array<{ search: string; replace: string }>
  }) => Promise<{ success: boolean; path?: string; replaceCount?: number; error?: string }>

  // Excel 读取（只读预览）
  excelOpen: (filePath: string) => Promise<ExcelOpenResponse>
  
  // Excel xls 转 xlsx
  excelConvertXlsToXlsx: (xlsPath: string) => Promise<ExcelConvertResponse>
  
  // 检查 LibreOffice 安装状态
  checkLibreOffice: () => Promise<{
    installed: boolean
    path: string | null
    downloadUrl: string | null
  }>

  // Excel 增删查改操作
  excelReadCells: (filePath: string, sheetName: string, range: string) => Promise<ExcelReadResult>
  excelSearch: (filePath: string, sheetName: string, searchText: string, options?: ExcelSearchOptions) => Promise<ExcelSearchResult>
  excelWriteCells: (filePath: string, sheetName: string, cellUpdates: ExcelCellUpdate[]) => Promise<ExcelWriteResult>
  excelInsertRows: (filePath: string, sheetName: string, startRow: number, count?: number, data?: any[][]) => Promise<ExcelInsertResult>
  excelInsertColumns: (filePath: string, sheetName: string, startCol: number, count?: number) => Promise<ExcelInsertResult>
  excelDeleteRows: (filePath: string, sheetName: string, startRow: number, count?: number) => Promise<ExcelDeleteResult>
  excelDeleteColumns: (filePath: string, sheetName: string, startCol: number, count?: number) => Promise<ExcelDeleteResult>
  excelAddSheet: (filePath: string, sheetName: string) => Promise<ExcelSheetResult>
  excelDeleteSheet: (filePath: string, sheetName: string) => Promise<ExcelSheetResult>
  excelListSheets: (filePath: string) => Promise<ExcelListSheetsResult>
  excelMergeCells: (filePath: string, sheetName: string, range: string) => Promise<ExcelMergeResult>
  excelUnmergeCells: (filePath: string, sheetName: string, range: string) => Promise<ExcelMergeResult>
  // 这些接口返回结构在主进程实现中可能包含更多统计字段（如 count、sortedRows 等），这里用宽松类型兼容
  excelSetFormula: (filePath: string, sheetName: string, formulas: any) => Promise<{ success: boolean; error?: string; [key: string]: any }>
  excelSort: (filePath: string, sheetName: string, options: any) => Promise<{ success: boolean; error?: string; [key: string]: any }>
  excelConditionalFormat: (filePath: string, sheetName: string, options: any) => Promise<{ success: boolean; error?: string; [key: string]: any }>
  excelAutoFill: (filePath: string, sheetName: string, options: any) => Promise<{ success: boolean; error?: string; [key: string]: any }>
  excelSetDimensions: (filePath: string, sheetName: string, options: any) => Promise<{ success: boolean; error?: string; [key: string]: any }>
  excelAddChart: (filePath: string, sheetName: string, options: any) => Promise<{ success: boolean; error?: string; [key: string]: any }>
  excelCalculate: (filePath: string, sheetName: string, addresses: any) => Promise<{ success: boolean; error?: string; [key: string]: any }>
  excelCreate: (filePath: string, options?: ExcelCreateOptions) => Promise<ExcelCreateResult>
  excelClose: (filePath: string) => Promise<{ success: boolean }>
  excelReload: (filePath: string) => Promise<ExcelOpenResponse>
  
  // 【新增】Excel 高级功能
  excelSetFilter: (filePath: string, sheetName: string, options: {
    range?: string
    remove?: boolean
  }) => Promise<{ success: boolean; message?: string; error?: string }>
  
  excelSetValidation: (filePath: string, sheetName: string, options: {
    range: string
    type?: 'list' | 'whole' | 'decimal' | 'date' | 'textLength'
    values?: string[]
    min?: number
    max?: number
    allowBlank?: boolean
    showError?: boolean
    errorTitle?: string
    errorMessage?: string
    remove?: boolean
  }) => Promise<{ success: boolean; message?: string; error?: string }>
  
  excelSetHyperlink: (filePath: string, sheetName: string, options: {
    cell: string
    url?: string
    text?: string
    tooltip?: string
    remove?: boolean
  }) => Promise<{ success: boolean; message?: string; error?: string }>
  
  excelFindReplace: (filePath: string, sheetName: string, options: {
    find: string
    replace?: string
    matchCase?: boolean
    matchWholeCell?: boolean
    allSheets?: boolean
  }) => Promise<{ success: boolean; count?: number; message?: string; details?: any[]; error?: string }>
  
  excelInsertChart: (filePath: string, sheetName: string, options: {
    type?: 'column' | 'bar' | 'line' | 'pie' | 'area' | 'scatter'
    dataRange: string
    title?: string
    position?: string
    width?: number
    height?: number
  }) => Promise<{ success: boolean; message?: string; chartConfig?: any; error?: string }>

  // Web 搜索（Brave MCP）
  webSearch: (options: {
    query: string
    locale?: string
    region?: string
    num?: number
    braveApiKey?: string
  }) => Promise<WebSearchResponse>

  // OpenRouter Gemini：大纲转文生图提示词
  openrouterGeminiPptPrompts: (options: {
    apiKey: string
    outline: string
    slideCount?: number
    theme?: string
    style?: string
    model?: string
    // 主模型回退参数（当没有 OpenRouter API Key 时使用）
    mainApiKey?: string
    mainBaseUrl?: string
    mainModel?: string
  }) => Promise<{
    success: boolean
    slides?: Array<{
      pageNumber: number
      pageType: string
      prompt: string
      negativePrompt: string
    }>
    designConcept?: string
    colorPalette?: string
    raw?: string
    error?: string
  }>

  // PPT 生成：DashScope 生图（并发=2）→ 后处理 1920x1200 → 打包 16:10 .pptx
  pptGenerateDeck: (options: {
    outputPath: string
    slides: Array<{
      prompt: string
      negativePrompt?: string
      originalChineseContent?: string
    }>
    // 主模型 API Key（用于 Gemini 生图）
    mainApiKey?: string
    dashscope?: {
      apiKey?: string
      region?: 'cn' | 'intl'
      size?: string // DashScope preset, default 1664*928
      model?: 'z-image-turbo' | 'qwen-image-plus' | 'gemini-image' // 生图模型
      promptExtend?: boolean
      watermark?: boolean
      negativePromptDefault?: string
    }
    postprocess?: {
      mode?: 'letterbox' | 'cover'
    }
    repair?: {
      enabled?: boolean
      openRouterApiKey?: string
      model?: string
      maxAttempts?: number
      deckContext?: {
        designConcept?: string
        colorPalette?: string
      }
    }
    outline?: unknown // 原始大纲（用于保存元数据）
  }) => Promise<{ success: boolean; path?: string; slideCount?: number; error?: string }>
  
  // PPT 编辑：整页重做 / 局部编辑
  pptEditSlides: (options: {
    pptxPath: string
    pageNumbers: number[] // 1-based 页码数组
    feedback?: string     // 用户反馈
    mode: 'regenerate' | 'partial_edit'
    openRouterApiKey: string
    dashscopeApiKey: string
    /** 主模型 API Key（当 pptImageModel=gemini-image 时用于 LinAPI 生图；否则可不传） */
    mainApiKey?: string
    /** 生图模型选择（与设置面板一致） */
    pptImageModel?: 'z-image-turbo' | 'qwen-image-plus' | 'gemini-image'
    deckContext?: {
      designConcept?: string
      colorPalette?: string
    }
    regionScreenshot?: string  // 用户框选区域的截图 base64
    regionRect?: {             // 框选区域坐标
      x: number
      y: number
      w: number
      h: number
    }
  }) => Promise<{
    success: boolean
    path?: string
    editedPages?: number[]
    logs?: Array<{ pageNum: number; success: boolean; error?: string }>
    error?: string
  }>

  // PPTX 预览：LibreOffice 渲染为 PNG（只读预览）
  pptxRenderPreview: (filePath: string) => Promise<{
    success: boolean
    images?: string[]
    cacheDir?: string
    cached?: boolean
    error?: string
    details?: string
    downloadUrl?: string | null
  }>
  
  // 平台信息
  platform: string
  isElectron: boolean
}

declare global {
  interface Window {
    electronAPI?: ElectronAPI
  }
}

export {}

