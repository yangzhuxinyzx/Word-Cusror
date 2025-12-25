const { contextBridge, ipcRenderer } = require('electron')

// 暴露安全的 API 给渲染进程
contextBridge.exposeInMainWorld('electronAPI', {
  // 文件夹操作
  selectFolder: () => ipcRenderer.invoke('select-folder'),
  readFolder: (folderPath) => ipcRenderer.invoke('read-folder', folderPath),
  
  // 文件操作
  readFile: (filePath) => ipcRenderer.invoke('read-file', filePath),
  writeFile: (filePath, content) => ipcRenderer.invoke('write-file', filePath, content),
  writeBinaryFile: (filePath, base64Data) => ipcRenderer.invoke('write-binary-file', filePath, base64Data),
  
  // 对话框
  saveFileDialog: (defaultName) => ipcRenderer.invoke('save-file-dialog', defaultName),
  
  // 文件管理
  createFile: (folderPath, fileName, content) => ipcRenderer.invoke('create-file', folderPath, fileName, content),
  deleteFile: (filePath) => ipcRenderer.invoke('delete-file', filePath),
  renameFile: (oldPath, newPath) => ipcRenderer.invoke('rename-file', oldPath, newPath),
  showInFolder: (filePath) => ipcRenderer.invoke('show-in-folder', filePath),
  getFileInfo: (filePath) => ipcRenderer.invoke('get-file-info', filePath),
  
  // ONLYOFFICE 文件服务
  getFileUrl: (filePath) => ipcRenderer.invoke('get-file-url', filePath),
  // 渲染进程直接使用的本地 URL（用于预览图片等）
  getLocalFileUrl: (filePath) => ipcRenderer.invoke('get-local-file-url', filePath),
  
  // ONLYOFFICE Document Builder - 创建带格式的文档
  createFormattedDocument: (options) => ipcRenderer.invoke('create-formatted-document', options),
  
  // 模板文档替换（使用 docxtemplater，需要 {{占位符}} 格式）
  fillTemplate: (options) => ipcRenderer.invoke('fill-template', options),
  
  // DOCX 直接搜索替换（不需要占位符，直接替换文本，完美保留格式）
  docxSearchReplace: (options) => ipcRenderer.invoke('docx-search-replace', options),
  
  // Excel 读取（只读预览）
  excelOpen: (filePath) => ipcRenderer.invoke('excel-open', filePath),
  
  // Excel xls 转 xlsx
  excelConvertXlsToXlsx: (xlsPath) => ipcRenderer.invoke('excel-convert-xls-to-xlsx', xlsPath),
  
  // 检查 LibreOffice 是否安装
  checkLibreOffice: () => ipcRenderer.invoke('check-libreoffice'),
  
  // Excel 增删查改操作
  excelReadCells: (filePath, sheetName, range) => ipcRenderer.invoke('excel-read-cells', filePath, sheetName, range),
  excelSearch: (filePath, sheetName, searchText, options) => ipcRenderer.invoke('excel-search', filePath, sheetName, searchText, options),
  excelWriteCells: (filePath, sheetName, cellUpdates) => ipcRenderer.invoke('excel-write-cells', filePath, sheetName, cellUpdates),
  excelInsertRows: (filePath, sheetName, startRow, count, data) => ipcRenderer.invoke('excel-insert-rows', filePath, sheetName, startRow, count, data),
  excelInsertColumns: (filePath, sheetName, startCol, count) => ipcRenderer.invoke('excel-insert-columns', filePath, sheetName, startCol, count),
  excelDeleteRows: (filePath, sheetName, startRow, count) => ipcRenderer.invoke('excel-delete-rows', filePath, sheetName, startRow, count),
  excelDeleteColumns: (filePath, sheetName, startCol, count) => ipcRenderer.invoke('excel-delete-columns', filePath, sheetName, startCol, count),
  excelAddSheet: (filePath, sheetName) => ipcRenderer.invoke('excel-add-sheet', filePath, sheetName),
  excelDeleteSheet: (filePath, sheetName) => ipcRenderer.invoke('excel-delete-sheet', filePath, sheetName),
  excelListSheets: (filePath) => ipcRenderer.invoke('excel-list-sheets', filePath),
  excelMergeCells: (filePath, sheetName, range) => ipcRenderer.invoke('excel-merge-cells', filePath, sheetName, range),
  excelUnmergeCells: (filePath, sheetName, range) => ipcRenderer.invoke('excel-unmerge-cells', filePath, sheetName, range),
  excelCreate: (filePath, options) => ipcRenderer.invoke('excel-create', filePath, options),
  excelSetFormula: (filePath, sheetName, formulas) => ipcRenderer.invoke('excel-set-formula', filePath, sheetName, formulas),
  excelSort: (filePath, sheetName, options) => ipcRenderer.invoke('excel-sort', filePath, sheetName, options),
  excelConditionalFormat: (filePath, sheetName, options) => ipcRenderer.invoke('excel-conditional-format', filePath, sheetName, options),
  excelAutoFill: (filePath, sheetName, options) => ipcRenderer.invoke('excel-auto-fill', filePath, sheetName, options),
  excelSetDimensions: (filePath, sheetName, options) => ipcRenderer.invoke('excel-set-dimensions', filePath, sheetName, options),
  excelAddChart: (filePath, sheetName, options) => ipcRenderer.invoke('excel-add-chart', filePath, sheetName, options),
  excelCalculate: (filePath, sheetName, addresses) => ipcRenderer.invoke('excel-calculate', filePath, sheetName, addresses),
  excelClose: (filePath) => ipcRenderer.invoke('excel-close', filePath),
  excelReload: (filePath) => ipcRenderer.invoke('excel-reload', filePath),
  
  // 【新增】Excel 高级功能
  excelSetFilter: (filePath, sheetName, options) => ipcRenderer.invoke('excel-set-filter', filePath, sheetName, options),
  excelSetValidation: (filePath, sheetName, options) => ipcRenderer.invoke('excel-set-validation', filePath, sheetName, options),
  excelSetHyperlink: (filePath, sheetName, options) => ipcRenderer.invoke('excel-set-hyperlink', filePath, sheetName, options),
  excelFindReplace: (filePath, sheetName, options) => ipcRenderer.invoke('excel-find-replace', filePath, sheetName, options),
  excelInsertChart: (filePath, sheetName, options) => ipcRenderer.invoke('excel-insert-chart', filePath, sheetName, options),

  // Web 搜索（Scrapeless MCP）
  webSearch: (options) => ipcRenderer.invoke('web-search', options),

  // PPT 生成：DashScope 生图（并发=2）→ 后处理 1920x1200 → 打包 16:10 .pptx
  pptGenerateDeck: (options) => ipcRenderer.invoke('ppt-generate-deck', options),
  
  // PPT 编辑：整页重做 / 局部编辑
  pptEditSlides: (options) => ipcRenderer.invoke('ppt-edit-slides', options),

  // OpenRouter Gemini：大纲转文生图提示词
  openrouterGeminiPptPrompts: (options) => ipcRenderer.invoke('openrouter-gemini-ppt-prompts', options),

  // PPTX 预览：LibreOffice 渲染为 PNG（只读预览）
  pptxRenderPreview: (filePath) => ipcRenderer.invoke('pptx-render-preview', filePath),
  
  // 平台信息
  platform: process.platform,
  isElectron: true
})

// 通知渲染进程 Electron 已就绪
window.addEventListener('DOMContentLoaded', () => {
  console.log('Word-Cursor Electron 已就绪')
})

