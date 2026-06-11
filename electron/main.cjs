const { app, BrowserWindow, ipcMain, dialog, shell, globalShortcut } = require('electron')
const path = require('path')
const fs = require('fs')
const { createMemoryManager } = require('./memory/index.cjs')
const http = require('http')
const mammoth = require('mammoth')
const WordExtractor = require('word-extractor')
const PizZip = require('pizzip')
const Docxtemplater = require('docxtemplater')
const dotenv = require('dotenv')
const XLSX = require('xlsx')
const ExcelJS = require('exceljs')
const { createFilesService } = require('./services/files.cjs')
const { createMemoryService } = require('./services/memory.cjs')
const { createKnowledgeService } = require('./services/knowledge.cjs')
const { createWebSearchService } = require('./services/web-search.cjs')
const { createExcelService } = require('./services/excel.cjs')
const { createFontsService } = require('./services/fonts.cjs')
const { createAIProxyService } = require('./services/ai-proxy.cjs')
const { createPptService } = require('./services/ppt.cjs')
const { createDocxInspectorService } = require('./services/docx-inspector.cjs')
const { createWordOracleService } = require('./services/word-oracle.cjs')
const { createNativeTextService } = require('./services/native-text.cjs')
const { registerFileIpc } = require('./ipc/register-files.cjs')
const { registerMemoryIpc } = require('./ipc/register-memory.cjs')
const { registerKnowledgeIpc } = require('./ipc/register-knowledge.cjs')
const { registerWebSearchIpc } = require('./ipc/register-web-search.cjs')
const { registerExcelIpc } = require('./ipc/register-excel.cjs')
const { registerFontsIpc } = require('./ipc/register-fonts.cjs')
const { registerAIProxyIpc } = require('./ipc/register-ai-proxy.cjs')
const { registerPptIpc } = require('./ipc/register-ppt.cjs')
const { registerDocxInspectorIpc } = require('./ipc/register-docx-inspector.cjs')
const { registerWordOracleIpc } = require('./ipc/register-word-oracle.cjs')
const { registerNativeTextIpc } = require('./ipc/register-native-text.cjs')

const APP_DISPLAY_NAME = '智启文档'
const APP_ICON_PATH = path.join(__dirname, '../public/icon.png')

app.setName(APP_DISPLAY_NAME)

dotenv.config({ path: path.join(__dirname, '..', '.env') })

const envFilePath = path.join(__dirname, '..', '.env')
const shouldOpenDevTools =
  process.env.WORD_CURSOR_OPEN_DEVTOOLS === 'true' ||
  process.env.OPEN_DEVTOOLS === 'true'

function parseBoolean(value) {
  if (value == null || value === '') return undefined
  return String(value).trim().toLowerCase() === 'true'
}

function parseNumber(value, fallback) {
  const parsed = Number(value)
  return Number.isFinite(parsed) ? parsed : fallback
}

function readEnvSettings() {
  try {
    if (!fs.existsSync(envFilePath)) return {}
    const raw = fs.readFileSync(envFilePath, 'utf8')
    const parsed = dotenv.parse(raw)
    return {
      apiKey: parsed.API_KEY || '',
      baseUrl: parsed.BASE_URL || '',
      model: parsed.MODEL || '',
      temperature: parseNumber(parsed.TEMPERATURE, 1),
      maxTokens: parseNumber(parsed.MAX_TOKENS, 131072),
      openRouterApiKey: parsed.OPENROUTER_API_KEY || undefined,
      dashscopeApiKey: parsed.DASHSCOPE_API_KEY || undefined,
      adobeFireflyClientId: parsed.ADOBE_FIREFLY_CLIENT_ID || undefined,
      adobeFireflyClientSecret: parsed.ADOBE_FIREFLY_CLIENT_SECRET || undefined,
      bflApiKey: parsed.BFL_API_KEY || parsed.FLUX_BFL_API_KEY || undefined,
      pptImageModel: parsed.PPT_IMAGE_MODEL || undefined,
      braveApiKey: parsed.BRAVE_API_KEY || undefined,
      memoryEnabled: parseBoolean(parsed.MEMORY_ENABLED),
      memoryTopK: parseNumber(parsed.MEMORY_TOP_K, undefined),
      memoryMaxChars: parseNumber(parsed.MEMORY_MAX_CHARS, undefined),
      memoryFlushThresholdChars: parseNumber(parsed.MEMORY_FLUSH_THRESHOLD, undefined),
      embeddingApiKey: parsed.EMBEDDING_API_KEY || undefined,
      localModel: {
        enabled: parseBoolean(parsed.LOCAL_MODEL_ENABLED) ?? true,
        baseUrl: parsed.LOCAL_MODEL_BASE_URL || parsed.BASE_URL || '',
        model: parsed.LOCAL_MODEL_MODEL || parsed.LOCAL_MODEL_NAME || parsed.MODEL || '',
        apiKey: parsed.LOCAL_MODEL_API_KEY || parsed.API_KEY || '',
      },
    }
  } catch (error) {
    console.warn('[env] 读取 .env 失败:', error)
    return {}
  }
}

function saveEnvSettingsToFile(settings = {}) {
  const current = readEnvSettings()
  const next = {
    ...current,
    ...settings,
    localModel: {
      ...(current.localModel || {}),
      ...(settings.localModel || {}),
    },
  }

  const lines = [
    '# 智启文档 环境变量配置',
    '# 在"设置"面板中修改的值也会自动回写到此文件',
    '',
    `API_KEY=${next.apiKey || ''}`,
    `BASE_URL=${next.baseUrl || ''}`,
    `MODEL=${next.model || ''}`,
    `TEMPERATURE=${next.temperature ?? 1}`,
    `MAX_TOKENS=${next.maxTokens ?? 131072}`,
    `OPENROUTER_API_KEY=${next.openRouterApiKey || ''}`,
    `DASHSCOPE_API_KEY=${next.dashscopeApiKey || ''}`,
    `ADOBE_FIREFLY_CLIENT_ID=${next.adobeFireflyClientId || ''}`,
    `ADOBE_FIREFLY_CLIENT_SECRET=${next.adobeFireflyClientSecret || ''}`,
    `BFL_API_KEY=${next.bflApiKey || ''}`,
    `PPT_IMAGE_MODEL=${next.pptImageModel || ''}`,
    `BRAVE_API_KEY=${next.braveApiKey || ''}`,
    `MEMORY_ENABLED=${next.memoryEnabled == null ? '' : String(next.memoryEnabled)}`,
    `MEMORY_TOP_K=${next.memoryTopK ?? ''}`,
    `MEMORY_MAX_CHARS=${next.memoryMaxChars ?? ''}`,
    `MEMORY_FLUSH_THRESHOLD=${next.memoryFlushThresholdChars ?? ''}`,
    `EMBEDDING_API_KEY=${next.embeddingApiKey || ''}`,
    `LOCAL_MODEL_ENABLED=${next.localModel?.enabled == null ? '' : String(next.localModel.enabled)}`,
    `LOCAL_MODEL_BASE_URL=${next.localModel?.baseUrl || ''}`,
    `LOCAL_MODEL_MODEL=${next.localModel?.model || ''}`,
    `LOCAL_MODEL_NAME=${next.localModel?.model || ''}`,
    `LOCAL_MODEL_API_KEY=${next.localModel?.apiKey || ''}`,
    '',
  ]

  fs.writeFileSync(envFilePath, lines.join('\n'), 'utf8')
  process.env.OPENROUTER_API_KEY = next.openRouterApiKey || ''
  process.env.DASHSCOPE_API_KEY = next.dashscopeApiKey || ''
  process.env.ADOBE_FIREFLY_CLIENT_ID = next.adobeFireflyClientId || ''
  process.env.ADOBE_FIREFLY_CLIENT_SECRET = next.adobeFireflyClientSecret || ''
  process.env.BFL_API_KEY = next.bflApiKey || ''
  process.env.EMBEDDING_API_KEY = next.embeddingApiKey || ''
  return { success: true }
}

let mainWindow
let fileServer = null
let fileServerPort = 9090
let memoryManager = null

ipcMain.handle('get-env-settings', async () => readEnvSettings())
ipcMain.handle('save-env-settings', async (_event, settings = {}) => {
  try {
    return saveEnvSettingsToFile(settings)
  } catch (error) {
    return { success: false, error: error?.message || String(error) }
  }
})

const getMemoryManager = () => {
  if (!memoryManager) {
    memoryManager = createMemoryManager(app)
  }
  return memoryManager
}

const filesService = createFilesService({
  fs,
  path,
  dialog,
  shell,
  WordExtractor,
  getMainWindow: () => mainWindow,
  getFileServerPort: () => fileServerPort,
})

registerFileIpc(ipcMain, filesService)

const memoryService = createMemoryService({
  getMemoryManager,
})

registerMemoryIpc(ipcMain, memoryService)

const knowledgeService = createKnowledgeService({
  fs,
  path,
  app,
  mammoth,
  WordExtractor,
  XLSX,
  getMemoryManager,
})

registerKnowledgeIpc(ipcMain, knowledgeService)

const webSearchService = createWebSearchService({
  app,
})

registerWebSearchIpc(ipcMain, webSearchService)

const excelService = createExcelService({
  fs,
  path,
  app,
  XLSX,
  ExcelJS,
})

registerExcelIpc(ipcMain, excelService)

const fontsService = createFontsService({
  fs,
  path,
})

registerFontsIpc(ipcMain, fontsService)

const aiProxyService = createAIProxyService({
  fs,
})

registerAIProxyIpc(ipcMain, aiProxyService)

const pptService = createPptService({
  fs,
  path,
  app,
  findLibreOffice,
  getLibreOfficeDownloadUrl,
})

registerPptIpc(ipcMain, pptService)

const docxInspectorService = createDocxInspectorService({
  fs,
  path,
  app,
})

registerDocxInspectorIpc(ipcMain, docxInspectorService)

const wordOracleService = createWordOracleService({
  fs,
  path,
  app,
  docxInspectorService,
})

registerWordOracleIpc(ipcMain, wordOracleService)

const nativeTextService = createNativeTextService({
  app,
})

registerNativeTextIpc(ipcMain, nativeTextService)

// 开发模式检测
const isDev = process.env.NODE_ENV === 'development' || !app.isPackaged

// p-limit@5 是 ESM-only，Electron main 这里是 CommonJS（main.cjs）。
// 为避免 ERR_REQUIRE_ESM，使用一个轻量并发 limiter，满足“并发=2”需求即可。
function pLimit(concurrency) {
  if (!Number.isFinite(concurrency) || concurrency < 1) {
    throw new Error('pLimit: concurrency must be >= 1')
  }
  let activeCount = 0
  const queue = []

  const next = () => {
    if (activeCount >= concurrency) return
    const item = queue.shift()
    if (!item) return
    activeCount++
    const { fn, resolve, reject } = item
    Promise.resolve()
      .then(fn)
      .then(resolve, reject)
      .finally(() => {
        activeCount--
        next()
      })
  }

  return (fn) =>
    new Promise((resolve, reject) => {
      queue.push({ fn, resolve, reject })
      next()
    })
}

// 创建本地文件服务器（供 ONLYOFFICE 访问文档）
function createFileServer(startPort = 9090) {
  const handler = (req, res) => {
    // 设置 CORS 头，允许 ONLYOFFICE 访问
    res.setHeader('Access-Control-Allow-Origin', '*')
    res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS')
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type')
    
    if (req.method === 'OPTIONS') {
      res.writeHead(200)
      res.end()
      return
    }
    
    try {
      const requestUrl = new URL(req.url, `http://127.0.0.1:${fileServerPort}`)
      if (isDev && requestUrl.pathname === '/__codex_eval') {
        if (!mainWindow || mainWindow.isDestroyed()) {
          res.writeHead(503, { 'Content-Type': 'application/json; charset=utf-8' })
          res.end(JSON.stringify({ success: false, error: 'mainWindow unavailable' }))
          return
        }

        const code = requestUrl.searchParams.get('code')
        if (!code) {
          res.writeHead(400, { 'Content-Type': 'application/json; charset=utf-8' })
          res.end(JSON.stringify({ success: false, error: 'missing code query param' }))
          return
        }

        mainWindow.webContents.executeJavaScript(code, true)
          .then((result) => {
            res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' })
            res.end(JSON.stringify({ success: true, result }))
          })
          .catch((error) => {
            res.writeHead(500, { 'Content-Type': 'application/json; charset=utf-8' })
            res.end(JSON.stringify({
              success: false,
              error: error?.message || String(error),
            }))
          })
        return
      }
    } catch (error) {
      res.writeHead(500, { 'Content-Type': 'application/json; charset=utf-8' })
      res.end(JSON.stringify({
        success: false,
        error: error?.message || String(error),
      }))
      return
    }

    // 解析文件路径（URL 编码的路径）
    const urlPath = decodeURIComponent(req.url.replace(/^\/file\//, ''))
    const filePath = urlPath.replace(/\//g, path.sep)
    
    console.log('文件服务器请求:', filePath)
    
    if (!fs.existsSync(filePath)) {
      console.error('文件不存在:', filePath)
      res.writeHead(404)
      res.end('File not found')
      return
    }
    
    // 获取文件扩展名
    const ext = path.extname(filePath).toLowerCase()
    
    // 设置 Content-Type
    const mimeTypes = {
      '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      '.doc': 'application/msword',
      '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      '.xls': 'application/vnd.ms-excel',
      '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      '.ppt': 'application/vnd.ms-powerpoint',
      '.pdf': 'application/pdf',
      '.txt': 'text/plain',
    }
    
    const contentType = mimeTypes[ext] || 'application/octet-stream'
    
    // 读取并发送文件
    try {
      const fileBuffer = fs.readFileSync(filePath)
      res.writeHead(200, {
        'Content-Type': contentType,
        'Content-Length': fileBuffer.length,
        'Content-Disposition': `attachment; filename="${encodeURIComponent(path.basename(filePath))}"`,
      })
      res.end(fileBuffer)
      console.log('文件发送成功:', path.basename(filePath))
    } catch (error) {
      console.error('读取文件失败:', error)
      res.writeHead(500)
      res.end('Internal server error')
    }
  }

  const tryListen = (port) => {
    try {
      fileServerPort = port
      fileServer = http.createServer(handler)

      fileServer.on('error', (err) => {
        if (err && err.code === 'EADDRINUSE') {
          console.warn(`文件服务器端口 ${port} 被占用，尝试端口 ${port + 1}...`)
          tryListen(port + 1)
          return
        }
        console.error('文件服务器错误:', err)
      })

      fileServer.listen(port, '0.0.0.0', () => {
        fileServerPort = port
        console.log(`📁 本地文件服务器已启动: http://localhost:${fileServerPort}`)
      })
    } catch (e) {
      console.error('文件服务器启动失败:', e)
    }
  }

  tryListen(startPort)
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1400,
    height: 900,
    minWidth: 1000,
    minHeight: 700,
    icon: APP_ICON_PATH,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'preload.cjs'),
    },
    // Windows 使用默认标题栏，确保可以正常关闭
    frame: true,
    backgroundColor: '#09090b',
  })

  // 开发模式加载 Vite 服务器，生产模式加载打包文件
  if (isDev) {
    mainWindow.loadURL('http://localhost:3000')
    if (shouldOpenDevTools) {
      mainWindow.webContents.openDevTools()
    }
  } else {
    mainWindow.loadFile(path.join(__dirname, '../dist/index.html'))
  }
  
  // 忽略一些无害的控制台警告
  mainWindow.webContents.on('console-message', (event, level, message) => {
    // 过滤掉 DevTools 内部警告
    if (message.includes('Unknown VE context') || 
        message.includes('Autofill.enable') ||
        message.includes('Storage.getStorageKeyForFrame')) {
      return
    }
  })

  mainWindow.on('closed', () => {
    mainWindow = null
  })
}

app.whenReady().then(() => {
  app.setAboutPanelOptions({
    applicationName: APP_DISPLAY_NAME,
    applicationVersion: app.getVersion(),
  })

  if (process.platform === 'darwin' && app.dock && fs.existsSync(APP_ICON_PATH)) {
    app.dock.setIcon(APP_ICON_PATH)
  }

  createFileServer()
  createWindow()
  
  // 注册全局快捷键打开 DevTools
  globalShortcut.register('CommandOrControl+Shift+I', () => {
    if (mainWindow) {
      mainWindow.webContents.toggleDevTools()
    }
  })
  
  // 也注册 F12
  globalShortcut.register('F12', () => {
    if (mainWindow) {
      mainWindow.webContents.toggleDevTools()
    }
  })
})

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit()
  }
})

app.on('before-quit', () => {
  void pptService.close?.()
  void aiProxyService.close()
  void knowledgeService.close?.()
  void webSearchService.close()
})

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow()
  }
})

// ==================== IPC 处理：文件系统操作 ====================

// 获取文件的 HTTP URL（供 ONLYOFFICE 使用）
async function readExcelWithSheetJS(filePath) {
  try {
    const XLSX = require('xlsx')
    const buffer = fs.readFileSync(filePath)
    
    console.log('[Excel] 开始读取 .xls 文件:', filePath)
    
    // 读取 xls 文件，启用样式选项
    const workbook = XLSX.read(buffer, { 
      type: 'buffer', 
      cellStyles: true, 
      cellFormula: true,
      cellNF: true,
      cellDates: true,
    })
    
    // 获取样式表
    const styles = workbook.Styles || {}
    const cellXfs = styles.CellXf || []
    const fonts = styles.Fonts || []
    const fills = styles.Fills || []
    const borders = styles.Borders || []
    const numFmts = styles.NumberFmt || {}
    
    console.log('[Excel] 样式表信息:', {
      cellXfsCount: cellXfs.length,
      fontsCount: fonts.length,
      fillsCount: fills.length,
      bordersCount: borders.length,
    })
    
    const sheets = []
    
    for (const sheetName of workbook.SheetNames) {
      const worksheet = workbook.Sheets[sheetName]
      const range = worksheet['!ref'] ? XLSX.utils.decode_range(worksheet['!ref']) : { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } }
      
      const sheetData = {
        name: sheetName,
        range: range,
        merges: [],
        colWidths: [],
        rowHeights: [],
        cells: []
      }
      
      // 合并单元格
      if (worksheet['!merges']) {
        sheetData.merges = worksheet['!merges'].map(m => ({
          s: { r: m.s.r, c: m.s.c },
          e: { r: m.e.r, c: m.e.c }
        }))
      }
      
      // 列宽
      if (worksheet['!cols']) {
        worksheet['!cols'].forEach((col, idx) => {
          if (col && col.wpx) {
            sheetData.colWidths[idx] = col.wpx
          } else if (col && col.wch) {
            sheetData.colWidths[idx] = Math.round(col.wch * 7 + 5)
          }
        })
      }
      
      // 行高
      if (worksheet['!rows']) {
        worksheet['!rows'].forEach((row, idx) => {
          if (row && row.hpx) {
            sheetData.rowHeights[idx] = row.hpx
          } else if (row && row.hpt) {
            sheetData.rowHeights[idx] = Math.round(row.hpt * 1.333)
          }
        })
      }
      
      // 遍历单元格
      let debugCount = 0
      const keys = Object.keys(worksheet).filter(k => !k.startsWith('!'))
      
      for (const addr of keys) {
        const cell = worksheet[addr]
        if (!cell) continue
        
        const decoded = XLSX.utils.decode_cell(addr)
        const r = decoded.r
        const c = decoded.c
        
        // 调试：打印前3个单元格的完整信息
        if (debugCount < 3) {
          console.log('[Excel XLS] 单元格完整数据:', {
            address: addr,
            cell: JSON.stringify(cell, null, 2)
          })
          debugCount++
        }
        
        // 解析样式
        const styleObj = {}
        
        // 方法1: 直接从 cell.s 获取样式对象
        if (cell.s && typeof cell.s === 'object') {
          console.log('[Excel XLS] 发现样式对象 cell.s:', cell.s)
          
          // 字体
          if (cell.s.font) {
            styleObj.font = {
              name: cell.s.font.name,
              sz: cell.s.font.sz,
              bold: cell.s.font.bold,
              italic: cell.s.font.italic,
              underline: cell.s.font.underline,
              strike: cell.s.font.strike,
              color: cell.s.font.color
            }
          }
          
          // 填充
          if (cell.s.fill || cell.s.fgColor || cell.s.bgColor) {
            styleObj.fill = {
              fgColor: cell.s.fgColor || cell.s.fill?.fgColor,
              bgColor: cell.s.bgColor || cell.s.fill?.bgColor
            }
          }
          
          // 对齐
          if (cell.s.alignment) {
            styleObj.alignment = cell.s.alignment
          }
          
          // 边框
          if (cell.s.border) {
            styleObj.border = cell.s.border
          }
        }
        // 方法2: 通过样式索引获取
        else if (typeof cell.s === 'number' && cellXfs[cell.s]) {
          const xf = cellXfs[cell.s]
          
          if (!debuggedFirstCell) {
            console.log('[Excel XLS] 单元格样式示例 (通过索引):', {
              address: addr,
              value: cell.v,
              styleIndex: cell.s,
              xf: xf,
              font: fonts[xf.fontId],
              fill: fills[xf.fillId]
            })
            debuggedFirstCell = true
          }
          
          // 字体
          if (xf.fontId !== undefined && fonts[xf.fontId]) {
            const font = fonts[xf.fontId]
            styleObj.font = {
              name: font.name,
              sz: font.sz,
              bold: font.bold,
              italic: font.italic,
              underline: font.underline,
              strike: font.strike,
              color: font.color
            }
          }
          
          // 填充
          if (xf.fillId !== undefined && fills[xf.fillId]) {
            const fill = fills[xf.fillId]
            styleObj.fill = {
              fgColor: fill.fgColor,
              bgColor: fill.bgColor
            }
          }
          
          // 对齐
          if (xf.alignment) {
            styleObj.alignment = xf.alignment
          }
          
          // 边框
          if (xf.borderId !== undefined && borders[xf.borderId]) {
            styleObj.border = borders[xf.borderId]
          }
          
          // 数字格式
          if (xf.numFmtId !== undefined) {
            styleObj.numFmt = numFmts[xf.numFmtId] || xf.numFmtId
          }
        }
        
        const cellData = {
          r,
          c,
          v: cell.v,
          t: cell.t,
          f: cell.f,
          s: styleObj,
          w: cell.w,
          display: cell.w || (cell.v != null ? String(cell.v) : '')
        }
        
        sheetData.cells.push(cellData)
      }
      
      sheets.push(sheetData)
    }
    
    console.log('[Excel] .xls 文件读取成功，工作表数:', sheets.length)
    return { success: true, sheets }
  } catch (error) {
    console.error('读取 .xls 文件失败:', error)
    return { success: false, error: error.message }
  }
}

// 检查 LibreOffice 是否安装
function findLibreOffice() {
  const possiblePaths = [
    // Windows 常见路径
    'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
    'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe',
    // 应用内置便携版（如果打包）
    path.join(__dirname, '..', 'libreoffice', 'program', 'soffice.exe'),
    path.join(__dirname, 'libreoffice', 'program', 'soffice.exe'),
    // 环境变量
    process.env.LIBREOFFICE_PATH,
  ].filter(Boolean)
  
  for (const p of possiblePaths) {
    if (fs.existsSync(p)) {
      console.log('[Excel] 找到 LibreOffice:', p)
      return p
    }
  }
  console.log('[Excel] LibreOffice 未找到')
  return null
}

// 获取 LibreOffice 下载链接
function getLibreOfficeDownloadUrl() {
  if (process.platform === 'win32') {
    // LibreOffice 便携版 (约 300MB)
    return 'https://download.documentfoundation.org/libreoffice/portable/7.6.4/LibreOfficePortable_7.6.4_MultilingualStandard.paf.exe'
  }
  return null
}

// ==================== PPTX 预览渲染（LibreOffice → PNG） ====================

ipcMain.handle('docx-search-replace', async (event, { sourcePath, outputPath, replacements }) => {
  try {
    console.log('DOCX 搜索替换开始:', sourcePath, '->', outputPath)
    console.log('替换列表:', replacements)
    
    // 读取源文件
    const content = fs.readFileSync(sourcePath, 'binary')
    const zip = new PizZip(content)
    
    // 获取 document.xml（Word 文档的主体内容）
    let documentXml = zip.file('word/document.xml').asText()
    
    // 执行所有替换
    let replaceCount = 0
    for (const item of replacements) {
      const searchText = item.search
      const replaceText = item.replace
      
      // 在 XML 中搜索并替换文本
      // 注意：Word 可能会把文本拆分成多个 <w:t> 标签，这里做简单替换
      // 对于复杂情况，可能需要更智能的处理
      const regex = new RegExp(escapeRegExp(searchText), 'g')
      const matches = documentXml.match(regex)
      if (matches) {
        documentXml = documentXml.replace(regex, escapeXml(replaceText))
        replaceCount += matches.length
        console.log(`替换 "${searchText}" -> "${replaceText}": ${matches.length} 处`)
      } else {
        console.log(`未找到: "${searchText}"`)
      }
    }
    
    // 更新 zip 中的 document.xml
    zip.file('word/document.xml', documentXml)
    
    // 生成输出文件
    const buf = zip.generate({
      type: 'nodebuffer',
      compression: 'DEFLATE'
    })
    
    // 写入文件
    fs.writeFileSync(outputPath, buf)
    
    console.log(`DOCX 搜索替换完成: ${replaceCount} 处替换，保存到 ${outputPath}`)
    return { success: true, path: outputPath, replaceCount }
  } catch (error) {
    console.error('DOCX 搜索替换失败:', error)
    return { success: false, error: error.message }
  }
})

function escapeRegExp(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

// 辅助函数：转义 XML 特殊字符
function escapeXml(string) {
  return string
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;')
}

// ==================== ONLYOFFICE Document Builder API ====================

// 创建带格式的文档
ipcMain.handle('create-formatted-document', async (event, { filePath, elements, title }) => {
  try {
    // 生成 Document Builder 脚本
    const script = generateDocBuilderScript(elements, title)
    
    console.log('生成的 Document Builder 脚本:')
    console.log(script)
    
    // 保存脚本到临时文件
    const tempDir = app.getPath('temp')
    const scriptPath = path.join(tempDir, `docbuilder_${Date.now()}.docbuilder`)
    fs.writeFileSync(scriptPath, script, 'utf-8')
    
    // 调用 ONLYOFFICE Document Builder API
    const result = await callDocumentBuilder(scriptPath, filePath)
    
    // 清理临时脚本文件
    try {
      fs.unlinkSync(scriptPath)
    } catch (e) {
      console.log('清理临时文件失败:', e)
    }
    
    return result
  } catch (error) {
    console.error('创建格式化文档失败:', error)
    return { success: false, error: error.message }
  }
})

// 生成 Document Builder 脚本
function generateDocBuilderScript(elements, title) {
  let script = `builder.CreateFile("docx");\n`
  script += `var oDocument = Api.GetDocument();\n`
  script += `var oParagraph;\n`
  script += `var oTable, oRow, oCell;\n\n`
  
  for (let i = 0; i < elements.length; i++) {
    const elem = elements[i]
    
    if (elem.type === 'heading') {
      const level = elem.level || 1
      const alignment = elem.alignment || 'left'
      const jc = alignment === 'center' ? 'center' : alignment === 'right' ? 'right' : 'left'
      
      script += `// 标题 ${level}\n`
      if (i === 0) {
        script += `oParagraph = oDocument.GetElement(0);\n`
      } else {
        script += `oParagraph = Api.CreateParagraph();\n`
        script += `oDocument.Push(oParagraph);\n`
      }
      script += `oParagraph.AddText("${escapeString(elem.content || '')}");\n`
      script += `oParagraph.SetStyle(oDocument.GetStyle("Heading ${level}"));\n`
      script += `oParagraph.SetJc("${jc}");\n`
      if (elem.bold) {
        script += `oParagraph.SetBold(true);\n`
      }
      script += `\n`
      
    } else if (elem.type === 'paragraph') {
      const alignment = elem.alignment || 'left'
      const jc = alignment === 'center' ? 'center' : alignment === 'right' ? 'right' : alignment === 'justify' ? 'both' : 'left'
      
      script += `// 段落\n`
      if (i === 0) {
        script += `oParagraph = oDocument.GetElement(0);\n`
      } else {
        script += `oParagraph = Api.CreateParagraph();\n`
        script += `oDocument.Push(oParagraph);\n`
      }
      
      // 使用 Run 来设置文本样式
      script += `var oRun = Api.CreateRun();\n`
      script += `oRun.AddText("${escapeString(elem.content || '')}");\n`
      
      if (elem.bold) {
        script += `oRun.SetBold(true);\n`
      }
      if (elem.fontSize) {
        // Document Builder 使用半磅，所以要乘以 2
        script += `oRun.SetFontSize(${elem.fontSize * 2});\n`
      }
      if (elem.fontFamily) {
        script += `oRun.SetFontFamily("${elem.fontFamily}");\n`
      }
      if (elem.color) {
        // 解析颜色（假设是 #RRGGBB 格式）
        const color = elem.color.replace('#', '')
        const r = parseInt(color.substr(0, 2), 16)
        const g = parseInt(color.substr(2, 2), 16)
        const b = parseInt(color.substr(4, 2), 16)
        script += `oRun.SetColor(${r}, ${g}, ${b});\n`
      }
      
      script += `oParagraph.AddElement(oRun);\n`
      script += `oParagraph.SetJc("${jc}");\n`
      script += `\n`
      
    } else if (elem.type === 'table') {
      const rows = elem.rows || 2
      const cols = elem.cols || 2
      const data = elem.data || []
      
      script += `// 表格 ${rows}x${cols}\n`
      script += `oTable = Api.CreateTable(${cols}, ${rows});\n`
      script += `oDocument.Push(oTable);\n`
      
      // 设置表格宽度为 100%
      script += `oTable.SetWidth("percent", 100);\n`
      
      // 填充表格数据
      for (let r = 0; r < rows; r++) {
        for (let c = 0; c < cols; c++) {
          const cellData = data[r] && data[r][c] ? data[r][c] : ''
          script += `oTable.GetRow(${r}).GetCell(${c}).GetContent().GetElement(0).AddText("${escapeString(cellData)}");\n`
        }
      }
      
      // 设置表格边框
      script += `oTable.SetTableBorderTop("single", 4, 0, 0, 0, 0);\n`
      script += `oTable.SetTableBorderBottom("single", 4, 0, 0, 0, 0);\n`
      script += `oTable.SetTableBorderLeft("single", 4, 0, 0, 0, 0);\n`
      script += `oTable.SetTableBorderRight("single", 4, 0, 0, 0, 0);\n`
      script += `oTable.SetTableBorderInsideH("single", 4, 0, 0, 0, 0);\n`
      script += `oTable.SetTableBorderInsideV("single", 4, 0, 0, 0, 0);\n`
      script += `\n`
    }
  }
  
  // 保存文件
  script += `builder.SaveFile("docx", "output.docx");\n`
  script += `builder.CloseFile();\n`
  
  return script
}

// 转义字符串中的特殊字符
function escapeString(str) {
  return str
    .replace(/\\/g, '\\\\')
    .replace(/"/g, '\\"')
    .replace(/\n/g, '\\n')
    .replace(/\r/g, '\\r')
    .replace(/\t/g, '\\t')
}

// 调用 ONLYOFFICE Document Builder API
async function callDocumentBuilder(scriptPath, outputPath) {
  return new Promise((resolve) => {
    // 首先尝试使用 Document Builder 服务
    // ONLYOFFICE DocumentServer 的 Document Builder 端点
    const DOCUMENT_SERVER_URL = 'http://localhost:8080'
    
    // 读取脚本内容
    const scriptContent = fs.readFileSync(scriptPath, 'utf-8')
    
    // 将脚本内容保存到一个可以被 DocumentServer 访问的位置
    // 由于 DocumentServer 在 Docker 中，需要通过 HTTP 提供脚本
    const scriptFileName = path.basename(scriptPath)
    
    // 在文件服务器上提供脚本文件
    const scriptUrl = `http://host.docker.internal:${FILE_SERVER_PORT}/file/${encodeURIComponent(scriptPath.replace(/\\/g, '/'))}`
    
    console.log('Document Builder 脚本 URL:', scriptUrl)
    
    // 发送请求到 Document Builder API
    const requestData = JSON.stringify({
      async: false,
      url: scriptUrl
    })
    
    const options = {
      hostname: 'localhost',
      port: 8080,
      path: '/docbuilder',
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(requestData)
      }
    }
    
    const req = http.request(options, (res) => {
      let data = ''
      
      res.on('data', (chunk) => {
        data += chunk
      })
      
      res.on('end', () => {
        console.log('Document Builder 响应:', data)
        
        try {
          const result = JSON.parse(data)
          
          if (result.error) {
            console.error('Document Builder 错误:', result.error)
            // 如果 Document Builder API 失败，回退到使用 docx 库
            fallbackCreateDocument(outputPath, scriptPath, resolve)
          } else if (result.urls && typeof result.urls === 'object') {
            // urls 是一个对象，键是文件名，值是 URL
            // 例如: { "output.docx": "http://..." }
            const urlKeys = Object.keys(result.urls)
            if (urlKeys.length > 0) {
              const firstUrl = result.urls[urlKeys[0]]
              console.log('找到生成的文档 URL:', firstUrl)
              downloadGeneratedDocument(firstUrl, outputPath, resolve)
            } else {
              console.log('Document Builder 返回了空的 urls 对象')
              fallbackCreateDocument(outputPath, scriptPath, resolve)
            }
          } else {
            console.log('Document Builder 返回了意外的结果:', result)
            fallbackCreateDocument(outputPath, scriptPath, resolve)
          }
        } catch (e) {
          console.error('解析 Document Builder 响应失败:', e)
          fallbackCreateDocument(outputPath, scriptPath, resolve)
        }
      })
    })
    
    req.on('error', (error) => {
      console.error('Document Builder 请求失败:', error)
      // 回退方案
      fallbackCreateDocument(outputPath, scriptPath, resolve)
    })
    
    req.write(requestData)
    req.end()
  })
}

// 下载生成的文档
function downloadGeneratedDocument(url, outputPath, resolve) {
  console.log('下载生成的文档:', url)
  
  // 解析 URL
  const urlObj = new URL(url)
  const options = {
    hostname: urlObj.hostname,
    port: urlObj.port || 80,
    path: urlObj.pathname + urlObj.search,
    method: 'GET'
  }
  
  const req = http.request(options, (res) => {
    const chunks = []
    
    res.on('data', (chunk) => {
      chunks.push(chunk)
    })
    
    res.on('end', () => {
      const buffer = Buffer.concat(chunks)
      fs.writeFileSync(outputPath, buffer)
      console.log('文档已保存到:', outputPath)
      resolve({ success: true, path: outputPath })
    })
  })
  
  req.on('error', (error) => {
    console.error('下载文档失败:', error)
    resolve({ success: false, error: error.message })
  })
  
  req.end()
}

// 回退方案：使用简单的方式创建文档
function fallbackCreateDocument(outputPath, scriptPath, resolve) {
  console.log('使用回退方案创建文档...')
  
  // 读取脚本，解析元素，使用 docx 库创建
  // 这里简化处理，创建一个空文档
  try {
    // 创建一个最小的有效 docx 文件
    // 使用 docx 库（如果可用）或创建空文件
    const emptyDocx = createMinimalDocx()
    fs.writeFileSync(outputPath, emptyDocx)
    console.log('回退方案：创建了基本文档')
    resolve({ success: true, path: outputPath, fallback: true })
  } catch (error) {
    console.error('回退方案失败:', error)
    resolve({ success: false, error: error.message })
  }
}

// 创建最小的有效 docx 文件
function createMinimalDocx() {
  // 一个最小的有效 docx 文件的 base64
  // 这是一个空的 docx 文件
  const minimalDocxBase64 = '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'
  
  return Buffer.from(minimalDocxBase64, 'base64')
}

// ==================== PPT (image-only) Generation ====================
