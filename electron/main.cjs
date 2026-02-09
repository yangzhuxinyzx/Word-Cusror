const { app, BrowserWindow, ipcMain, dialog, shell, globalShortcut } = require('electron')
const path = require('path')
const fs = require('fs')
const { createMemoryManager } = require('./memory/index.cjs')
const http = require('http')
const https = require('https')
const mammoth = require('mammoth')
const WordExtractor = require('word-extractor')
const PizZip = require('pizzip')
const Docxtemplater = require('docxtemplater')
const dotenv = require('dotenv')
const XLSX = require('xlsx')
const ExcelJS = require('exceljs')
const sharp = require('sharp')
const PptxGenJS = require('pptxgenjs')
const crypto = require('crypto')
const sdkBase = path.join(__dirname, '..', 'node_modules', '@modelcontextprotocol', 'sdk', 'dist', 'cjs')
const { Client: McpClient } = require(path.join(sdkBase, 'client', 'index.js'))
const { InMemoryTransport } = require(path.join(sdkBase, 'inMemory.js'))

dotenv.config({ path: path.join(__dirname, '..', '.env') })

let mainWindow
let fileServer = null
let fileServerPort = 9090
let memoryManager = null

const getMemoryManager = () => {
  if (!memoryManager) {
    memoryManager = createMemoryManager(app)
  }
  return memoryManager
}

// 开发模式检测
const isDev = process.env.NODE_ENV === 'development' || !app.isPackaged
const DEFAULT_RESULT_FILTER = ['web', 'query', 'faq', 'news', 'videos', 'discussions']
const braveServerModulePromise = import('@brave/brave-search-mcp-server/dist/server.js')

let braveMcpConnection = null
let braveMcpInitPromise = null
let braveMcpApiKey = null // 记录当前使用的 API Key

const COUNTRY_CODES = new Set([
  'ALL','AR','AU','AT','BE','BR','CA','CL','DK','FI','FR','DE','HK','IN','ID','IT','JP','KR','MY','MX','NL','NZ','NO',
  'CN','PL','PT','PH','RU','SA','ZA','ES','SE','CH','TW','TR','GB','US'
])

const UI_LANG_OPTIONS = new Set([
  'es-AR','en-AU','de-AT','nl-BE','fr-BE','pt-BR','en-CA','fr-CA','es-CL','da-DK','fi-FI','fr-FR','de-DE','el-GR','zh-HK',
  'en-IN','en-ID','it-IT','ja-JP','ko-KR','en-MY','es-MX','nl-NL','en-NZ','no-NO','zh-CN','pl-PL','en-PH','ru-RU','en-ZA',
  'es-ES','sv-SE','fr-CH','de-CH','zh-TW','tr-TR','en-GB','en-US','es-US'
])

const SEARCH_LANG_OPTIONS = new Set([
  'ar','eu','bn','bg','ca','zh-hans','zh-hant','hr','cs','da','nl','en','en-gb','et','fi','fr','gl','de','el','gu','he','hi',
  'hu','is','it','jp','kn','ko','lv','lt','ms','ml','mr','nb','pl','pt-br','pt-pt','pa','ro','ru','sr','sk','sl','es','sv',
  'ta','te','th','tr','uk','vi'
])

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

async function ensureBraveMcpClient(apiKeyOverride) {
  // 获取要使用的 API Key（优先使用传入的，其次是环境变量）
  const apiKey = apiKeyOverride || process.env.BRAVE_API_KEY
  
  // 如果 API Key 变化了，需要重新初始化
  if (braveMcpConnection && braveMcpApiKey === apiKey) {
    return braveMcpConnection
  }
  
  // 如果有旧连接且 API Key 变了，关闭旧连接
  if (braveMcpConnection && braveMcpApiKey !== apiKey) {
    try {
      braveMcpConnection.client?.close?.()
      braveMcpConnection.server?.close?.()
    } catch {}
    braveMcpConnection = null
    braveMcpInitPromise = null
  }
  
  if (braveMcpInitPromise) return braveMcpInitPromise

  braveMcpInitPromise = (async () => {
    if (!apiKey) {
      throw new Error('请在设置中配置 Brave Search API Key，或在 .env 中配置 BRAVE_API_KEY')
    }

    const serverModule = await braveServerModulePromise
    const createServer = serverModule?.default || serverModule
    const server = createServer({ config: { braveApiKey: apiKey } })

    const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair()
    await server.connect(serverTransport)

    const client = new McpClient({
      name: 'word-cursor',
      version: app?.getVersion?.() || 'dev'
    })
    await client.connect(clientTransport)
    await client.listTools({})

    braveMcpConnection = { client, server }
    braveMcpApiKey = apiKey // 记录使用的 API Key
    return braveMcpConnection
  })().catch((error) => {
    braveMcpInitPromise = null
    throw error
  })

  return braveMcpInitPromise
}

function buildBraveWebArguments(query, options = {}) {
  const count = Math.max(1, Math.min(parseInt(options.num ?? 5, 10) || 5, 20))
  const args = {
    query,
    count,
    safesearch: 'moderate',
    spellcheck: true,
    text_decorations: true,
    summary: false,
    extra_snippets: true,
    result_filter: ['web', 'news', 'faq', 'videos', 'discussions']
  }

  const locale = typeof options.locale === 'string' ? options.locale.trim() : ''
  const region = typeof options.region === 'string' ? options.region.trim() : ''

  const uiLang = normalizeUiLang(locale) || 'en-US'
  if (uiLang) {
    args.ui_lang = uiLang
  }

  const searchLang = normalizeSearchLang(locale) || 'en'
  if (searchLang) {
    args.search_lang = searchLang
  }

  const country = normalizeCountry(region || (uiLang?.split('-')[1] || 'US'))
  if (country) {
    args.country = country
  }

  return args
}

function normalizeUiLang(locale) {
  if (!locale) return null
  const normalized = locale.replace('_', '-')
  const [lang, region] = normalized.split('-')
  if (!lang) return null
  const candidate = region ? `${lang.toLowerCase()}-${region.toUpperCase()}` : `${lang.toLowerCase()}`
  if (UI_LANG_OPTIONS.has(candidate)) {
    return candidate
  }
  if (region) {
    const fallback = `${lang.toLowerCase()}-${region.toUpperCase()}`
    if (UI_LANG_OPTIONS.has(fallback)) return fallback
  }
  return UI_LANG_OPTIONS.has('en-US') ? 'en-US' : null
}

function normalizeSearchLang(locale) {
  if (!locale) return null
  const normalized = locale.toLowerCase()
  if (SEARCH_LANG_OPTIONS.has(normalized)) return normalized

  if (normalized.startsWith('zh')) {
    return normalized.includes('tw') || normalized.includes('hk') ? 'zh-hant' : 'zh-hans'
  }

  const base = normalized.split(/[-_]/)[0]
  if (SEARCH_LANG_OPTIONS.has(base)) return base
  return 'en'
}

function normalizeCountry(region) {
  if (!region) return 'US'
  const upper = region.toUpperCase()
  return COUNTRY_CODES.has(upper) ? upper : 'US'
}

function transformBraveContent(content = [], maxWebCount = 5) {
  const sections = {
    web: [],
    faq: [],
    news: [],
    videos: [],
    discussions: []
  }
  let summarizerKey = null

  for (const block of content || []) {
    if (!block || block.type !== 'text' || !block.text) continue
    const textBlock = block.text.trim()
    if (!textBlock) continue

    if (textBlock.startsWith('Summarizer key:')) {
      summarizerKey = textBlock.split(':').slice(1).join(':').trim()
      continue
    }

    let data
    try {
      data = JSON.parse(textBlock)
    } catch (error) {
      continue
    }

    if (isFaqResult(data)) {
      sections.faq.push({
        question: data.question,
        answer: data.answer,
        title: data.title,
        link: data.url
      })
      continue
    }

    if (isNewsResult(data)) {
      sections.news.push({
        title: data.title,
        link: data.url,
        source: data.source,
        description: data.description,
        breaking: Boolean(data.breaking),
        isLive: Boolean(data.is_live),
        age: data.age
      })
      continue
    }

    if (isVideoResult(data)) {
      sections.videos.push({
        title: data.title,
        link: data.url,
        description: data.description,
        duration: data.duration,
        thumbnail: data.thumbnail_url,
        viewCount: data.view_count,
        creator: data.creator,
        publisher: data.publisher
      })
      continue
    }

    if (isDiscussionResult(data)) {
      sections.discussions.push({
        link: data.url,
        forumName: data.data?.forum_name,
        question: data.data?.question,
        topComment: data.data?.top_comment
      })
      continue
    }

    if (isWebResult(data)) {
      sections.web.push({
        title: data.title || '未命名结果',
        link: data.url || '',
        snippet: data.description || '',
        extraSnippets: Array.isArray(data.extra_snippets) ? data.extra_snippets : undefined
      })
      continue
    }
  }

  sections.web = sections.web.slice(0, maxWebCount)

  return {
    sections,
    summarizerKey
  }
}

function isFaqResult(data) {
  return data && typeof data.question === 'string' && typeof data.answer === 'string'
}

function isNewsResult(data) {
  return data && typeof data.source === 'string' && Object.prototype.hasOwnProperty.call(data, 'breaking')
}

function isVideoResult(data) {
  return data && (Object.prototype.hasOwnProperty.call(data, 'thumbnail_url') || Object.prototype.hasOwnProperty.call(data, 'duration'))
}

function isDiscussionResult(data) {
  return data && data.data && typeof data.data.forum_name === 'string'
}

function isWebResult(data) {
  if (!data || typeof data !== 'object') return false
  if (typeof data.title !== 'string' || typeof data.url !== 'string') return false
  if (isFaqResult(data) || isNewsResult(data) || isVideoResult(data) || isDiscussionResult(data)) {
    return false
  }
  return true
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
    icon: path.join(__dirname, '../public/favicon.svg'),
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
    // 打开 DevTools 查看调试日志
    mainWindow.webContents.openDevTools()
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
  if (braveMcpConnection) {
    braveMcpConnection.client?.close?.().catch((err) => console.error('关闭 MCP 客户端失败:', err))
    braveMcpConnection.server?.close?.().catch((err) => console.error('关闭 MCP 服务失败:', err))
  }
})

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow()
  }
})

async function performBraveWebSearch(query, options = {}) {
  const { client } = await ensureBraveMcpClient(options.braveApiKey)
  const args = buildBraveWebArguments(query, options)

  console.log('[Brave MCP] 调用 brave_web_search:', args)
  const result = await client.callTool({
    name: 'brave_web_search',
    arguments: args
  })

  if (result.isError) {
    const errorMessage = Array.isArray(result.content)
      ? result.content.map((item) => item?.text || '').join('\n')
      : 'Brave 搜索失败'
    throw new Error(errorMessage || 'Brave 搜索失败')
  }

  const parsedContent = transformBraveContent(result.content, args.count || 5)
  if (parsedContent.sections.web.length === 0) {
    return { success: false, message: 'Brave 搜索未返回结果' }
  }

  return {
    success: true,
    results: parsedContent.sections.web,
    sections: parsedContent.sections,
    summarizerKey: parsedContent.summarizerKey,
    raw: result.content,
  }
}


// ==================== IPC 处理：文件系统操作 ====================

// 获取文件的 HTTP URL（供 ONLYOFFICE 使用）
ipcMain.handle('get-file-url', async (event, filePath) => {
  // 将本地文件路径转换为 HTTP URL
  // 使用 host.docker.internal 让 Docker 容器能访问宿主机
  const encodedPath = encodeURIComponent(filePath.replace(/\\/g, '/'))
  return `http://host.docker.internal:${fileServerPort}/file/${encodedPath}`
})

// 获取文件的 HTTP URL（供渲染进程直接使用）
ipcMain.handle('get-local-file-url', async (_event, filePath) => {
  const encodedPath = encodeURIComponent(filePath.replace(/\\/g, '/'))
  return `http://localhost:${fileServerPort}/file/${encodedPath}`
})

// 选择文件夹
ipcMain.handle('select-folder', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ['openDirectory'],
    title: '选择工作文件夹'
  })
  
  if (result.canceled) return null
  return result.filePaths[0]
})

// 读取文件夹内容（递归）
ipcMain.handle('read-folder', async (event, folderPath) => {
  try {
    const items = await readFolderRecursive(folderPath, folderPath)
    return { success: true, data: items }
  } catch (error) {
    return { success: false, error: error.message }
  }
})

async function readFolderRecursive(basePath, currentPath, depth = 0) {
  if (depth > 5) return [] // 限制递归深度
  
  const items = []
  const entries = fs.readdirSync(currentPath, { withFileTypes: true })
  
  for (const entry of entries) {
    // 跳过隐藏文件和 node_modules
    if (entry.name.startsWith('.') || entry.name === 'node_modules') continue
    
    const fullPath = path.join(currentPath, entry.name)
    const relativePath = path.relative(basePath, fullPath)
    
    if (entry.isDirectory()) {
      const children = await readFolderRecursive(basePath, fullPath, depth + 1)
      items.push({
        name: entry.name,
        path: fullPath,
        relativePath: relativePath,
        type: 'folder',
        children
      })
    } else {
      // 只显示支持的文件类型
      const ext = path.extname(entry.name).toLowerCase()
      const supportedExts = ['.docx', '.doc', '.txt', '.md', '.json', '.xml', '.xlsx', '.xls', '.pptx', '.ppt']
      
      if (supportedExts.includes(ext)) {
        items.push({
          name: entry.name,
          path: fullPath,
          relativePath: relativePath,
          type: 'file',
          extension: ext
        })
      }
    }
  }
  
  // 文件夹优先，然后按名称排序
  items.sort((a, b) => {
    if (a.type !== b.type) return a.type === 'folder' ? -1 : 1
    return a.name.localeCompare(b.name)
  })
  
  return items
}

// 读取文件内容
ipcMain.handle('read-file', async (event, filePath) => {
  try {
    const ext = path.extname(filePath).toLowerCase()
    const fileName = path.basename(filePath)
    
    // 跳过临时文件（以 ~$ 开头的文件）
    if (fileName.startsWith('~$')) {
      console.log('跳过临时文件:', filePath)
      return { 
        success: true, 
        data: '<p style="text-align: center; color: #888; padding: 40px;">这是一个 Word 临时文件，无法打开。</p>', 
        type: 'html' 
      }
    }
    
    if (ext === '.docx') {
      // .docx 文件返回 base64，让前端用自定义解析器处理（保留更多样式）
      console.log('读取 .docx 文件:', filePath)
      const buffer = fs.readFileSync(filePath)
      return { success: true, data: buffer.toString('base64'), type: 'docx' }
    } else if (ext === '.pptx') {
      // .pptx 文件返回 base64，让前端用纯 JS 渲染（无需 LibreOffice）
      console.log('读取 .pptx 文件:', filePath)
      const buffer = fs.readFileSync(filePath)
      return { success: true, data: buffer.toString('base64'), type: 'pptx' }
    } else if (ext === '.doc') {
      // .doc 文件（旧版 Word 97-2003）使用 word-extractor 解析
      console.log('使用 word-extractor 解析 .doc 文件:', filePath)
      
      try {
        const extractor = new WordExtractor()
        const extracted = await extractor.extract(filePath)
        
        // 获取文档内容
        const body = extracted.getBody() || ''
        
        console.log('word-extractor 提取成功，内容长度:', body.length)
        
        // 将纯文本转换为 HTML - 保持简单格式
        let html = ''
        
        // 处理正文 - 按段落分割（两个或更多换行）
        const paragraphs = body.split(/\n\n+/)
        for (const para of paragraphs) {
          const trimmed = para.trim()
          if (trimmed) {
            // 处理段落内的单个换行
            const lines = trimmed.split(/\n/)
            const formattedPara = lines.map(line => escapeHtml(line)).join('<br>')
            html += `<p>${formattedPara}</p>`
          }
        }
        
        if (!html) {
          html = '<p></p>'
        }
        
        // 返回为 doc-html 类型，前端可以区分处理
        return { success: true, data: html, type: 'doc-html' }
      } catch (extractorError) {
        console.error('word-extractor 解析 .doc 失败:', extractorError)
        
        return { 
          success: true, 
          data: `<div style="padding: 40px; text-align: center; color: #888;">
            <p style="font-size: 18px; margin-bottom: 10px;">⚠️ 无法解析此 .doc 文件</p>
            <p style="font-size: 14px;">此文件可能已损坏或使用了不支持的格式。</p>
            <p style="font-size: 12px; margin-top: 15px; color: #666;">
              建议：使用 Microsoft Word 打开此文件，然后另存为 .docx 格式。
            </p>
          </div>`, 
          type: 'doc-html' 
        }
      }
    } else {
      // 读取文本文件
      const content = fs.readFileSync(filePath, 'utf-8')
      return { success: true, data: content, type: 'text' }
    }
  } catch (error) {
    console.error('读取文件失败:', error)
    return { success: false, error: error.message }
  }
})

// ==================== IPC 处理：记忆系统 ====================

ipcMain.handle('memory-search', async (_event, options) => {
  try {
    const mgr = getMemoryManager()
    const query = options?.query || ''
    const topK = options?.topK || 5
    const textWeight = typeof options?.textWeight === 'number' ? options.textWeight : 0.6
    const vectorWeight = typeof options?.vectorWeight === 'number' ? options.vectorWeight : 0.4
    const workspaceKey = options?.workspaceKey || ''
    const sources = options?.sources
    return mgr.search({ query, topK, textWeight, vectorWeight, workspaceKey, sources })
  } catch (error) {
    return { success: false, error: error.message, results: [] }
  }
})

ipcMain.handle('memory-append', async (_event, payload) => {
  try {
    const mgr = getMemoryManager()
    const text = payload?.text || ''
    if (!text.trim()) return { success: false, error: '内容为空' }
    return mgr.appendDaily({
      text,
      source: payload?.source || 'chat',
      tags: Array.isArray(payload?.tags) ? payload.tags : [],
    })
  } catch (error) {
    return { success: false, error: error.message }
  }
})

ipcMain.handle('memory-append-session', async (_event, payload) => {
  try {
    const mgr = getMemoryManager()
    const sessionId = payload?.sessionId || ''
    const text = payload?.text || ''
    if (!sessionId || !text.trim()) return { success: false, error: 'sessionId 或内容为空' }
    return mgr.appendSession({
      sessionId,
      text,
      meta: payload?.meta || {},
    })
  } catch (error) {
    return { success: false, error: error.message }
  }
})

ipcMain.handle('memory-status', async () => {
  try {
    const mgr = getMemoryManager()
    return mgr.getStatus()
  } catch (error) {
    return { success: false, error: error.message }
  }
})

ipcMain.handle('memory-status-detail', async () => {
  try {
    const mgr = getMemoryManager()
    return mgr.getStatusDetail()
  } catch (error) {
    return { success: false, error: error.message }
  }
})

ipcMain.handle('memory-clear', async (_event, payload) => {
  try {
    const mgr = getMemoryManager()
    const scope = payload?.scope || 'all'
    return mgr.clear(scope)
  } catch (error) {
    return { success: false, error: error.message }
  }
})

ipcMain.handle('memory-rebuild-index', async () => {
  try {
    const mgr = getMemoryManager()
    mgr.rebuildAll()
    return mgr.getStatusDetail()
  } catch (error) {
    return { success: false, error: error.message }
  }
})

// 直接使用 SheetJS 读取 xls 文件（提取尽可能多的样式信息）
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

function hashForFileCache(filePath) {
  const st = fs.statSync(filePath)
  const key = `${filePath}|${st.size}|${st.mtimeMs}`
  return crypto.createHash('sha1').update(key).digest('hex')
}

function getPptxPreviewCacheDir(filePath) {
  const hash = hashForFileCache(filePath)
  const tempDir = app.getPath('temp')
  return path.join(tempDir, 'word-cursor-ppt-preview', hash)
}

function listPngFilesSorted(dir) {
  const files = fs.readdirSync(dir).filter((f) => f.toLowerCase().endsWith('.png'))
  const withMeta = files.map((name) => {
    const m = name.match(/(\d+)(?=\.png$)/)
    const idx = m ? parseInt(m[1], 10) : 0
    return { name, idx }
  })
  withMeta.sort((a, b) => (a.idx - b.idx) || a.name.localeCompare(b.name))
  return withMeta.map((x) => path.join(dir, x.name))
}

async function renderPptxToPngsWithLibreOffice(pptxPath, outDir) {
  const libreOfficePath = findLibreOffice()
  if (!libreOfficePath) {
    return { success: false, error: 'LibreOffice 未安装', downloadUrl: getLibreOfficeDownloadUrl() }
  }

  if (!fs.existsSync(outDir)) {
    fs.mkdirSync(outDir, { recursive: true })
  }

  const { execFile } = require('child_process')
  // LibreOffice 将每页导出为 PNG（文件名规则依版本不同，导出后我们扫描目录排序）
  return new Promise((resolve) => {
    execFile(
      libreOfficePath,
      ['--headless', '--nologo', '--nolockcheck', '--norestore', '--convert-to', 'png', '--outdir', outDir, pptxPath],
      { timeout: 180000 },
      (error, stdout, stderr) => {
        if (error) {
          console.error('[PPTX] LibreOffice 转换失败:', error)
          resolve({ success: false, error: 'LibreOffice 转换失败', details: stderr || stdout })
          return
        }
        const pngs = listPngFilesSorted(outDir)
        if (!pngs.length) {
          resolve({ success: false, error: 'LibreOffice 转换未生成 PNG' })
          return
        }
        resolve({ success: true, images: pngs })
      }
    )
  })
}

ipcMain.handle('pptx-render-preview', async (_event, filePath) => {
  try {
    if (!filePath || typeof filePath !== 'string') {
      return { success: false, error: '缺少 filePath' }
    }
    if (!fs.existsSync(filePath)) {
      return { success: false, error: '文件不存在' }
    }
    if (path.extname(filePath).toLowerCase() !== '.pptx') {
      return { success: false, error: '仅支持 .pptx' }
    }

    const cacheDir = getPptxPreviewCacheDir(filePath)
    if (fs.existsSync(cacheDir)) {
      const cached = listPngFilesSorted(cacheDir)
      if (cached.length > 0) {
        return { success: true, images: cached, cacheDir, cached: true }
      }
    }

    const result = await renderPptxToPngsWithLibreOffice(filePath, cacheDir)
    if (!result.success) {
      return result
    }
    return { success: true, images: result.images, cacheDir, cached: false }
  } catch (error) {
    console.error('[PPTX] render preview failed:', error)
    return { success: false, error: error.message || String(error) }
  }
})

// 检查是否需要安装 LibreOffice 的 IPC
ipcMain.handle('check-libreoffice', async () => {
  const path = findLibreOffice()
  return {
    installed: !!path,
    path: path,
    downloadUrl: !path ? getLibreOfficeDownloadUrl() : null
  }
})

// 使用 LibreOffice 进行无损转换（开源方案）
async function convertWithLibreOffice(xlsPath) {
  const libreOfficePath = findLibreOffice()
  if (!libreOfficePath) {
    return { success: false, error: 'LibreOffice 未安装' }
  }
  
  const xlsxPath = xlsPath.replace(/\.xls$/i, '.xlsx')
  const outputDir = path.dirname(xlsPath)
  
  if (fs.existsSync(xlsxPath)) {
    return { 
      success: false, 
      error: `文件 ${path.basename(xlsxPath)} 已存在。请先删除或重命名现有文件。` 
    }
  }
  
  const { execFile } = require('child_process')
  
  return new Promise((resolve) => {
    // LibreOffice 命令行转换
    execFile(libreOfficePath, [
      '--headless',
      '--convert-to', 'xlsx',
      '--outdir', outputDir,
      xlsPath
    ], { timeout: 60000 }, (error, stdout, stderr) => {
      if (error) {
        console.error('[Excel] LibreOffice 转换失败:', error)
        resolve({ success: false, error: 'LibreOffice 转换失败', details: stderr })
      } else if (fs.existsSync(xlsxPath)) {
        console.log('[Excel] LibreOffice 转换成功:', xlsxPath)
        resolve({ 
          success: true, 
          xlsxPath,
          message: `已使用 LibreOffice 转换为 ${path.basename(xlsxPath)}，所有样式已完整保留！`
        })
      } else {
        resolve({ success: false, error: 'LibreOffice 转换后文件不存在' })
      }
    })
  })
}

// 使用系统安装的 Excel 进行无损转换（保留所有样式）
async function convertWithExcel(xlsPath) {
  const xlsxPath = xlsPath.replace(/\.xls$/i, '.xlsx')
  
  // 检查输出文件是否已存在
  if (fs.existsSync(xlsxPath)) {
    return { 
      success: false, 
      error: `文件 ${path.basename(xlsxPath)} 已存在。请先删除或重命名现有文件。` 
    }
  }
  
  // 使用 PowerShell 调用 Excel COM 对象
  const { exec } = require('child_process')
  
  // 转义路径中的特殊字符
  const escapedXlsPath = xlsPath.replace(/'/g, "''")
  const escapedXlsxPath = xlsxPath.replace(/'/g, "''")
  
  const psScript = `
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try {
      $workbook = $excel.Workbooks.Open('${escapedXlsPath}')
      $workbook.SaveAs('${escapedXlsxPath}', 51)
      $workbook.Close($false)
      Write-Output "SUCCESS"
    } catch {
      Write-Output "ERROR: $($_.Exception.Message)"
    } finally {
      $excel.Quit()
      [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
  `
  
  return new Promise((resolve) => {
    exec(`powershell -Command "${psScript.replace(/"/g, '\\"').replace(/\n/g, ' ')}"`, 
      { encoding: 'utf8', maxBuffer: 1024 * 1024, timeout: 60000 },
      (error, stdout, stderr) => {
        if (error || !stdout.includes('SUCCESS')) {
          console.error('[Excel] PowerShell 转换失败:', error || stderr || stdout)
          resolve({ 
            success: false, 
            error: '调用 Excel 失败',
            details: stderr || stdout
          })
        } else {
          console.log('[Excel] Excel COM 转换成功:', xlsxPath)
          resolve({ 
            success: true, 
            xlsxPath,
            message: `已使用 Microsoft Excel 转换为 ${path.basename(xlsxPath)}，所有样式已完整保留！`
          })
        }
      }
    )
  })
}

// 使用 SheetJS 转换（数据转换，样式可能丢失）
async function convertWithSheetJS(xlsPath) {
  const XLSX = require('xlsx')
  const xlsxPath = xlsPath.replace(/\.xls$/i, '.xlsx')
  
  if (fs.existsSync(xlsxPath)) {
    return { 
      success: false, 
      error: `文件 ${path.basename(xlsxPath)} 已存在。请先删除或重命名现有文件。` 
    }
  }
  
  const buffer = fs.readFileSync(xlsPath)
  const workbook = XLSX.read(buffer, { 
    type: 'buffer',
    cellFormula: true,
    cellNF: true,
    cellDates: true
  })
  
  XLSX.writeFile(workbook, xlsxPath, { bookType: 'xlsx' })
  
  return { 
    success: true, 
    xlsxPath,
    message: `已转换为 ${path.basename(xlsxPath)}。注意：由于技术限制，样式信息可能丢失。`
  }
}

// 将 xls 转换为 xlsx（优先级：LibreOffice > Excel > SheetJS）
ipcMain.handle('excel-convert-xls-to-xlsx', async (_event, xlsPath) => {
  try {
    console.log('[Excel] 开始转换 xls 到 xlsx:', xlsPath)
    
    // 1. 优先尝试 LibreOffice（开源，跨平台）
    console.log('[Excel] 尝试 LibreOffice...')
    const libreResult = await convertWithLibreOffice(xlsPath)
    if (libreResult.success) {
      return libreResult
    }
    console.log('[Excel] LibreOffice 不可用:', libreResult.error)
    
    // 2. Windows 上尝试 Excel COM
    if (process.platform === 'win32') {
      console.log('[Excel] 尝试 Microsoft Excel...')
      const excelResult = await convertWithExcel(xlsPath)
      if (excelResult.success) {
        return excelResult
      }
      console.log('[Excel] Excel COM 不可用:', excelResult.error)
    }
    
    // 3. 最后使用 SheetJS（数据转换，样式可能丢失）
    console.log('[Excel] 使用 SheetJS 进行基础转换（样式可能丢失）...')
    return await convertWithSheetJS(xlsPath)
  } catch (error) {
    console.error('xls 转 xlsx 失败:', error)
    return { success: false, error: error.message }
  }
})

// 读取 Excel（高保真只读预览数据）
// .xlsx 使用 ExcelJS（更好的样式支持），.xls 使用 SheetJS
ipcMain.handle('excel-open', async (_event, filePath) => {
  if (!filePath) {
    return { success: false, error: '缺少 filePath 参数' }
  }

  const ext = path.extname(filePath).toLowerCase()
  
  // .xls 文件使用 SheetJS 直接读取
  // 注意：SheetJS 免费版对 xls 样式支持有限
  if (ext === '.xls') {
    const result = await readExcelWithSheetJS(filePath)
    result.isXls = true  // 标记为 xls 文件
    result.originalPath = filePath
    // 添加警告信息，提示用户样式可能不完整
    result.warning = '提示：.xls 格式的样式支持有限。建议在 Microsoft Excel 中打开原文件，另存为 .xlsx 格式后重新打开，即可完整显示所有样式。'
    return result
  }
  
  // .xlsx 文件使用 ExcelJS 读取（更好的样式支持）
  try {
    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.readFile(filePath)
    
    const sheets = []
    const names = workbook.definedNames?.model || []
    
    workbook.eachSheet((worksheet, sheetId) => {
      // 注意：ExcelJS 的 worksheet.rowCount/columnCount 可能因为“整表格式/模板”变成 1048576/16384
      // 前端会按 range 构造矩阵，导致 OOM/白屏。这里按真实 used-range（非空单元格 + 合并区域）计算。
      let maxR = -1
      let maxC = -1

      const decodeCell = (addr) => {
        const match = String(addr || '').toUpperCase().match(/^(\$?)([A-Z]+)(\$?)(\d+)$/)
        if (!match) return null
        let col = 0
        for (let i = 0; i < match[2].length; i++) {
          col = col * 26 + (match[2].charCodeAt(i) - 64)
        }
        return { c: col - 1, r: parseInt(match[4], 10) - 1 }
      }

      const sheetData = {
        name: worksheet.name,
        range: { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } },
        merges: [],
        colWidths: [],
        rowHeights: [],
        autoFilter: worksheet.autoFilter || null,
        printArea: null,
        margins: null,
        dataValidations: null,
        cells: []
      }
      
      // 合并单元格
      if (worksheet.model && worksheet.model.merges) {
        worksheet.model.merges.forEach((mergeRange) => {
          const parts = String(mergeRange || '').split(':')
          if (parts.length !== 2) return
          const start = decodeCell(parts[0])
          const end = decodeCell(parts[1])
          if (!start || !end) return
          sheetData.merges.push({ s: { r: start.r, c: start.c }, e: { r: end.r, c: end.c } })
          maxR = Math.max(maxR, end.r)
          maxC = Math.max(maxC, end.c)
        })
      }
      
      // 列宽
      if (worksheet.columns) {
        worksheet.columns.forEach((col, idx) => {
          if (col && col.width) {
            // ExcelJS 列宽是字符数，转为像素（约 7px/字符 + 5px padding）
            sheetData.colWidths[idx] = Math.round(col.width * 7 + 5)
          }
        })
      }
      
      // 行高和单元格
      worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        // 行高（ExcelJS 返回 points，转为像素）
        if (row.height) {
          sheetData.rowHeights[rowNumber - 1] = Math.round(row.height * 1.333)
        }
        
        row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
          const r = rowNumber - 1
          const c = colNumber - 1
          maxR = Math.max(maxR, r)
          maxC = Math.max(maxC, c)
          
          // 提取样式
          const styleObj = {}
          
          // 字体
          if (cell.font) {
            styleObj.font = {
              name: cell.font.name,
              sz: cell.font.size,
              bold: cell.font.bold,
              italic: cell.font.italic,
              underline: cell.font.underline,
              strike: cell.font.strike,
              color: cell.font.color ? { argb: cell.font.color.argb, rgb: cell.font.color.argb?.slice(2) } : null
            }
          }
          
          // 填充/背景色
          if (cell.fill) {
            styleObj.fill = {}
            if (cell.fill.type === 'pattern' && cell.fill.fgColor) {
              styleObj.fill.fgColor = { argb: cell.fill.fgColor.argb, rgb: cell.fill.fgColor.argb?.slice(2) }
            }
            if (cell.fill.bgColor) {
              styleObj.fill.bgColor = { argb: cell.fill.bgColor.argb, rgb: cell.fill.bgColor.argb?.slice(2) }
            }
          }
          
          // 对齐
          if (cell.alignment) {
            styleObj.alignment = {
              horizontal: cell.alignment.horizontal,
              vertical: cell.alignment.vertical,
              wrapText: cell.alignment.wrapText,
              shrinkToFit: cell.alignment.shrinkToFit,
              indent: cell.alignment.indent,
              textRotation: cell.alignment.textRotation
            }
          }
          
          // 边框
          if (cell.border) {
            styleObj.border = {}
            ;['top', 'bottom', 'left', 'right'].forEach((side) => {
              if (cell.border[side]) {
                styleObj.border[side] = {
                  style: cell.border[side].style,
                  color: cell.border[side].color ? { argb: cell.border[side].color.argb, rgb: cell.border[side].color.argb?.slice(2) } : null
                }
              }
            })
          }
          
          // 数字格式
          if (cell.numFmt) {
            styleObj.numFmt = cell.numFmt
          }
          
          // 获取显示值（安全处理，避免 null 值和合并单元格错误）
          let display = ''
          try {
            // 先尝试获取 value，因为 text getter 在合并单元格时会报错
            const cellValue = cell.value
            if (cellValue != null) {
              if (typeof cellValue === 'object') {
                // 富文本 { richText: [...] }
                if (cellValue.richText && Array.isArray(cellValue.richText)) {
                  display = cellValue.richText.map(rt => rt.text || '').join('')
                }
                // 公式 { formula: '...', result: ... }
                else if (cellValue.formula) {
                  // 如果有计算结果，显示结果
                  if (cellValue.result != null) {
                    display = String(cellValue.result)
                  } else {
                    // 尝试计算公式（传入 workbook 支持跨工作表引用）
                    const calculated = evaluateSimpleFormula(cellValue.formula, worksheet, workbook)
                    if (calculated != null) {
                      display = String(calculated)
                    } else {
                      // 无法计算时显示公式本身
                      display = '=' + cellValue.formula
                    }
                  }
                }
                // 超链接 { text: '...', hyperlink: '...' }
                else if (cellValue.text != null) {
                  display = String(cellValue.text)
                }
                // 其他对象（可能有 result 但没有 formula）
                else if (cellValue.result != null) {
                  display = String(cellValue.result)
                }
                // 其他对象
                else {
                  display = String(cellValue)
                }
              } else {
                display = String(cellValue)
              }
            }
          } catch (e) {
            // 如果还是失败，返回空字符串
            console.warn(`[Excel Read] 单元格 ${colNumber}:${rowNumber} 读取失败:`, e.message)
            display = ''
          }
          
          // 公式
          const formula = cell.formula || (cell.value && cell.value.formula) || null
          
          // 超链接
          const hyperlink = cell.hyperlink || null
          
          // 批注
          let comment = null
          if (cell.note) {
            comment = typeof cell.note === 'string' ? cell.note : (cell.note.texts ? cell.note.texts.map(t => t.text || t).join('') : '')
          }
          
          sheetData.cells.push({
            r,
            c,
            v: cell.value,
            t: cell.type,
            w: display, // 使用安全计算的 display 值，避免 cell.text getter 错误
            f: formula,
            l: hyperlink,
            z: cell.numFmt,
            cmt: comment,
            display,
            s: styleObj
          })
        })
      })

      // 修正 range：使用真实 used-range，避免 rowCount/columnCount 造成超大范围
      if (maxR >= 0 && maxC >= 0) {
        sheetData.range.e = { r: maxR, c: maxC }
      } else {
        sheetData.range.e = { r: 0, c: 0 }
      }
      
      sheets.push(sheetData)
    })

    return { success: true, sheets, names }
  } catch (error) {
    console.error('读取 Excel 失败:', error)
    return { success: false, error: error.message || '读取 Excel 失败' }
  }
})

// ==================== Excel 增删查改操作 ====================

// 缓存打开的工作簿，避免每次操作都重新加载
const openWorkbooks = new Map()

// 获取或加载工作簿
async function getWorkbook(filePath) {
  if (openWorkbooks.has(filePath)) {
    return openWorkbooks.get(filePath)
  }
  
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(filePath)
  openWorkbooks.set(filePath, workbook)
  return workbook
}

// 保存工作簿
async function saveWorkbook(filePath) {
  const workbook = openWorkbooks.get(filePath)
  if (workbook) {
    await workbook.xlsx.writeFile(filePath)
    return true
  }
  return false
}

// 清除工作簿缓存
function clearWorkbookCache(filePath) {
  openWorkbooks.delete(filePath)
}

// ============================================================
// Excel 公式计算引擎 - 支持跨工作表引用和完整函数库
// ============================================================

/**
 * 创建一个公式计算器实例
 * @param {Object} workbook - ExcelJS 工作簿对象
 * @param {Object} currentWorksheet - 当前工作表
 */
function createFormulaEngine(workbook, currentWorksheet) {
  // 缓存已计算的单元格，防止循环引用
  const calculationCache = new Map()
  const calculationStack = new Set()
  
  // 解析单元格地址 (如 "A1" -> { r: 0, c: 0 })
  // 也支持纯列引用 "A" -> { r: null, c: 0, isColumn: true }
  const parseCellAddr = (address) => {
    const upperAddr = address.toUpperCase()
    
    // 尝试匹配带行号的地址 (如 A1, $B$2)
    const match = upperAddr.match(/^(\$?)([A-Z]+)(\$?)(\d+)$/)
    if (match) {
      let col = 0
      for (let i = 0; i < match[2].length; i++) {
        col = col * 26 + (match[2].charCodeAt(i) - 64)
      }
      return { r: parseInt(match[4], 10) - 1, c: col - 1 }
    }
    
    // 尝试匹配纯列引用 (如 A, B, $C)
    const colMatch = upperAddr.match(/^(\$?)([A-Z]+)$/)
    if (colMatch) {
      let col = 0
      for (let i = 0; i < colMatch[2].length; i++) {
        col = col * 26 + (colMatch[2].charCodeAt(i) - 64)
      }
      return { r: null, c: col - 1, isColumn: true }
    }
    
    return null
  }
  
  // 获取工作表（支持跨工作表引用）
  const getWorksheet = (sheetName) => {
    if (!sheetName) return currentWorksheet
    // 移除引号
    const cleanName = sheetName.replace(/^'|'$/g, '')
    const targetSheet = workbook.getWorksheet(cleanName)
    
    // 调试日志
    console.log(`[Formula Debug] getWorksheet: sheetName="${sheetName}", cleanName="${cleanName}", found=${!!targetSheet}`)
    if (!targetSheet) {
      // 列出所有可用的工作表名称
      const availableSheets = []
      workbook.eachSheet((ws) => availableSheets.push(ws.name))
      console.log(`[Formula Debug] 可用工作表: ${availableSheets.join(', ')}`)
    }
    
    return targetSheet || currentWorksheet
  }
  
  // 解析带工作表引用的单元格地址 (如 "'Sheet1'!A1" 或 "A1")
  const parseFullReference = (ref) => {
    const sheetMatch = ref.match(/^'?([^'!]+)'?!(.+)$/)
    if (sheetMatch) {
      return { sheetName: sheetMatch[1], cellRef: sheetMatch[2] }
    }
    return { sheetName: null, cellRef: ref }
  }
  
  // 获取单元格的原始值（不计算）
  const getRawCellValue = (ws, row, col) => {
    const cell = ws.getCell(row, col)
    return cell.value
  }
  
  // 当前计算上下文的工作表（用于嵌套公式计算）
  let activeWorksheet = currentWorksheet
  
  // 获取单元格的计算值
  const getCellValue = (ref, defaultWs = null) => {
    const { sheetName, cellRef } = parseFullReference(ref)
    // 优先级：1. ref 中指定的工作表 2. 传入的 defaultWs 3. activeWorksheet
    const ws = sheetName ? getWorksheet(sheetName) : (defaultWs || activeWorksheet)
    const addr = parseCellAddr(cellRef)
    if (!addr) return 0
    
    const cacheKey = `${ws.name || 'default'}!${cellRef}`
    
    // 检查循环引用
    if (calculationStack.has(cacheKey)) {
      console.warn(`[Formula] 检测到循环引用: ${cacheKey}`)
      return 0
    }
    
    // 检查缓存
    if (calculationCache.has(cacheKey)) {
      return calculationCache.get(cacheKey)
    }
    
    const cell = ws.getCell(addr.r + 1, addr.c + 1)
    const value = cell.value
    
    if (value == null) return 0
    if (typeof value === 'number') return value
    if (typeof value === 'string') {
      const num = parseFloat(value)
      return isNaN(num) ? value : num
    }
    if (typeof value === 'object') {
      if (value.result != null) return value.result
      if (value.formula) {
        calculationStack.add(cacheKey)
        // 关键修复：临时切换活动工作表上下文，确保嵌套公式在正确的工作表中计算
        const previousActiveWs = activeWorksheet
        activeWorksheet = ws
        const result = evaluateFormula(value.formula, ws)
        activeWorksheet = previousActiveWs  // 恢复之前的上下文
        calculationStack.delete(cacheKey)
        if (result != null) {
          calculationCache.set(cacheKey, result)
          return result
        }
      }
      if (value.richText) {
        return value.richText.map(t => t.text || '').join('')
      }
      if (value.text != null) return value.text
    }
    return 0
  }
  
  // 获取单元格的文本值（用于文本函数）
  const getCellText = (ref, defaultWs = currentWorksheet) => {
    const val = getCellValue(ref, defaultWs)
    return String(val)
  }
  
  // 解析范围并获取所有值
  const getRangeValues = (rangeStr, ws = currentWorksheet) => {
    const { sheetName, cellRef } = parseFullReference(rangeStr)
    const targetWs = getWorksheet(sheetName)
    
    console.log(`[Formula Debug] getRangeValues: rangeStr="${rangeStr}", sheetName="${sheetName}", cellRef="${cellRef}", targetWs="${targetWs?.name}"`)
    
    const parts = cellRef.split(':')
    if (parts.length !== 2) {
      // 单个单元格
      return [getCellValue(rangeStr, ws)]
    }
    
    const start = parseCellAddr(parts[0])
    const end = parseCellAddr(parts[1])
    if (!start || !end) return []
    
    // 处理整列范围（如 E:E）
    let startRow = start.r
    let endRow = end.r
    if (start.isColumn || end.isColumn) {
      // 整列范围：只遍历有数据的行
      startRow = 0
      endRow = Math.max((targetWs?.rowCount || 100) - 1, 0)
      // 限制最大行数，避免遍历太多空行
      endRow = Math.min(endRow, 999)
    }
    
    const values = []
    for (let r = startRow; r <= endRow; r++) {
      for (let c = start.c; c <= end.c; c++) {
        const val = getCellValue(`${getColumnLabel(c)}${r + 1}`, targetWs)
        values.push(val)
      }
    }
    
    // 显示前10个值用于调试
    console.log(`[Formula Debug] getRangeValues 结果: 共${values.length}个值, 非0值数量: ${values.filter(v => v !== 0 && v !== '').length}`)
    
    return values
  }
  
  // 解析范围并获取所有单元格信息（包含位置）
  const getRangeCells = (rangeStr, ws = currentWorksheet) => {
    const { sheetName, cellRef } = parseFullReference(rangeStr)
    const targetWs = getWorksheet(sheetName)
    
    const parts = cellRef.split(':')
    if (parts.length !== 2) return []
    
    const start = parseCellAddr(parts[0])
    const end = parseCellAddr(parts[1])
    if (!start || !end) return []
    
    // 处理整列范围（如 E:E, H:H）
    let startRow = start.r
    let endRow = end.r
    if (start.isColumn || end.isColumn) {
      startRow = 0
      endRow = Math.max((targetWs?.rowCount || 100) - 1, 0)
      endRow = Math.min(endRow, 999)
    }
    
    const cells = []
    for (let r = startRow; r <= endRow; r++) {
      for (let c = start.c; c <= end.c; c++) {
        const ref = `${getColumnLabel(c)}${r + 1}`
        cells.push({
          row: r,
          col: c,
          ref,
          value: getCellValue(ref, targetWs),
          rawValue: getRawCellValue(targetWs, r + 1, c + 1)
        })
      }
    }
    return cells
  }
  
  // 获取列标签
  const getColumnLabel = (colIndex) => {
    let label = ''
    let n = colIndex
    while (n >= 0) {
      label = String.fromCharCode(65 + (n % 26)) + label
      n = Math.floor(n / 26) - 1
    }
    return label
  }
  
  // 解析函数参数（处理嵌套括号和逗号）
  const parseFunctionArgs = (argsStr) => {
    const args = []
    let depth = 0
    let current = ''
    
    for (let i = 0; i < argsStr.length; i++) {
      const char = argsStr[i]
      if (char === '(') depth++
      else if (char === ')') depth--
      else if (char === ',' && depth === 0) {
        args.push(current.trim())
        current = ''
        continue
      }
      current += char
    }
    if (current.trim()) args.push(current.trim())
    return args
  }
  
  // ============================================================
  // Excel 函数实现
  // ============================================================
  
  const functions = {
    // -------------------- 基础数学函数 --------------------
    
    // SUM - 求和
    SUM: (args) => {
      let total = 0
      for (const arg of args) {
        if (arg.includes(':')) {
          const values = getRangeValues(arg)
          total += values.filter(v => typeof v === 'number').reduce((a, b) => a + b, 0)
        } else {
          const val = evaluateExpression(arg)
          if (typeof val === 'number') total += val
        }
      }
      return total
    },
    
    // SUMIF - 条件求和
    SUMIF: (args) => {
      if (args.length < 2) return 0
      const [rangeStr, criteria, sumRangeStr] = args
      const cells = getRangeCells(rangeStr)
      const sumCells = sumRangeStr ? getRangeCells(sumRangeStr) : cells
      
      const criteriaValue = evaluateExpression(criteria.replace(/^"|"$/g, ''))
      let total = 0
      
      cells.forEach((cell, idx) => {
        if (matchCriteria(cell.value, criteriaValue)) {
          const sumVal = sumCells[idx]?.value
          if (typeof sumVal === 'number') total += sumVal
        }
      })
      return total
    },
    
    // SUMIFS - 多条件求和
    SUMIFS: (args) => {
      if (args.length < 3) return 0
      const sumRangeStr = args[0]
      const sumCells = getRangeCells(sumRangeStr)
      
      // 解析条件对
      const conditions = []
      for (let i = 1; i < args.length; i += 2) {
        if (i + 1 < args.length) {
          conditions.push({
            cells: getRangeCells(args[i]),
            criteria: evaluateExpression(args[i + 1].replace(/^"|"$/g, ''))
          })
        }
      }
      
      let total = 0
      sumCells.forEach((sumCell, idx) => {
        const allMatch = conditions.every(cond => {
          const cell = cond.cells[idx]
          return cell && matchCriteria(cell.value, cond.criteria)
        })
        if (allMatch && typeof sumCell.value === 'number') {
          total += sumCell.value
        }
      })
      return total
    },
    
    // AVERAGE - 平均值（只计算非空单元格中的数字）
    AVERAGE: (args) => {
      const values = []
      for (const arg of args) {
        if (arg.includes(':')) {
          // 使用 getRangeCells 获取原始值，排除真正的空单元格
          const cells = getRangeCells(arg)
          cells.forEach(c => {
            if (c.rawValue != null && c.rawValue !== '' && typeof c.value === 'number') {
              values.push(c.value)
            }
          })
        } else {
          const val = evaluateExpression(arg)
          if (typeof val === 'number') values.push(val)
        }
      }
      return values.length > 0 ? values.reduce((a, b) => a + b, 0) / values.length : 0
    },
    
    // AVERAGEIF - 条件平均值
    AVERAGEIF: (args) => {
      if (args.length < 2) return 0
      const [rangeStr, criteria, avgRangeStr] = args
      const cells = getRangeCells(rangeStr)
      const avgCells = avgRangeStr ? getRangeCells(avgRangeStr) : cells
      
      const criteriaValue = evaluateExpression(criteria.replace(/^"|"$/g, ''))
      const values = []
      
      cells.forEach((cell, idx) => {
        if (matchCriteria(cell.value, criteriaValue)) {
          const avgVal = avgCells[idx]?.value
          if (typeof avgVal === 'number') values.push(avgVal)
        }
      })
      return values.length > 0 ? values.reduce((a, b) => a + b, 0) / values.length : 0
    },
    
    // MAX - 最大值
    MAX: (args) => {
      const values = []
      for (const arg of args) {
        if (arg.includes(':')) {
          values.push(...getRangeValues(arg).filter(v => typeof v === 'number'))
        } else {
          const val = evaluateExpression(arg)
          if (typeof val === 'number') values.push(val)
        }
      }
      return values.length > 0 ? Math.max(...values) : 0
    },
    
    // MIN - 最小值（只计算非空单元格中的数字）
    MIN: (args) => {
      const values = []
      for (const arg of args) {
        if (arg.includes(':')) {
          // 使用 getRangeCells 获取原始值，排除真正的空单元格
          const cells = getRangeCells(arg)
          cells.forEach(c => {
            if (c.rawValue != null && c.rawValue !== '' && typeof c.value === 'number') {
              values.push(c.value)
            }
          })
        } else {
          const val = evaluateExpression(arg)
          if (typeof val === 'number') values.push(val)
        }
      }
      return values.length > 0 ? Math.min(...values) : 0
    },
    
    // ROUND - 四舍五入
    ROUND: (args) => {
      const num = evaluateExpression(args[0])
      const digits = args[1] ? evaluateExpression(args[1]) : 0
      if (typeof num !== 'number') return 0
      const factor = Math.pow(10, digits)
      return Math.round(num * factor) / factor
    },
    
    // ABS - 绝对值
    ABS: (args) => Math.abs(evaluateExpression(args[0]) || 0),
    
    // SQRT - 平方根
    SQRT: (args) => Math.sqrt(evaluateExpression(args[0]) || 0),
    
    // POWER - 幂运算
    POWER: (args) => Math.pow(evaluateExpression(args[0]) || 0, evaluateExpression(args[1]) || 0),
    
    // MOD - 取余
    MOD: (args) => {
      const num = evaluateExpression(args[0])
      const divisor = evaluateExpression(args[1])
      if (divisor === 0) return 0
      return num % divisor
    },
    
    // -------------------- 统计函数 --------------------
    
    // COUNT - 计数（仅数字）
    COUNT: (args) => {
      let count = 0
      for (const arg of args) {
        if (arg.includes(':')) {
          count += getRangeValues(arg).filter(v => typeof v === 'number').length
        } else {
          const val = evaluateExpression(arg)
          if (typeof val === 'number') count++
        }
      }
      return count
    },
    
    // COUNTA - 计数（非空单元格）
    COUNTA: (args) => {
      let count = 0
      for (const arg of args) {
        if (arg.includes(':')) {
          // 使用 getRangeCells 获取原始值，正确判断空单元格
          const cells = getRangeCells(arg)
          count += cells.filter(c => c.rawValue != null && c.rawValue !== '').length
        } else {
          const val = evaluateExpression(arg)
          if (val != null && val !== '') count++
        }
      }
      return count
    },
    
    // COUNTBLANK - 计数空单元格
    COUNTBLANK: (args) => {
      let count = 0
      for (const arg of args) {
        if (arg.includes(':')) {
          const cells = getRangeCells(arg)
          count += cells.filter(c => c.rawValue == null || c.rawValue === '').length
        }
      }
      return count
    },
    
    // COUNTIF - 条件计数
    COUNTIF: (args) => {
      if (args.length < 2) return 0
      const [rangeStr, criteria] = args
      const cells = getRangeCells(rangeStr)
      const criteriaValue = criteria.replace(/^"|"$/g, '')
      
      return cells.filter(cell => matchCriteria(cell.value, criteriaValue)).length
    },
    
    // COUNTIFS - 多条件计数
    COUNTIFS: (args) => {
      if (args.length < 2) return 0
      
      // 获取第一个范围作为基准
      const baseCells = getRangeCells(args[0])
      
      // 解析所有条件对
      const conditions = []
      for (let i = 0; i < args.length; i += 2) {
        if (i + 1 < args.length) {
          conditions.push({
            cells: getRangeCells(args[i]),
            criteria: args[i + 1].replace(/^"|"$/g, '')
          })
        }
      }
      
      let count = 0
      for (let idx = 0; idx < baseCells.length; idx++) {
        const allMatch = conditions.every(cond => {
          const cell = cond.cells[idx]
          return cell && matchCriteria(cell.value, cond.criteria)
        })
        if (allMatch) count++
      }
      return count
    },
    
    // -------------------- 逻辑函数 --------------------
    
    // IF - 条件判断
    IF: (args) => {
      const condition = evaluateExpression(args[0])
      const trueValue = args[1] ? evaluateExpression(args[1]) : true
      const falseValue = args[2] ? evaluateExpression(args[2]) : false
      return condition ? trueValue : falseValue
    },
    
    // AND - 逻辑与
    AND: (args) => args.every(arg => !!evaluateExpression(arg)),
    
    // OR - 逻辑或
    OR: (args) => args.some(arg => !!evaluateExpression(arg)),
    
    // NOT - 逻辑非
    NOT: (args) => !evaluateExpression(args[0]),
    
    // IFERROR - 错误处理
    IFERROR: (args) => {
      try {
        const result = evaluateExpression(args[0])
        if (result == null || (typeof result === 'number' && isNaN(result))) {
          return evaluateExpression(args[1])
        }
        return result
      } catch {
        return evaluateExpression(args[1])
      }
    },
    
    // -------------------- 查找/引用函数 --------------------
    
    // VLOOKUP - 垂直查找
    VLOOKUP: (args) => {
      const lookupValue = evaluateExpression(args[0])
      const tableRangeStr = args[1]
      const colIndex = evaluateExpression(args[2])
      const exactMatch = args[3] ? evaluateExpression(args[3]) === false : true
      
      const cells = getRangeCells(tableRangeStr)
      if (cells.length === 0) return '#N/A'
      
      // 确定表格的列数
      const { sheetName, cellRef } = parseFullReference(tableRangeStr)
      const parts = cellRef.split(':')
      const start = parseCellAddr(parts[0])
      const end = parseCellAddr(parts[1])
      const numCols = end.c - start.c + 1
      const numRows = end.r - start.r + 1
      
      // 查找匹配行
      for (let r = 0; r < numRows; r++) {
        const firstColValue = cells[r * numCols]?.value
        
        if (exactMatch) {
          if (firstColValue === lookupValue || String(firstColValue) === String(lookupValue)) {
            const targetIdx = r * numCols + (colIndex - 1)
            return cells[targetIdx]?.value ?? '#N/A'
          }
        } else {
          // 近似匹配（假设已排序）
          if (firstColValue <= lookupValue) {
            const nextRowValue = cells[(r + 1) * numCols]?.value
            if (nextRowValue == null || nextRowValue > lookupValue) {
              const targetIdx = r * numCols + (colIndex - 1)
              return cells[targetIdx]?.value ?? '#N/A'
            }
          }
        }
      }
      return '#N/A'
    },
    
    // INDEX - 返回指定位置的值
    INDEX: (args) => {
      const rangeStr = args[0]
      const rowNum = evaluateExpression(args[1])
      const colNum = args[2] ? evaluateExpression(args[2]) : 1
      
      const { sheetName, cellRef } = parseFullReference(rangeStr)
      const parts = cellRef.split(':')
      const start = parseCellAddr(parts[0])
      const end = parseCellAddr(parts[1])
      const numCols = end.c - start.c + 1
      
      const cells = getRangeCells(rangeStr)
      const idx = (rowNum - 1) * numCols + (colNum - 1)
      return cells[idx]?.value ?? '#REF!'
    },
    
    // MATCH - 查找匹配位置
    MATCH: (args) => {
      const lookupValue = evaluateExpression(args[0])
      const rangeStr = args[1]
      const matchType = args[2] ? evaluateExpression(args[2]) : 1
      
      const values = getRangeValues(rangeStr)
      
      if (matchType === 0) {
        // 精确匹配
        const idx = values.findIndex(v => v === lookupValue || String(v) === String(lookupValue))
        return idx >= 0 ? idx + 1 : '#N/A'
      } else if (matchType === 1) {
        // 小于或等于
        let lastIdx = -1
        for (let i = 0; i < values.length; i++) {
          if (values[i] <= lookupValue) lastIdx = i
          else break
        }
        return lastIdx >= 0 ? lastIdx + 1 : '#N/A'
      } else {
        // 大于或等于
        for (let i = 0; i < values.length; i++) {
          if (values[i] >= lookupValue) return i + 1
        }
        return '#N/A'
      }
    },
    
    // OFFSET - 偏移引用
    OFFSET: (args) => {
      const refStr = args[0]
      const rowOffset = evaluateExpression(args[1])
      const colOffset = evaluateExpression(args[2])
      const height = args[3] ? evaluateExpression(args[3]) : 1
      const width = args[4] ? evaluateExpression(args[4]) : 1
      
      const { sheetName, cellRef } = parseFullReference(refStr)
      const addr = parseCellAddr(cellRef.split(':')[0])
      if (!addr) return '#REF!'
      
      const newRow = addr.r + rowOffset
      const newCol = addr.c + colOffset
      
      if (height === 1 && width === 1) {
        return getCellValue(`${getColumnLabel(newCol)}${newRow + 1}`)
      }
      
      // 返回范围的值（求和）
      const values = []
      for (let r = 0; r < height; r++) {
        for (let c = 0; c < width; c++) {
          values.push(getCellValue(`${getColumnLabel(newCol + c)}${newRow + r + 1}`))
        }
      }
      return values.filter(v => typeof v === 'number').reduce((a, b) => a + b, 0)
    },
    
    // -------------------- 文本函数 --------------------
    
    // LEFT - 左侧字符
    LEFT: (args) => {
      const text = String(evaluateExpression(args[0]) || '')
      const numChars = args[1] ? evaluateExpression(args[1]) : 1
      return text.substring(0, numChars)
    },
    
    // RIGHT - 右侧字符
    RIGHT: (args) => {
      const text = String(evaluateExpression(args[0]) || '')
      const numChars = args[1] ? evaluateExpression(args[1]) : 1
      return text.substring(text.length - numChars)
    },
    
    // MID - 中间字符
    MID: (args) => {
      const text = String(evaluateExpression(args[0]) || '')
      const startNum = evaluateExpression(args[1])
      const numChars = evaluateExpression(args[2])
      return text.substring(startNum - 1, startNum - 1 + numChars)
    },
    
    // LEN - 字符长度
    LEN: (args) => String(evaluateExpression(args[0]) || '').length,
    
    // EXACT - 精确比较
    EXACT: (args) => {
      const text1 = String(evaluateExpression(args[0]) || '')
      const text2 = String(evaluateExpression(args[1]) || '')
      return text1 === text2
    },
    
    // CONCATENATE / CONCAT - 连接文本
    CONCATENATE: (args) => args.map(a => String(evaluateExpression(a) || '')).join(''),
    CONCAT: (args) => args.map(a => String(evaluateExpression(a) || '')).join(''),
    
    // TEXT - 格式化文本
    TEXT: (args) => {
      const value = evaluateExpression(args[0])
      const format = String(args[1] || '').replace(/^"|"$/g, '')
      const valueStr = String(value)
      
      // 日期格式化：如 "0000-00-00" 将 "19950315" 转为 "1995-03-15"
      if (format.match(/^0+-0+-0+$/) && /^\d{8}$/.test(valueStr)) {
        return `${valueStr.substring(0, 4)}-${valueStr.substring(4, 6)}-${valueStr.substring(6, 8)}`
      }
      
      // 日期格式化：如 "yyyy-mm-dd" 将 "19950315" 转为 "1995-03-15"
      if (format.toLowerCase().match(/^y+-m+-d+$/) && /^\d{8}$/.test(valueStr)) {
        return `${valueStr.substring(0, 4)}-${valueStr.substring(4, 6)}-${valueStr.substring(6, 8)}`
      }
      
      // 日期格式化：如 "yyyy/mm/dd"
      if (format.toLowerCase().match(/^y+\/m+\/d+$/) && /^\d{8}$/.test(valueStr)) {
        return `${valueStr.substring(0, 4)}/${valueStr.substring(4, 6)}/${valueStr.substring(6, 8)}`
      }
      
      // 日期格式化：如 "yyyy年mm月dd日"
      if (format.includes('年') && format.includes('月') && /^\d{8}$/.test(valueStr)) {
        return `${valueStr.substring(0, 4)}年${valueStr.substring(4, 6)}月${valueStr.substring(6, 8)}日`
      }
      
      if (typeof value === 'number') {
        // 简单的数字格式化
        if (format.includes('0') && !format.includes('-')) {
          const decimals = (format.split('.')[1] || '').length
          return value.toFixed(decimals)
        }
        if (format.includes('%')) {
          return (value * 100).toFixed(0) + '%'
        }
        // 千位分隔符格式 #,##0
        if (format.includes(',')) {
          return value.toLocaleString('en-US')
        }
      }
      return String(value)
    },
    
    // TRIM - 去除空格
    TRIM: (args) => String(evaluateExpression(args[0]) || '').trim(),
    
    // UPPER - 转大写
    UPPER: (args) => String(evaluateExpression(args[0]) || '').toUpperCase(),
    
    // LOWER - 转小写
    LOWER: (args) => String(evaluateExpression(args[0]) || '').toLowerCase(),
    
    // -------------------- 日期函数 --------------------
    
    // TODAY - 今天日期（返回 Date 对象，便于 YEAR/MONTH/DAY 处理）
    TODAY: () => {
      const now = new Date()
      now.setHours(0, 0, 0, 0) // 只保留日期部分
      return now
    },
    
    // NOW - 当前日期时间
    NOW: () => new Date(),
    
    // YEAR - 获取年份
    YEAR: (args) => {
      const val = evaluateExpression(args[0])
      // 如果是 Date 对象
      if (val instanceof Date) return val.getFullYear()
      // 如果是字符串格式的日期 "2025-12-08"
      if (typeof val === 'string') {
        // 尝试 YYYY-MM-DD 格式
        const match = val.match(/^(\d{4})-(\d{2})-(\d{2})/)
        if (match) return parseInt(match[1], 10)
        // 尝试 Date 解析
        const date = new Date(val)
        if (!isNaN(date.getTime())) return date.getFullYear()
      }
      // 如果是 Excel 日期序列号
      if (typeof val === 'number' && val > 1000 && val < 100000) {
        // Excel 日期从 1900-01-01 开始
        const excelEpoch = new Date(1900, 0, 1)
        const date = new Date(excelEpoch.getTime() + (val - 1) * 24 * 60 * 60 * 1000)
        return date.getFullYear()
      }
      return new Date().getFullYear() // 默认返回当前年份
    },
    
    // MONTH - 获取月份
    MONTH: (args) => {
      const val = evaluateExpression(args[0])
      if (val instanceof Date) return val.getMonth() + 1
      if (typeof val === 'string') {
        const match = val.match(/^(\d{4})-(\d{2})-(\d{2})/)
        if (match) return parseInt(match[2], 10)
        const date = new Date(val)
        if (!isNaN(date.getTime())) return date.getMonth() + 1
      }
      if (typeof val === 'number' && val > 1000 && val < 100000) {
        const excelEpoch = new Date(1900, 0, 1)
        const date = new Date(excelEpoch.getTime() + (val - 1) * 24 * 60 * 60 * 1000)
        return date.getMonth() + 1
      }
      return new Date().getMonth() + 1
    },
    
    // DAY - 获取日期
    DAY: (args) => {
      const val = evaluateExpression(args[0])
      if (val instanceof Date) return val.getDate()
      if (typeof val === 'string') {
        const match = val.match(/^(\d{4})-(\d{2})-(\d{2})/)
        if (match) return parseInt(match[3], 10)
        const date = new Date(val)
        if (!isNaN(date.getTime())) return date.getDate()
      }
      if (typeof val === 'number' && val > 1000 && val < 100000) {
        const excelEpoch = new Date(1900, 0, 1)
        const date = new Date(excelEpoch.getTime() + (val - 1) * 24 * 60 * 60 * 1000)
        return date.getDate()
      }
      return new Date().getDate()
    },
    
    // -------------------- 信息函数 --------------------
    
    // ISBLANK - 是否为空
    ISBLANK: (args) => {
      const val = evaluateExpression(args[0])
      return val == null || val === ''
    },
    
    // ISNUMBER - 是否为数字
    ISNUMBER: (args) => typeof evaluateExpression(args[0]) === 'number',
    
    // ISTEXT - 是否为文本
    ISTEXT: (args) => typeof evaluateExpression(args[0]) === 'string'
  }
  
  // 条件匹配函数（支持通配符和比较运算符）
  const matchCriteria = (value, criteria) => {
    const criteriaStr = String(criteria)
    
    // 比较运算符
    if (criteriaStr.startsWith('>=')) {
      return value >= parseFloat(criteriaStr.slice(2))
    }
    if (criteriaStr.startsWith('<=')) {
      return value <= parseFloat(criteriaStr.slice(2))
    }
    if (criteriaStr.startsWith('<>')) {
      return String(value) !== criteriaStr.slice(2)
    }
    if (criteriaStr.startsWith('>')) {
      return value > parseFloat(criteriaStr.slice(1))
    }
    if (criteriaStr.startsWith('<')) {
      return value < parseFloat(criteriaStr.slice(1))
    }
    if (criteriaStr.startsWith('=')) {
      return String(value) === criteriaStr.slice(1)
    }
    
    // 通配符匹配
    if (criteriaStr.includes('*') || criteriaStr.includes('?')) {
      const regex = new RegExp('^' + criteriaStr.replace(/\*/g, '.*').replace(/\?/g, '.') + '$', 'i')
      return regex.test(String(value))
    }
    
    // 精确匹配
    return String(value) === criteriaStr || value === criteria
  }
  
  // 解析并计算表达式
  const evaluateExpression = (expr) => {
    if (expr == null) return 0
    expr = String(expr).trim()
    
    // 字符串字面量
    if ((expr.startsWith('"') && expr.endsWith('"')) || (expr.startsWith("'") && expr.endsWith("'"))) {
      return expr.slice(1, -1)
    }
    
    // 数字
    if (/^-?\d+\.?\d*$/.test(expr)) {
      return parseFloat(expr)
    }
    
    // 布尔值
    if (expr.toUpperCase() === 'TRUE') return true
    if (expr.toUpperCase() === 'FALSE') return false
    
    // 单元格引用（包括跨工作表）- 必须是完整的引用，不是表达式的一部分
    if (/^'?[^'!]*'?![A-Z]+\d+$/i.test(expr) || /^[A-Z]+\d+$/i.test(expr)) {
      return getCellValue(expr)
    }
    
    // ============================================================
    // 复合表达式处理 - 支持 FUNC1()-FUNC2()+... 格式
    // ============================================================
    
    // 将表达式分解为标记（函数调用、运算符、数字、单元格引用）
    const tokenizeExpression = (expression) => {
      const tokens = []
      let i = 0
      
      while (i < expression.length) {
        // 跳过空格
        if (expression[i] === ' ') {
          i++
          continue
        }
        
        // 运算符
        if ('+-*/'.includes(expression[i])) {
          tokens.push({ type: 'operator', value: expression[i] })
          i++
          continue
        }
        
        // 数字
        if (/\d/.test(expression[i]) || (expression[i] === '-' && i === 0)) {
          let numStr = ''
          if (expression[i] === '-') {
            numStr = '-'
            i++
          }
          while (i < expression.length && /[\d.]/.test(expression[i])) {
            numStr += expression[i]
            i++
          }
          tokens.push({ type: 'number', value: parseFloat(numStr) })
          continue
        }
        
        // 字符串
        if (expression[i] === '"') {
          let str = ''
          i++ // 跳过开始引号
          while (i < expression.length && expression[i] !== '"') {
            str += expression[i]
            i++
          }
          i++ // 跳过结束引号
          tokens.push({ type: 'string', value: str })
          continue
        }
        
        // 函数调用或单元格引用
        if (/[A-Z']/i.test(expression[i])) {
          let token = ''
          
          // 处理带引号的工作表名（如 'Sheet1'!A1）
          if (expression[i] === "'") {
            while (i < expression.length && expression[i] !== '!') {
              token += expression[i]
              i++
            }
            if (expression[i] === '!') {
              token += expression[i]
              i++
            }
          }
          
          // 继续读取字母/数字
          while (i < expression.length && /[A-Z0-9_]/i.test(expression[i])) {
            token += expression[i]
            i++
          }
          
          // 检查是否是函数调用
          if (i < expression.length && expression[i] === '(') {
            // 找到匹配的右括号
            let depth = 1
            i++ // 跳过开始括号
            let argsStr = ''
            while (i < expression.length && depth > 0) {
              if (expression[i] === '(') depth++
              else if (expression[i] === ')') depth--
              if (depth > 0) argsStr += expression[i]
              i++
            }
            
            // 调用函数
            const funcName = token.toUpperCase()
            if (functions[funcName]) {
              const args = parseFunctionArgs(argsStr)
              const result = functions[funcName](args)
              tokens.push({ type: 'value', value: result })
            } else {
              tokens.push({ type: 'value', value: 0 })
            }
          } else {
            // 单元格引用
            const cellValue = getCellValue(token)
            // 如果是字符串形式的数字，转换为数字用于计算
            if (typeof cellValue === 'string' && /^-?\d+\.?\d*$/.test(cellValue)) {
              tokens.push({ type: 'value', value: parseFloat(cellValue) })
            } else {
              tokens.push({ type: 'value', value: cellValue })
            }
          }
          continue
        }
        
        // 括号
        if (expression[i] === '(') {
          let depth = 1
          i++
          let subExpr = ''
          while (i < expression.length && depth > 0) {
            if (expression[i] === '(') depth++
            else if (expression[i] === ')') depth--
            if (depth > 0) subExpr += expression[i]
            i++
          }
          tokens.push({ type: 'value', value: evaluateExpression(subExpr) })
          continue
        }
        
        i++ // 跳过未知字符
      }
      
      return tokens
    }
    
    // 计算标记序列
    const calculateTokens = (tokens) => {
      if (tokens.length === 0) return 0
      if (tokens.length === 1) {
        const t = tokens[0]
        return t.type === 'value' || t.type === 'number' ? t.value : 0
      }
      
      // 先处理乘除
      let i = 0
      while (i < tokens.length) {
        if (tokens[i].type === 'operator' && (tokens[i].value === '*' || tokens[i].value === '/')) {
          const left = tokens[i - 1]?.value ?? 0
          const right = tokens[i + 1]?.value ?? 0
          const leftNum = typeof left === 'string' ? (parseFloat(left) || 0) : (left || 0)
          const rightNum = typeof right === 'string' ? (parseFloat(right) || 0) : (right || 0)
          
          let result
          if (tokens[i].value === '*') {
            result = leftNum * rightNum
          } else {
            result = rightNum !== 0 ? leftNum / rightNum : 0
          }
          tokens.splice(i - 1, 3, { type: 'value', value: result })
          i = Math.max(0, i - 1)
        } else {
          i++
        }
      }
      
      // 再处理加减
      i = 0
      while (i < tokens.length) {
        if (tokens[i].type === 'operator' && (tokens[i].value === '+' || tokens[i].value === '-')) {
          const left = tokens[i - 1]?.value ?? 0
          const right = tokens[i + 1]?.value ?? 0
          const leftNum = typeof left === 'string' ? (parseFloat(left) || 0) : (left || 0)
          const rightNum = typeof right === 'string' ? (parseFloat(right) || 0) : (right || 0)
          
          let result
          if (tokens[i].value === '+') {
            result = leftNum + rightNum
          } else {
            result = leftNum - rightNum
          }
          tokens.splice(i - 1, 3, { type: 'value', value: result })
          i = Math.max(0, i - 1)
        } else {
          i++
        }
      }
      
      return tokens[0]?.value ?? 0
    }
    
    // 检测是否是复合表达式（包含运算符或多个函数）
    const hasOperator = /[+\-*/]/.test(expr.replace(/'[^']+'/g, '')) // 排除工作表名中的引号
    const hasFunctionCall = /[A-Z]+\(/i.test(expr)
    
    if (hasOperator || hasFunctionCall) {
      try {
        const tokens = tokenizeExpression(expr)
        if (tokens.length > 0) {
          return calculateTokens(tokens)
        }
      } catch (e) {
        console.warn('[Formula] 表达式解析错误:', expr, e.message)
      }
    }
    
    // 比较表达式
    const compareMatch = expr.match(/^(.+)(>=|<=|<>|>|<|=)(.+)$/)
    if (compareMatch) {
      const left = evaluateExpression(compareMatch[1])
      const right = evaluateExpression(compareMatch[3])
      switch (compareMatch[2]) {
        case '>=': return left >= right
        case '<=': return left <= right
        case '<>': return left !== right
        case '>': return left > right
        case '<': return left < right
        case '=': return left === right
      }
    }
    
    return expr
  }
  
  // 主计算函数
  const evaluateFormula = (formula, ws = currentWorksheet) => {
    try {
      return evaluateExpression(formula)
    } catch (e) {
      console.warn('[Formula Engine] 计算失败:', formula, e.message)
      return null
    }
  }
  
  return { evaluateFormula, getCellValue, getRangeValues }
}

// 简单公式计算器 - 兼容旧接口
function evaluateSimpleFormula(formula, worksheet, workbook = null) {
  // 如果没有 workbook，创建一个简单的包装
  const wb = workbook || { 
    getWorksheet: () => worksheet,
    worksheets: [worksheet]
  }
  const engine = createFormulaEngine(wb, worksheet)
  return engine.evaluateFormula(formula)
}

// 解析单元格地址（如 "A1" -> { r: 0, c: 0 }）
function parseCellAddress(address) {
  const match = address.toUpperCase().match(/^([A-Z]+)(\d+)$/)
  if (!match) return null
  
  let col = 0
  for (let i = 0; i < match[1].length; i++) {
    col = col * 26 + (match[1].charCodeAt(i) - 64)
  }
  return { r: parseInt(match[2], 10) - 1, c: col - 1 }
}

// 生成列标（如 0 -> "A", 25 -> "Z", 26 -> "AA"）
function getColumnLabel(i) {
  let label = ''
  let n = i
  while (n >= 0) {
    label = String.fromCharCode((n % 26) + 65) + label
    n = Math.floor(n / 26) - 1
  }
  return label
}

// 格式化单元格地址
function formatCellAddress(r, c) {
  return `${getColumnLabel(c)}${r + 1}`
}

// 【查询】读取单元格/区域
ipcMain.handle('excel-read-cells', async (_event, filePath, sheetName, rangeOrCell) => {
  try {
    const workbook = await getWorkbook(filePath)
    const worksheet = workbook.getWorksheet(sheetName)
    if (!worksheet) {
      return { success: false, error: `工作表 "${sheetName}" 不存在` }
    }
    
    // 解析范围：可以是单个单元格 "A1" 或范围 "A1:C5"
    const parts = rangeOrCell.toUpperCase().split(':')
    const start = parseCellAddress(parts[0])
    const end = parts.length > 1 ? parseCellAddress(parts[1]) : start
    
    if (!start || !end) {
      return { success: false, error: `无效的单元格地址: ${rangeOrCell}` }
    }
    
    const cells = []
    for (let r = start.r; r <= end.r; r++) {
      for (let c = start.c; c <= end.c; c++) {
        const cell = worksheet.getCell(r + 1, c + 1)
        // 安全获取文本值
        let textValue = ''
        try {
          const v = cell.value
          if (v != null) {
            if (typeof v === 'object' && v.richText) {
              textValue = v.richText.map(rt => rt.text || '').join('')
            } else if (typeof v === 'object' && v.result != null) {
              textValue = String(v.result)
            } else if (typeof v === 'object' && v.text != null) {
              textValue = String(v.text)
            } else {
              textValue = String(v)
            }
          }
        } catch (e) {
          textValue = ''
        }
        cells.push({
          address: formatCellAddress(r, c),
          r, c,
          value: cell.value,
          text: textValue,
          formula: cell.formula,
          type: cell.type
        })
      }
    }
    
    return { success: true, cells, range: rangeOrCell }
  } catch (error) {
    console.error('[Excel Read] 失败:', error)
    return { success: false, error: error.message }
  }
})

// 【查询】搜索单元格内容
ipcMain.handle('excel-search', async (_event, filePath, sheetName, searchText, options = {}) => {
  try {
    const workbook = await getWorkbook(filePath)
    const worksheet = workbook.getWorksheet(sheetName)
    if (!worksheet) {
      return { success: false, error: `工作表 "${sheetName}" 不存在` }
    }
    
    const results = []
    const { caseSensitive = false, matchWholeCell = false } = options
    const searchLower = caseSensitive ? searchText : searchText.toLowerCase()
    
    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        // 安全获取单元格文本
        let cellText = ''
        try {
          const v = cell.value
          if (v != null) {
            if (typeof v === 'object' && v.richText) {
              cellText = v.richText.map(rt => rt.text || '').join('')
            } else if (typeof v === 'object' && v.result != null) {
              cellText = String(v.result)
            } else if (typeof v === 'object' && v.text != null) {
              cellText = String(v.text)
            } else {
              cellText = String(v)
            }
          }
        } catch (e) {
          cellText = ''
        }
        const compareText = caseSensitive ? cellText : cellText.toLowerCase()
        
        let match = false
        if (matchWholeCell) {
          match = compareText === searchLower
        } else {
          match = compareText.includes(searchLower)
        }
        
        if (match) {
          results.push({
            address: formatCellAddress(rowNumber - 1, colNumber - 1),
            r: rowNumber - 1,
            c: colNumber - 1,
            value: cell.value,
            text: cellText
          })
        }
      })
    })
    
    return { success: true, results, count: results.length }
  } catch (error) {
    console.error('[Excel Search] 失败:', error)
    return { success: false, error: error.message }
  }
})

// 【修改】写入单元格
ipcMain.handle('excel-write-cells', async (_event, filePath, sheetName, cellUpdates) => {
  try {
    // 检查文件是否被锁定
    try {
      const fd = fs.openSync(filePath, 'r+')
      fs.closeSync(fd)
    } catch (lockErr) {
      if (lockErr.code === 'EBUSY' || lockErr.code === 'EACCES') {
        return { 
          success: false, 
          error: '文件被其他程序占用（可能是 Excel 正在打开此文件）。请关闭 Excel 后重试。' 
        }
      }
    }
    
    // 清除缓存，重新加载文件
    clearWorkbookCache(filePath)
    const workbook = await getWorkbook(filePath)
    const worksheet = workbook.getWorksheet(sheetName)
    if (!worksheet) {
      return { success: false, error: `工作表 "${sheetName}" 不存在` }
    }
    
    console.log(`[Excel Write] 写入 ${cellUpdates.length} 个单元格到 ${sheetName}`)
    
    // cellUpdates: [{ address: "A1", value: "new value", style?: {...} }, ...]
    const updatedCells = []
    for (const update of cellUpdates) {
      const addr = parseCellAddress(update.address)
      if (!addr) {
        console.warn(`[Excel Write] 跳过无效地址: ${update.address}`)
        continue
      }
      
      const cell = worksheet.getCell(addr.r + 1, addr.c + 1)
      
      // 设置值（支持公式）
      if (update.value !== undefined) {
        if (typeof update.value === 'string' && update.value.startsWith('=')) {
          cell.value = { formula: update.value.slice(1) }
        } else {
          cell.value = update.value
        }
      }
      
      // 设置样式
      if (update.style) {
        if (update.style.font) {
          cell.font = { ...cell.font, ...update.style.font }
        }
        if (update.style.fill) {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: update.style.fill.fgColor || update.style.fill
          }
        }
        if (update.style.alignment) {
          cell.alignment = { ...cell.alignment, ...update.style.alignment }
        }
        if (update.style.border) {
          cell.border = { ...cell.border, ...update.style.border }
        }
        if (update.style.numFmt) {
          cell.numFmt = update.style.numFmt
        }
      }
      
      updatedCells.push(update.address)
    }
    
    // 保存文件
    await saveWorkbook(filePath)
    clearWorkbookCache(filePath) // 清除缓存以便重新读取
    
    console.log(`[Excel Write] 成功写入 ${updatedCells.length} 个单元格`)
    return { success: true, updatedCells, count: updatedCells.length }
  } catch (error) {
    console.error('[Excel Write] 失败:', error)
    return { success: false, error: error.message }
  }
})

// 【新增】插入行
ipcMain.handle('excel-insert-rows', async (_event, filePath, sheetName, startRow, count = 1, data = null) => {
  try {
    // 清除缓存，重新加载文件
    clearWorkbookCache(filePath)
    const workbook = await getWorkbook(filePath)
    const worksheet = workbook.getWorksheet(sheetName)
    if (!worksheet) {
      return { success: false, error: `工作表 "${sheetName}" 不存在` }
    }
    
    console.log(`[Excel Insert Rows] 在第 ${startRow} 行插入 ${count} 行`)
    
    // 准备要插入的行数据
    let rowsToInsert = []
    if (data && Array.isArray(data) && data.length > 0) {
      // 使用提供的数据
      rowsToInsert = data.slice(0, count)
      // 如果数据不够，填充空行
      while (rowsToInsert.length < count) {
        rowsToInsert.push([])
      }
    } else {
      // 创建空行
      for (let i = 0; i < count; i++) {
        rowsToInsert.push([])
      }
    }
    
    // ExcelJS insertRows: 第二个参数是行数据数组
    worksheet.insertRows(startRow, rowsToInsert)
    
    await saveWorkbook(filePath)
    clearWorkbookCache(filePath) // 清除缓存以便重新读取
    
    console.log(`[Excel Insert Rows] 成功插入 ${count} 行`)
    return { success: true, insertedAt: startRow, count }
  } catch (error) {
    console.error('[Excel Insert Rows] 失败:', error)
    return { success: false, error: error.message }
  }
})

// 【新增】插入列
ipcMain.handle('excel-insert-columns', async (_event, filePath, sheetName, startCol, count = 1) => {
  try {
    // 清除缓存，重新加载文件
    clearWorkbookCache(filePath)
    const workbook = await getWorkbook(filePath)
    const worksheet = workbook.getWorksheet(sheetName)
    if (!worksheet) {
      return { success: false, error: `工作表 "${sheetName}" 不存在` }
    }
    
    console.log(`[Excel Insert Columns] 在第 ${startCol} 列插入 ${count} 列`)
    
    // ExcelJS spliceColumns(start, deleteCount, ...insert)
    // 第二个参数 0 表示不删除，后面的参数是要插入的列数据
    // 每个列数据是一个数组，代表该列所有行的值
    const emptyColumns = []
    for (let i = 0; i < count; i++) {
      emptyColumns.push([]) // 空列
    }
    worksheet.spliceColumns(startCol, 0, ...emptyColumns)
    
    await saveWorkbook(filePath)
    clearWorkbookCache(filePath)
    
    console.log(`[Excel Insert Columns] 成功插入 ${count} 列`)
    return { success: true, insertedAt: startCol, count }
  } catch (error) {
    console.error('[Excel Insert Columns] 失败:', error)
    return { success: false, error: error.message }
  }
})

// 【新增】新建工作表
ipcMain.handle('excel-add-sheet', async (_event, filePath, sheetName) => {
  try {
    // 清除缓存，重新加载文件
    clearWorkbookCache(filePath)
    const workbook = await getWorkbook(filePath)
    
    // 检查是否已存在
    if (workbook.getWorksheet(sheetName)) {
      return { success: false, error: `工作表 "${sheetName}" 已存在` }
    }
    
    console.log(`[Excel Add Sheet] 新建工作表: ${sheetName}`)
    
    workbook.addWorksheet(sheetName)
    await saveWorkbook(filePath)
    clearWorkbookCache(filePath)
    
    console.log(`[Excel Add Sheet] 成功创建工作表: ${sheetName}`)
    return { success: true, sheetName }
  } catch (error) {
    console.error('[Excel Add Sheet] 失败:', error)
    return { success: false, error: error.message }
  }
})

// 【删除】删除行
ipcMain.handle('excel-delete-rows', async (_event, filePath, sheetName, startRow, count = 1) => {
  try {
    // 清除缓存，重新加载文件
    clearWorkbookCache(filePath)
    const workbook = await getWorkbook(filePath)
    const worksheet = workbook.getWorksheet(sheetName)
    if (!worksheet) {
      return { success: false, error: `工作表 "${sheetName}" 不存在` }
    }
    
    console.log(`[Excel Delete Rows] 删除第 ${startRow} 行开始的 ${count} 行`)
    
    // ExcelJS spliceRows(start, count) - 从 start 行开始删除 count 行
    worksheet.spliceRows(startRow, count)
    
    await saveWorkbook(filePath)
    clearWorkbookCache(filePath) // 清除缓存以便重新读取
    
    console.log(`[Excel Delete Rows] 成功删除 ${count} 行`)
    return { success: true, deletedFrom: startRow, count }
  } catch (error) {
    console.error('[Excel Delete Rows] 失败:', error)
    return { success: false, error: error.message }
  }
})

// 【删除】删除列
ipcMain.handle('excel-delete-columns', async (_event, filePath, sheetName, startCol, count = 1) => {
  try {
    // 清除缓存，重新加载文件
    clearWorkbookCache(filePath)
    const workbook = await getWorkbook(filePath)
    const worksheet = workbook.getWorksheet(sheetName)
    if (!worksheet) {
      return { success: false, error: `工作表 "${sheetName}" 不存在` }
    }
    
    console.log(`[Excel Delete Columns] 删除第 ${startCol} 列开始的 ${count} 列`)
    
    // ExcelJS spliceColumns(start, deleteCount)
    worksheet.spliceColumns(startCol, count)
    
    await saveWorkbook(filePath)
    clearWorkbookCache(filePath)
    
    console.log(`[Excel Delete Columns] 成功删除 ${count} 列`)
    return { success: true, deletedFrom: startCol, count }
  } catch (error) {
    console.error('[Excel Delete Columns] 失败:', error)
    return { success: false, error: error.message }
  }
})

// 【删除】删除工作表
ipcMain.handle('excel-delete-sheet', async (_event, filePath, sheetName) => {
  try {
    // 清除缓存，重新加载文件
    clearWorkbookCache(filePath)
    const workbook = await getWorkbook(filePath)
    const worksheet = workbook.getWorksheet(sheetName)
    if (!worksheet) {
      return { success: false, error: `工作表 "${sheetName}" 不存在` }
    }
    
    console.log(`[Excel Delete Sheet] 删除工作表: ${sheetName}, id: ${worksheet.id}`)
    
    workbook.removeWorksheet(worksheet.id)
    await saveWorkbook(filePath)
    clearWorkbookCache(filePath)
    
    console.log(`[Excel Delete Sheet] 成功删除工作表: ${sheetName}`)
    return { success: true, deletedSheet: sheetName }
  } catch (error) {
    console.error('[Excel Delete Sheet] 失败:', error)
    return { success: false, error: error.message }
  }
})

// 【查询】获取工作表列表
ipcMain.handle('excel-list-sheets', async (_event, filePath) => {
  try {
    const workbook = await getWorkbook(filePath)
    const sheets = []
    
    workbook.eachSheet((worksheet) => {
      sheets.push({
        name: worksheet.name,
        rowCount: worksheet.rowCount,
        columnCount: worksheet.columnCount
      })
    })
    
    return { success: true, sheets }
  } catch (error) {
    console.error('[Excel List Sheets] 失败:', error)
    return { success: false, error: error.message }
  }
})

// 【修改】合并单元格
ipcMain.handle('excel-merge-cells', async (_event, filePath, sheetName, range) => {
  try {
    // 清除缓存，重新加载文件
    clearWorkbookCache(filePath)
    const workbook = await getWorkbook(filePath)
    const worksheet = workbook.getWorksheet(sheetName)
    if (!worksheet) {
      return { success: false, error: `工作表 "${sheetName}" 不存在` }
    }
    
    console.log(`[Excel Merge Cells] 合并单元格: ${range}`)
    
    worksheet.mergeCells(range)
    await saveWorkbook(filePath)
    clearWorkbookCache(filePath)
    
    console.log(`[Excel Merge Cells] 成功合并: ${range}`)
    return { success: true, mergedRange: range }
  } catch (error) {
    console.error('[Excel Merge Cells] 失败:', error)
    return { success: false, error: error.message }
  }
})

// 【修改】取消合并单元格
ipcMain.handle('excel-unmerge-cells', async (_event, filePath, sheetName, range) => {
  try {
    // 清除缓存，重新加载文件
    clearWorkbookCache(filePath)
    const workbook = await getWorkbook(filePath)
    const worksheet = workbook.getWorksheet(sheetName)
    if (!worksheet) {
      return { success: false, error: `工作表 "${sheetName}" 不存在` }
    }
    
    console.log(`[Excel Unmerge Cells] 取消合并: ${range}`)
    
    worksheet.unMergeCells(range)
    await saveWorkbook(filePath)
    clearWorkbookCache(filePath)
    
    console.log(`[Excel Unmerge Cells] 成功取消合并: ${range}`)
    return { success: true, unmergedRange: range }
  } catch (error) {
    console.error('[Excel Unmerge Cells] 失败:', error)
    return { success: false, error: error.message }
  }
})

// 【新增】批量设置公式
ipcMain.handle('excel-set-formula', async (_event, filePath, sheetName, formulas) => {
  try {
    clearWorkbookCache(filePath)
    const workbook = await getWorkbook(filePath)
    const worksheet = workbook.getWorksheet(sheetName)
    if (!worksheet) {
      return { success: false, error: `工作表 "${sheetName}" 不存在` }
    }
    
    console.log(`[Excel Formula] 设置 ${formulas.length} 个公式到 ${sheetName}`)
    
    const setFormulas = []
    for (const item of formulas) {
      const { address, formula, numberFormat } = item
      const addr = parseCellAddress(address)
      if (!addr) continue
      
      const cell = worksheet.getCell(addr.r + 1, addr.c + 1)
      
      // 设置公式（去掉开头的 = 如果有的话）
      const formulaText = formula.startsWith('=') ? formula.slice(1) : formula
      cell.value = { formula: formulaText }
      
      // 设置数字格式（可选）
      if (numberFormat) {
        cell.numFmt = numberFormat
      }
      
      setFormulas.push({ address, formula: formulaText })
    }
    
    await saveWorkbook(filePath)
    clearWorkbookCache(filePath)
    
    console.log(`[Excel Formula] 成功设置 ${setFormulas.length} 个公式`)
    return { success: true, formulas: setFormulas, count: setFormulas.length }
  } catch (error) {
    console.error('[Excel Formula] 失败:', error)
    return { success: false, error: error.message }
  }
})

// 【新增】排序数据
ipcMain.handle('excel-sort', async (_event, filePath, sheetName, options) => {
  try {
    const { range, column, ascending = true, hasHeader = true } = options
    
    clearWorkbookCache(filePath)
    const workbook = await getWorkbook(filePath)
    const worksheet = workbook.getWorksheet(sheetName)
    if (!worksheet) {
      return { success: false, error: `工作表 "${sheetName}" 不存在` }
    }
    
    console.log(`[Excel Sort] 排序 ${sheetName} 范围 ${range} 按列 ${column}`)
    
    // 解析范围
    const rangeMatch = range.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/)
    if (!rangeMatch) {
      return { success: false, error: `无效的范围格式: ${range}` }
    }
    
    const startCol = columnToNumber(rangeMatch[1])
    const startRow = parseInt(rangeMatch[2])
    const endCol = columnToNumber(rangeMatch[3])
    const endRow = parseInt(rangeMatch[4])
    
    // 确定排序列的索引
    const sortColIndex = columnToNumber(column) - startCol
    
    // 收集数据
    const rows = []
    const dataStartRow = hasHeader ? startRow + 1 : startRow
    
    for (let r = dataStartRow; r <= endRow; r++) {
      const rowData = []
      for (let c = startCol; c <= endCol; c++) {
        const cell = worksheet.getCell(r, c)
        rowData.push({
          value: cell.value,
          style: {
            font: cell.font,
            fill: cell.fill,
            alignment: cell.alignment,
            border: cell.border,
            numFmt: cell.numFmt
          }
        })
      }
      rows.push(rowData)
    }
    
    // 排序
    rows.sort((a, b) => {
      let valA = a[sortColIndex]?.value
      let valB = b[sortColIndex]?.value
      
      // 处理公式结果
      if (valA && typeof valA === 'object' && valA.result !== undefined) valA = valA.result
      if (valB && typeof valB === 'object' && valB.result !== undefined) valB = valB.result
      
      // 处理 null/undefined
      if (valA == null && valB == null) return 0
      if (valA == null) return ascending ? 1 : -1
      if (valB == null) return ascending ? -1 : 1
      
      // 数字比较
      const numA = typeof valA === 'number' ? valA : parseFloat(valA)
      const numB = typeof valB === 'number' ? valB : parseFloat(valB)
      
      if (!isNaN(numA) && !isNaN(numB)) {
        return ascending ? numA - numB : numB - numA
      }
      
      // 字符串比较
      const strA = String(valA).toLowerCase()
      const strB = String(valB).toLowerCase()
      return ascending ? strA.localeCompare(strB, 'zh-CN') : strB.localeCompare(strA, 'zh-CN')
    })
    
    // 写回数据
    for (let i = 0; i < rows.length; i++) {
      const rowData = rows[i]
      const r = dataStartRow + i
      for (let j = 0; j < rowData.length; j++) {
        const c = startCol + j
        const cell = worksheet.getCell(r, c)
        const data = rowData[j]
        
        cell.value = data.value
        if (data.style.font) cell.font = data.style.font
        if (data.style.fill) cell.fill = data.style.fill
        if (data.style.alignment) cell.alignment = data.style.alignment
        if (data.style.border) cell.border = data.style.border
        if (data.style.numFmt) cell.numFmt = data.style.numFmt
      }
    }
    
    await saveWorkbook(filePath)
    clearWorkbookCache(filePath)
    
    console.log(`[Excel Sort] 成功排序 ${rows.length} 行`)
    return { success: true, sortedRows: rows.length, column, ascending }
  } catch (error) {
    console.error('[Excel Sort] 失败:', error)
    return { success: false, error: error.message }
  }
})

// 辅助函数：列字母转数字
function columnToNumber(col) {
  let result = 0
  for (let i = 0; i < col.length; i++) {
    result = result * 26 + (col.charCodeAt(i) - 64)
  }
  return result
}

// 【新增】设置条件格式
ipcMain.handle('excel-conditional-format', async (_event, filePath, sheetName, options) => {
  try {
    const { range, rules } = options
    
    clearWorkbookCache(filePath)
    const workbook = await getWorkbook(filePath)
    const worksheet = workbook.getWorksheet(sheetName)
    if (!worksheet) {
      return { success: false, error: `工作表 "${sheetName}" 不存在` }
    }
    
    console.log(`[Excel ConditionalFormat] 设置条件格式到 ${sheetName} 范围 ${range}`)
    
    // ExcelJS 支持的条件格式
    const conditionalFormattings = []
    
    for (const rule of rules) {
      const cfRule = {
        ref: range,
        rules: []
      }
      
      if (rule.type === 'cellIs') {
        // 单元格值条件
        cfRule.rules.push({
          type: 'cellIs',
          operator: rule.operator, // greaterThan, lessThan, equal, between, etc.
          formulae: Array.isArray(rule.value) ? rule.value : [rule.value],
          style: {
            fill: rule.fill ? {
              type: 'pattern',
              pattern: 'solid',
              bgColor: rule.fill.bgColor || rule.fill
            } : undefined,
            font: rule.font
          }
        })
      } else if (rule.type === 'colorScale') {
        // 色阶
        cfRule.rules.push({
          type: 'colorScale',
          cfvo: [
            { type: 'min' },
            { type: 'max' }
          ],
          color: [
            { argb: rule.minColor || 'FFF8696B' },
            { argb: rule.maxColor || 'FF63BE7B' }
          ]
        })
      } else if (rule.type === 'dataBar') {
        // 数据条
        cfRule.rules.push({
          type: 'dataBar',
          minLength: 0,
          maxLength: 100,
          showValue: true,
          gradient: true,
          color: { argb: rule.color || 'FF638EC6' }
        })
      } else if (rule.type === 'containsText') {
        // 包含文本
        cfRule.rules.push({
          type: 'containsText',
          operator: 'containsText',
          text: rule.text,
          style: {
            fill: rule.fill ? {
              type: 'pattern',
              pattern: 'solid',
              bgColor: rule.fill.bgColor || rule.fill
            } : undefined,
            font: rule.font
          }
        })
      }
      
      conditionalFormattings.push(cfRule)
    }
    
    // 添加条件格式
    worksheet.addConditionalFormatting(...conditionalFormattings)
    
    await saveWorkbook(filePath)
    clearWorkbookCache(filePath)
    
    console.log(`[Excel ConditionalFormat] 成功设置 ${rules.length} 条规则`)
    return { success: true, rulesApplied: rules.length }
  } catch (error) {
    console.error('[Excel ConditionalFormat] 失败:', error)
    return { success: false, error: error.message }
  }
})

// 【新增】自动填充/序列填充
ipcMain.handle('excel-auto-fill', async (_event, filePath, sheetName, options) => {
  try {
    const { sourceRange, targetRange, fillType = 'copy' } = options
    
    clearWorkbookCache(filePath)
    const workbook = await getWorkbook(filePath)
    const worksheet = workbook.getWorksheet(sheetName)
    if (!worksheet) {
      return { success: false, error: `工作表 "${sheetName}" 不存在` }
    }
    
    console.log(`[Excel AutoFill] 从 ${sourceRange} 填充到 ${targetRange}`)
    
    // 解析源范围
    const srcMatch = sourceRange.match(/([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?/)
    if (!srcMatch) {
      return { success: false, error: `无效的源范围: ${sourceRange}` }
    }
    
    const srcStartCol = columnToNumber(srcMatch[1])
    const srcStartRow = parseInt(srcMatch[2])
    const srcEndCol = srcMatch[3] ? columnToNumber(srcMatch[3]) : srcStartCol
    const srcEndRow = srcMatch[4] ? parseInt(srcMatch[4]) : srcStartRow
    
    // 解析目标范围
    const tgtMatch = targetRange.match(/([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?/)
    if (!tgtMatch) {
      return { success: false, error: `无效的目标范围: ${targetRange}` }
    }
    
    const tgtStartCol = columnToNumber(tgtMatch[1])
    const tgtStartRow = parseInt(tgtMatch[2])
    const tgtEndCol = tgtMatch[3] ? columnToNumber(tgtMatch[3]) : tgtStartCol
    const tgtEndRow = tgtMatch[4] ? parseInt(tgtMatch[4]) : tgtStartRow
    
    // 收集源数据
    const sourceData = []
    for (let r = srcStartRow; r <= srcEndRow; r++) {
      const rowData = []
      for (let c = srcStartCol; c <= srcEndCol; c++) {
        const cell = worksheet.getCell(r, c)
        rowData.push({
          value: cell.value,
          style: {
            font: cell.font,
            fill: cell.fill,
            alignment: cell.alignment,
            border: cell.border,
            numFmt: cell.numFmt
          }
        })
      }
      sourceData.push(rowData)
    }
    
    // 填充目标范围
    let filledCount = 0
    const srcRows = sourceData.length
    const srcCols = sourceData[0]?.length || 0
    
    for (let r = tgtStartRow; r <= tgtEndRow; r++) {
      for (let c = tgtStartCol; c <= tgtEndCol; c++) {
        const srcRowIdx = (r - tgtStartRow) % srcRows
        const srcColIdx = (c - tgtStartCol) % srcCols
        const srcCell = sourceData[srcRowIdx]?.[srcColIdx]
        
        if (srcCell) {
          const cell = worksheet.getCell(r, c)
          
          if (fillType === 'series' && typeof srcCell.value === 'number') {
            // 序列填充：数字递增
            const increment = r - tgtStartRow + 1
            cell.value = srcCell.value + increment
          } else if (fillType === 'formula' && srcCell.value?.formula) {
            // 公式填充：调整相对引用（简化处理）
            const rowOffset = r - srcStartRow
            const colOffset = c - srcStartCol
            let formula = srcCell.value.formula
            
            // 简单调整行号（更复杂的需要完整的公式解析器）
            formula = formula.replace(/([A-Z]+)(\d+)/g, (match, col, row) => {
              const newRow = parseInt(row) + rowOffset
              return col + newRow
            })
            
            cell.value = { formula }
          } else {
            // 复制填充
            cell.value = srcCell.value
          }
          
          // 复制样式
          if (srcCell.style.font) cell.font = srcCell.style.font
          if (srcCell.style.fill) cell.fill = srcCell.style.fill
          if (srcCell.style.alignment) cell.alignment = srcCell.style.alignment
          if (srcCell.style.border) cell.border = srcCell.style.border
          if (srcCell.style.numFmt) cell.numFmt = srcCell.style.numFmt
          
          filledCount++
        }
      }
    }
    
    await saveWorkbook(filePath)
    clearWorkbookCache(filePath)
    
    console.log(`[Excel AutoFill] 成功填充 ${filledCount} 个单元格`)
    return { success: true, filledCells: filledCount, fillType }
  } catch (error) {
    console.error('[Excel AutoFill] 失败:', error)
    return { success: false, error: error.message }
  }
})

// 【新增】设置列宽和行高
ipcMain.handle('excel-set-dimensions', async (_event, filePath, sheetName, options) => {
  try {
    const { columns = [], rows = [] } = options
    
    clearWorkbookCache(filePath)
    const workbook = await getWorkbook(filePath)
    const worksheet = workbook.getWorksheet(sheetName)
    if (!worksheet) {
      return { success: false, error: `工作表 "${sheetName}" 不存在` }
    }
    
    console.log(`[Excel Dimensions] 设置 ${columns.length} 列宽, ${rows.length} 行高`)
    
    // 设置列宽
    for (const col of columns) {
      const colNum = typeof col.column === 'string' ? columnToNumber(col.column) : col.column
      const column = worksheet.getColumn(colNum)
      if (col.width !== undefined) column.width = col.width
      if (col.hidden !== undefined) column.hidden = col.hidden
      if (col.style) column.style = col.style
    }
    
    // 设置行高
    for (const row of rows) {
      const rowObj = worksheet.getRow(row.row)
      if (row.height !== undefined) rowObj.height = row.height
      if (row.hidden !== undefined) rowObj.hidden = row.hidden
    }
    
    await saveWorkbook(filePath)
    clearWorkbookCache(filePath)
    
    return { success: true, columnsSet: columns.length, rowsSet: rows.length }
  } catch (error) {
    console.error('[Excel Dimensions] 失败:', error)
    return { success: false, error: error.message }
  }
})

// 【新增】创建图表（简化版）
ipcMain.handle('excel-add-chart', async (_event, filePath, sheetName, options) => {
  try {
    const { 
      type = 'column', // column, bar, line, pie, scatter, area
      dataRange,
      title = '',
      position = { col: 1, row: 1 },
      size = { width: 600, height: 400 }
    } = options
    
    clearWorkbookCache(filePath)
    const workbook = await getWorkbook(filePath)
    const worksheet = workbook.getWorksheet(sheetName)
    if (!worksheet) {
      return { success: false, error: `工作表 "${sheetName}" 不存在` }
    }
    
    console.log(`[Excel Chart] 添加 ${type} 图表到 ${sheetName}`)
    
    // ExcelJS 对图表的支持有限，这里我们创建一个基本的图表配置
    // 实际上 ExcelJS 不直接支持图表创建，需要通过其他方式
    // 这里我们记录图表配置，用户可以在 Excel 中手动创建
    
    // 作为替代，我们可以在指定位置添加一个注释说明
    const cell = worksheet.getCell(position.row, position.col)
    cell.note = {
      texts: [
        { text: `图表配置:\n类型: ${type}\n数据范围: ${dataRange}\n标题: ${title || '无'}`, font: { size: 10 } }
      ]
    }
    
    await saveWorkbook(filePath)
    clearWorkbookCache(filePath)
    
    // 返回图表信息（实际图表需要用 Excel 打开后手动创建）
    return { 
      success: true, 
      message: 'ExcelJS 不直接支持图表创建，已在指定位置添加配置说明。请在 Excel 中手动创建图表。',
      chartConfig: { type, dataRange, title, position, size }
    }
  } catch (error) {
    console.error('[Excel Chart] 失败:', error)
    return { success: false, error: error.message }
  }
})

// 【新增】计算公式（获取公式计算结果）
ipcMain.handle('excel-calculate', async (_event, filePath, sheetName, addresses) => {
  try {
    clearWorkbookCache(filePath)
    const workbook = await getWorkbook(filePath)
    const worksheet = workbook.getWorksheet(sheetName)
    if (!worksheet) {
      return { success: false, error: `工作表 "${sheetName}" 不存在` }
    }
    
    console.log(`[Excel Calculate] 获取 ${addresses.length} 个单元格的计算结果`)
    
    const results = []
    for (const address of addresses) {
      const addr = parseCellAddress(address)
      if (!addr) continue
      
      const cell = worksheet.getCell(addr.r + 1, addr.c + 1)
      const value = cell.value
      
      let result = {
        address,
        value: null,
        formula: null,
        type: 'unknown'
      }
      
      if (value && typeof value === 'object') {
        if (value.formula) {
          result.formula = value.formula
          result.value = value.result !== undefined ? value.result : '计算中...'
          result.type = 'formula'
        } else if (value.richText) {
          result.value = value.richText.map(t => t.text).join('')
          result.type = 'richText'
        } else if (value.hyperlink) {
          result.value = value.text || value.hyperlink
          result.type = 'hyperlink'
        }
      } else {
        result.value = value
        result.type = typeof value
      }
      
      results.push(result)
    }
    
    return { success: true, results }
  } catch (error) {
    console.error('[Excel Calculate] 失败:', error)
    return { success: false, error: error.message }
  }
})

// 【新增】创建新的 Excel 文件
ipcMain.handle('excel-create', async (_event, filePath, options = {}) => {
  try {
    const { 
      sheets = [{ name: 'Sheet1', data: [] }], 
      openAfterCreate = true,
      defaultStyle = null,  // 全局默认样式
      headerStyle = null    // 表头默认样式
    } = options
    
    console.log(`[Excel Create] 创建新文件: ${filePath}`)
    
    // 检查文件是否已存在
    if (fs.existsSync(filePath)) {
      console.log(`[Excel Create] 文件已存在，将覆盖: ${filePath}`)
    }
    
    // 创建新工作簿
    const workbook = new ExcelJS.Workbook()
    workbook.creator = 'Word-Cursor AI'
    workbook.created = new Date()
    
    // 默认表头样式（如果用户没有指定）
    const defaultHeaderStyle = headerStyle || {
      font: { bold: true, size: 12, color: { argb: 'FFFFFFFF' } },
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } },
      alignment: { horizontal: 'center', vertical: 'middle' },
      border: {
        top: { style: 'thin', color: { argb: 'FF000000' } },
        bottom: { style: 'thin', color: { argb: 'FF000000' } },
        left: { style: 'thin', color: { argb: 'FF000000' } },
        right: { style: 'thin', color: { argb: 'FF000000' } }
      }
    }
    
    // 默认数据单元格样式
    const defaultCellStyle = defaultStyle || {
      font: { size: 11 },
      alignment: { vertical: 'middle' },
      border: {
        top: { style: 'thin', color: { argb: 'FFD0D0D0' } },
        bottom: { style: 'thin', color: { argb: 'FFD0D0D0' } },
        left: { style: 'thin', color: { argb: 'FFD0D0D0' } },
        right: { style: 'thin', color: { argb: 'FFD0D0D0' } }
      }
    }
    
    // 辅助函数：解析简化的样式参数
    const parseSimpleStyle = (styleStr) => {
      if (!styleStr || typeof styleStr !== 'string') return null
      const style = {}
      // 解析类似 "bold,center,#FF0000,14" 的简化格式
      const parts = styleStr.split(',').map(s => s.trim())
      for (const part of parts) {
        if (part === 'bold') {
          style.font = style.font || {}
          style.font.bold = true
        } else if (part === 'italic') {
          style.font = style.font || {}
          style.font.italic = true
        } else if (part === 'underline') {
          style.font = style.font || {}
          style.font.underline = true
        } else if (part === 'center') {
          style.alignment = style.alignment || {}
          style.alignment.horizontal = 'center'
        } else if (part === 'left') {
          style.alignment = style.alignment || {}
          style.alignment.horizontal = 'left'
        } else if (part === 'right') {
          style.alignment = style.alignment || {}
          style.alignment.horizontal = 'right'
        } else if (part.startsWith('#')) {
          // 颜色
          style.font = style.font || {}
          style.font.color = { argb: 'FF' + part.slice(1) }
        } else if (part.startsWith('bg#')) {
          // 背景色
          style.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + part.slice(3) } }
        } else if (/^\d+$/.test(part)) {
          // 字号
          style.font = style.font || {}
          style.font.size = parseInt(part)
        }
      }
      return Object.keys(style).length > 0 ? style : null
    }
    
    // 添加工作表和数据
    for (const sheetConfig of sheets) {
      const worksheet = workbook.addWorksheet(sheetConfig.name || 'Sheet1')
      
      // 是否应用默认样式（默认开启）
      const applyDefaultStyles = sheetConfig.applyDefaultStyles !== false
      // 第一行是否为表头（默认是）
      const firstRowIsHeader = sheetConfig.firstRowIsHeader !== false
      
      // 如果有数据，填充数据
      if (sheetConfig.data && Array.isArray(sheetConfig.data)) {
        sheetConfig.data.forEach((rowData, rowIndex) => {
          if (Array.isArray(rowData)) {
            const row = worksheet.getRow(rowIndex + 1)
            const isHeaderRow = rowIndex === 0 && firstRowIsHeader
            
            // 设置行高
            if (isHeaderRow) {
              row.height = sheetConfig.headerHeight || 25
            } else {
              row.height = sheetConfig.rowHeight || 20
            }
            
            rowData.forEach((cellValue, colIndex) => {
              const cell = row.getCell(colIndex + 1)
              
              // 支持对象格式 { value: ..., style: ... } 或 { v: ..., s: ... }
              if (cellValue && typeof cellValue === 'object' && ('value' in cellValue || 'v' in cellValue)) {
                cell.value = cellValue.value ?? cellValue.v
                
                // 应用样式
                const cellStyle = cellValue.style || cellValue.s
                if (cellStyle) {
                  // 如果是字符串，解析简化格式
                  const parsedStyle = typeof cellStyle === 'string' ? parseSimpleStyle(cellStyle) : cellStyle
                  if (parsedStyle) {
                    if (parsedStyle.font) cell.font = { ...cell.font, ...parsedStyle.font }
                    if (parsedStyle.fill) cell.fill = parsedStyle.fill
                    if (parsedStyle.alignment) cell.alignment = { ...cell.alignment, ...parsedStyle.alignment }
                    if (parsedStyle.border) cell.border = parsedStyle.border
                    if (parsedStyle.numFmt) cell.numFmt = parsedStyle.numFmt
                  }
                }
              } else {
                // 检测公式字符串（以=开头）
                if (typeof cellValue === 'string' && cellValue.startsWith('=')) {
                  cell.value = { formula: cellValue.slice(1) }
                } else {
                  cell.value = cellValue
                }
              }
              
              // 应用默认样式
              if (applyDefaultStyles) {
                if (isHeaderRow) {
                  // 表头样式（如果单元格没有自定义样式）
                  if (!cell.font || !cell.font.bold) {
                    cell.font = { ...defaultHeaderStyle.font, ...cell.font }
                  }
                  if (!cell.fill) {
                    cell.fill = defaultHeaderStyle.fill
                  }
                  if (!cell.alignment) {
                    cell.alignment = defaultHeaderStyle.alignment
                  }
                  if (!cell.border) {
                    cell.border = defaultHeaderStyle.border
                  }
                } else {
                  // 数据行样式
                  if (!cell.font) {
                    cell.font = defaultCellStyle.font
                  }
                  if (!cell.alignment) {
                    cell.alignment = defaultCellStyle.alignment
                  }
                  if (!cell.border) {
                    cell.border = defaultCellStyle.border
                  }
                }
              }
            })
            row.commit()
          }
        })
      }
      
      // 设置列宽（如果提供）
      if (sheetConfig.columnWidths && Array.isArray(sheetConfig.columnWidths)) {
        sheetConfig.columnWidths.forEach((width, index) => {
          if (width) {
            worksheet.getColumn(index + 1).width = width
          }
        })
      } else if (sheetConfig.data && sheetConfig.data.length > 0) {
        // 自动计算列宽
        const firstRow = sheetConfig.data[0]
        if (Array.isArray(firstRow)) {
          firstRow.forEach((_, colIndex) => {
            // 根据内容计算列宽，最小10，最大50
            let maxWidth = 10
            sheetConfig.data.forEach(rowData => {
              if (Array.isArray(rowData) && rowData[colIndex] != null) {
                const val = rowData[colIndex]
                const text = typeof val === 'object' ? String(val.value ?? val.v ?? '') : String(val)
                // 中文字符算2个宽度
                const len = text.split('').reduce((acc, char) => acc + (char.charCodeAt(0) > 127 ? 2 : 1), 0)
                maxWidth = Math.max(maxWidth, Math.min(len + 2, 50))
              }
            })
            worksheet.getColumn(colIndex + 1).width = maxWidth
          })
        }
      }
      
      // 设置合并单元格（如果提供）
      if (sheetConfig.merges && Array.isArray(sheetConfig.merges)) {
        sheetConfig.merges.forEach(range => {
          try {
            worksheet.mergeCells(range)
          } catch (e) {
            console.warn(`[Excel Create] 合并单元格失败: ${range}`, e.message)
          }
        })
      }
      
      // 冻结表头
      if (firstRowIsHeader && sheetConfig.freezeHeader !== false) {
        worksheet.views = [{ state: 'frozen', ySplit: 1 }]
      }
    }
    
    // 确保目录存在
    const dir = path.dirname(filePath)
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true })
    }
    
    // 保存文件
    await workbook.xlsx.writeFile(filePath)
    
    console.log(`[Excel Create] 文件创建成功: ${filePath}`)
    
    return { 
      success: true, 
      filePath,
      sheetsCreated: sheets.map(s => s.name || 'Sheet1'),
      openAfterCreate
    }
  } catch (error) {
    console.error('[Excel Create] 失败:', error)
    return { success: false, error: error.message }
  }
})

// 关闭文件时清除缓存
ipcMain.handle('excel-close', async (_event, filePath) => {
  clearWorkbookCache(filePath)
  return { success: true }
})

// 重新加载 Excel 文件（刷新缓存）
ipcMain.handle('excel-reload', async (_event, filePath) => {
  clearWorkbookCache(filePath)
  // 触发重新打开
  return await ipcMain.handlers.get('excel-open')({ sender: mainWindow.webContents }, filePath)
})

// 【新增】设置自动筛选 (AutoFilter)
ipcMain.handle('excel-set-filter', async (_event, filePath, sheetName, options) => {
  try {
    const { range, remove = false } = options || {}
    
    clearWorkbookCache(filePath)
    const workbook = await getWorkbook(filePath)
    const worksheet = workbook.getWorksheet(sheetName)
    if (!worksheet) {
      return { success: false, error: `工作表 "${sheetName}" 不存在` }
    }
    
    if (remove) {
      worksheet.autoFilter = undefined
      console.log(`[Excel Filter] 清除 ${sheetName} 的自动筛选`)
    } else if (range) {
      worksheet.autoFilter = range
      console.log(`[Excel Filter] 设置 ${sheetName} 的自动筛选范围: ${range}`)
    } else {
      // 如果没有指定范围，自动检测数据范围
      const dimensions = worksheet.dimensions
      if (dimensions) {
        const autoRange = `${dimensions.top}:${dimensions.bottom}`.replace(/(\d+):(\d+)/, (m, t, b) => {
          const topAddr = worksheet.getCell(parseInt(t), 1).address
          const bottomAddr = worksheet.getCell(parseInt(t), dimensions.right).address
          return `${topAddr}:${bottomAddr}`
        })
        worksheet.autoFilter = { from: dimensions.tl, to: { row: 1, col: dimensions.right } }
        console.log(`[Excel Filter] 自动设置筛选范围`)
      }
    }
    
    await saveWorkbook(filePath)
    clearWorkbookCache(filePath)
    
    return { 
      success: true, 
      message: remove ? '已清除自动筛选' : `已设置自动筛选范围: ${range || '自动检测'}`
    }
  } catch (error) {
    console.error('[Excel Filter] 失败:', error)
    return { success: false, error: error.message }
  }
})

// 【新增】设置数据验证 (Data Validation)
ipcMain.handle('excel-set-validation', async (_event, filePath, sheetName, options) => {
  try {
    const { 
      range, 
      type = 'list', // list, whole, decimal, date, textLength
      values,        // 对于 list 类型
      min,           // 对于数值类型
      max,           // 对于数值类型
      allowBlank = true,
      showError = true,
      errorTitle = '输入错误',
      errorMessage = '请输入有效的值',
      remove = false
    } = options || {}
    
    if (!range) {
      return { success: false, error: '请指定单元格范围 (range)' }
    }
    
    clearWorkbookCache(filePath)
    const workbook = await getWorkbook(filePath)
    const worksheet = workbook.getWorksheet(sheetName)
    if (!worksheet) {
      return { success: false, error: `工作表 "${sheetName}" 不存在` }
    }
    
    // 解析范围并应用到每个单元格
    const rangeMatch = range.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/i)
    if (!rangeMatch && !range.match(/^[A-Z]+\d+$/i)) {
      return { success: false, error: `无效的范围格式: ${range}` }
    }
    
    const applyValidation = (cell) => {
      if (remove) {
        cell.dataValidation = undefined
        return
      }
      
      const validation = {
        type: type,
        allowBlank: allowBlank,
        showErrorMessage: showError,
        errorTitle: errorTitle,
        error: errorMessage
      }
      
      if (type === 'list' && values) {
        // 列表类型
        const listValues = Array.isArray(values) ? values : [values]
        validation.formulae = ['"' + listValues.join(',') + '"']
        validation.showDropDown = true
      } else if (type === 'whole' || type === 'decimal') {
        // 数值类型
        validation.operator = 'between'
        validation.formulae = [min !== undefined ? min : 0, max !== undefined ? max : 999999999]
      } else if (type === 'textLength') {
        // 文本长度
        validation.operator = 'between'
        validation.formulae = [min !== undefined ? min : 0, max !== undefined ? max : 255]
      }
      
      cell.dataValidation = validation
    }
    
    if (rangeMatch) {
      // 范围格式 A1:B10
      const startCol = rangeMatch[1].toUpperCase()
      const startRow = parseInt(rangeMatch[2])
      const endCol = rangeMatch[3].toUpperCase()
      const endRow = parseInt(rangeMatch[4])
      
      for (let row = startRow; row <= endRow; row++) {
        for (let colCode = startCol.charCodeAt(0); colCode <= endCol.charCodeAt(0); colCode++) {
          const col = String.fromCharCode(colCode)
          const cell = worksheet.getCell(`${col}${row}`)
          applyValidation(cell)
        }
      }
    } else {
      // 单个单元格
      const cell = worksheet.getCell(range)
      applyValidation(cell)
    }
    
    await saveWorkbook(filePath)
    clearWorkbookCache(filePath)
    
    console.log(`[Excel Validation] ${remove ? '清除' : '设置'}数据验证: ${range}, 类型: ${type}`)
    
    return { 
      success: true, 
      message: remove ? `已清除 ${range} 的数据验证` : `已设置 ${range} 的${type === 'list' ? '下拉列表' : '数据'}验证`
    }
  } catch (error) {
    console.error('[Excel Validation] 失败:', error)
    return { success: false, error: error.message }
  }
})

// 【新增】设置超链接 (Hyperlink)
ipcMain.handle('excel-set-hyperlink', async (_event, filePath, sheetName, options) => {
  try {
    const { 
      cell, 
      url, 
      text,
      tooltip,
      remove = false
    } = options || {}
    
    if (!cell) {
      return { success: false, error: '请指定单元格地址 (cell)' }
    }
    
    clearWorkbookCache(filePath)
    const workbook = await getWorkbook(filePath)
    const worksheet = workbook.getWorksheet(sheetName)
    if (!worksheet) {
      return { success: false, error: `工作表 "${sheetName}" 不存在` }
    }
    
    const targetCell = worksheet.getCell(cell)
    
    if (remove) {
      // 清除超链接，保留文本
      const currentText = targetCell.text || targetCell.value
      targetCell.value = currentText
      targetCell.font = { ...targetCell.font, color: undefined, underline: false }
    } else {
      if (!url) {
        return { success: false, error: '请指定链接地址 (url)' }
      }
      
      // 设置超链接
      targetCell.value = {
        text: text || url,
        hyperlink: url,
        tooltip: tooltip || url
      }
      
      // 设置超链接样式
      targetCell.font = {
        ...targetCell.font,
        color: { argb: 'FF0000FF' },
        underline: true
      }
    }
    
    await saveWorkbook(filePath)
    clearWorkbookCache(filePath)
    
    console.log(`[Excel Hyperlink] ${remove ? '清除' : '设置'}超链接: ${cell}`)
    
    return { 
      success: true, 
      message: remove ? `已清除 ${cell} 的超链接` : `已在 ${cell} 设置超链接: ${url}`
    }
  } catch (error) {
    console.error('[Excel Hyperlink] 失败:', error)
    return { success: false, error: error.message }
  }
})

// 【新增】查找替换 (Find and Replace)
ipcMain.handle('excel-find-replace', async (_event, filePath, sheetName, options) => {
  try {
    const { 
      find, 
      replace = '',
      matchCase = false,
      matchWholeCell = false,
      allSheets = false
    } = options || {}
    
    if (!find) {
      return { success: false, error: '请指定要查找的内容 (find)' }
    }
    
    clearWorkbookCache(filePath)
    const workbook = await getWorkbook(filePath)
    
    let totalCount = 0
    const results = []
    
    const processSheet = (worksheet) => {
      let sheetCount = 0
      
      worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
          let cellValue = cell.value
          
          // 处理富文本
          if (cellValue && typeof cellValue === 'object' && cellValue.richText) {
            cellValue = cellValue.richText.map(r => r.text).join('')
          }
          
          // 处理超链接
          if (cellValue && typeof cellValue === 'object' && cellValue.text) {
            cellValue = cellValue.text
          }
          
          if (typeof cellValue === 'string') {
            const searchValue = matchCase ? find : find.toLowerCase()
            const compareValue = matchCase ? cellValue : cellValue.toLowerCase()
            
            let shouldReplace = false
            if (matchWholeCell) {
              shouldReplace = compareValue === searchValue
            } else {
              shouldReplace = compareValue.includes(searchValue)
            }
            
            if (shouldReplace) {
              // 执行替换
              const regex = new RegExp(
                find.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'),
                matchCase ? 'g' : 'gi'
              )
              
              if (matchWholeCell) {
                cell.value = replace
              } else {
                cell.value = cellValue.replace(regex, replace)
              }
              
              sheetCount++
              results.push({
                sheet: worksheet.name,
                cell: cell.address,
                oldValue: cellValue,
                newValue: cell.value
              })
            }
          }
        })
      })
      
      return sheetCount
    }
    
    if (allSheets) {
      workbook.eachSheet((worksheet) => {
        totalCount += processSheet(worksheet)
      })
    } else {
      const worksheet = workbook.getWorksheet(sheetName)
      if (!worksheet) {
        return { success: false, error: `工作表 "${sheetName}" 不存在` }
      }
      totalCount = processSheet(worksheet)
    }
    
    if (totalCount > 0) {
      await saveWorkbook(filePath)
      clearWorkbookCache(filePath)
    }
    
    console.log(`[Excel Find/Replace] 替换了 ${totalCount} 处: "${find}" → "${replace}"`)
    
    return { 
      success: true, 
      count: totalCount,
      message: totalCount > 0 
        ? `已将 ${totalCount} 处 "${find}" 替换为 "${replace}"`
        : `未找到 "${find}"`,
      details: results.slice(0, 20) // 最多返回20条详情
    }
  } catch (error) {
    console.error('[Excel Find/Replace] 失败:', error)
    return { success: false, error: error.message }
  }
})

// 【新增】插入图表（生成图片版本 - 使用 QuickChart API）
ipcMain.handle('excel-insert-chart', async (_event, filePath, sheetName, options) => {
  try {
    const { 
      type = 'column', // column, bar, line, pie, area, scatter, doughnut
      dataRange,
      title = '',
      position = 'E1',
      width = 500,
      height = 300,
      backgroundColor = '#ffffff'
    } = options || {}
    
    if (!dataRange) {
      return { success: false, error: '请指定数据范围 (dataRange)' }
    }
    
    clearWorkbookCache(filePath)
    const workbook = await getWorkbook(filePath)
    const worksheet = workbook.getWorksheet(sheetName)
    if (!worksheet) {
      return { success: false, error: `工作表 "${sheetName}" 不存在` }
    }
    
    console.log(`[Excel Chart] 图表请求: 类型=${type}, 数据=${dataRange}, 位置=${position}`)
    
    // 1. 解析数据范围并读取数据
    const rangeMatch = dataRange.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/i)
    if (!rangeMatch) {
      return { success: false, error: `无效的数据范围格式: ${dataRange}` }
    }
    
    const startCol = rangeMatch[1].toUpperCase()
    const startRow = parseInt(rangeMatch[2])
    const endCol = rangeMatch[3].toUpperCase()
    const endRow = parseInt(rangeMatch[4])
    
    // 读取数据
    const labels = []
    const datasets = []
    const dataColumns = {}
    
    // 假设第一行是标题，第一列是标签
    for (let row = startRow; row <= endRow; row++) {
      const labelCell = worksheet.getCell(`${startCol}${row}`)
      let labelValue = labelCell.value
      if (labelValue && typeof labelValue === 'object') {
        labelValue = labelValue.text || labelValue.result || String(labelValue)
      }
      
      if (row === startRow) {
        // 第一行是系列标题
        for (let colCode = startCol.charCodeAt(0) + 1; colCode <= endCol.charCodeAt(0); colCode++) {
          const col = String.fromCharCode(colCode)
          const headerCell = worksheet.getCell(`${col}${row}`)
          let headerValue = headerCell.value
          if (headerValue && typeof headerValue === 'object') {
            headerValue = headerValue.text || headerValue.result || String(headerValue)
          }
          dataColumns[col] = {
            label: headerValue || `系列${col}`,
            data: []
          }
        }
      } else {
        // 数据行
        labels.push(labelValue || `行${row}`)
        for (let colCode = startCol.charCodeAt(0) + 1; colCode <= endCol.charCodeAt(0); colCode++) {
          const col = String.fromCharCode(colCode)
          const dataCell = worksheet.getCell(`${col}${row}`)
          let cellValue = dataCell.value
          if (cellValue && typeof cellValue === 'object') {
            cellValue = cellValue.result || cellValue.text || 0
          }
          const numValue = typeof cellValue === 'number' ? cellValue : parseFloat(cellValue) || 0
          if (dataColumns[col]) {
            dataColumns[col].data.push(numValue)
          }
        }
      }
    }
    
    // 构建 datasets
    const colors = [
      'rgba(54, 162, 235, 0.8)',
      'rgba(255, 99, 132, 0.8)',
      'rgba(75, 192, 192, 0.8)',
      'rgba(255, 206, 86, 0.8)',
      'rgba(153, 102, 255, 0.8)',
      'rgba(255, 159, 64, 0.8)',
      'rgba(199, 199, 199, 0.8)',
      'rgba(83, 102, 255, 0.8)'
    ]
    
    const borderColors = colors.map(c => c.replace('0.8', '1'))
    
    let colorIndex = 0
    for (const col in dataColumns) {
      datasets.push({
        label: dataColumns[col].label,
        data: dataColumns[col].data,
        backgroundColor: type === 'pie' || type === 'doughnut' 
          ? colors.slice(0, dataColumns[col].data.length)
          : colors[colorIndex % colors.length],
        borderColor: type === 'pie' || type === 'doughnut'
          ? borderColors.slice(0, dataColumns[col].data.length)
          : borderColors[colorIndex % borderColors.length],
        borderWidth: 1
      })
      colorIndex++
    }
    
    // 如果只有一列数据（没有标题行），直接用第一列作为标签
    if (datasets.length === 0 && labels.length > 0) {
      // 单列数据，第一列作为标签，需要重新解析
      labels.length = 0
      const singleData = []
      for (let row = startRow; row <= endRow; row++) {
        const labelCell = worksheet.getCell(`${startCol}${row}`)
        const valueCell = worksheet.getCell(`${endCol}${row}`)
        let labelValue = labelCell.value
        let dataValue = valueCell.value
        
        if (labelValue && typeof labelValue === 'object') {
          labelValue = labelValue.text || labelValue.result || String(labelValue)
        }
        if (dataValue && typeof dataValue === 'object') {
          dataValue = dataValue.result || dataValue.text || 0
        }
        
        labels.push(labelValue || `项${row}`)
        singleData.push(typeof dataValue === 'number' ? dataValue : parseFloat(dataValue) || 0)
      }
      
      datasets.push({
        label: title || '数据',
        data: singleData,
        backgroundColor: type === 'pie' || type === 'doughnut'
          ? colors.slice(0, singleData.length)
          : colors[0],
        borderColor: type === 'pie' || type === 'doughnut'
          ? borderColors.slice(0, singleData.length)
          : borderColors[0],
        borderWidth: 1
      })
    }
    
    console.log(`[Excel Chart] 标签: ${labels.length} 个, 数据系列: ${datasets.length} 个`)
    
    // 2. 构建 QuickChart 配置
    const chartTypeMap = {
      'column': 'bar',
      'bar': 'horizontalBar',
      'line': 'line',
      'pie': 'pie',
      'doughnut': 'doughnut',
      'area': 'line',
      'scatter': 'scatter'
    }
    
    const chartConfig = {
      type: chartTypeMap[type] || 'bar',
      data: {
        labels: labels,
        datasets: datasets
      },
      options: {
        title: {
          display: !!title,
          text: title,
          fontSize: 16
        },
        legend: {
          display: datasets.length > 1 || type === 'pie' || type === 'doughnut'
        },
        plugins: {
          datalabels: {
            display: type === 'pie' || type === 'doughnut',
            color: '#fff',
            font: { weight: 'bold' }
          }
        }
      }
    }
    
    // 面积图特殊处理
    if (type === 'area') {
      chartConfig.data.datasets = chartConfig.data.datasets.map(ds => ({
        ...ds,
        fill: true
      }))
    }
    
    // 3. 调用 QuickChart API 生成图片
    // 使用 GET 方法更稳定
    const chartConfigEncoded = encodeURIComponent(JSON.stringify(chartConfig))
    const quickChartUrl = `https://quickchart.io/chart?c=${chartConfigEncoded}&w=${width}&h=${height}&bkg=${encodeURIComponent(backgroundColor)}&f=png`
    
    console.log('[Excel Chart] 调用 QuickChart API...')
    console.log('[Excel Chart] 图表配置:', JSON.stringify(chartConfig).substring(0, 200))
    
    const response = await fetch(quickChartUrl)
    
    if (!response.ok) {
      const errorText = await response.text()
      console.error('[Excel Chart] API 错误:', errorText)
      throw new Error(`QuickChart API 返回错误: ${response.status} ${response.statusText}`)
    }
    
    const arrayBuffer = await response.arrayBuffer()
    const imageBuffer = Buffer.from(arrayBuffer)
    
    if (imageBuffer.length < 1000) {
      // 图片太小，可能是错误响应
      console.error('[Excel Chart] 图片数据太小，可能生成失败:', imageBuffer.length)
      throw new Error('图表生成失败：返回数据异常')
    }
    
    console.log(`[Excel Chart] 图片生成成功, 大小: ${imageBuffer.length} bytes`)
    
    // 4. 保存图片到临时文件（ExcelJS 对 buffer 支持有时不稳定）
    const tempDir = require('os').tmpdir()
    const tempImagePath = path.join(tempDir, `chart_${Date.now()}.png`)
    fs.writeFileSync(tempImagePath, imageBuffer)
    console.log(`[Excel Chart] 临时图片保存到: ${tempImagePath}`)
    
    // 5. 将图片插入到 Excel（使用文件路径而不是 buffer）
    const imageId = workbook.addImage({
      filename: tempImagePath,
      extension: 'png'
    })
    
    // 解析位置
    const posMatch = position.match(/([A-Z]+)(\d+)/i)
    if (!posMatch) {
      // 清理临时文件
      try { fs.unlinkSync(tempImagePath) } catch {}
      return { success: false, error: `无效的位置格式: ${position}` }
    }
    
    const posCol = posMatch[1].toUpperCase().charCodeAt(0) - 64 // A=1, B=2...
    const posRow = parseInt(posMatch[2])
    
    // 使用 tl + br 方式定位（更稳定）
    // 计算结束位置
    const imgEndCol = posCol - 1 + Math.ceil(width / 72)  // 假设每列约 72 像素
    const imgEndRow = posRow - 1 + Math.ceil(height / 20) // 假设每行约 20 像素
    
    worksheet.addImage(imageId, {
      tl: { col: posCol - 1, row: posRow - 1 },
      br: { col: imgEndCol, row: imgEndRow }
    })
    
    await saveWorkbook(filePath)
    clearWorkbookCache(filePath)
    
    // 清理临时文件
    try { fs.unlinkSync(tempImagePath) } catch {}
    
    console.log(`[Excel Chart] 图表图片已插入到 ${position}`)
    
    return { 
      success: true, 
      message: `已在 ${position} 插入${type === 'column' ? '柱状' : type === 'line' ? '折线' : type === 'pie' ? '饼' : type}图`,
      chartConfig: { type, dataRange, title, position, width, height, labelsCount: labels.length, datasetsCount: datasets.length }
    }
  } catch (error) {
    console.error('[Excel Chart] 失败:', error)
    return { success: false, error: error.message }
  }
})

// HTML 转义函数
function escapeHtml(text) {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;')
}

// 写入文件
ipcMain.handle('write-file', async (event, filePath, content) => {
  try {
    // 确保目录存在
    const dir = path.dirname(filePath)
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true })
    }
    
    fs.writeFileSync(filePath, content, 'utf-8')
    return { success: true }
  } catch (error) {
    return { success: false, error: error.message }
  }
})

// 写入二进制文件（用于 docx）
ipcMain.handle('write-binary-file', async (event, filePath, base64Data) => {
  try {
    const buffer = Buffer.from(base64Data, 'base64')
    fs.writeFileSync(filePath, buffer)
    return { success: true }
  } catch (error) {
    return { success: false, error: error.message }
  }
})

// 保存文件对话框
ipcMain.handle('save-file-dialog', async (event, defaultName) => {
  const result = await dialog.showSaveDialog(mainWindow, {
    defaultPath: defaultName,
    filters: [
      { name: 'Word 文档', extensions: ['docx'] },
      { name: 'Markdown', extensions: ['md'] },
      { name: '文本文件', extensions: ['txt'] },
      { name: '所有文件', extensions: ['*'] }
    ]
  })
  
  if (result.canceled) return null
  return result.filePath
})

// 创建新文件
ipcMain.handle('create-file', async (event, folderPath, fileName, content = '') => {
  try {
    const filePath = path.join(folderPath, fileName)
    fs.writeFileSync(filePath, content, 'utf-8')
    return { success: true, path: filePath }
  } catch (error) {
    return { success: false, error: error.message }
  }
})

// 删除文件
ipcMain.handle('delete-file', async (event, filePath) => {
  try {
    fs.unlinkSync(filePath)
    return { success: true }
  } catch (error) {
    return { success: false, error: error.message }
  }
})

// 重命名文件
ipcMain.handle('rename-file', async (event, oldPath, newPath) => {
  try {
    fs.renameSync(oldPath, newPath)
    return { success: true }
  } catch (error) {
    return { success: false, error: error.message }
  }
})

// 在系统文件管理器中显示
ipcMain.handle('show-in-folder', async (event, filePath) => {
  shell.showItemInFolder(filePath)
  return { success: true }
})

// 获取文件信息
ipcMain.handle('get-file-info', async (event, filePath) => {
  try {
    const stats = fs.statSync(filePath)
    return {
      success: true,
      data: {
        size: stats.size,
        created: stats.birthtime,
        modified: stats.mtime,
        isFile: stats.isFile(),
        isDirectory: stats.isDirectory()
      }
    }
  } catch (error) {
    return { success: false, error: error.message }
  }
})

// ==================== 环境变量配置读写 ====================

// 读取 .env 文件中的设置
ipcMain.handle('get-env-settings', async () => {
  const envPath = path.join(__dirname, '..', '.env')
  const result = {}
  try {
    if (fs.existsSync(envPath)) {
      const envContent = fs.readFileSync(envPath, 'utf-8')
      const parsed = dotenv.parse(envContent)
      // 映射 env 变量到设置对象
      if (parsed.API_KEY) result.apiKey = parsed.API_KEY
      if (parsed.BASE_URL) result.baseUrl = parsed.BASE_URL
      if (parsed.MODEL) result.model = parsed.MODEL
      if (parsed.TEMPERATURE) result.temperature = parseFloat(parsed.TEMPERATURE)
      if (parsed.MAX_TOKENS) result.maxTokens = parseInt(parsed.MAX_TOKENS, 10)
      if (parsed.DASHSCOPE_API_KEY) result.dashscopeApiKey = parsed.DASHSCOPE_API_KEY
      if (parsed.OPENROUTER_API_KEY) result.openRouterApiKey = parsed.OPENROUTER_API_KEY
      if (parsed.PPT_IMAGE_MODEL) result.pptImageModel = parsed.PPT_IMAGE_MODEL
      if (parsed.BRAVE_API_KEY) result.braveApiKey = parsed.BRAVE_API_KEY
      // 本地模型
      const hasLocal = parsed.LOCAL_MODEL_BASE_URL || parsed.LOCAL_MODEL_NAME
      if (hasLocal) {
        result.localModel = {
          enabled: parsed.LOCAL_MODEL_ENABLED !== 'false',
          baseUrl: parsed.LOCAL_MODEL_BASE_URL || '',
          model: parsed.LOCAL_MODEL_NAME || '',
          apiKey: parsed.LOCAL_MODEL_API_KEY || '',
        }
      }
      // 记忆系统
      if (parsed.MEMORY_ENABLED !== undefined) result.memoryEnabled = parsed.MEMORY_ENABLED !== 'false'
      if (parsed.MEMORY_TOP_K) result.memoryTopK = parseInt(parsed.MEMORY_TOP_K, 10)
      if (parsed.MEMORY_MAX_CHARS) result.memoryMaxChars = parseInt(parsed.MEMORY_MAX_CHARS, 10)
      if (parsed.MEMORY_FLUSH_THRESHOLD) result.memoryFlushThresholdChars = parseInt(parsed.MEMORY_FLUSH_THRESHOLD, 10)
    }
  } catch (e) {
    console.warn('读取 .env 设置失败:', e)
  }
  return result
})

// 保存设置到 .env 文件
ipcMain.handle('save-env-settings', async (_event, settings) => {
  const envPath = path.join(__dirname, '..', '.env')
  try {
    const lines = [
      '# Word-Cursor 配置文件（由设置界面自动生成）',
      '# 复制到新电脑即可直接使用，无需重新配置',
      '',
      '# ===== LLM 主模型 =====',
      `API_KEY=${settings.apiKey || ''}`,
      `BASE_URL=${settings.baseUrl || ''}`,
      `MODEL=${settings.model || ''}`,
      `TEMPERATURE=${settings.temperature ?? 1}`,
      `MAX_TOKENS=${settings.maxTokens ?? 4096}`,
      '',
      '# ===== PPT 图像生成 =====',
      `DASHSCOPE_API_KEY=${settings.dashscopeApiKey || ''}`,
      `OPENROUTER_API_KEY=${settings.openRouterApiKey || ''}`,
      `PPT_IMAGE_MODEL=${settings.pptImageModel || 'gemini-image'}`,
      '',
      '# ===== 联网搜索 =====',
      `BRAVE_API_KEY=${settings.braveApiKey || ''}`,
      '',
      '# ===== 本地模型（Tab 补全）=====',
      `LOCAL_MODEL_ENABLED=${settings.localModel?.enabled ?? true}`,
      `LOCAL_MODEL_BASE_URL=${settings.localModel?.baseUrl || ''}`,
      `LOCAL_MODEL_NAME=${settings.localModel?.model || ''}`,
      `LOCAL_MODEL_API_KEY=${settings.localModel?.apiKey || ''}`,
      '',
      '# ===== 记忆系统 =====',
      `MEMORY_ENABLED=${settings.memoryEnabled ?? true}`,
      `MEMORY_TOP_K=${settings.memoryTopK ?? 5}`,
      `MEMORY_MAX_CHARS=${settings.memoryMaxChars ?? 2000}`,
      `MEMORY_FLUSH_THRESHOLD=${settings.memoryFlushThresholdChars ?? 12000}`,
      '',
    ]
    fs.writeFileSync(envPath, lines.join('\n'), 'utf-8')
    // 同步更新当前进程环境变量
    dotenv.config({ path: envPath, override: true })
    return { success: true }
  } catch (e) {
    console.warn('保存 .env 设置失败:', e)
    return { success: false, error: e.message }
  }
})

// ==================== 模板文档替换（保留完整格式）====================

// 使用 docxtemplater 进行模板替换 - 完美保留所有格式
ipcMain.handle('fill-template', async (event, { templatePath, outputPath, replacements }) => {
  try {
    console.log('模板替换开始:', templatePath, '->', outputPath)
    console.log('替换内容:', replacements)
    
    // 读取模板文件
    const content = fs.readFileSync(templatePath, 'binary')
    const zip = new PizZip(content)
    
    // 创建 docxtemplater 实例
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
      // 自定义分隔符（可选，默认是 { }）
      delimiters: { start: '{{', end: '}}' }
    })
    
    // 设置替换数据
    doc.setData(replacements)
    
    // 渲染文档
    doc.render()
    
    // 生成输出
    const buf = doc.getZip().generate({
      type: 'nodebuffer',
      compression: 'DEFLATE'
    })
    
    // 写入文件
    fs.writeFileSync(outputPath, buf)
    
    console.log('模板替换成功:', outputPath)
    return { success: true, path: outputPath }
  } catch (error) {
    console.error('模板替换失败:', error)
    return { success: false, error: error.message }
  }
})

// 直接在 docx 文件中进行文本替换（不需要占位符，直接搜索替换）
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

// ==================== Web 搜索（Brave MCP） ====================
ipcMain.handle('web-search', async (event, options = {}) => {
  const query = (options.query || '').trim()
  if (!query) {
    return { success: false, message: '缺少 query 参数' }
  }

  try {
    const result = await performBraveWebSearch(query, {
      locale: options.locale,
      region: options.region,
      num: options.num,
      braveApiKey: options.braveApiKey,
    })
    return result
  } catch (error) {
    console.error('Brave Web 搜索失败:', error)
    return { success: false, message: error.message || 'Brave Web 搜索失败，请在设置中配置 Brave Search API Key' }
  }
})

// ==================== AI 请求代理（绕开渲染进程 HTTP/2） ====================
// 说明：渲染进程使用浏览器 fetch（Chromium/HTTP2）时，部分网络环境会出现 ERR_HTTP2_PING_FAILED。
// 将请求放到主进程使用 Node fetch（通常 HTTP/1.1）可显著提升稳定性，并且不受 CORS 影响。
const aiAbortControllers = new Map() // requestId -> AbortController

ipcMain.handle('ai-chat-completions', async (event, payload = {}) => {
  const {
    requestId,
    baseUrl,
    apiKey,
    model,
    messages,
    temperature,
    maxTokens,
  } = payload || {}
  if (!requestId) return { success: false, error: '缺少 requestId' }
  if (!baseUrl) return { success: false, error: '缺少 baseUrl' }
  if (!model) return { success: false, error: '缺少 model' }
  if (!Array.isArray(messages) || messages.length === 0) return { success: false, error: '缺少 messages' }

  const controller = new AbortController()
  aiAbortControllers.set(requestId, controller)

  const sendDelta = (contentDelta, reasoningDelta, fullContent, fullReasoning) => {
    try {
      if (event?.sender && !event.sender.isDestroyed()) {
        event.sender.send('ai-stream-delta', { 
          requestId, 
          delta: contentDelta,           // 内容增量
          reasoningDelta: reasoningDelta, // 思考增量（kimi-k2.5 等思考模型）
          fullContent, 
          fullReasoning 
        })
      }
    } catch (e) {
      // ignore send errors
    }
  }

  try {
    const headers = {
      'Content-Type': 'application/json',
    }
    if (apiKey) headers['Authorization'] = `Bearer ${apiKey}`

    const url = `${baseUrl.replace(/\/$/, '')}/chat/completions`
    console.log(`[AI API] 请求: ${url}, model=${model}, temperature=${temperature}, messages=${messages?.length}条`)
    const resp = await fetch(url, {
      method: 'POST',
      headers,
      signal: controller.signal,
      body: JSON.stringify({
        model,
        messages,
        temperature,
        max_tokens: maxTokens,
        stream: true,
      }),
    })

    if (!resp.ok) {
      const errorText = await resp.text().catch(() => '')
      console.error(`[AI API] 请求失败: HTTP ${resp.status}, ${errorText?.slice(0, 200)}`)
      return { success: false, error: errorText || `请求失败（HTTP ${resp.status}）` }
    }
    console.log(`[AI API] 请求成功，开始流式读取...`)

    const reader = resp.body?.getReader?.()
    if (!reader) {
      const text = await resp.text().catch(() => '')
      return { success: false, error: text || '无法读取响应流' }
    }

    const decoder = new TextDecoder()
    let fullContent = ''
    let fullReasoning = ''  // 思考内容（kimi-k2.5 等思考模型）
    let buffer = ''

    const extractStreamText = (value) => {
      if (!value) return ''
      if (typeof value === 'string') return value
      if (Array.isArray(value)) {
        return value.map(part => {
          if (!part) return ''
          if (typeof part === 'string') return part
          if (typeof part.text === 'string') return part.text
          if (typeof part.content === 'string') return part.content
          return ''
        }).join('')
      }
      if (typeof value === 'object') {
        if (typeof value.text === 'string') return value.text
        if (typeof value.content === 'string') return value.content
      }
      return ''
    }

    while (true) {
      const { done, value } = await reader.read()
      if (done) break

      buffer += decoder.decode(value, { stream: true })
      const lines = buffer.split('\n')
      buffer = lines.pop() || ''

      for (const line of lines) {
        if (!line.startsWith('data: ')) continue
        const data = line.slice(6).trim()
        if (!data || data === '[DONE]') continue
        try {
          const json = JSON.parse(data)
          const choice = json.choices?.[0] || {}
          const delta = choice?.delta || {}
          const message = choice?.message || {}
          let contentDelta = delta?.content || delta?.text || ''
          // 尝试多种可能的思考内容字段名
          let reasoningDelta = delta?.reasoning_content || delta?.thinking_content || delta?.reasoning || delta?.thinking || delta?.thought || message?.reasoning_content || json?.reasoning_content || ''

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

          // 调试日志：记录收到的数据结构（首次 delta，便于确认 kimi-k2.5 等模型的字段名）
          if (!fullContent && !fullReasoning && delta) {
            const debugPayload = { deltaKeys: Object.keys(delta || {}), hasReasoning: !!reasoningDelta, sample: JSON.stringify(delta).slice(0, 300) }
            console.log('[AI stream] delta structure:', debugPayload)
            try { fs.appendFileSync(debugLogPath, JSON.stringify({location:'main.cjs:ai-stream-delta',message:'delta structure',data:debugPayload,timestamp:Date.now()}) + '\n') } catch {}
          }
          
          // 只要有任何增量就发送
          if (contentDelta || reasoningDelta) {
            if (contentDelta) fullContent += contentDelta
            if (reasoningDelta) fullReasoning += reasoningDelta
            
            // 从 fullContent 中提取 <think> 标签内的思考内容
            const thinkMatch = fullContent.match(/<think>([\s\S]*?)<\/think>/g)
            if (thinkMatch) {
              const extractedThinking = thinkMatch.map(m => m.replace(/<\/?think>/g, '')).join('\n')
              if (extractedThinking && extractedThinking !== fullReasoning) {
                fullReasoning = extractedThinking
              }
            }
            
            sendDelta(contentDelta, reasoningDelta, fullContent, fullReasoning)
          }
        } catch {
          // ignore parse errors
        }
      }
    }

    console.log(`[AI API] 流式完成: content=${fullContent.length}字, reasoning=${fullReasoning.length}字`)
    return { success: true, content: fullContent }
  } catch (error) {
    console.error(`[AI API] 错误: ${error?.name} - ${error?.message}`)
    console.error(error?.stack || error)
    const msg = error?.name === 'AbortError' ? '请求已取消' : (error?.message || String(error))
    return { success: false, error: msg }
  } finally {
    aiAbortControllers.delete(requestId)
  }
})

ipcMain.handle('ai-cancel', async (_event, requestId) => {
  try {
    const controller = aiAbortControllers.get(requestId)
    if (controller) controller.abort()
    aiAbortControllers.delete(requestId)
    return { success: true }
  } catch (error) {
    return { success: false, error: error?.message || String(error) }
  }
})

// 辅助函数：转义正则表达式特殊字符
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
  const minimalDocxBase64 = 'UEsDBBQAAAAIAAAAAACHTuJAXgAAAGIAAAALAAAAX3JlbHMvLnJlbHONzrEKwjAQBuC9T3Hc3TQOIiKmLuIqOkq8hpg2hSaB+xb79YYO4uLq8P3/z5G6+Bqt+KCPg2MN2VKBQGqcHajT8FqsF3sQMRkyxjnCDX4Y5MJkbLiGfkirqacc4yJJYu2RJi7yDNRPm3+OJjIHamxGa1pTI7YiV8/aHvb/DEiD5e9z0Vu3Bq9Yc5Q0HXWQ7xf4AQAAAP//AwBQSwMEFAAAAAgAAAAAAOaFjPVNAQAA7AIAABAAAABQT0NQcm9wcy9hcHAueG1snVLLTsMwELwj8Q+R71HSCqQKNT0gISQOCFEQZ8vZpLH8kNdJ6d+zTlMehRM+rWdnZ7zjXV6+O1ttICZjfcnms5xV4JWtjW/K9vnp5mTBqkTSa2mtB8l2kNjl6miZLCJEMJ4qyvBJ8pZIfCF4Ui04mYYJfKLWNjpJdIwNjzW8wRaCW+T5OYcOJGpYn0/B/4J2u3v3YPSQPWB8gNgnxgIhKCuJlPbJ+v6+pMN/Yb2L6sF6aRp6n6Qy2Nh6uMEIXYSk/dCnPzCQKFDaRAP2lUTXJPp5Pt9j/QL4bnb7SsYHqPYxpD6Zk+fDuP0fOTH7u4qGSJhp+9/o5v0HAAD//wMAUEsDBBQAAAAIAAAAAABRBQlhsAAAACkBAAARAAAQT0NQcm9wcy9jb3JlLnhtbE2QQU7DMBBF90jcwfIeOQlCCKG4G6QuWLBBHMCyJ4nVeGx5XNreHidQwWr+zNf8P1rc7N0kPiFR8L6BuqpBgDfeBj808Lq7v7wGQVl6KyfvoYEjENy0Z4vOYGwCvqQRBJd4akBnHW8lkdHgJFUhgufLPiQnM8c0SCfNRg4gL6rqSjrI0kosZZ7A+BMl7z+2+BM2r8bGUbqU9j/kPLuWKLGPzpJmgM8J/r7Kz7f/0H4DAAD//wMAUEsDBBQAAAAIAAAAAAC4/U5pVwEAAJkCAAAQAAAAZG9jUHJvcHMvYXBwLnhtbJ2SwU7DMBBE70j8g+U7TZpSIUSTHkBCnBAqgvNibRKr9jryuqH9e5ykVKAnTt7ZnfWMvVrdH1y97iFmG3xJF/OCEvAmmOC3Jd1s7q9uKU1ofK5d8FDSIyS6qs4WKWKAiB4TxRqfEo+0TJn4QrJkwYmMQiJBwmBKmjLGO8mS3YEP0i7y/JpDC9rA8mIK/hfUHvbuI9jDJHuA+ADxLzEOCMFYQ6S1T87P9yXt/gvrQ9IPwUvbst+TVCbrcJ4hQh9h2v3Bh/0fGCgMKG2igfpKouuSfT2b7zH+AfjT7PaVjA9Q7VPIYzIn18dx+z9yYvZ3FY2RMNPuv9HN+w8AAP//AwBQSwMEFAAAAAgAAAAAAI0oZfPdAAAASgIAABEAAAB3b3JkL2RvY3VtZW50LnhtbJ2Sy07DMBBF9/0Ko+xJnBahCjXdICEkFggVwdqyJ4nVeCzZLi1/j5OWRxfs5o7u9Yzn6PZw8G7xBomsh4ouqpIuwGgfbOgq+rK5v7qhixQ1BO08BKjoERK9bS8WXRJxn8CkBRf4lHikTcrEV1IkCx6SWQQIvLQPMWnkZ4xQJ/u6HYNTcJL6KKJJ2Pcf+y+sD1HvgzO2YX+mqEz2Yb6miH2EeftHH/Z/YKAwobSJBppXEl2X7Mf5fI/xC+BP2e0rGR6g2qeQ+2ROno/j9n/kxOzvKhoj4Ua7/0Y37z8AAAD//wMAUEsDBBQAAAAIAAAAAADWsxKqvgAAAC8BAAASAAAAd29yZC9mb250VGFibGUueG1sbY9BDoIwEEX3niLpHlrcGGMKG+PGnTuPMNABGmgnnSr19lIhGl3Nn5n/8jOqOo+9+ICQC8HCttyCQHKhJuosfN4f1nsQnDXVug+EFq7IUFVLFROFETyLXFg4ao4HKdmNOGqehBEpKE0YNXMbOilHd4E/DmCN3O12G6wQNJTbpeAfkJT0/kfYjKLZBOfs0PxMYTjZxvkOEYYIy/pHj/t/YCBTQukiDVRfkug2Zz8v5/cY/wF+V9y9kvYBqiGFMiZ98nwc7/9LTsx+V9EYCTfa/Te6efkFAAD//wMAUEsDBBQAAAAIAAAAAABzPjMmuwAAAC0BAAARAAAAd29yZC9zZXR0aW5ncy54bWxtkE0OwiAQhe+egtC9ULsxxpQuNO7cuXMPMEApEJgJjFZvL/in0c3Me/ne5DGq+Rq9eANJl8DCutqCQPKxJRos3F+vNnsQnDW1uo+EFu7IUFdLlRKFCTyLXFg4ao5HKdmNOGuexhEpKG0YNXMbBilnd4U/jmCN3G63G6wQNJTbpeAfkJT0/kfYjKLdBufs2P5OYTzZxoUBEcYIy+5Hj4d/YCBTQukiDdRfkug2Z78u5/cY/wF+V9y9kvYBqiGFMiZ98nwc7/9LTsx+V9EYCTfa/Te6efsJAAD//wMAUEsDBBQAAAAIAAAAAACKIflUvAAAACwBAAASAAAAd29yZC9zdHlsZXMueG1sbZBBDoIwEEX3nsLpHlrcGGMKG+PGnTuPMNABGmgnnSr19lIhGl3Nn5n/8jOqOo+9+ICQC8HCttyCQHKhJuosfN4f1nsQnDXVug+EFq7IUFVLFROFETyLXFg4ao4HKdmNOGqehBEpKE0YNXMbOilHd4E/DmCN3O12G6wQNJTbpeAfkJT0/kfYjKLZBOfs0PxMYTjZxvkOEYYIy/pHj/t/YCBTQukiDVRfkug2Zz8v5/cY/wF+V9y9kvYBqiGFMiZ98nwc7/9LTsx+V9EYCTfa/Te6efkFAAD//wMAUEsDBBQAAAAIAAAAAACNKGXz3QAAAEoCAAARAAAAd29yZC9kb2N1bWVudC54bWydkstOwzAQRff9CqPsSZwWoQo13SAhJBYIFcHasieJ1Xgs2S4tf4+TlkcX7OaO7vWM5+j2cPBusQaJrIeKLqqSLsBoH2zoKvqyub+6oYsUNQTtPASo6BESvW0vFl0ScZ/ApAUX+JR4pE3KxFdSJAseklkECLy0DzFp5GeMUCf7uh2DU3CS+iiiSdj3H/svrA9R74MztmF/pqhM9mG+poh9hHn7Rx/2f2CgMKG0iQaaVxJdl+zH+XyP8QvgT9ntKxkeoNqnkPtkTp6P4/Z/5MTs7yoaI+FGu/9GN+8/AAAA//8DAFBLAQItABQAAAAIAAAAAACHTuJAXgAAAGIAAAALAAAAAAAAAAAAAACAAAAAAAAAAF9yZWxzLy5yZWxzUEsBAi0AFAAAAAgAAAAAAOaFjPVNAQAA7AIAABAAAAAAAAAAAAAAIIAAAACHAAAAZG9jUHJvcHMvYXBwLnhtbFBLAQItABQAAAAIAAAAAABRBQlhsAAAACkBAAARAAAAAAAAAAAAAACAgQACAABkb2NQcm9wcy9jb3JlLnhtbFBLAQItABQAAAAIAAAAAAC4/U5pVwEAAJkCAAAQAAAAAAAAAAAAAACAgd8CAABkb2NQcm9wcy9hcHAueG1sUEsBAi0AFAAAAAgAAAAAAI0oZfPdAAAASgIAABEAAAAAAAAAAAAAAICBZAQAAHdvcmQvZG9jdW1lbnQueG1sUEsBAi0AFAAAAAgAAAAAANazEqq+AAAALwEAABIAAAAAAAAAAAAAAICBcAUAAHdvcmQvZm9udFRhYmxlLnhtbFBLAQItABQAAAAIAAAAAABzPjMmuwAAAC0BAAARAAAAAAAAAAAAAACAQWwGAAB3b3JkL3NldHRpbmdzLnhtbFBLAQItABQAAAAIAAAAAACKIflUvAAAACwBAAASAAAAAAAAAAAAAACBgVYHAAB3b3JkL3N0eWxlcy54bWxQSwECLQAUAAAACAAAAAABjShl890AAABKAgAAEQAAAAAAAAAAAAAAgYFACAAAd29yZC9kb2N1bWVudC54bWxQSwUGAAAAAAkACQA0AgAAzAkAAAAA'
  
  return Buffer.from(minimalDocxBase64, 'base64')
}

// ==================== PPT (image-only) Generation ====================

function getDashScopeEndpoint(region = 'cn') {
  return region === 'intl'
    ? 'https://dashscope-intl.aliyuncs.com/api/v1/services/aigc/multimodal-generation/generation'
    : 'https://dashscope.aliyuncs.com/api/v1/services/aigc/multimodal-generation/generation'
}

// 负面词基线：用于"去水印/去UI/去乱码/去廉价/去AI味/防字体畸变"
const NEGATIVE_PROMPT_BASELINE =
  // 防止字体/文字畸变（最重要）
  'deformed text, broken text, malformed letters, illegible text, unreadable text, distorted characters, corrupted text, warped text, melted text, stretched text, squished text, overlapping text, cropped text, cut off text, truncated text, incomplete text, missing letters, extra letters, wrong stroke order, bad stroke, messy strokes, ' +
  // 防止中文乱码/错字
  'garbled Chinese, wrong Chinese characters, simplified-traditional mix, mojibake, wrong characters, misspelling, random letters, gibberish, extra text, unwanted text, english text mixed, ' +
  // 防止排版问题
  'ugly typography, amateur typography, bad kerning, bad tracking, uneven spacing, inconsistent font size, font size mismatch, bad line height, crowded text, text too small, ' +
  // 去水印/去UI/去品牌
  'watermark, logo, brand name, badge, QR code, UI elements, screenshot, buttons, interface, HUD, sci-fi interface, holographic UI, futuristic dashboard, ' +
  // 去廉价科技风
  'neon cyberpunk, neon cyan, bright cyan, fluorescent cyan, neon teal, cheap turquoise, neon glow, laser lines, circuit board, generic isometric city, isometric cityscape, circuit-board city, cheap sci-fi, ' +
  // 去AI味/低质量
  'lowres, low resolution, blurry, jpeg artifacts, compression artifacts, noisy, grainy, pixelated, worst quality, low quality, normal quality, bad quality, amateur, unprofessional, amateur layout, noisy background, oversaturated, cheap plastic, toy-like, glossy, harsh specular, overbloom, stock 3d icons, generic template, ai artifacts, uncanny, artificial looking, cgi looking, ' +
  // 去结构问题
  'bad composition, cluttered, messy layout, unbalanced, asymmetric in bad way, empty space, too much whitespace, boring layout, generic layout'

function mergeNegativePrompt(userNegativePrompt) {
  const set = new Set()
  const add = (s) => {
    String(s || '')
      .split(',')
      .map((t) => t.trim())
      .filter(Boolean)
      .forEach((t) => set.add(t))
  }
  add(userNegativePrompt)
  add(NEGATIVE_PROMPT_BASELINE)
  return Array.from(set).join(', ')
}

function requestJson(urlStr, { method = 'GET', headers = {}, body } = {}) {
  return new Promise((resolve, reject) => {
    const urlObj = new URL(urlStr)
    const isHttps = urlObj.protocol === 'https:'
    const lib = isHttps ? https : http

    const req = lib.request(
      {
        protocol: urlObj.protocol,
        hostname: urlObj.hostname,
        port: urlObj.port || (isHttps ? 443 : 80),
        path: urlObj.pathname + urlObj.search,
        method,
        headers,
      },
      (res) => {
        const chunks = []
        res.on('data', (c) => chunks.push(c))
        res.on('end', () => {
          const text = Buffer.concat(chunks).toString('utf-8')
          resolve({ statusCode: res.statusCode || 0, headers: res.headers, text })
        })
      }
    )
    req.on('error', reject)
    if (body) req.write(body)
    req.end()
  })
}

async function downloadToBuffer(urlStr, redirectLeft = 5) {
  const urlObj = new URL(urlStr)
  const lib = urlObj.protocol === 'https:' ? https : http

  return new Promise((resolve, reject) => {
    const req = lib.request(
      {
        protocol: urlObj.protocol,
        hostname: urlObj.hostname,
        port: urlObj.port || (urlObj.protocol === 'https:' ? 443 : 80),
        path: urlObj.pathname + urlObj.search,
        method: 'GET',
        headers: {
          'User-Agent': 'word-cursor/1.0',
        },
      },
      (res) => {
        const status = res.statusCode || 0
        const location = res.headers.location
        if ([301, 302, 303, 307, 308].includes(status) && location && redirectLeft > 0) {
          res.resume()
          const nextUrl = new URL(location, urlStr).toString()
          downloadToBuffer(nextUrl, redirectLeft - 1).then(resolve).catch(reject)
          return
        }
        if (status < 200 || status >= 300) {
          const chunks = []
          res.on('data', (c) => chunks.push(c))
          res.on('end', () => reject(new Error(`下载失败: HTTP ${status} ${Buffer.concat(chunks).toString('utf-8').slice(0, 200)}`)))
          return
        }
        const chunks = []
        res.on('data', (c) => chunks.push(c))
        res.on('end', () => resolve(Buffer.concat(chunks)))
      }
    )
    req.on('error', reject)
    req.end()
  })
}

function extractDashScopeImageUrl(json) {
  // sync multimodal-generation format
  const maybe1 = json?.output?.choices?.[0]?.message?.content?.find?.((c) => c?.image)?.image
  if (maybe1) return maybe1
  // async / ImageSynthesis format
  const maybe2 = json?.output?.results?.[0]?.url
  if (maybe2) return maybe2
  // task query format
  const maybe3 = json?.output?.results?.[0]?.url || json?.output?.results?.[0]?.image
  if (maybe3) return maybe3
  return null
}

function extractDashScopeTaskId(json) {
  return (
    json?.output?.task_id ||
    json?.output?.taskId ||
    json?.output?.taskID ||
    json?.task_id ||
    json?.taskId ||
    null
  )
}

function getDashScopeTaskEndpoint(region = 'cn', taskId) {
  const origin =
    region === 'intl' ? 'https://dashscope-intl.aliyuncs.com' : 'https://dashscope.aliyuncs.com'
  return `${origin}/api/v1/tasks/${encodeURIComponent(String(taskId))}`
}

async function dashscopeWaitForImageUrlByTaskId({ taskId, region, apiKey, timeoutMs = 120000 }) {
  const started = Date.now()
  let delay = 800
  let lastText = ''
  while (Date.now() - started < timeoutMs) {
    const endpoint = getDashScopeTaskEndpoint(region, taskId)
    const { statusCode, text } = await requestJson(endpoint, {
      method: 'GET',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${apiKey}`,
      },
    })
    lastText = text
    if (statusCode >= 200 && statusCode < 300) {
      try {
        const json = JSON.parse(text)
        const url = extractDashScopeImageUrl(json)
        if (url) return { url, raw: json }
        const status =
          json?.output?.task_status || json?.output?.taskStatus || json?.output?.status || json?.status
        if (String(status).toUpperCase().includes('FAILED')) {
          const msg = json?.message || json?.output?.message || 'DashScope 任务失败'
          throw new Error(`DashScope 任务失败: ${msg}`)
        }
      } catch {
        // ignore JSON parse errors; will retry
      }
    }
    await new Promise((r) => setTimeout(r, delay))
    delay = Math.min(Math.floor(delay * 1.35), 5000)
  }
  throw new Error(`DashScope 异步任务超时，taskId=${taskId}，last=${String(lastText).slice(0, 200)}`)
}

async function dashscopeGenerateImageUrl({
  prompt,
  negativePrompt = '',
  size = '2048*1152',
  promptExtend = false,
  watermark = false,
  model = 'z-image-turbo',
  region = 'cn',
  apiKey: apiKeyOverride,
}) {
  const apiKey =
    apiKeyOverride ||
    process.env.DASHSCOPE_API_KEY ||
    process.env.BAILIAN_API_KEY ||
    process.env.DASHSCOPE_KEY ||
    process.env.API_KEY
  if (!apiKey) {
    throw new Error('缺少 DashScope API Key：请在“AI 设置”里填写 apiKey（与 LLM 相同也可），或在 .env 中配置 DASHSCOPE_API_KEY')
  }
  if (!prompt || !String(prompt).trim()) {
    throw new Error('缺少 prompt')
  }

  const endpoint = getDashScopeEndpoint(region)
  const payload = {
    model,
    input: {
      messages: [
        {
          role: 'user',
          content: [{ text: String(prompt) }],
        },
      ],
    },
    parameters: {
      negative_prompt: String(negativePrompt || ''),
      prompt_extend: !!promptExtend,
      watermark: !!watermark,
      size,
    },
  }

  const { statusCode, text } = await requestJson(endpoint, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${apiKey}`,
    },
    body: JSON.stringify(payload),
  })

  let json
  try {
    json = JSON.parse(text)
  } catch {
    throw new Error(`DashScope 返回非 JSON: HTTP ${statusCode} ${text.slice(0, 200)}`)
  }

  if (statusCode < 200 || statusCode >= 300) {
    const msg = json?.message || json?.error?.message || text
    throw new Error(`DashScope 调用失败: HTTP ${statusCode} ${msg}`)
  }

  const url = extractDashScopeImageUrl(json)
  if (url) return { url, raw: json }

  // 兼容异步任务：如果返回 task_id，则轮询直到拿到图片 URL
  const taskId = extractDashScopeTaskId(json)
  if (taskId) {
    return await dashscopeWaitForImageUrlByTaskId({ taskId, region, apiKey })
  }

  throw new Error(`DashScope 返回中未找到 image url/task_id: ${text.slice(0, 500)}`)
}

/**
 * DashScope 图像编辑 API（qwen-image-edit-plus）
 * 用于局部编辑 PPT 页面（换背景、改文字等）
 */
async function dashscopeImageEdit({
  imageBase64,         // 当前页图片 base64（不含 data:... 前缀）
  prompt,              // 编辑指令
  negativePrompt = '',
  n = 1,
  watermark = false,
  model = 'qwen-image-edit-plus',
  region = 'cn',
  apiKey: apiKeyOverride,
}) {
  const apiKey =
    apiKeyOverride ||
    process.env.DASHSCOPE_API_KEY ||
    process.env.BAILIAN_API_KEY ||
    process.env.DASHSCOPE_KEY ||
    process.env.API_KEY
  if (!apiKey) {
    throw new Error('缺少 DashScope API Key')
  }
  if (!prompt || !String(prompt).trim()) {
    throw new Error('缺少编辑 prompt')
  }
  if (!imageBase64 || !String(imageBase64).trim()) {
    throw new Error('缺少待编辑的图片 base64')
  }

  const endpoint = getDashScopeEndpoint(region)
  
  // qwen-image-edit-plus 使用 MultiModalConversation 格式
  // 图片可以是 URL 或 data URI
  const imageDataUri = imageBase64.startsWith('data:')
    ? imageBase64
    : `data:image/png;base64,${imageBase64}`

  const payload = {
    model,
    input: {
      messages: [
        {
          role: 'user',
          content: [
            { image: imageDataUri },
            { text: String(prompt) },
          ],
        },
      ],
    },
    parameters: {
      negative_prompt: String(negativePrompt || ''),
      n: Math.max(1, Math.min(4, n)),
      watermark: !!watermark,
    },
  }

  const { statusCode, text } = await requestJson(endpoint, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${apiKey}`,
    },
    body: JSON.stringify(payload),
  })

  let json
  try {
    json = JSON.parse(text)
  } catch {
    throw new Error(`DashScope ImageEdit 返回非 JSON: HTTP ${statusCode} ${text.slice(0, 200)}`)
  }

  if (statusCode < 200 || statusCode >= 300) {
    const msg = json?.message || json?.error?.message || text
    throw new Error(`DashScope ImageEdit 调用失败: HTTP ${statusCode} ${msg}`)
  }

  const url = extractDashScopeImageUrl(json)
  if (url) return { url, raw: json }

  // 兼容异步任务
  const taskId = extractDashScopeTaskId(json)
  if (taskId) {
    return await dashscopeWaitForImageUrlByTaskId({ taskId, region, apiKey })
  }

  throw new Error(`DashScope ImageEdit 返回中未找到 image url/task_id: ${text.slice(0, 500)}`)
}

/**
 * 保存 PPT 生成的元数据到 _assets 目录
 * 用于后续编辑时恢复上下文
 */
function saveDeckMetadata(assetsDir, metadata) {
  try {
    if (!fs.existsSync(assetsDir)) {
      fs.mkdirSync(assetsDir, { recursive: true })
    }

    if (metadata.deckContext) {
      fs.writeFileSync(
        path.join(assetsDir, 'deck_context.json'),
        JSON.stringify(metadata.deckContext, null, 2)
      )
    }

    if (metadata.slidesPrompts) {
      fs.writeFileSync(
        path.join(assetsDir, 'slides_prompts.json'),
        JSON.stringify(metadata.slidesPrompts, null, 2)
      )
    }

    if (metadata.outline) {
      fs.writeFileSync(
        path.join(assetsDir, 'outline.json'),
        JSON.stringify(metadata.outline, null, 2)
      )
    }

    console.log('[PPTX] 元数据已保存到:', assetsDir)
  } catch (e) {
    console.warn('[PPTX] 保存元数据失败:', e?.message || e)
  }
}

/**
 * 从 _assets 目录加载 PPT 元数据
 */
function loadDeckMetadata(assetsDir) {
  const result = {
    deckContext: null,
    slidesPrompts: null,
    outline: null,
  }

  try {
    const contextPath = path.join(assetsDir, 'deck_context.json')
    if (fs.existsSync(contextPath)) {
      result.deckContext = JSON.parse(fs.readFileSync(contextPath, 'utf-8'))
    }
  } catch {}

  try {
    const promptsPath = path.join(assetsDir, 'slides_prompts.json')
    if (fs.existsSync(promptsPath)) {
      result.slidesPrompts = JSON.parse(fs.readFileSync(promptsPath, 'utf-8'))
    }
  } catch {}

  try {
    const outlinePath = path.join(assetsDir, 'outline.json')
    if (fs.existsSync(outlinePath)) {
      result.outline = JSON.parse(fs.readFileSync(outlinePath, 'utf-8'))
    }
  } catch {}

  return result
}

/**
 * 从 PPTX 或 _assets 目录读取指定页的图片
 * @returns {Promise<Buffer|null>}
 */
async function getSlideImageFromPptx(pptxPath, pageIndex, assetsDir) {
  // 优先从 _assets 读取最新的 processed PNG
  if (assetsDir) {
    const seq = String(pageIndex + 1).padStart(2, '0')
    // 查找最新 attempt 的 1920x1080 PNG（兼容旧的 1920x1200）
    const files = fs.existsSync(assetsDir) ? fs.readdirSync(assetsDir) : []
    const pngFiles = files
      .filter((f) => (f.startsWith(`slide_${seq}_1920x1080_`) || f.startsWith(`slide_${seq}_1920x1200_`)) && f.endsWith('.png'))
      .sort()
      .reverse()
    if (pngFiles.length > 0) {
      const pngPath = path.join(assetsDir, pngFiles[0])
      return fs.readFileSync(pngPath)
    }
  }

  // 从 PPTX 解压读取
  try {
    const JSZip = require('jszip')
    const pptxBuffer = fs.readFileSync(pptxPath)
    const zip = await JSZip.loadAsync(pptxBuffer)

    // 找到对应页的图片
    const slideNum = pageIndex + 1
    const relPath = `ppt/slides/_rels/slide${slideNum}.xml.rels`
    const relFile = zip.file(relPath)
    if (!relFile) return null

    const relXml = await relFile.async('string')
    // 找第一个 image 关系
    const match = relXml.match(/Relationship[^>]*Type="[^"]*image[^"]*"[^>]*Target="([^"]+)"/)
    if (!match) return null

    let imagePath = match[1]
    // 解析相对路径
    if (imagePath.startsWith('..')) {
      imagePath = 'ppt/' + imagePath.replace(/^\.\.\//g, '')
    } else if (!imagePath.startsWith('ppt/')) {
      imagePath = 'ppt/slides/' + imagePath
    }

    const imgFile = zip.file(imagePath)
    if (!imgFile) return null

    return Buffer.from(await imgFile.async('arraybuffer'))
  } catch (e) {
    console.warn('[PPTX] 从 PPTX 读取图片失败:', e?.message || e)
    return null
  }
}

/**
 * 替换 PPTX 中指定页的图片并覆盖写回
 * @param {string} pptxPath - PPTX 文件路径
 * @param {Array<{pageIndex: number, imageBuffer: Buffer}>} replacements - 替换列表
 * @param {boolean} backup - 是否备份原文件
 */
async function replaceSlideImagesInPptx(pptxPath, replacements, backup = true) {
  const JSZip = require('jszip')
  const pptxBuffer = fs.readFileSync(pptxPath)
  const zip = await JSZip.loadAsync(pptxBuffer)

  // 备份原文件
  if (backup) {
    const dir = path.dirname(pptxPath)
    const baseName = path.basename(pptxPath, '.pptx')
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19)
    const backupPath = path.join(dir, `${baseName}_backup_${timestamp}.pptx`)
    fs.copyFileSync(pptxPath, backupPath)
    console.log('[PPTX] 已备份到:', backupPath)
  }

  for (const { pageIndex, imageBuffer } of replacements) {
    const slideNum = pageIndex + 1
    const relPath = `ppt/slides/_rels/slide${slideNum}.xml.rels`
    const relFile = zip.file(relPath)
    if (!relFile) {
      console.warn(`[PPTX] 未找到 slide${slideNum} 的 rels`)
      continue
    }

    const relXml = await relFile.async('string')
    const match = relXml.match(/Relationship[^>]*Type="[^"]*image[^"]*"[^>]*Target="([^"]+)"/)
    if (!match) {
      console.warn(`[PPTX] slide${slideNum} 未找到图片关系`)
      continue
    }

    let imagePath = match[1]
    if (imagePath.startsWith('..')) {
      imagePath = 'ppt/' + imagePath.replace(/^\.\.\//g, '')
    } else if (!imagePath.startsWith('ppt/')) {
      imagePath = 'ppt/slides/' + imagePath
    }

    // 替换图片
    zip.file(imagePath, imageBuffer)
    console.log(`[PPTX] 已替换 slide${slideNum} 图片: ${imagePath}`)
  }

  // 写回 PPTX
  const newBuffer = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
  fs.writeFileSync(pptxPath, newBuffer)
  console.log('[PPTX] 已覆盖写回:', pptxPath)

  return { success: true, path: pptxPath }
}

async function postprocessTo1920x1200(buffer, mode = 'letterbox') {
  // 更新为 16:9 分辨率以匹配 z-image-turbo 输出 (2048*1152)
  const targetW = 1920
  const targetH = 1080 // 16:9 比例
  if (mode === 'cover') {
    return await sharp(buffer).resize(targetW, targetH, { fit: 'cover', position: 'attention' }).png().toBuffer()
  }
  // default: letterbox (no crop)
  return await sharp(buffer)
    .resize(targetW, targetH, { fit: 'contain', background: { r: 0, g: 0, b: 0, alpha: 1 } })
    .png()
    .toBuffer()
}

function makePptx16x10FromImagesBase64(imageBase64List, outputPath) {
  const pptx = new PptxGenJS()
  // 更新为 16:9 布局以匹配 z-image-turbo 输出
  const w = 13.333 // 10 英寸 * 1.333
  const h = 7.5    // 16:9 比例 (10 英寸宽 * 9/16 = 5.625 英寸，但 PPT 标准是 13.333 x 7.5)
  pptx.defineLayout({ name: 'LAYOUT_16X9', width: w, height: h })
  pptx.layout = 'LAYOUT_16X9'
  pptx.author = 'Word-Cursor'

  for (const img of imageBase64List) {
    const slide = pptx.addSlide()
    // 显式设置背景，避免部分前端预览器解析时出现 background undefined
    slide.background = { color: '000000' }
    slide.addImage({ data: img, x: 0, y: 0, w, h })
  }
  return pptx.writeFile({ fileName: outputPath, compression: true })
}

// OpenRouter Gemini: 调用（支持 messages 或 system+user）
async function callOpenRouterGemini({ apiKey, model, systemPrompt, userPrompt, messages }) {
  const baseUrl = 'https://openrouter.ai/api/v1/chat/completions'
  // 使用 Gemini 3 Pro Preview（最新最强）
  const selectedModel = model || 'google/gemini-3-pro-preview'
  const finalMessages = Array.isArray(messages) && messages.length > 0
    ? messages
    : [
        { role: 'system', content: systemPrompt },
        { role: 'user', content: userPrompt }
      ]
  const body = {
    model: selectedModel,
    messages: finalMessages,
    temperature: 0.7,
    max_tokens: 16000, // Gemini 3 Pro 支持更长输出，PPT 提示词需要足够空间
  }
  console.log('[OpenRouter] Calling Gemini 3 Pro:', selectedModel)
  const res = await fetch(baseUrl, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${apiKey}`,
      'HTTP-Referer': 'https://word-cursor.app',
      'X-Title': 'Word-Cursor PPT Generator',
    },
    body: JSON.stringify(body),
  })
  if (!res.ok) {
    const text = await res.text()
    throw new Error(`OpenRouter API error: ${res.status} - ${text}`)
  }
  const data = await res.json()
  console.log('[OpenRouter] Gemini 3 Pro response, finish_reason:', data.choices?.[0]?.finish_reason, 'tokens:', data.usage?.total_tokens)
  return data.choices?.[0]?.message?.content || ''
}

// 通用 OpenAI 兼容 API 调用（用于主模型回退）
async function callOpenAICompatible({ apiKey, baseUrl, model, systemPrompt, userPrompt, messages }) {
  // 清理 baseUrl，确保正确格式
  let endpoint = String(baseUrl || 'https://api.openai.com/v1').trim()
  if (endpoint.endsWith('/')) {
    endpoint = endpoint.slice(0, -1)
  }
  if (!endpoint.endsWith('/chat/completions')) {
    endpoint = `${endpoint}/chat/completions`
  }
  
  const finalMessages = Array.isArray(messages) && messages.length > 0
    ? messages
    : [
        { role: 'system', content: systemPrompt },
        { role: 'user', content: userPrompt }
      ]
  
  const body = {
    model: model || 'gpt-4',
    messages: finalMessages,
    temperature: 0.7,
    max_tokens: 16000,
  }
  
  console.log('[OpenAI Compatible] Calling:', model, 'at', endpoint)
  const res = await fetch(endpoint, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${apiKey}`,
    },
    body: JSON.stringify(body),
  })
  
  if (!res.ok) {
    const text = await res.text()
    throw new Error(`API error: ${res.status} - ${text}`)
  }
  
  const data = await res.json()
  console.log('[OpenAI Compatible] Response, finish_reason:', data.choices?.[0]?.finish_reason, 'tokens:', data.usage?.total_tokens)
  return data.choices?.[0]?.message?.content || ''
}

// LinAPI Gemini: 调用 chat/completions 接口（用于 PPT 提示词生成）
async function callLinAPIGemini({ apiKey, model, systemPrompt, userPrompt, messages }) {
  const baseUrl = 'https://api.linapi.net/v1/chat/completions'
  // 默认使用 gemini-3-pro-preview 生成 PPT 提示词
  const selectedModel = model || 'gemini-3-pro-preview'
  const finalMessages = Array.isArray(messages) && messages.length > 0
    ? messages
    : [
        { role: 'system', content: systemPrompt },
        { role: 'user', content: userPrompt }
      ]
  const body = {
    model: selectedModel,
    messages: finalMessages,
    temperature: 0.7,
    max_tokens: 16000,
  }
  
  // 添加重试逻辑（网络不稳定时最多重试 3 次）
  const maxRetries = 3
  let lastError = null
  
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      console.log(`[LinAPI] Calling Gemini: ${selectedModel} (attempt ${attempt}/${maxRetries})`)
      const res = await fetch(baseUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${apiKey}`,
        },
        body: JSON.stringify(body),
      })
      if (!res.ok) {
        const text = await res.text()
        throw new Error(`LinAPI error: ${res.status} - ${text}`)
      }
      const data = await res.json()
      console.log('[LinAPI] Gemini response, finish_reason:', data.choices?.[0]?.finish_reason, 'tokens:', data.usage?.total_tokens)
      return data.choices?.[0]?.message?.content || ''
    } catch (err) {
      lastError = err
      const isNetworkError = err?.cause?.code === 'ECONNRESET' || 
                             err?.cause?.code === 'UND_ERR_SOCKET' ||
                             err?.message?.includes('fetch failed')
      if (isNetworkError && attempt < maxRetries) {
        console.warn(`[LinAPI] 网络错误，${attempt}s 后重试... (${err?.cause?.code || err.message})`)
        await new Promise(r => setTimeout(r, attempt * 1000)) // 递增等待时间
        continue
      }
      throw err
    }
  }
  throw lastError
}

// LinAPI Gemini 生图: 调用 gemini-3-pro-image-preview-2K 生成图片（带重试）
async function linapiGenerateImage({ apiKey, prompt, aspectRatio = '16:9' }) {
  const endpoint = 'https://api.linapi.net/v1beta/models/gemini-3-pro-image-preview-2K:generateContent'
  
  const body = {
    contents: [{
      parts: [{ text: prompt }]
    }],
    generationConfig: {
      imageConfig: {
        aspectRatio: aspectRatio,
        imageSize: '1K'
      }
    }
  }
  
  // 添加重试逻辑（网络不稳定时最多重试 3 次）
  const maxRetries = 3
  let lastError = null
  
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      console.log(`\n${'-'.repeat(40)}`)
      console.log(`[LinAPI Image] Generating image (attempt ${attempt}/${maxRetries})`)
      console.log(`[LinAPI Image] FULL PROMPT:\n${prompt}`)
      console.log(`${'-'.repeat(40)}\n`)
      
      const res = await fetch(endpoint, {
        method: 'POST',
        headers: {
          'x-goog-api-key': apiKey,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(body),
      })
      
      if (!res.ok) {
        const text = await res.text()
        throw new Error(`LinAPI Image error: ${res.status} - ${text}`)
      }
      
      const data = await res.json()
      
      // 提取生成的图片 Base64 数据
      const candidate = data.candidates?.[0]
      const parts = candidate?.content?.parts || []
      
      for (const part of parts) {
        if (part.inlineData?.data) {
          const mimeType = part.inlineData.mimeType || 'image/png'
          const base64Data = part.inlineData.data
          console.log('[LinAPI Image] Got image, mimeType:', mimeType, 'size:', base64Data.length)
          return {
            url: `data:${mimeType};base64,${base64Data}`,
            base64: base64Data,
            mimeType
          }
        }
      }
      
      throw new Error('LinAPI Image: 未在响应中找到图片数据')
    } catch (err) {
      lastError = err
      const isNetworkError = err?.cause?.code === 'ECONNRESET' || 
                             err?.cause?.code === 'UND_ERR_SOCKET' ||
                             err?.message?.includes('fetch failed') ||
                             err?.message?.includes('socket')
      if (isNetworkError && attempt < maxRetries) {
        console.warn(`[LinAPI Image] 网络错误，${attempt * 2}s 后重试... (${err?.cause?.code || err.message})`)
        await new Promise(r => setTimeout(r, attempt * 2000)) // 生图需要更长等待时间
        continue
      }
      throw err
    }
  }
  throw lastError
}

function enhancePromptForGeminiImage({ prompt, negativePrompt }) {
  const safePrompt = String(prompt || '').trim()
  const safeNeg = String(negativePrompt || '').trim()
  // Gemini 生图对"负面词"没有单独字段，这里把 negativePrompt 作为 Avoid 列表融入同一条文本指令里
  // 目标：高端杂志级设计、丰富细节、精致纹理、专业排版
  
  const designDirectives = `## IMAGE GENERATION REQUIREMENTS (STRICT)

### 1. FORMAT & PURPOSE
You are generating a PRESENTATION SLIDE image (16:9 aspect ratio).
This is FLAT GRAPHIC DESIGN for business/editorial use — NOT a 3D render, NOT a game scene, NOT dark cyberpunk art.

### 2. COLOR PALETTE (MANDATORY)
- PRIMARY BACKGROUND: Off-white (#F8F7F4), warm gray (#E8E6E1), soft ivory (#FDFBF7), or pale cream (#F5F3EE)
- Background must have SUBTLE TEXTURE: fine paper grain, linen weave, very light noise, or faint watercolor wash — NEVER flat solid color
- TEXT COLOR: Rich charcoal (#2D2D2D) or warm dark gray (#3A3A3A) — NOT pure black
- ACCENT COLOR: ONE sophisticated accent only (muted blue #4A7C9B, terracotta #C4785A, sage green #7D9B76, or warm gold #B8976A) — used sparingly (≤5% of area)
- FORBIDDEN: Neon colors, saturated blues/purples, glowing effects, gradients that look like cheap stock photos

### 3. LAYOUT & COMPOSITION (MANDATORY)
- Use strict GRID SYSTEM: Bento grid, modular grid, or classic editorial columns
- Strong ALIGNMENT: all elements snap to baseline grid
- Generous WHITE SPACE: minimum 15% margins, breathing room between elements
- Clear VISUAL HIERARCHY: primary → secondary → tertiary information levels
- Professional TYPOGRAPHY: elegant sans-serif for Chinese text, proper kerning, comfortable line-height (1.4-1.6)

### 4. MICRO-DETAILS (CRITICAL — This creates richness)
Add these LOW-OPACITY decorative elements throughout the design:
- Ultra-thin grid lines (0.5px, 5-10% opacity)
- Corner registration marks (like print crop marks)
- Small page numbers or serial codes (No.01, VOL.25)
- Tiny geometric accents: dots, crosses, small squares, subtle lines
- Abstract data visualization elements (thin connecting lines, small nodes)
- Faint geometric patterns in background (hexagons, circles, triangles at 3-5% opacity)
- Subtle shadow layers for depth
- Delicate dividing lines between content sections
- Small iconographic elements relevant to the topic (minimal line-art style)
- Grain/noise texture overlay (very subtle, 2-5% opacity)

### 5. MATERIALITY & TEXTURE
- Frosted glass cards (glassmorphism) for content containers — with realistic blur and soft edges
- Soft drop shadows (not harsh, offset 0-4px, blur 8-16px, 8-15% opacity)
- Paper-like texture on background
- Subtle embossing or debossing effects on key elements
- Matte finish aesthetic — no glossy/plastic look

### 6. TYPOGRAPHY REQUIREMENTS
- ALL Chinese characters must be PERFECTLY LEGIBLE, correctly formed, elegantly spaced
- Use modern Chinese sans-serif aesthetic (like PingFang, Source Han Sans style)
- Proper text hierarchy through size, weight, and spacing — NOT color variation
- Headlines: bold/semibold, generous tracking
- Body text: regular weight, comfortable line-height
- NO random English text, NO gibberish, NO Lorem Ipsum — only the exact content provided

### 7. STRICT AVOIDANCE LIST
NEVER include these elements:
- 3D rendered spheres, cubes, or geometric shapes that look like stock 3D assets
- Dark/black backgrounds
- Neon glows, lens flares, or light leaks
- Cyberpunk/sci-fi hologram aesthetics
- Cheap-looking gradients (especially blue-purple)
- Generic stock photo elements (handshake, lightbulb, puzzle pieces)
- Watermarks, logos, or brand marks
- Cluttered layouts with no breathing room
- Toy-like or plastic textures
- Overly complex 3D scenes or realistic photo compositions
- Random English abbreviations or placeholder text

---

## CONTENT TO VISUALIZE:`
  
  const avoidList = [
    '3D spheres', '3D balls', '3D cubes', '3D geometric primitives',
    'dark background', 'black background', 'neon glow', 'cyberpunk',
    'hologram', 'sci-fi UI', 'circuit board', 'matrix code',
    'cheap gradient', 'stock photo', 'plastic texture', 'toy-like',
    'blurry text', 'deformed text', 'broken text', 'garbled Chinese',
    'wrong characters', 'illegible text', 'ugly typography',
    'watermark', 'logo', 'brand mark', 'lowres', 'amateur design',
    safeNeg
  ].filter(Boolean).join(', ')

  return [
    designDirectives,
    '',
    safePrompt,
    '',
    `## MUST AVOID: ${avoidList}`
  ].join('\n')
}

function isDashscopeInappropriateContentError(err) {
  const msg = String(err?.message || err || '').toLowerCase()
  return msg.includes('inappropriate content') || msg.includes('inappropriate-content')
}

function extractHttpStatusFromErrorMessage(err) {
  const msg = String(err?.message || err || '')
  const m = msg.match(/\bHTTP\s+(\d{3})\b/i)
  return m ? Number(m[1]) : null
}

function parseJsonFromModelText(text) {
  if (!text) return null
  try {
    const jsonMatch = String(text).match(/```json\s*([\s\S]*?)\s*```/i)
    if (jsonMatch?.[1]) return JSON.parse(jsonMatch[1])
    return JSON.parse(String(text))
  } catch {
    return null
  }
}

// IPC: 调用 Gemini 生成文生图提示词（统一使用主模型 API）
ipcMain.handle('openrouter-gemini-ppt-prompts', async (_event, options = {}) => {
  try {
    const { outline, theme, style, mainApiKey } = options
    
    if (!mainApiKey) {
      return { success: false, error: '缺少主模型 API Key，请在设置中配置' }
    }
    if (!outline) {
      return { success: false, error: '缺少 PPT 大纲' }
    }

    const systemPrompt = `你是一位专精于 **高端品牌视觉设计** 的顶级艺术总监。你的任务是编写 **极其详细、视觉元素丰富** 的 AI 绘画提示词，用于生成一张 **直接包含完整 PPT 内容** 的幻灯片成片。

⚠️ **核心目标：视觉丰富度（VISUAL RICHNESS）** ⚠️
- 每张图片必须像 **Dribbble/Behance 上获奖的品牌提案** 那样精致、层次丰富、细节饱满。
- 绝不能是"文字+简单背景"的单调设计，必须有 **多层次视觉元素堆叠**。

## 🎨 视觉丰富度铁律（每页必须全部满足）

### 1) 多层次构图（Layered Composition）- 必须 5+ 层
每页至少包含以下层次（从后到前）：
- **Layer 1 背景层**：渐变/纹理/图案（绝不能是纯色）
- **Layer 2 氛围层**：大面积模糊光斑、柔和渐变云、抽象几何形状
- **Layer 3 装饰层**：网格线、几何图形、抽象元素、图标阵列
- **Layer 4 主视觉层**：与主题相关的核心插图/图形/3D元素
- **Layer 5 内容层**：毛玻璃卡片承载的文字内容

### 2) 主题创意元素（Thematic Visual Elements）- 必须 2-3 个
根据 PPT 主题，必须加入 **与内容直接相关的创意视觉元素**：
- 科技主题：电路线条、数据流粒子、代码片段装饰、芯片纹理、光纤线
- 商业主题：图表元素、上升箭头、齿轮连接、网络节点、增长曲线
- 教育主题：书本元素、灯泡图标、知识树、公式装饰、学术符号
- 创意主题：画笔笔触、色彩飞溅、艺术纹理、创意工具图标
- 自然主题：植物剪影、水波纹理、有机曲线、自然光影
- **必须在 prompt 中明确描述这些元素的位置、大小、颜色和透明度**

### 3) 微元素密度（Micro-Detail Density）- 必须 8+ 种
每页必须包含大量低透明度装饰元素（10-30% opacity），从以下清单中选择至少 8 种：
- □ 极细网格线（ultra-thin grid lines, 0.5px）
- □ 角标/裁切标记（corner marks, registration marks）
- □ 页码序列号（No.01, VOL.25, SLIDE 03）
- □ 微型图标阵列（tiny icons array, 16px）
- □ 抽象条形码/二维码装饰（abstract barcode pattern）
- □ 点阵图案（dot matrix pattern）
- □ 细分隔线/引导线（thin dividers, guide lines）
- □ 浮动几何小块（floating geometric shapes）
- □ 数据可视化元素（mini charts, data points, progress bars）
- □ 渐变光晕/光斑（gradient orbs, soft glows）
- □ 纹理叠加（noise texture, paper grain, fabric weave）
- □ 连接线/流程线（connecting lines, flow paths）
- □ 时间轴元素（timeline markers, date stamps）
- □ 标签/徽章装饰（label badges, status indicators）
- □ 波形/脉冲线（waveforms, pulse lines）

### 4) 色彩层次（Color Depth）
- 主背景：Off-white (#F5F3EE) 到 Warm Gray (#E8E4DF) 的微妙渐变
- 必须有 2-3 个不同透明度的装饰色层
- 一个鲜明但克制的强调色（面积≤8%）
- 阴影必须是暖灰色调，不能是纯黑

### 5) 材质与质感（Materiality）
- 毛玻璃卡片：blur 20-40px, 白色 60-80% 透明度, 1px 白色边框
- 多层柔和阴影：近影 + 远影 创造立体感
- 背景必须有可见纹理：纸纹/布纹/噪点（5-15% opacity）

### 6) 文字规范
- 中文必须清晰可读（crisp Chinese text, elegant sans-serif）
- 只包含大纲提供的文字，禁止随机内容
- 文字有呼吸空间，行距 1.5+

## ✅ Prompt 格式要求

每条 prompt 必须 **600-900 字符**，结构如下：
1. 整体场景描述（overall scene）
2. 背景层详细描述（background layer details）
3. 装饰元素详细描述（decorative elements with positions）
4. 主题视觉元素描述（thematic visual elements）
5. 内容卡片描述（content card with glassmorphism）
6. 完整的中文文字内容（exact Chinese text）
7. 色彩和光影描述（colors, lighting, shadows）
8. 风格关键词（style keywords）

## ✅ 输出格式（严格 JSON）
你必须只输出 JSON（可用 \`\`\`json 代码块包裹），结构：
{
  "designConcept": "整体视觉策略：选用的风格 + 核心视觉元素 + 统一的配色方案",
  "colorPalette": "具体色值：背景色 + 装饰色 + 强调色",
  "slides": [
    {
      "pageNumber": 1,
      "pageType": "cover/content/summary",
      "visualConcept": "本页视觉创意：使用哪些主题元素 + 如何体现内容",
      "prompt": "极其详细的英文提示词（600-900字符）",
      "negativePrompt": "负面词"
    }
  ]
}

## 负面词（必须包含）
negativePrompt 必须包含：deformed text, broken text, malformed letters, illegible text, garbled Chinese, wrong Chinese characters, ugly typography, dark background, pure black, neon glow, cyberpunk, hologram, messy layout, cluttered, watermark, logo, brand mark, lowres, blurry, cheap, plastic, amateur, empty, minimal, simple, plain, boring, flat design without depth

---

下面是你要处理的具体大纲与风格偏好。`

    const userPrompt = `请为以下 PPT 大纲设计视觉方案并生成文生图提示词（每页一张成片）：

## PPT 主题/用途
${theme || '（未指定，请根据大纲内容判断）'}

## 用户期望的风格倾向
${style || '你可在风格库 A-F 中自动选择一个最匹配大纲的高级风格；也可以在同体系内做少量变奏（保持统一审美）'}

## PPT 大纲内容
${outline}

---
要求：
1) 先选择最匹配该大纲的主风格 preset（A-F），在 designConcept 里说明原因；必要时可“同体系变奏”，但不要乱混风格  
2) **反廉价/反AI味（强制）**：避免“塑料感/玩具感/廉价霓虹/模板化等距城市/素材库风 3D 图标”。整体要像品牌 KV / 杂志海报  
3) **配色（强制）**：给出“美术生审美”的配色——低饱和主色 + 中性色 + 1 个点睛色；避免过饱和、刺眼荧光、廉价蓝紫霓虹  
4) **主题创意元素（强制）**：每页除背景+文字外，至少加入 1-2 个与主题直接相关的创意视觉元素/隐喻（例如：城市路网拓扑线、地图纹理、时间轴丝带、印章纹理、建筑剖面线稿、数据流粒子等），而不是随机几何装饰  
5) 每页 prompt 必须包含该页所有中文文案（标题/副标题/要点/页脚），并强调中文清晰可读；**禁止出现大纲之外的任何文字**（尤其随机英文缩写/乱码）  
6) 图片中只能有设计元素和用户内容：**禁止任何品牌/Logo/软件界面/水印/角标**  
7) 每条 prompt ≤900 chars，negativePrompt ≤400 chars，并在 negativePrompt 中**必须**加入：deformed text, broken text, malformed letters, illegible text, garbled Chinese, wrong Chinese characters, ugly typography, cheap plastic, toy-like, lowres, blurry, amateur, neon cyberpunk, circuit board, watermark, logo  
8) **最终输出必须是严格的 JSON 格式**（参考 system prompt 中的格式），不要输出任何解释性文字，只输出 JSON 代码块。`

    // 统一使用主模型 API（LinAPI）调用 gemini-3-pro-preview
    let response = ''
    
    if (!mainApiKey) {
      return { success: false, error: '缺少主模型 API Key，请在设置中配置' }
    }
    
    console.log('[PPT Prompts] 使用主模型 API (gemini-3-pro-preview)')
    response = await callLinAPIGemini({
      apiKey: mainApiKey,
      model: 'gemini-3-pro-preview',
      systemPrompt,
      userPrompt,
    })

    // 解析JSON响应
    let parsed = null
    try {
      // 尝试提取JSON块
      const jsonMatch = response.match(/```json\s*([\s\S]*?)\s*```/)
      if (jsonMatch) {
        parsed = JSON.parse(jsonMatch[1])
      } else {
        // 尝试直接解析
        parsed = JSON.parse(response)
      }
    } catch (parseError) {
      console.error('Gemini response parse error:', parseError)
      return { success: false, error: 'Gemini 返回的内容无法解析为JSON', raw: response }
    }

    const normalizedSlides = Array.isArray(parsed?.slides)
      ? parsed.slides.map((s, idx) => ({
          pageNumber: Number(s?.pageNumber) || idx + 1,
          pageType: String(s?.pageType || 'content'),
          visualConcept: typeof s?.visualConcept === 'string' ? s.visualConcept : '',
          prompt: String(s?.prompt || ''),
          negativePrompt: mergeNegativePrompt(s?.negativePrompt),
        }))
      : parsed?.slides

    return {
      success: true,
      slides: normalizedSlides,
      designConcept: parsed.designConcept || '',
      colorPalette: parsed.colorPalette || '',
      raw: response,
    }
  } catch (error) {
    console.error('openrouter-gemini-ppt-prompts error:', error)
    return { success: false, error: error.message || String(error) }
  }
})

ipcMain.handle('ppt-generate-deck', async (_event, options = {}) => {
  try {
    const {
      outputPath,
      slides = [],
      mainApiKey = '', // 主模型 API Key（用于 Gemini 生图）
      dashscope = {},
      postprocess = { mode: 'letterbox' },
      repair = {},
      outline = null, // 原始大纲（用于保存元数据）
    } = options
    
    // 用于收集每页最终使用的 prompt（含修复后的）
    const finalSlidesPrompts = []

    if (!outputPath || typeof outputPath !== 'string') {
      return { success: false, error: '缺少 outputPath' }
    }
    if (!outputPath.toLowerCase().endsWith('.pptx')) {
      return { success: false, error: 'outputPath 必须以 .pptx 结尾' }
    }
    if (!Array.isArray(slides) || slides.length === 0) {
      return { success: false, error: 'slides 不能为空' }
    }

    const limit = pLimit(2) // DashScope: RPS=2 且并发=2（两张两张生成）
    const geminiRepairLimit = pLimit(1) // Gemini 维修：串行，保证上下文一致
    // 用户选择的图像生成模型
    const imageModel = dashscope.model || 'z-image-turbo'
    // 根据模型选择默认分辨率
    const defaultSize = imageModel === 'z-image-turbo' ? '2048*1152' : '1664*928'
    const size = dashscope.size || defaultSize
    const promptExtend = !!dashscope.promptExtend
    const watermark = dashscope.watermark === true
    const negativePromptDefault = dashscope.negativePromptDefault || ''
    const region = dashscope.region || 'cn'
    const apiKey = dashscope.apiKey || ''
    const saveImages = dashscope.saveImages !== false // 默认保存，便于排查"是否真的生成了图片"
    
    console.log(`[PPT Generate] 使用模型: ${imageModel}, 分辨率: ${size}`)

    const repairEnabled =
      repair?.enabled !== false &&
      !!repair?.openRouterApiKey &&
      typeof repair?.openRouterApiKey === 'string' &&
      repair.openRouterApiKey.trim().length > 10
    const repairMaxAttempts = Math.max(0, Math.min(5, Number(repair?.maxAttempts ?? 2)))
    const repairModel = repair?.model || 'google/gemini-3-pro-preview'
    const deckContext = repair?.deckContext || {}

    // 维修会话上下文：用于“只修失败页”的连续对话（串行执行，避免并发污染上下文）
    const geminiRepairMessages = repairEnabled
      ? [
          {
            role: 'system',
            content:
              'You are a world-class presentation designer and prompt engineer. ' +
              'We are generating poster-style PPT slide images (text is part of the image). ' +
              'Some slides may fail DashScope safety moderation (inappropriate content). ' +
              'Your job: REWRITE ONLY the failed slide prompt to pass moderation while keeping the same deck style.\n' +
              '\n' +
              'Rules:\n' +
              '- Keep the overall style consistent with the deck design concept and color palette.\n' +
              '- Keep Chinese text crisp & legible. Prefer keeping the exact Chinese copy; if any phrase is likely to trigger moderation, paraphrase into neutral, compliant wording while preserving meaning.\n' +
              '- Avoid any violence/politics/sensitive content. Avoid brand names, logos, UI, watermarks.\n' +
              '- Output JSON ONLY: {"prompt":"...","negativePrompt":"...","textEdits":[{"from":"...","to":"..."}]}\n' +
              '- prompt <= 800 chars, negativePrompt <= 300 chars.',
          },
          {
            role: 'user',
            content:
              'Deck context (keep consistent):\n' +
              `- designConcept: ${String(deckContext.designConcept || '').slice(0, 800)}\n` +
              `- colorPalette: ${String(deckContext.colorPalette || '').slice(0, 200)}\n` +
              'Remember: do NOT regenerate the whole deck. We will request single-slide repairs as needed.',
          },
        ]
      : []

    async function repairSlidePromptWithGemini({ idx, attempt, prompt, negativePrompt, errorMessage }) {
      return await geminiRepairLimit(async () => {
        const slideNo = idx + 1
        const userMsg =
          `Slide repair request:\n` +
          `- slideNumber: ${slideNo}\n` +
          `- dashscopeError: ${String(errorMessage).slice(0, 800)}\n` +
          `- previousPrompt: ${String(prompt).slice(0, 4000)}\n` +
          `- previousNegativePrompt: ${String(negativePrompt || '').slice(0, 1200)}\n` +
          '\n' +
          'Rewrite a safer prompt that preserves layout and typography, keeps deck style consistent, and avoids moderation triggers. Output JSON only.'

        geminiRepairMessages.push({ role: 'user', content: userMsg })
        
        let responseText = ''
        try {
          responseText = await callOpenRouterGemini({
            apiKey: repair.openRouterApiKey,
            model: repairModel,
            messages: geminiRepairMessages,
          })
        } catch (geminiErr) {
          console.warn(`[PPT Repair] Gemini 调用异常 (slide=${slideNo}, attempt=${attempt}):`, geminiErr?.message || geminiErr)
          throw new Error(`Gemini 修复调用失败（slide=${slideNo}）: ${geminiErr?.message || '网络错误'}`)
        }
        
        // 检查空响应（Gemini 有时会返回空内容）
        if (!responseText || responseText.trim().length < 10) {
          console.warn(`[PPT Repair] Gemini 返回空响应 (slide=${slideNo}, attempt=${attempt})`)
          throw new Error(`Gemini 修复返回空响应（slide=${slideNo}），请重试`)
        }
        
        geminiRepairMessages.push({ role: 'assistant', content: responseText })

        const parsed = parseJsonFromModelText(responseText)
        const newPrompt = parsed?.prompt
        const newNegative = parsed?.negativePrompt
        if (!newPrompt || typeof newPrompt !== 'string' || newPrompt.trim().length < 20) {
          console.warn(`[PPT Repair] Gemini 返回无效 JSON (slide=${slideNo}):`, responseText?.slice(0, 500))
          throw new Error(`Gemini 修复提示词失败：无法解析 prompt（slide=${slideNo}, attempt=${attempt}）`)
        }
        return {
          prompt: String(newPrompt).trim(),
          negativePrompt: typeof newNegative === 'string' ? String(newNegative).trim() : String(negativePrompt || '').trim(),
          textEdits: Array.isArray(parsed?.textEdits) ? parsed.textEdits : [],
          raw: responseText,
        }
      })
    }

    // 把每页下载到的原始图片 & 后处理后的 1920x1080 PNG 保存到本地，便于排查
    const outDir = path.dirname(outputPath)
    const baseName = path.basename(outputPath, path.extname(outputPath))
    const assetsDir = path.join(outDir, `${baseName}_assets`)
    if (saveImages && !fs.existsSync(assetsDir)) {
      fs.mkdirSync(assetsDir, { recursive: true })
    }
    const results = await Promise.all(
      slides.map((s, idx) =>
        limit(async () => {
          let prompt = s.prompt || s.finalPrompt || s.finalPromptCNorEN || ''
          let negativePrompt = s.negativePrompt ?? negativePromptDefault

          const seq = String(idx + 1).padStart(2, '0')
          const promptPathBase = saveImages ? path.join(assetsDir, `slide_${seq}_prompt`) : null

          // 单页重试：遇到审核失败（inappropriate content）→ 把该页失败信息交给 Gemini 改写提示词 → 仅重试该页
          let attempt = 0
          while (true) {
            try {
              // 保存当前尝试的 prompt（便于对比）
              if (saveImages && promptPathBase) {
                try {
                  fs.writeFileSync(`${promptPathBase}_attempt${attempt}.txt`, String(prompt))
                  fs.writeFileSync(`${promptPathBase}_neg_attempt${attempt}.txt`, String(negativePrompt || ''))
                } catch {}
              }

              let raw
              let imageSource = '' // 记录图片来源（URL 或 'gemini-base64'）
              if (imageModel === 'gemini-image') {
                // 使用 LinAPI Gemini 生图（需要主模型 API Key）
                // 直接使用 Gemini 3 Pro 生成的原始提示词，不做额外增强
                if (!mainApiKey) {
                  throw new Error('使用 Gemini 生图需要配置主模型 API Key（LinAPI）')
                }
                console.log(`\n${'='.repeat(60)}`)
                console.log(`[PPT Generate] ✅ 使用 gemini-3-pro-image-preview-2K 生图`)
                console.log(`[PPT Generate] Slide ${idx + 1}/${slides.length}`)
                console.log(`[PPT Generate] 提示词长度: ${prompt.length} chars`)
                console.log(`${'='.repeat(60)}\n`)
                
                const geminiResult = await linapiGenerateImage({
                  apiKey: mainApiKey,
                  prompt: prompt, // 直接使用原始提示词
                  aspectRatio: '16:9',
                })
                raw = Buffer.from(geminiResult.base64, 'base64')
                imageSource = 'gemini-base64'
                console.log(`[PPT Generate] Slide ${idx + 1} 生图完成，图片大小: ${raw.length} bytes`)
              } else {
                // 使用 DashScope 生图
                const { url } = await dashscopeGenerateImageUrl({
                  prompt,
                  negativePrompt,
                  size,
                  promptExtend,
                  watermark,
                  model: imageModel,
                  region,
                  apiKey,
                })
                raw = await downloadToBuffer(url)
                imageSource = url // 记录 URL 来源
              }
              if (!raw || raw.length === 0) {
                throw new Error(`图片下载失败或为空（idx=${idx}）`)
              }
              const processed = await postprocessTo1920x1200(raw, postprocess?.mode || 'letterbox')
              if (!processed || processed.length === 0) {
                throw new Error(`图片后处理失败或为空（idx=${idx}）`)
              }

              // 保存图片到本地（用于排查是否真实生成/下载/后处理成功）
              if (saveImages) {
                // 根据图片来源决定文件扩展名
                let ext = '.jpg' // 默认为 jpg
                if (imageSource && imageSource !== 'gemini-base64') {
                  try {
                    const u = new URL(imageSource)
                    ext = path.extname(u.pathname).toLowerCase() || '.jpg'
                  } catch {}
                }
                if (!ext || ext.length > 5) ext = '.jpg'
                const rawPath = path.join(assetsDir, `slide_${seq}_raw_attempt${attempt}${ext}`)
                const pngPath = path.join(assetsDir, `slide_${seq}_1920x1080_attempt${attempt}.png`)
                const sourcePath = path.join(assetsDir, `slide_${seq}_source_attempt${attempt}.txt`)
                try {
                  fs.writeFileSync(rawPath, raw)
                  fs.writeFileSync(pngPath, processed)
                  fs.writeFileSync(sourcePath, imageSource === 'gemini-base64' ? 'gemini-3-pro-image-preview-2K (base64)' : String(imageSource))
                } catch (e) {
                  console.warn('[PPTX] 保存图片失败:', e?.message || e)
                }
              }

              const base64 = processed.toString('base64')
              const dataUri = `image/png;base64,${base64}`

              return { idx, dataUri, finalPrompt: prompt, finalNegativePrompt: negativePrompt, attempts: attempt + 1 }
            } catch (e) {
              const errorMessage = e?.message || String(e)
              const status = extractHttpStatusFromErrorMessage(e)
              const isInappropriate = status === 400 && isDashscopeInappropriateContentError(e)

              if (saveImages) {
                try {
                  fs.writeFileSync(path.join(assetsDir, `slide_${seq}_error_attempt${attempt}.txt`), String(errorMessage))
                } catch {}
              }

              if (!repairEnabled || !isInappropriate || attempt >= repairMaxAttempts) {
                throw e
              }

              // 触发 Gemini 单页修复
              const repairRes = await repairSlidePromptWithGemini({
                idx,
                attempt,
                prompt,
                negativePrompt,
                errorMessage,
              })

              if (saveImages) {
                try {
                  fs.writeFileSync(path.join(assetsDir, `slide_${seq}_repair_response_attempt${attempt}.txt`), String(repairRes.raw || ''))
                  if (Array.isArray(repairRes.textEdits) && repairRes.textEdits.length) {
                    fs.writeFileSync(
                      path.join(assetsDir, `slide_${seq}_repair_text_edits_attempt${attempt}.json`),
                      JSON.stringify(repairRes.textEdits, null, 2)
                    )
                  }
                } catch {}
              }

              prompt = repairRes.prompt
              negativePrompt = repairRes.negativePrompt || negativePrompt
              attempt += 1
              continue
            }
          }
        })
      )
    )

    results.sort((a, b) => a.idx - b.idx)
    const images = results.map((r) => r.dataUri)
    
    // 收集每页最终的 prompt 信息
    const slidesPromptsData = results.map((r, i) => ({
      pageNumber: i + 1,
      prompt: r.finalPrompt || slides[r.idx]?.prompt || '',
      negativePrompt: r.finalNegativePrompt || slides[r.idx]?.negativePrompt || '',
      attempts: r.attempts || 1,
      originalChineseContent: slides[r.idx]?.originalChineseContent || '',
    }))

    // Ensure directory exists
    const dir = path.dirname(outputPath)
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true })
    }

    await makePptx16x10FromImagesBase64(images, outputPath)
    
    // 保存元数据到 _assets 目录（用于后续编辑）
    if (saveImages) {
      saveDeckMetadata(assetsDir, {
        deckContext: deckContext,
        slidesPrompts: slidesPromptsData,
        outline: outline,
      })
    }

    return { success: true, path: outputPath, slideCount: slides.length }
  } catch (error) {
    console.error('ppt-generate-deck failed:', error)
    return { success: false, error: error.message || String(error) }
  }
})

// ==================== PPT 编辑（整页重做 / 局部编辑）====================

ipcMain.handle('ppt-edit-slides', async (_event, options = {}) => {
  try {
    const {
      pptxPath,           // PPTX 文件路径
      pageNumbers = [],   // 要编辑的页码数组（1-based）
      feedback = '',      // 用户反馈
      mode = 'regenerate', // 'regenerate' = 整页重做，'partial_edit' = 局部编辑
      openRouterApiKey,   // Gemini API Key
      dashscopeApiKey,    // DashScope API Key
      mainApiKey,         // 主模型 API Key（用于 LinAPI Gemini 生图）
      pptImageModel = 'z-image-turbo', // 生图模型选择
      deckContext: providedDeckContext, // 可选，优先使用提供的
      regionScreenshot,   // 新增：用户框选区域的截图 base64
      regionRect,         // 新增：框选区域坐标 {x, y, w, h}
    } = options

    if (!pptxPath || typeof pptxPath !== 'string') {
      return { success: false, error: '缺少 pptxPath' }
    }
    if (!fs.existsSync(pptxPath)) {
      return { success: false, error: `PPTX 文件不存在: ${pptxPath}` }
    }
    if (!Array.isArray(pageNumbers) || pageNumbers.length === 0) {
      return { success: false, error: '缺少要编辑的页码' }
    }
    if (!feedback || typeof feedback !== 'string' || !feedback.trim()) {
      return { success: false, error: '缺少用户反馈' }
    }
    if (!openRouterApiKey) {
      return { success: false, error: '缺少 OpenRouter API Key' }
    }
    if (pptImageModel === 'gemini-image') {
      // Gemini 生图走 LinAPI，需要主模型 key（或复用 dashscopeApiKey 兜底，但推荐传 mainApiKey）
      if (!mainApiKey && !dashscopeApiKey) {
        return { success: false, error: '缺少 API Key：Gemini 生图需要 mainApiKey（或至少提供 dashscopeApiKey 兜底）' }
      }
    } else {
      // DashScope 生图/编辑仍需要 DashScope Key
      if (!dashscopeApiKey) {
        return { success: false, error: '缺少 DashScope API Key' }
      }
    }

    // 读取 _assets 元数据
    const baseName = path.basename(pptxPath, '.pptx')
    const assetsDir = path.join(path.dirname(pptxPath), `${baseName}_assets`)
    const metadata = loadDeckMetadata(assetsDir)
    
    const deckContext = providedDeckContext || metadata.deckContext || {}
    const slidesPrompts = metadata.slidesPrompts || []
    const outline = metadata.outline || {}
    
    // 获取大纲中的 slides 数组
    const outlineSlides = outline.slides || outline.pages || outline.content || []

    console.log(`[PPT Edit] 模式: ${mode}, 页码: ${pageNumbers.join(', ')}, 反馈: ${feedback.slice(0, 100)}...`)
    console.log(`[PPT Edit] 大纲页数: ${outlineSlides.length}, slidesPrompts: ${slidesPrompts.length}`)

    const editLimit = pLimit(1) // 串行编辑，保证 Gemini 上下文一致
    const replacements = []
    const editLogs = []

    for (const pageNum of pageNumbers) {
      await editLimit(async () => {
        const pageIndex = pageNum - 1
        if (pageIndex < 0) {
          editLogs.push({ pageNum, success: false, error: '页码无效' })
          return
        }

        // 获取该页的原始图片
        const originalImage = await getSlideImageFromPptx(pptxPath, pageIndex, assetsDir)
        if (!originalImage) {
          editLogs.push({ pageNum, success: false, error: '无法读取该页图片' })
          return
        }

        const originalImageBase64 = originalImage.toString('base64')
        const slidePromptInfo = slidesPrompts[pageIndex] || {}
        
        // 获取该页的大纲内容（非常重要：确保生成的图片包含正确的文字）
        const outlineSlide = outlineSlides[pageIndex] || {}
        const slideHeadline = outlineSlide.headline || outlineSlide.title || outlineSlide.heading || outlineSlide.pageTitle || outlineSlide.page_title || ''
        const slideSubheadline = outlineSlide.subheadline || outlineSlide.subtitle || outlineSlide.sub_title || ''
        const slideBullets = outlineSlide.bullets || outlineSlide.points || outlineSlide.content_points || outlineSlide.keyPoints || []
        const slideFooter = outlineSlide.footerNote || outlineSlide.footer || ''
        const slidePageType = outlineSlide.pageType || outlineSlide.page_type || 'content'
        const slideLayoutIntent = outlineSlide.layoutIntent || outlineSlide.layout_intent || ''
        
        // 构建该页的完整中文内容（用于 Gemini）
        let slideChineseContent = ''
        if (slideHeadline) slideChineseContent += `标题: "${slideHeadline}"\n`
        if (slideSubheadline) slideChineseContent += `副标题: "${slideSubheadline}"\n`
        if (Array.isArray(slideBullets) && slideBullets.length > 0) {
          slideChineseContent += `要点:\n${slideBullets.map((b, i) => `  ${i + 1}. "${b}"`).join('\n')}\n`
        }
        if (slideFooter) slideChineseContent += `页脚: "${slideFooter}"\n`
        
        // 如果大纲内容为空，尝试使用 slidesPrompts 中保存的内容
        if (!slideChineseContent.trim() && slidePromptInfo.originalChineseContent) {
          slideChineseContent = slidePromptInfo.originalChineseContent
        }
        
        console.log(`[PPT Edit] 第 ${pageNum} 页大纲内容:\n${slideChineseContent.slice(0, 500)}`)

        let newImageBuffer = null

        if (mode === 'regenerate') {
          // ========== 整页重做 ==========
          // 1. 让 Gemini 根据反馈重写 prompt（必须包含该页的中文内容 + 高级设计感）
          const geminiSystemPrompt = 
            'You are a world-class presentation designer creating PREMIUM, AWARD-WINNING slide visuals. ' +
            'The user is not satisfied with their current slide and wants it redesigned with MUCH BETTER aesthetics. ' +
            '\n\n' +
            '## YOUR DESIGN PHILOSOPHY\n' +
            '- Think like top-tier keynote design teams - clean, sophisticated, highly curated\n' +
            '- Use LAYERED DEPTH: paper/cards/scroll-tabs/panels with soft shadows and clear hierarchy (not sci-fi HUD)\n' +
            '- Apply PREMIUM MATERIALS: matte ceramic, fine paper (xuan paper), silk, lacquer, jade/bronze accents, restrained gold foil lines\n' +
            '- Create VISUAL HIERARCHY: clear focal point, breathing space, balanced composition\n' +
            '- Add REFINED DETAILS: filmic tone mapping, subtle grain, micro-texture, restrained highlights\n' +
            '\n' +
            '## CRITICAL RULES\n' +
            '1. ALL Chinese text from slide content MUST appear in the image (title, bullets, footer)\n' +
            '2. Chinese text: crisp, high-contrast, elegant typography (not plain black on white!)\n' +
            '3. Layout: asymmetric balance, golden ratio, generous margins\n' +
            '4. For AGENDA/TOC pages: use creative layouts like numbered cards, timeline, floating panels\n' +
            '5. Quote each Chinese text explicitly in your prompt\n' +
            '6. Add 1-2 THEME-RELATED creative motifs (not random decoration)\n' +
            '7. If user feedback asks for "古风/国风/东方/典雅/宋韵/新中式": MUST switch to a premium neo-Chinese heritage aesthetic (xuan paper, ink wash, seal stamp, antique map lines) and AVOID any tech/HUD/neon look\n' +
            '\n' +
            '## AESTHETIC TECHNIQUES TO USE\n' +
            '- Neo-Chinese heritage (when requested): xuan paper texture, ink wash gradients, subtle cloud patterns, antique gold foil dividers, cinnabar seal accents\n' +
            '- Depth layering: foreground text on refined panels over textured background (paper/ink wash)\n' +
            '- Premium color harmony: low-chroma main hue + neutrals + one accent (art-school palette, filmic grading)\n' +
            '- Elegant motifs: thin contour lines, map topology, architectural line art, seals, minimal ornaments tied to the topic\n' +
            '- Lighting: soft global illumination, gentle vignetting, subtle shadowing (avoid neon glow rings)\n' +
            '\n' +
            '## HIGH-END DESIGN VOCAB (USE SELECTIVELY, NOT KEYWORD STUFFING)\n' +
            '- Layout/Grid: International Typographic Style (Swiss), typographic grid, baseline grid, modular grid, strong alignment, consistent gutters, generous margins\n' +
            '- Microtypography: microtypography, optical alignment, kerning, tracking, leading, typographic scale, clean line breaks\n' +
            '- Premium finishes/material cues: soft-touch matte lamination, paper grain, spot UV varnish, hot foil stamping, emboss/deboss, debossed foil linework, letterpress impression, duotone, spot color, subtle halftone\n' +
            '- Cinematic lighting: three-point lighting, key light, fill light ratio, rim light, kicker light, bounce light, softbox diffusion, gentle falloff, volumetric light rays, ambient occlusion\n' +
            '- Filmic color: filmic tone mapping, split toning (warm highlights + cool shadows), matte blacks, highlight roll-off, subtle halation, fine film grain, restrained bloom\n' +
            '- Use 6-12 of these terms per prompt at most; keep the prompt actionable.\n' +
            '\n' +
            '## AVOID CHEAP/AI LOOK (MANDATORY)\n' +
            '- Avoid: cheap plastic, toy-like, glossy, harsh specular, over-bloom, over-saturated neon\n' +
            '- Avoid: HUD, sci-fi interface, holographic UI, futuristic dashboards, glowing rings/dials\n' +
            '- Avoid: generic isometric city / stock 3D icon templates / cliché circuit-board city\n' +
            '- Prefer: matte, textured, editorial poster vibe, restrained highlights, elegant palettes\n' +
            '\n' +
            'Output JSON ONLY: {"prompt":"...","negativePrompt":"..."}'

          // 构建框选区域描述（如果有）
          const regionHint = regionRect 
            ? `\n\n## User Selected Region\nThe user specifically highlighted a region at: x=${regionRect.x}, y=${regionRect.y}, width=${regionRect.w}, height=${regionRect.h}.\nPlease pay special attention to improving this area in your redesign.`
            : ''
          
          const geminiUserPrompt = 
            `## Deck Style Context\n` +
            `- Design Concept: ${String(deckContext.designConcept || 'Premium neo-Chinese heritage editorial with refined textures').slice(0, 800)}\n` +
            `- Color Palette: ${String(deckContext.colorPalette || 'Ink black, warm parchment, cinnabar accent, antique gold').slice(0, 200)}\n\n` +
            `## Page ${pageNum} Information\n` +
            `- Page Type: ${slidePageType}\n` +
            `- Layout Intent: ${slideLayoutIntent || 'balanced asymmetric layout with visual hierarchy'}\n\n` +
            `## SLIDE CONTENT (Chinese text that MUST appear):\n` +
            `${slideChineseContent || '(No content provided)'}\n\n` +
            `## User Feedback (the problem to solve):\n${feedback}${regionHint}\n\n` +
            `## Original Prompt (what went wrong - AVOID these issues):\n${String(slidePromptInfo.prompt || '').slice(0, 1000)}\n\n` +
            '## YOUR TASK\n' +
            'Create a COMPLETELY NEW, VISUALLY STUNNING prompt that:\n' +
            '1. Addresses the user feedback (more design, more polish, more visual interest)\n' +
            '2. Uses premium materials + theme-related motifs (neo-Chinese heritage if requested)\n' +
            '3. Includes ALL the Chinese text content with elegant typography\n' +
            '4. Creates a slide that looks like it belongs in a Fortune 500 keynote' +
            (regionRect ? '\n5. Especially focus on improving the user-highlighted region' : '')

          // 带重试的 Gemini 调用（网络不稳定时自动重试）
          let geminiResponse = null
          let geminiRetries = 0
          const maxGeminiRetries = 3
          while (geminiRetries < maxGeminiRetries) {
            try {
              geminiResponse = await callOpenRouterGemini({
                apiKey: openRouterApiKey,
                model: 'google/gemini-3-pro-preview',
                systemPrompt: geminiSystemPrompt,
                userPrompt: geminiUserPrompt,
              })
              break // 成功则跳出
            } catch (geminiErr) {
              geminiRetries++
              const errMsg = geminiErr?.message || String(geminiErr)
              console.warn(`[PPT Edit] Gemini 调用失败 (尝试 ${geminiRetries}/${maxGeminiRetries}): ${errMsg.slice(0, 200)}`)
              if (geminiRetries >= maxGeminiRetries) {
                throw new Error(`Gemini 调用失败（已重试 ${maxGeminiRetries} 次）: ${errMsg}`)
              }
              // 等待后重试
              await new Promise(r => setTimeout(r, 1000 * geminiRetries))
            }
          }

          const parsed = parseJsonFromModelText(geminiResponse)
          if (!parsed?.prompt) {
            editLogs.push({ pageNum, success: false, error: 'Gemini 返回的 prompt 无效' })
            return
          }

          // 2. 生成新图（根据模型选择不同接口）
          let raw
          const maxRetries = 2
          let retries = 0
          
          while (retries < maxRetries) {
            try {
              if (pptImageModel === 'gemini-image') {
                // 使用 LinAPI Gemini 生图
                const enhancedPrompt = enhancePromptForGeminiImage({
                  prompt: parsed.prompt,
                  negativePrompt: mergeNegativePrompt(parsed.negativePrompt),
                })
                const geminiResult = await linapiGenerateImage({
                  apiKey: mainApiKey || dashscopeApiKey,
                  prompt: enhancedPrompt,
                  aspectRatio: '16:9',
                })
                raw = Buffer.from(geminiResult.base64, 'base64')
              } else {
                // 使用 DashScope 生图
                const result = await dashscopeGenerateImageUrl({
                  prompt: parsed.prompt,
                  negativePrompt: mergeNegativePrompt(parsed.negativePrompt),
                  size: '2048*1152',
                  promptExtend: false,
                  watermark: false,
                  model: pptImageModel || 'z-image-turbo',
                  region: 'cn',
                  apiKey: dashscopeApiKey,
                })
                raw = await downloadToBuffer(result.url)
              }
              break
            } catch (imgErr) {
              retries++
              const errMsg = imgErr?.message || String(imgErr)
              console.warn(`[PPT Edit] 生图失败 (尝试 ${retries}/${maxRetries}): ${errMsg.slice(0, 200)}`)
              if (retries >= maxRetries) {
                throw new Error(`生图失败（已重试 ${maxRetries} 次）: ${errMsg}`)
              }
              await new Promise(r => setTimeout(r, 1500 * retries))
            }
          }
          newImageBuffer = await postprocessTo1920x1200(raw, 'letterbox')

          // 保存编辑记录
          if (fs.existsSync(assetsDir)) {
            const seq = String(pageNum).padStart(2, '0')
            const timestamp = Date.now()
            try {
              fs.writeFileSync(path.join(assetsDir, `slide_${seq}_edit_${timestamp}_prompt.txt`), parsed.prompt)
              fs.writeFileSync(path.join(assetsDir, `slide_${seq}_edit_${timestamp}_after.png`), newImageBuffer)
            } catch {}
          }

          // 更新 slidesPrompts
          if (slidesPrompts[pageIndex]) {
            slidesPrompts[pageIndex].prompt = parsed.prompt
            slidesPrompts[pageIndex].negativePrompt = parsed.negativePrompt || ''
          }

        } else if (mode === 'partial_edit') {
          // ========== 局部编辑 ==========
          // 1. 让 Gemini 生成编辑指令（给 qwen-image-edit-plus）
          const geminiSystemPrompt = 
            'You are an expert at image editing prompts for AI image editors. ' +
            'The user wants to make SPECIFIC partial edits to a PPT slide image. ' +
            '\n\n' +
            '## EDITING GUIDELINES\n' +
            '- Be PRECISE about what to change and what to keep\n' +
            '- If changing background: describe the NEW background style (gradient, abstract, etc.)\n' +
            '- If changing colors: specify exact color transitions\n' +
            '- If changing text style: describe the NEW typography style\n' +
            '- PRESERVE all Chinese text unless user explicitly wants to change it\n' +
            '\n' +
            '## QUALITY REQUIREMENTS\n' +
            '- Maintain premium design aesthetic\n' +
            '- Chinese text must remain crisp and readable\n' +
            '- Changes should enhance, not diminish the design\n' +
            '\n' +
            '## HIGH-END DESIGN VOCAB (USE SELECTIVELY)\n' +
            '- Microtypography: typographic grid, baseline grid, microtypography, optical alignment, kerning, tracking, leading\n' +
            '- Premium finishes/material cues: soft-touch matte, paper grain, spot UV varnish, hot foil stamping, emboss/deboss, letterpress impression, duotone/spot color\n' +
            '- Cinematic lighting: key light, fill ratio, rim/kicker light, softbox diffusion, bounce light, gentle falloff, volumetric rays, subtle vignetting\n' +
            '- Filmic color: filmic tone mapping, split toning, matte blacks, highlight roll-off, fine grain, restrained bloom\n' +
            '- Do NOT keyword-stuff. Use only what helps the requested edit.\n' +
            '\n' +
            'Output JSON ONLY: {"editPrompt":"...","negativePrompt":"..."}'

          const geminiUserPrompt = 
            `## Current Slide Design\n` +
            `- Design Concept: ${String(deckContext.designConcept || '').slice(0, 600)}\n` +
            `- Color Palette: ${String(deckContext.colorPalette || '').slice(0, 200)}\n\n` +
            `## Chinese Text Content (PRESERVE unless asked to change):\n` +
            `${slideChineseContent || String(slidePromptInfo.originalChineseContent || '').slice(0, 800)}\n\n` +
            `## User Edit Request:\n${feedback}\n\n` +
            'Create an edit prompt that makes ONLY the requested changes while keeping everything else intact.'

          // 带重试的 Gemini 调用
          let geminiResponse = null
          let geminiRetries = 0
          const maxGeminiRetries = 3
          while (geminiRetries < maxGeminiRetries) {
            try {
              geminiResponse = await callOpenRouterGemini({
                apiKey: openRouterApiKey,
                model: 'google/gemini-3-pro-preview',
                systemPrompt: geminiSystemPrompt,
                userPrompt: geminiUserPrompt,
              })
              break
            } catch (geminiErr) {
              geminiRetries++
              console.warn(`[PPT Edit] Gemini 调用失败 (尝试 ${geminiRetries}/${maxGeminiRetries})`)
              if (geminiRetries >= maxGeminiRetries) {
                throw new Error(`Gemini 调用失败（已重试 ${maxGeminiRetries} 次）`)
              }
              await new Promise(r => setTimeout(r, 1000 * geminiRetries))
            }
          }

          const parsed = parseJsonFromModelText(geminiResponse)
          if (!parsed?.editPrompt) {
            editLogs.push({ pageNum, success: false, error: 'Gemini 返回的 editPrompt 无效' })
            return
          }

          // 2. 带重试的 DashScope 图像编辑
          let dashscopeUrl = null
          let dashscopeRetries = 0
          const maxDashscopeRetries = 2
          while (dashscopeRetries < maxDashscopeRetries) {
            try {
              const result = await dashscopeImageEdit({
                imageBase64: originalImageBase64,
                prompt: parsed.editPrompt,
                negativePrompt: mergeNegativePrompt(parsed.negativePrompt),
                n: 1,
                watermark: false,
                model: 'qwen-image-edit-plus',
                region: 'cn',
                apiKey: dashscopeApiKey,
              })
              dashscopeUrl = result.url
              break
            } catch (dsErr) {
              dashscopeRetries++
              console.warn(`[PPT Edit] DashScope 编辑失败 (尝试 ${dashscopeRetries}/${maxDashscopeRetries})`)
              if (dashscopeRetries >= maxDashscopeRetries) {
                throw new Error(`DashScope 图像编辑失败（已重试 ${maxDashscopeRetries} 次）`)
              }
              await new Promise(r => setTimeout(r, 1500 * dashscopeRetries))
            }
          }
          
          const { url } = { url: dashscopeUrl }

          const raw = await downloadToBuffer(url)
          newImageBuffer = await postprocessTo1920x1200(raw, 'letterbox')

          // 保存编辑记录
          if (fs.existsSync(assetsDir)) {
            const seq = String(pageNum).padStart(2, '0')
            const timestamp = Date.now()
            try {
              fs.writeFileSync(path.join(assetsDir, `slide_${seq}_partialedit_${timestamp}_prompt.txt`), parsed.editPrompt)
              fs.writeFileSync(path.join(assetsDir, `slide_${seq}_partialedit_${timestamp}_before.png`), originalImage)
              fs.writeFileSync(path.join(assetsDir, `slide_${seq}_partialedit_${timestamp}_after.png`), newImageBuffer)
            } catch {}
          }
        }

        if (newImageBuffer && newImageBuffer.length > 0) {
          replacements.push({ pageIndex, imageBuffer: newImageBuffer })
          editLogs.push({ pageNum, success: true })
        } else {
          editLogs.push({ pageNum, success: false, error: '生成的图片为空' })
        }
      })
    }

    if (replacements.length === 0) {
      return { success: false, error: '没有成功编辑任何页面', logs: editLogs }
    }

    // 替换 PPTX 中的图片并覆盖写回
    await replaceSlideImagesInPptx(pptxPath, replacements, true)

    // 更新 slides_prompts.json
    if (fs.existsSync(assetsDir) && slidesPrompts.length > 0) {
      try {
        fs.writeFileSync(
          path.join(assetsDir, 'slides_prompts.json'),
          JSON.stringify(slidesPrompts, null, 2)
        )
      } catch {}
    }

    return {
      success: true,
      path: pptxPath,
      editedPages: replacements.map((r) => r.pageIndex + 1),
      logs: editLogs,
    }
  } catch (error) {
    console.error('ppt-edit-slides failed:', error)
    return { success: false, error: error.message || String(error) }
  }
})

// ===================== Fonts IPC =====================
// 列出 Fonts/ 目录下可用字体文件
ipcMain.handle('fonts-list', async () => {
  try {
    // Fonts 目录相对于项目根目录
    const fontsDir = path.join(__dirname, '..', 'Fonts')
    if (!fs.existsSync(fontsDir)) {
      console.log('[fonts-list] Fonts/ 目录不存在:', fontsDir)
      return { success: true, fonts: [] }
    }

    const entries = fs.readdirSync(fontsDir, { withFileTypes: true })
    const fontFiles = []
    const supportedExts = ['.ttf', '.otf', '.ttc', '.woff', '.woff2']

    for (const entry of entries) {
      if (!entry.isFile()) continue
      const ext = path.extname(entry.name).toLowerCase()
      if (supportedExts.includes(ext)) {
        fontFiles.push({
          name: entry.name,
          ext,
          // 文件大小（用于 UI 显示 / 优先加载小字体）
          size: fs.statSync(path.join(fontsDir, entry.name)).size,
        })
      }
    }

    console.log(`[fonts-list] 找到 ${fontFiles.length} 个字体文件`)
    return { success: true, fonts: fontFiles }
  } catch (error) {
    console.error('[fonts-list] 失败:', error)
    return { success: false, error: error.message, fonts: [] }
  }
})

// 读取单个字体文件为 base64（限制只能读 Fonts/ 下）
ipcMain.handle('fonts-read', async (event, fileName) => {
  try {
    // 安全检查：只允许读取 Fonts/ 目录下的文件
    const fontsDir = path.join(__dirname, '..', 'Fonts')
    const safeName = path.basename(String(fileName || ''))
    if (!safeName) {
      return { success: false, error: '无效的文件名' }
    }

    const filePath = path.join(fontsDir, safeName)
    // 确保路径没有逃逸出 Fonts 目录
    if (!filePath.startsWith(fontsDir)) {
      return { success: false, error: '路径不合法' }
    }

    if (!fs.existsSync(filePath)) {
      return { success: false, error: '字体文件不存在: ' + safeName }
    }

    const buffer = fs.readFileSync(filePath)
    const base64 = buffer.toString('base64')
    // 日志已移除，减少控制台输出
    return { success: true, base64, size: buffer.length }
  } catch (error) {
    console.error('[fonts-read] 失败:', safeName, error.message)
    return { success: false, error: error.message }
  }
})
