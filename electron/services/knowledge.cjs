const chokidar = require('chokidar')
const crypto = require('crypto')

const SUPPORTED_EXTENSIONS = new Set([
  '.docx',
  '.doc',
  '.xlsx',
  '.xls',
  '.pdf',
  '.md',
  '.txt',
  '.json',
  '.xml',
])

const DEFAULT_CHUNK_SIZE = 1000
const DEFAULT_CHUNK_OVERLAP = 150
const DEFAULT_EMBEDDING_BASE_URL = 'https://api.siliconflow.cn/v1'
const DEFAULT_EMBEDDING_MODEL = 'BAAI/bge-m3'

function ensureDir(fs, dirPath) {
  fs.mkdirSync(dirPath, { recursive: true })
}

function hashText(text) {
  return crypto.createHash('sha1').update(text || '').digest('hex')
}

function normalizeText(text) {
  return String(text || '').replace(/\r/g, '').replace(/\s+/g, ' ').trim()
}

function stripHtml(html) {
  return String(html || '')
    .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, ' ')
    .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, ' ')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n')
    .replace(/<\/div>/gi, '\n')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/<[^>]+>/g, ' ')
    .replace(/\n{3,}/g, '\n\n')
    .replace(/[ \t]+\n/g, '\n')
    .trim()
}

function truncate(text, maxChars) {
  if (!text || text.length <= maxChars) return text
  return `${text.slice(0, maxChars)}...`
}

function buildSnippet(text, query, maxChars = 280) {
  const normalized = String(text || '').replace(/\s+/g, ' ').trim()
  if (!normalized) return ''
  const target = String(query || '').trim().toLowerCase()
  if (!target) return truncate(normalized, maxChars)
  const index = normalized.toLowerCase().indexOf(target)
  if (index < 0) return truncate(normalized, maxChars)
  const start = Math.max(0, index - Math.floor(maxChars / 3))
  const end = Math.min(normalized.length, start + maxChars)
  const prefix = start > 0 ? '...' : ''
  const suffix = end < normalized.length ? '...' : ''
  return `${prefix}${normalized.slice(start, end)}${suffix}`
}

function computeCharOverlapScore(query, text) {
  const queryChars = Array.from(String(query || '').replace(/\s+/g, ''))
  const textSet = new Set(Array.from(String(text || '').replace(/\s+/g, '')))
  if (!queryChars.length || !textSet.size) return 0
  const uniqueQueryChars = Array.from(new Set(queryChars))
  const matched = uniqueQueryChars.filter((char) => textSet.has(char)).length
  return matched / uniqueQueryChars.length
}

function safeJsonParse(fs, filePath, fallback) {
  if (!fs.existsSync(filePath)) return fallback
  try {
    return JSON.parse(fs.readFileSync(filePath, 'utf8'))
  } catch {
    return fallback
  }
}

function safeParseJsonValue(value, fallback) {
  if (!value) return fallback
  try {
    return JSON.parse(value)
  } catch {
    return fallback
  }
}

function nowIso() {
  return new Date().toISOString()
}

function normalizeRootPath(rootPath) {
  return String(rootPath || '').trim()
}

function getRelativePath(path, rootPath, filePath) {
  return path.relative(rootPath, filePath).replace(/\\/g, '/')
}

function getExtension(path, filePath) {
  return path.extname(filePath).toLowerCase()
}

function getTitleFromFile(path, filePath) {
  return path.basename(filePath, path.extname(filePath))
}

function listSupportedFiles(fs, path, rootPath) {
  const queue = [rootPath]
  const files = []

  while (queue.length > 0) {
    const current = queue.shift()
    let entries = []
    try {
      entries = fs.readdirSync(current, { withFileTypes: true })
    } catch {
      continue
    }

    for (const entry of entries) {
      if (!entry || entry.name.startsWith('.')) continue
      if (entry.name === 'node_modules') continue

      const fullPath = path.join(current, entry.name)
      if (entry.isDirectory()) {
        queue.push(fullPath)
        continue
      }

      const ext = getExtension(path, fullPath)
      if (!SUPPORTED_EXTENSIONS.has(ext)) continue
      try {
        const stats = fs.statSync(fullPath)
        files.push({
          filePath: fullPath,
          relativePath: getRelativePath(path, rootPath, fullPath),
          ext,
          size: stats.size,
          modifiedAt: stats.mtime.toISOString(),
          modifiedMs: stats.mtimeMs,
        })
      } catch {
        // ignore unreadable files
      }
    }
  }

  files.sort((left, right) => left.relativePath.localeCompare(right.relativePath))
  return files
}

function fallbackSplitText(text, chunkSize = DEFAULT_CHUNK_SIZE, chunkOverlap = DEFAULT_CHUNK_OVERLAP) {
  const normalized = String(text || '').trim()
  if (!normalized) return []
  const chunks = []
  let start = 0
  while (start < normalized.length) {
    const end = Math.min(normalized.length, start + chunkSize)
    const chunk = normalized.slice(start, end).trim()
    if (chunk) chunks.push(chunk)
    if (end >= normalized.length) break
    start = Math.max(end - chunkOverlap, start + 1)
  }
  return chunks
}

function toQueryText(item) {
  if (!item) return ''
  return [
    item.category ? `[${item.category}]` : '',
    item.statement || '',
    item.evidenceText || '',
  ]
    .filter(Boolean)
    .join(' ')
    .trim()
}

async function dynamicDefault(modulePromise) {
  const loaded = await modulePromise
  return loaded?.default || loaded
}

function createKnowledgeService(options = {}) {
  const {
    fs,
    path,
    app,
    mammoth,
    WordExtractor,
    XLSX,
    getMemoryManager,
  } = options

  const baseDir = path.join(app.getPath('userData'), 'word-cursor', 'knowledge')
  const statePath = path.join(baseDir, 'state.json')
  const lanceDir = path.join(baseDir, 'lancedb')
  ensureDir(fs, baseDir)
  ensureDir(fs, lanceDir)

  let depsPromise = null
  let dbPromise = null
  let workspaceWatcher = null
  let globalWatcher = null
  let activeWorkspacePath = ''
  const rootTimers = new Map()
  const tableCache = new Map()

  const runtimeState = safeJsonParse(fs, statePath, {
    roots: {},
  })

  const config = {
    knowledgeEnabled: true,
    workspaceKnowledgeEnabled: true,
    globalKnowledgePath: '',
    profileMemoryEnabled: true,
    embeddingBaseUrl: DEFAULT_EMBEDDING_BASE_URL,
    embeddingApiKey: '',
    embeddingModel: DEFAULT_EMBEDDING_MODEL,
    knowledgeTopK: 8,
  }

  function saveState() {
    fs.writeFileSync(statePath, JSON.stringify(runtimeState, null, 2), 'utf8')
  }

  function getRootKey(sourceType, rootPath) {
    return hashText(`${sourceType}:${normalizeRootPath(rootPath)}`)
  }

  function getTableName(rootKey) {
    return `knowledge_${rootKey.slice(0, 16)}`
  }

  function getRootState(sourceType, rootPath) {
    const normalizedRoot = normalizeRootPath(rootPath)
    if (!normalizedRoot) return null
    const rootKey = getRootKey(sourceType, normalizedRoot)
    const existing = runtimeState.roots[rootKey]
    if (existing) return existing

    const next = {
      rootKey,
      sourceType,
      rootPath: normalizedRoot,
      tableName: getTableName(rootKey),
      lastIndexedAt: null,
      lastError: '',
      lastSkippedReason: '',
      status: 'idle',
      fileCount: 0,
      indexedFileCount: 0,
      chunkCount: 0,
      files: {},
    }
    runtimeState.roots[rootKey] = next
    saveState()
    return next
  }

  async function ensureDeps() {
    if (!depsPromise) {
      depsPromise = (async () => {
        const [lancedb, pdfParseModule, textSplitters] = await Promise.all([
          dynamicDefault(import('@lancedb/lancedb')),
          import('pdf-parse'),
          import('@langchain/textsplitters'),
        ])
        return {
          lancedb,
          PDFParse: pdfParseModule.PDFParse || pdfParseModule.default?.PDFParse || null,
          RecursiveCharacterTextSplitter: textSplitters.RecursiveCharacterTextSplitter,
        }
      })()
    }
    return depsPromise
  }

  async function getDb() {
    if (!dbPromise) {
      dbPromise = (async () => {
        const deps = await ensureDeps()
        return deps.lancedb.connect(lanceDir)
      })()
    }
    return dbPromise
  }

  async function getTable(rootState, createIfMissing = false, initialRows = []) {
    if (!rootState?.tableName) return null
    if (tableCache.has(rootState.tableName)) {
      return tableCache.get(rootState.tableName)
    }

    const db = await getDb()
    let table = null
    try {
      table = await db.openTable(rootState.tableName)
    } catch {
      if (!createIfMissing || !initialRows.length) {
        return null
      }
      table = await db.createTable(rootState.tableName, initialRows)
    }

    if (table) {
      tableCache.set(rootState.tableName, table)
    }
    return table
  }

  async function splitText(text) {
    const normalized = String(text || '').trim()
    if (!normalized) return []
    try {
      const deps = await ensureDeps()
      const splitter = new deps.RecursiveCharacterTextSplitter({
        chunkSize: DEFAULT_CHUNK_SIZE,
        chunkOverlap: DEFAULT_CHUNK_OVERLAP,
      })
      const chunks = await splitter.splitText(normalized)
      return chunks.filter(Boolean)
    } catch {
      return fallbackSplitText(normalized)
    }
  }

  async function embedTexts(texts) {
    const apiKey = String(config.embeddingApiKey || '').trim()
    if (!apiKey) {
      throw new Error('未配置 Embedding API Key')
    }
    const model = String(config.embeddingModel || DEFAULT_EMBEDDING_MODEL).trim()
    const baseUrl = String(config.embeddingBaseUrl || DEFAULT_EMBEDDING_BASE_URL).replace(/\/$/, '')
    const response = await fetch(`${baseUrl}/embeddings`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${apiKey}`,
      },
      body: JSON.stringify({
        model,
        input: texts,
      }),
    })

    if (!response.ok) {
      const errorText = await response.text().catch(() => '')
      throw new Error(errorText || `Embedding 请求失败 (${response.status})`)
    }

    const payload = await response.json()
    const vectors = Array.isArray(payload?.data)
      ? payload.data
          .map((item) => item?.embedding)
          .filter((item) => Array.isArray(item))
      : []

    if (vectors.length !== texts.length) {
      throw new Error('Embedding 返回数量异常')
    }
    return vectors
  }

  async function extractTextFromDocx(filePath) {
    const buffer = fs.readFileSync(filePath)
    const result = await mammoth.convertToHtml({ buffer })
    return stripHtml(result.value || '')
  }

  async function extractTextFromDoc(filePath) {
    const extractor = new WordExtractor()
    const extracted = await extractor.extract(filePath)
    return normalizeText(extracted.getBody() || '')
  }

  async function extractWorkbookDocuments(fileMeta) {
    const workbook = XLSX.read(fs.readFileSync(fileMeta.filePath), {
      type: 'buffer',
      cellText: true,
      cellFormula: false,
      cellDates: true,
    })
    const documents = []

    for (const sheetName of workbook.SheetNames || []) {
      const worksheet = workbook.Sheets[sheetName]
      if (!worksheet || !worksheet['!ref']) continue
      const rows = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
        raw: false,
        defval: '',
      })
      const bodyText = rows
        .map((row, rowIndex) => {
          if (!Array.isArray(row)) return ''
          const normalizedRow = row.map((cell) => String(cell || '').trim())
          if (!normalizedRow.some(Boolean)) return ''
          return `Row ${rowIndex + 1}: ${normalizedRow.join(' | ')}`
        })
        .filter(Boolean)
        .join('\n')
        .trim()

      if (!bodyText) continue
      documents.push({
        sourceType: fileMeta.sourceType,
        sourceRoot: fileMeta.sourceRoot,
        relativePath: fileMeta.relativePath,
        fileType: fileMeta.fileType,
        title: `${fileMeta.title} - ${sheetName}`,
        bodyText,
        metadata: {
          sheetName,
          path: fileMeta.filePath,
        },
        hash: hashText(`${fileMeta.hash}:${sheetName}:${bodyText.length}`),
        modifiedAt: fileMeta.modifiedAt,
      })
    }

    return documents
  }

  async function extractTextFromPdf(filePath) {
    const PDFParse = (await ensureDeps()).PDFParse
    if (!PDFParse) {
      throw new Error('PDFParse class unavailable')
    }
    const parser = new PDFParse({ data: fs.readFileSync(filePath) })
    try {
      const result = await parser.getText()
      return normalizeText(result?.text || '')
    } finally {
      await parser.destroy().catch(() => {})
    }
  }

  async function buildKnowledgeDocuments(fileMeta) {
    const ext = fileMeta.fileType
    if (ext === '.xlsx' || ext === '.xls') {
      return extractWorkbookDocuments(fileMeta)
    }

    let bodyText = ''
    if (ext === '.docx') {
      bodyText = await extractTextFromDocx(fileMeta.filePath)
    } else if (ext === '.doc') {
      bodyText = await extractTextFromDoc(fileMeta.filePath)
    } else if (ext === '.pdf') {
      bodyText = await extractTextFromPdf(fileMeta.filePath)
    } else {
      bodyText = fs.readFileSync(fileMeta.filePath, 'utf8')
    }

    const normalizedBody = normalizeText(bodyText)
    if (!normalizedBody) {
      return []
    }

    return [{
      sourceType: fileMeta.sourceType,
      sourceRoot: fileMeta.sourceRoot,
      relativePath: fileMeta.relativePath,
      fileType: fileMeta.fileType,
      title: fileMeta.title,
      bodyText: normalizedBody,
      metadata: {
        path: fileMeta.filePath,
      },
      hash: fileMeta.hash,
      modifiedAt: fileMeta.modifiedAt,
    }]
  }

  async function buildChunkRows(fileMeta) {
    const documents = await buildKnowledgeDocuments(fileMeta)
    const rows = []

    for (const document of documents) {
      const chunks = await splitText(document.bodyText)
      if (!chunks.length) continue
      const vectors = await embedTexts(chunks)
      chunks.forEach((chunk, index) => {
        rows.push({
          id: `${fileMeta.hash}:${index}:${hashText(chunk).slice(0, 12)}`,
          fileHash: fileMeta.hash,
          relativePath: document.relativePath,
          title: document.title,
          text: chunk,
          vector: vectors[index],
          sourceType: document.sourceType,
          sourceRoot: document.sourceRoot,
          fileType: document.fileType,
          modifiedAt: document.modifiedAt,
          metadataJson: JSON.stringify(document.metadata || {}),
          path: fileMeta.filePath,
          sheetName: document.metadata?.sheetName || '',
          pageRange: document.metadata?.pageRange || '',
        })
      })
    }

    return rows
  }

  async function deleteFileRows(rootState, fileHash) {
    const table = await getTable(rootState)
    if (!table) return
    try {
      await table.delete(`fileHash = "${fileHash}"`)
    } catch {
      // ignore delete failures, they will be fixed on full rebuild
    }
  }

  async function upsertFileRows(rootState, fileMeta, rows) {
    if (!rows.length) return
    let table = await getTable(rootState)
    if (!table) {
      table = await getTable(rootState, true, rows)
      return
    }
    await table.add(rows)
  }

  async function reconcileRoot(sourceType, rootPath, options = {}) {
    const normalizedRoot = normalizeRootPath(rootPath)
    if (!normalizedRoot) {
      return { success: true, skipped: true, reason: 'missing root path' }
    }
    if (!fs.existsSync(normalizedRoot)) {
      const rootState = getRootState(sourceType, normalizedRoot)
      if (rootState) {
        rootState.status = 'error'
        rootState.lastError = '目录不存在'
        saveState()
      }
      return { success: false, error: '目录不存在' }
    }

    if (!config.knowledgeEnabled) {
      const rootState = getRootState(sourceType, normalizedRoot)
      if (rootState) {
        rootState.status = 'idle'
        rootState.lastSkippedReason = 'knowledge disabled'
        saveState()
      }
      return { success: true, skipped: true, reason: 'knowledge disabled' }
    }

    const rootState = getRootState(sourceType, normalizedRoot)
    rootState.status = 'indexing'
    rootState.lastError = ''
    rootState.lastSkippedReason = ''
    saveState()

    try {
      const files = listSupportedFiles(fs, path, normalizedRoot).map((file) => ({
        sourceType,
        sourceRoot: normalizedRoot,
        filePath: file.filePath,
        relativePath: file.relativePath,
        fileType: file.ext,
        title: getTitleFromFile(path, file.filePath),
        hash: hashText(`${file.filePath}:${file.modifiedMs}:${file.size}`),
        modifiedAt: file.modifiedAt,
        modifiedMs: file.modifiedMs,
        size: file.size,
      }))

      const nextPaths = new Set(files.map((file) => file.relativePath))
      const staleRelativePaths = Object.keys(rootState.files || {}).filter(
        (relativePath) => !nextPaths.has(relativePath),
      )

      for (const relativePath of staleRelativePaths) {
        const stale = rootState.files[relativePath]
        if (stale?.hash) {
          await deleteFileRows(rootState, stale.hash)
        }
        delete rootState.files[relativePath]
      }

      for (const fileMeta of files) {
        const existing = rootState.files[fileMeta.relativePath]
        if (
          !options.force &&
          existing &&
          existing.hash === fileMeta.hash &&
          existing.modifiedAt === fileMeta.modifiedAt
        ) {
          continue
        }

        if (existing?.hash) {
          await deleteFileRows(rootState, existing.hash)
        }

        try {
          const rows = await buildChunkRows(fileMeta)
          await upsertFileRows(rootState, fileMeta, rows)
          rootState.files[fileMeta.relativePath] = {
            hash: fileMeta.hash,
            modifiedAt: fileMeta.modifiedAt,
            size: fileMeta.size,
            chunkCount: rows.length,
            fileType: fileMeta.fileType,
            status: rows.length ? 'indexed' : 'skipped',
            error: rows.length ? '' : '文件无可索引文本',
            updatedAt: nowIso(),
            title: fileMeta.title,
          }
        } catch (error) {
          rootState.files[fileMeta.relativePath] = {
            hash: fileMeta.hash,
            modifiedAt: fileMeta.modifiedAt,
            size: fileMeta.size,
            chunkCount: 0,
            fileType: fileMeta.fileType,
            status: 'error',
            error: (error && error.message) || String(error),
            updatedAt: nowIso(),
            title: fileMeta.title,
          }
        }
      }

      const fileEntries = Object.values(rootState.files || {})
      rootState.fileCount = fileEntries.length
      rootState.indexedFileCount = fileEntries.filter((file) => file.status === 'indexed').length
      rootState.chunkCount = fileEntries.reduce((sum, file) => sum + (file.chunkCount || 0), 0)
      rootState.lastIndexedAt = nowIso()
      rootState.status = 'ready'
      rootState.lastError = ''
      saveState()
      return { success: true, rootKey: rootState.rootKey }
    } catch (error) {
      rootState.status = 'error'
      rootState.lastError = (error && error.message) || String(error)
      rootState.lastIndexedAt = nowIso()
      saveState()
      return { success: false, error: rootState.lastError }
    }
  }

  function scheduleRootReconcile(sourceType, rootPath, options = {}) {
    const normalizedRoot = normalizeRootPath(rootPath)
    if (!normalizedRoot) return
    const rootKey = getRootKey(sourceType, normalizedRoot)
    const existing = rootTimers.get(rootKey)
    if (existing) clearTimeout(existing)
    rootTimers.set(
      rootKey,
      setTimeout(() => {
        rootTimers.delete(rootKey)
        void reconcileRoot(sourceType, normalizedRoot, options)
      }, 600),
    )
  }

  function restartWorkspaceWatcher() {
    if (workspaceWatcher) {
      workspaceWatcher.close().catch(() => {})
      workspaceWatcher = null
    }
    if (!config.knowledgeEnabled || !config.workspaceKnowledgeEnabled || !activeWorkspacePath) return
    workspaceWatcher = chokidar.watch(activeWorkspacePath, { ignoreInitial: true })
    const onChange = () => scheduleRootReconcile('workspace', activeWorkspacePath)
    workspaceWatcher.on('add', onChange)
    workspaceWatcher.on('change', onChange)
    workspaceWatcher.on('unlink', onChange)
  }

  function restartGlobalWatcher() {
    if (globalWatcher) {
      globalWatcher.close().catch(() => {})
      globalWatcher = null
    }
    const globalRoot = normalizeRootPath(config.globalKnowledgePath)
    if (!config.knowledgeEnabled || !globalRoot) return
    globalWatcher = chokidar.watch(globalRoot, { ignoreInitial: true })
    const onChange = () => scheduleRootReconcile('global', globalRoot)
    globalWatcher.on('add', onChange)
    globalWatcher.on('change', onChange)
    globalWatcher.on('unlink', onChange)
  }

  function toKnowledgeSearchResult(item, query) {
    return {
      sourceScope: item.sourceScope || item.sourceType || item.source || '',
      sourcePath: item.path || item.sourcePath || '',
      relativePath: item.relativePath || '',
      fileType: item.fileType || '',
      title: item.title || '',
      score: typeof item.score === 'number'
        ? item.score
        : (typeof item._distance === 'number' ? 1 - item._distance : 0),
      snippet: buildSnippet(item.text || item.snippet || item.statement || '', query),
      metadata: safeParseJsonValue(item.metadataJson, {}),
      category: item.category || '',
      statement: item.statement || '',
    }
  }

  async function searchRoot(sourceType, rootPath, query, topK) {
    const rootState = getRootState(sourceType, rootPath)
    if (!rootState) return []
    if (!rootState.fileCount) {
      await reconcileRoot(sourceType, rootPath)
    }

    const table = await getTable(rootState)
    if (!table) return []
    const [queryVector] = await embedTexts([query])
    const rows = await table.search(queryVector).limit(Math.max(topK, 1)).toArray()
    return rows
      .map((row) => {
        const baseScore = typeof row._distance === 'number' ? 1 - row._distance : 0
        const charScore = computeCharOverlapScore(query, row.text || '')
        const score = Math.max(0, baseScore * 0.75 + charScore * 0.25)
        return toKnowledgeSearchResult({
          sourceType,
          path: row.path || '',
          relativePath: row.relativePath || '',
          fileType: row.fileType || '',
          title: row.title || '',
          text: row.text || '',
          metadataJson: row.metadataJson || '{}',
          _distance: row._distance,
          score,
          charScore,
        }, query)
      })
      .filter((item) => item.score >= 0.12 || computeCharOverlapScore(query, item.snippet || '') >= 0.08)
  }

  async function retrieve(options = {}) {
    const query = String(options.query || '').trim()
    if (!query) {
      return { success: false, results: [], error: 'query 为空' }
    }

    const topK = Math.max(1, Number(options.topK) || config.knowledgeTopK || 8)
    const workspaceHits = config.knowledgeEnabled && config.workspaceKnowledgeEnabled && activeWorkspacePath
      ? await searchRoot('workspace', activeWorkspacePath, query, Math.min(5, topK))
      : []
    const globalRoot = normalizeRootPath(config.globalKnowledgePath)
    const globalHits = config.knowledgeEnabled && globalRoot
      ? await searchRoot('global', globalRoot, query, Math.min(4, topK))
      : []

    const memoryManager = getMemoryManager()
    const profileHits = config.profileMemoryEnabled
      ? memoryManager.search({
          query,
          topK: 2,
          textWeight: 0.7,
          vectorWeight: 0.3,
          sources: ['profile'],
        }).results || []
      : []
    const sessionHits = memoryManager.search({
      query,
      topK: 1,
      textWeight: 0.7,
      vectorWeight: 0.3,
      sources: ['daily', 'sessions'],
    }).results || []

    const results = [
      ...workspaceHits.map((item) => ({ ...item, priority: 0 })),
      ...profileHits.map((item) =>
        toKnowledgeSearchResult({
          sourceScope: 'profile',
          sourcePath: item.path,
          title: '用户画像',
          statement: item.snippet,
          score: item.score,
          category: 'profile',
          snippet: item.snippet,
        }, query),
      ).map((item) => ({ ...item, priority: 1 })),
      ...globalHits.map((item) => ({ ...item, priority: 2 })),
      ...sessionHits.map((item) =>
        toKnowledgeSearchResult({
          sourceScope: item.source,
          sourcePath: item.path,
          title: '历史记忆',
          snippet: item.snippet,
          score: item.score,
        }, query),
      ).map((item) => ({ ...item, priority: 3 })),
    ]

    results.sort((left, right) => left.priority - right.priority || right.score - left.score)
    const seen = new Set()
    const deduped = []
    for (const item of results) {
      const key = [
        item.sourceScope,
        item.sourcePath,
        item.relativePath,
        item.category,
        item.statement,
        item.snippet,
      ].join('::')
      if (seen.has(key)) continue
      seen.add(key)
      deduped.push(item)
    }
    return {
      success: true,
      results: deduped.slice(0, Math.min(12, Math.max(4, topK))).map(({ priority, ...rest }) => rest),
    }
  }

  async function configure(nextConfig = {}) {
    Object.assign(config, {
      knowledgeEnabled: nextConfig.knowledgeEnabled !== false,
      workspaceKnowledgeEnabled: nextConfig.workspaceKnowledgeEnabled !== false,
      globalKnowledgePath: normalizeRootPath(nextConfig.globalKnowledgePath || ''),
      profileMemoryEnabled: nextConfig.profileMemoryEnabled !== false,
      embeddingBaseUrl: String(nextConfig.embeddingBaseUrl || DEFAULT_EMBEDDING_BASE_URL).trim(),
      embeddingApiKey: String(nextConfig.embeddingApiKey || '').trim(),
      embeddingModel: String(nextConfig.embeddingModel || DEFAULT_EMBEDDING_MODEL).trim(),
      knowledgeTopK: Math.max(1, Number(nextConfig.knowledgeTopK) || 8),
    })

    restartWorkspaceWatcher()
    restartGlobalWatcher()

    if (config.globalKnowledgePath) {
      scheduleRootReconcile('global', config.globalKnowledgePath)
    }

    return status()
  }

  async function setActiveWorkspace(payload = {}) {
    activeWorkspacePath = normalizeRootPath(payload.workspacePath || '')
    restartWorkspaceWatcher()
    if (activeWorkspacePath && config.workspaceKnowledgeEnabled) {
      scheduleRootReconcile('workspace', activeWorkspacePath)
    }
    return { success: true, workspacePath: activeWorkspacePath }
  }

  async function status() {
    const memoryManager = getMemoryManager()
    const pending = memoryManager.listPendingProfile()
    const facts = memoryManager.listProfileFacts()

    const workspaceRoot = activeWorkspacePath ? getRootState('workspace', activeWorkspacePath) : null
    const globalRoot = config.globalKnowledgePath ? getRootState('global', config.globalKnowledgePath) : null
    const missingDeps = []
    try {
      await ensureDeps()
    } catch (error) {
      missingDeps.push((error && error.message) || String(error))
    }

    return {
      success: missingDeps.length === 0,
      configured: {
        knowledgeEnabled: config.knowledgeEnabled,
        workspaceKnowledgeEnabled: config.workspaceKnowledgeEnabled,
        profileMemoryEnabled: config.profileMemoryEnabled,
        globalKnowledgePath: config.globalKnowledgePath,
        embeddingBaseUrl: config.embeddingBaseUrl,
        embeddingModel: config.embeddingModel,
        embeddingConfigured: !!config.embeddingApiKey,
        knowledgeTopK: config.knowledgeTopK,
      },
      workspace: workspaceRoot ? {
        rootPath: workspaceRoot.rootPath,
        status: workspaceRoot.status,
        fileCount: workspaceRoot.fileCount,
        indexedFileCount: workspaceRoot.indexedFileCount,
        chunkCount: workspaceRoot.chunkCount,
        lastIndexedAt: workspaceRoot.lastIndexedAt,
        lastError: workspaceRoot.lastError || '',
      } : null,
      global: globalRoot ? {
        rootPath: globalRoot.rootPath,
        status: globalRoot.status,
        fileCount: globalRoot.fileCount,
        indexedFileCount: globalRoot.indexedFileCount,
        chunkCount: globalRoot.chunkCount,
        lastIndexedAt: globalRoot.lastIndexedAt,
        lastError: globalRoot.lastError || '',
      } : null,
      profile: {
        pendingCount: pending.success ? pending.items.length : 0,
        factCount: facts.success ? facts.items.length : 0,
      },
      error: missingDeps.join('; ') || undefined,
    }
  }

  async function rebuild(payload = {}) {
    const scope = String(payload.scope || 'all')
    const force = true
    if ((scope === 'all' || scope === 'workspace') && activeWorkspacePath && config.workspaceKnowledgeEnabled) {
      await reconcileRoot('workspace', activeWorkspacePath, { force })
    }
    if ((scope === 'all' || scope === 'global') && config.globalKnowledgePath) {
      await reconcileRoot('global', config.globalKnowledgePath, { force })
    }
    return status()
  }

  async function listPendingProfile() {
    return getMemoryManager().listPendingProfile()
  }

  async function resolvePendingProfile(payload = {}) {
    const action = String(payload.action || '').trim()
    const ids = Array.isArray(payload.ids)
      ? payload.ids.map((id) => String(id)).filter(Boolean)
      : (payload.id ? [String(payload.id)] : [])
    if (!ids.length || !action) {
      return { success: false, error: 'id/action 缺失' }
    }
    const items = []
    for (const id of ids) {
      const result = getMemoryManager().resolvePendingProfile({ id, action })
      if (!result.success) return result
      if (result.item) items.push(result.item)
    }
    return { success: true, items }
  }

  async function listProfileFacts() {
    return getMemoryManager().listProfileFacts()
  }

  async function queueProfileCandidates(payload = {}) {
    const items = Array.isArray(payload.items) ? payload.items : []
    return getMemoryManager().queueProfileCandidates(items)
  }

  async function close() {
    if (workspaceWatcher) {
      await workspaceWatcher.close().catch(() => {})
      workspaceWatcher = null
    }
    if (globalWatcher) {
      await globalWatcher.close().catch(() => {})
      globalWatcher = null
    }
    for (const timer of rootTimers.values()) {
      clearTimeout(timer)
    }
    rootTimers.clear()
  }

  return {
    configure,
    setActiveWorkspace,
    status,
    retrieve,
    rebuild,
    listPendingProfile,
    resolvePendingProfile,
    listProfileFacts,
    queueProfileCandidates,
    close,
  }
}

module.exports = {
  createKnowledgeService,
}
