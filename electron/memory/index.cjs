const fs = require('fs')
const path = require('path')
const crypto = require('crypto')
const chokidar = require('chokidar')

let Database
try {
  Database = require('better-sqlite3')
} catch (err) {
  Database = null
}

const DEFAULT_CHUNK_SIZE = 1200
const DEFAULT_CHUNK_OVERLAP_LINES = 2
const DEFAULT_SNIPPET_MAX = 600
const SESSION_CURSOR_FILE = 'session-cursors.json'
const DEFAULT_WORKSPACE_KEY = 'global'
const DEFAULT_PROFILE_REJECTION_TTL_DAYS = 30
const PROFILE_SOURCE = 'profile'

const normalizeText = (text) => (text || '').replace(/\s+/g, ' ').trim()

const ensureDir = (dir) => {
  fs.mkdirSync(dir, { recursive: true })
}

const hashText = (text) =>
  crypto.createHash('sha1').update(text || '').digest('hex')

const nowIso = () => new Date().toISOString()

const addDaysIso = (days) => {
  const date = new Date()
  date.setDate(date.getDate() + days)
  return date.toISOString()
}

const safeJsonParse = (value, fallback) => {
  if (!value) return fallback
  try {
    return JSON.parse(value)
  } catch {
    return fallback
  }
}

const safeReadFile = (filePath) => {
  try {
    return fs.readFileSync(filePath, 'utf-8')
  } catch (e) {
    return ''
  }
}

const buildSnippet = (text, maxLen = DEFAULT_SNIPPET_MAX) => {
  const cleaned = normalizeText(text)
  if (cleaned.length <= maxLen) return cleaned
  return cleaned.slice(0, maxLen) + '...'
}

const chunkTextByLines = (text, options = {}) => {
  const chunkSize = options.chunkSize || DEFAULT_CHUNK_SIZE
  const overlapLines = options.overlapLines ?? DEFAULT_CHUNK_OVERLAP_LINES
  const lines = text.split(/\r?\n/)
  const chunks = []
  let i = 0
  let lastChunkEnd = -1  // 防止无限循环
  while (i < lines.length) {
    const startLine = i + 1
    let length = 0
    const buffer = []
    while (i < lines.length) {
      const line = lines[i]
      const nextLen = length + line.length + 1
      if (buffer.length > 0 && nextLen > chunkSize) break
      buffer.push(line)
      length = nextLen
      i += 1
    }
    if (buffer.length === 0 && i < lines.length) {
      buffer.push(lines[i])
      i += 1
    }
    const endLine = startLine + buffer.length - 1
    const content = buffer.join('\n')
    if (normalizeText(content)) {
      chunks.push({ content, startLine, endLine })
    }
    // 防止无限循环：只有当 chunk 结束位置前进时才应用 overlap
    if (overlapLines > 0 && i > lastChunkEnd + overlapLines) {
      const newI = i - overlapLines
      if (newI > lastChunkEnd) {
        i = newI
      }
    }
    lastChunkEnd = Math.max(lastChunkEnd, i)
    // 额外安全检查：如果 i 没有前进，强制前进
    if (i <= startLine - 1) {
      i = startLine
    }
  }
  return chunks
}

const embedTextLocal = (text, dims = 128) => {
  const vec = new Array(dims).fill(0)
  const tokens = (text || '').toLowerCase().match(/[\p{L}\p{N}]+/gu) || []
  for (const token of tokens) {
    let hash = 0
    for (let i = 0; i < token.length; i += 1) {
      hash = (hash * 31 + token.charCodeAt(i)) >>> 0
    }
    const idx = hash % dims
    vec[idx] += 1
  }
  const norm = Math.sqrt(vec.reduce((sum, v) => sum + v * v, 0)) || 1
  return vec.map((v) => v / norm)
}

const cosineSimilarity = (a, b) => {
  if (!a || !b || a.length !== b.length) return 0
  let dot = 0
  let normA = 0
  let normB = 0
  for (let i = 0; i < a.length; i += 1) {
    dot += a[i] * b[i]
    normA += a[i] * a[i]
    normB += b[i] * b[i]
  }
  const denom = Math.sqrt(normA) * Math.sqrt(normB)
  return denom ? dot / denom : 0
}

const buildFtsQuery = (text) => {
  const tokens = (text || '').match(/[\p{L}\p{N}]+/gu) || []
  return tokens.join(' ')
}

const ensureFtsTable = (db) => {
  const columns = db.prepare('PRAGMA table_info(chunks_fts)').all()
  const hasTable = columns.length > 0
  const needsUpgrade = hasTable && !columns.some((c) => c.name === 'workspaceKey')
  if (needsUpgrade) {
    db.exec('DROP TABLE IF EXISTS chunks_fts;')
  }
  if (!hasTable || needsUpgrade) {
    db.exec(`
      CREATE VIRTUAL TABLE IF NOT EXISTS chunks_fts USING fts5(
        content,
        path UNINDEXED,
        source UNINDEXED,
        workspaceKey UNINDEXED,
        sessionId UNINDEXED,
        startLine UNINDEXED,
        endLine UNINDEXED,
        embedding UNINDEXED
      );
    `)
  }
  return needsUpgrade
}

const initDb = (db) => {
  db.exec(`
    CREATE TABLE IF NOT EXISTS memory_meta (
      key TEXT PRIMARY KEY,
      value TEXT
    );
  `)
  db.exec(`
    CREATE TABLE IF NOT EXISTS memory_files (
      path TEXT PRIMARY KEY,
      mtimeMs INTEGER,
      size INTEGER,
      hash TEXT,
      source TEXT,
      workspaceKey TEXT,
      sessionId TEXT
    );
  `)
  db.exec(`
    CREATE TABLE IF NOT EXISTS profile_pending (
      id TEXT PRIMARY KEY,
      category TEXT NOT NULL,
      statement TEXT NOT NULL,
      evidenceHash TEXT NOT NULL,
      evidenceText TEXT NOT NULL,
      sourceScope TEXT,
      sourcePath TEXT,
      metadataJson TEXT,
      createdAt TEXT NOT NULL,
      updatedAt TEXT NOT NULL
    );
  `)
  db.exec(`
    CREATE TABLE IF NOT EXISTS profile_facts (
      id TEXT PRIMARY KEY,
      category TEXT NOT NULL,
      statement TEXT NOT NULL,
      evidenceHash TEXT NOT NULL,
      evidenceText TEXT NOT NULL,
      sourceScope TEXT,
      sourcePath TEXT,
      metadataJson TEXT,
      createdAt TEXT NOT NULL,
      updatedAt TEXT NOT NULL
    );
  `)
  db.exec(`
    CREATE TABLE IF NOT EXISTS profile_rejections (
      id TEXT PRIMARY KEY,
      category TEXT NOT NULL,
      statement TEXT NOT NULL,
      evidenceHash TEXT NOT NULL,
      sourceScope TEXT,
      sourcePath TEXT,
      metadataJson TEXT,
      rejectedUntil TEXT NOT NULL,
      createdAt TEXT NOT NULL,
      updatedAt TEXT NOT NULL
    );
  `)
  const fileCols = db.prepare('PRAGMA table_info(memory_files)').all()
  const addCol = (name) => {
    if (!fileCols.some((c) => c.name === name)) {
      db.exec(`ALTER TABLE memory_files ADD COLUMN ${name} TEXT`)
    }
  }
  addCol('source')
  addCol('workspaceKey')
  addCol('sessionId')

  db.exec(`
    CREATE VIRTUAL TABLE IF NOT EXISTS profile_facts_fts USING fts5(
      statement,
      category UNINDEXED,
      factId UNINDEXED,
      sourceScope UNINDEXED,
      sourcePath UNINDEXED,
      embedding UNINDEXED
    );
  `)

  return ensureFtsTable(db)
}

class MemoryManager {
  constructor({ baseDir }) {
    this.baseDir = baseDir
    this.dailyDir = path.join(baseDir, 'daily')
    this.sessionsDir = path.join(baseDir, 'sessions')
    this.longMemoryPath = path.join(baseDir, 'MEMORY.md')
    this.dbPath = path.join(baseDir, 'index.sqlite')
    this.cursorPath = path.join(baseDir, SESSION_CURSOR_FILE)
    this.dirty = true
    ensureDir(baseDir)
    ensureDir(this.dailyDir)
    ensureDir(this.sessionsDir)
    if (Database) {
      this.db = new Database(this.dbPath)
      const upgraded = initDb(this.db)
      if (upgraded) {
        this.dirty = true
      }
    } else {
      this.db = null
    }
    this.watchTimer = null
    this.watcher = null
    this.startWatch()
  }

  getMemoryFiles() {
    const files = []
    if (fs.existsSync(this.dailyDir)) {
      const daily = fs.readdirSync(this.dailyDir)
      for (const name of daily) {
        if (name.endsWith('.md')) {
          files.push({ source: 'daily', path: path.join(this.dailyDir, name) })
        }
      }
    }
    if (fs.existsSync(this.longMemoryPath)) {
      files.push({ source: 'long', path: this.longMemoryPath })
    }
    if (fs.existsSync(this.sessionsDir)) {
      const sessions = fs.readdirSync(this.sessionsDir)
      for (const name of sessions) {
        if (name.endsWith('.jsonl')) {
          files.push({ source: 'sessions', path: path.join(this.sessionsDir, name) })
        }
      }
    }
    return files
  }

  markDirty() {
    this.dirty = true
  }

  appendDaily({ text, source = 'chat', tags = [] }) {
    const date = new Date()
    const fileName = date.toISOString().slice(0, 10) + '.md'
    const filePath = path.join(this.dailyDir, fileName)
    const stamp = date.toISOString()
    const header = `\n## ${stamp} [${source}${tags.length ? ` | ${tags.join(',')}` : ''}]\n`
    fs.appendFileSync(filePath, header + text.trim() + '\n')
    this.markDirty()
    return { success: true, filePath }
  }

  clear(scope = 'all') {
    if (scope === 'all' || scope === 'daily') {
      if (fs.existsSync(this.dailyDir)) {
        const daily = fs.readdirSync(this.dailyDir)
        for (const name of daily) {
          if (name.endsWith('.md')) {
            fs.unlinkSync(path.join(this.dailyDir, name))
          }
        }
      }
    }
    if (scope === 'all' || scope === 'long') {
      if (fs.existsSync(this.longMemoryPath)) {
        fs.unlinkSync(this.longMemoryPath)
      }
    }
    if (scope === 'all' || scope === 'sessions') {
      if (fs.existsSync(this.sessionsDir)) {
        const sessions = fs.readdirSync(this.sessionsDir)
        for (const name of sessions) {
          if (name.endsWith('.jsonl')) {
            fs.unlinkSync(path.join(this.sessionsDir, name))
          }
        }
      }
      if (fs.existsSync(this.cursorPath)) {
        fs.unlinkSync(this.cursorPath)
      }
    }
    if (this.db) {
      this.db.exec('DELETE FROM chunks_fts;')
      this.db.exec('DELETE FROM memory_files;')
      this.db.exec('DELETE FROM memory_meta;')
      this.db.exec('DELETE FROM profile_pending;')
      this.db.exec('DELETE FROM profile_facts;')
      this.db.exec('DELETE FROM profile_rejections;')
      this.db.exec('DELETE FROM profile_facts_fts;')
    }
    this.markDirty()
    return { success: true }
  }

  getStatus() {
    if (!this.db) {
      return {
        success: false,
        message: 'better-sqlite3 未安装',
        memoryDir: this.baseDir,
        fileCount: this.getMemoryFiles().length,
      }
    }
    const chunkCount = this.db.prepare('SELECT count(1) as count FROM chunks_fts').get()?.count || 0
    const fileCount = this.db.prepare('SELECT count(1) as count FROM memory_files').get()?.count || 0
    const lastIndexedAt = this.db.prepare('SELECT value FROM memory_meta WHERE key = ?').get('lastIndexedAt')?.value || null
    return { success: true, memoryDir: this.baseDir, fileCount, chunkCount, lastIndexedAt }
  }

  startWatch() {
    if (!this.db) return
    if (this.watcher) return
    const watchTargets = [this.dailyDir, this.longMemoryPath, this.sessionsDir]
    this.watcher = chokidar.watch(watchTargets, { ignoreInitial: true })
    const onChange = () => {
      this.markDirty()
      this.scheduleReindex()
    }
    this.watcher.on('add', onChange)
    this.watcher.on('change', onChange)
    this.watcher.on('unlink', onChange)
  }

  scheduleReindex() {
    if (this.watchTimer) clearTimeout(this.watchTimer)
    this.watchTimer = setTimeout(() => {
      try {
        this.ensureIndex()
      } catch {
        // ignore
      }
    }, 1500)
  }

  getStatusDetail() {
    if (!this.db) {
      return {
        success: false,
        message: 'better-sqlite3 未安装',
      }
    }
    const sources = this.db.prepare('SELECT source, count(1) as count FROM chunks_fts GROUP BY source').all()
    const fileSources = this.db.prepare('SELECT source, count(1) as count FROM memory_files GROUP BY source').all()
    const lastIndexedAt = this.db.prepare('SELECT value FROM memory_meta WHERE key = ?').get('lastIndexedAt')?.value || null
    return {
      success: true,
      memoryDir: this.baseDir,
      chunkSources: sources,
      fileSources,
      lastIndexedAt,
    }
  }

  listPendingProfile() {
    if (!this.db) {
      return { success: false, items: [], error: 'better-sqlite3 未安装' }
    }
    const rows = this.db.prepare(
      `SELECT id, category, statement, evidenceHash, evidenceText, sourceScope, sourcePath, metadataJson, createdAt, updatedAt
       FROM profile_pending
       ORDER BY datetime(createdAt) DESC`,
    ).all()
    return {
      success: true,
      items: rows.map((row) => ({
        id: row.id,
        category: row.category,
        statement: row.statement,
        evidenceHash: row.evidenceHash,
        evidenceText: row.evidenceText,
        sourceScope: row.sourceScope || '',
        sourcePath: row.sourcePath || '',
        metadata: safeJsonParse(row.metadataJson, {}),
        createdAt: row.createdAt,
        updatedAt: row.updatedAt,
      })),
    }
  }

  listProfileFacts() {
    if (!this.db) {
      return { success: false, items: [], error: 'better-sqlite3 未安装' }
    }
    const rows = this.db.prepare(
      `SELECT id, category, statement, evidenceHash, evidenceText, sourceScope, sourcePath, metadataJson, createdAt, updatedAt
       FROM profile_facts
       ORDER BY datetime(updatedAt) DESC, datetime(createdAt) DESC`,
    ).all()
    return {
      success: true,
      items: rows.map((row) => ({
        id: row.id,
        category: row.category,
        statement: row.statement,
        evidenceHash: row.evidenceHash,
        evidenceText: row.evidenceText,
        sourceScope: row.sourceScope || '',
        sourcePath: row.sourcePath || '',
        metadata: safeJsonParse(row.metadataJson, {}),
        createdAt: row.createdAt,
        updatedAt: row.updatedAt,
      })),
    }
  }

  queueProfileCandidates(candidates = []) {
    if (!this.db) {
      return { success: false, created: 0, skipped: 0, error: 'better-sqlite3 未安装' }
    }

    const selectPending = this.db.prepare(
      'SELECT id FROM profile_pending WHERE evidenceHash = ? OR (category = ? AND statement = ?)',
    )
    const selectFact = this.db.prepare(
      'SELECT id FROM profile_facts WHERE evidenceHash = ? OR (category = ? AND statement = ?)',
    )
    const selectRejection = this.db.prepare(
      'SELECT id, rejectedUntil FROM profile_rejections WHERE evidenceHash = ? OR (category = ? AND statement = ?)',
    )
    const insertPending = this.db.prepare(
      `INSERT INTO profile_pending (
        id, category, statement, evidenceHash, evidenceText, sourceScope, sourcePath, metadataJson, createdAt, updatedAt
      ) VALUES (
        @id, @category, @statement, @evidenceHash, @evidenceText, @sourceScope, @sourcePath, @metadataJson, @createdAt, @updatedAt
      )`,
    )

    let created = 0
    let skipped = 0
    const createdItems = []

    for (const item of candidates) {
      const category = String(item.category || '').trim()
      const statement = String(item.statement || '').replace(/\s+/g, ' ').trim()
      const evidenceText = String(item.evidenceText || '').replace(/\s+/g, ' ').trim()
      const sourceScope = String(item.sourceScope || '').trim()
      const sourcePath = String(item.sourcePath || '').trim()
      const metadata = item.metadata && typeof item.metadata === 'object' ? item.metadata : {}
      if (!category || !statement || !evidenceText) {
        skipped += 1
        continue
      }

      const evidenceHash = hashText([category, statement, evidenceText].join('\n'))
      if (selectPending.get(evidenceHash, category, statement)) {
        skipped += 1
        continue
      }
      if (selectFact.get(evidenceHash, category, statement)) {
        skipped += 1
        continue
      }
      const rejection = selectRejection.get(evidenceHash, category, statement)
      if (rejection?.rejectedUntil && new Date(rejection.rejectedUntil).getTime() > Date.now()) {
        skipped += 1
        continue
      }

      const createdAt = nowIso()
      const next = {
        id: item.id || `profile-pending-${Date.now()}-${Math.random().toString(16).slice(2)}`,
        category,
        statement,
        evidenceHash,
        evidenceText,
        sourceScope,
        sourcePath,
        metadataJson: JSON.stringify(metadata),
        createdAt,
        updatedAt: createdAt,
      }
      insertPending.run(next)
      created += 1
      createdItems.push({
        id: next.id,
        category,
        statement,
        evidenceHash,
        evidenceText,
        sourceScope,
        sourcePath,
        metadata,
        createdAt,
        updatedAt: createdAt,
      })
    }

    return { success: true, created, skipped, items: createdItems }
  }

  resolvePendingProfile({ id, action }) {
    if (!this.db) {
      return { success: false, error: 'better-sqlite3 未安装' }
    }
    const row = this.db.prepare(
      `SELECT id, category, statement, evidenceHash, evidenceText, sourceScope, sourcePath, metadataJson, createdAt, updatedAt
       FROM profile_pending WHERE id = ?`,
    ).get(id)
    if (!row) {
      return { success: false, error: '未找到待确认画像' }
    }

    const metadataJson = row.metadataJson || '{}'
    const now = nowIso()

    if (action === 'accept') {
      this.db.prepare(
        `INSERT OR REPLACE INTO profile_facts (
          id, category, statement, evidenceHash, evidenceText, sourceScope, sourcePath, metadataJson, createdAt, updatedAt
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
      ).run(
        row.id,
        row.category,
        row.statement,
        row.evidenceHash,
        row.evidenceText,
        row.sourceScope || '',
        row.sourcePath || '',
        metadataJson,
        row.createdAt || now,
        now,
      )
      this.db.prepare('DELETE FROM profile_rejections WHERE evidenceHash = ? OR id = ?').run(row.evidenceHash, row.id)
      this.db.prepare('DELETE FROM profile_facts_fts WHERE factId = ?').run(row.id)
      const embedding = embedTextLocal(`${row.category} ${row.statement}`)
      this.db.prepare(
        `INSERT INTO profile_facts_fts (statement, category, factId, sourceScope, sourcePath, embedding)
         VALUES (?, ?, ?, ?, ?, ?)`,
      ).run(
        row.statement,
        row.category,
        row.id,
        row.sourceScope || '',
        row.sourcePath || '',
        JSON.stringify(embedding),
      )
    } else if (action === 'reject') {
      this.db.prepare(
        `INSERT OR REPLACE INTO profile_rejections (
          id, category, statement, evidenceHash, sourceScope, sourcePath, metadataJson, rejectedUntil, createdAt, updatedAt
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
      ).run(
        row.id,
        row.category,
        row.statement,
        row.evidenceHash,
        row.sourceScope || '',
        row.sourcePath || '',
        metadataJson,
        addDaysIso(DEFAULT_PROFILE_REJECTION_TTL_DAYS),
        row.createdAt || now,
        now,
      )
    } else {
      return { success: false, error: '不支持的处理动作' }
    }

    this.db.prepare('DELETE FROM profile_pending WHERE id = ?').run(row.id)
    return {
      success: true,
      item: {
        id: row.id,
        category: row.category,
        statement: row.statement,
        evidenceHash: row.evidenceHash,
        evidenceText: row.evidenceText,
        sourceScope: row.sourceScope || '',
        sourcePath: row.sourcePath || '',
        metadata: safeJsonParse(metadataJson, {}),
        createdAt: row.createdAt,
        updatedAt: now,
      },
    }
  }

  appendSession({ sessionId, text, meta }) {
    if (!sessionId || !text) {
      return { success: false, error: 'sessionId 或 text 为空' }
    }
    const filePath = path.join(this.sessionsDir, `${sessionId}.jsonl`)
    const payload = {
      ts: new Date().toISOString(),
      text,
      meta: meta || {},
    }
    fs.appendFileSync(filePath, JSON.stringify(payload) + '\n')
    this.markDirty()
    return { success: true, filePath }
  }

  rebuildAll() {
    const files = this.getMemoryFiles()
    this.rebuildIndex(files)
    this.dirty = false
    return { success: true }
  }
  parseSessionContent(content, fallbackWorkspaceKey, sessionId) {
    const entries = []
    const lines = content.split(/\r?\n/).filter(Boolean)
    for (const line of lines) {
      try {
        const item = JSON.parse(line)
        const text = item?.text || item?.content || ''
        if (!text) continue
        const workspaceKey = item?.meta?.workspaceKey || fallbackWorkspaceKey || DEFAULT_WORKSPACE_KEY
        entries.push({
          text: String(text),
          workspaceKey,
          sessionId,
        })
      } catch {
        // ignore invalid lines
      }
    }
    return entries
  }


  loadSessionCursors() {
    if (!fs.existsSync(this.cursorPath)) return {}
    try {
      return JSON.parse(fs.readFileSync(this.cursorPath, 'utf-8'))
    } catch {
      return {}
    }
  }

  saveSessionCursors(cursors) {
    fs.writeFileSync(this.cursorPath, JSON.stringify(cursors, null, 2))
  }

  rebuildIndex(files) {
    if (!this.db) return
    this.db.exec('DELETE FROM chunks_fts;')
    this.db.exec('DELETE FROM memory_files;')
    const insertChunk = this.db.prepare(
      'INSERT INTO chunks_fts (content, path, source, workspaceKey, sessionId, startLine, endLine, embedding) VALUES (@content, @path, @source, @workspaceKey, @sessionId, @startLine, @endLine, @embedding)'
    )
    const insertFile = this.db.prepare(
      'INSERT INTO memory_files (path, mtimeMs, size, hash, source, workspaceKey, sessionId) VALUES (@path, @mtimeMs, @size, @hash, @source, @workspaceKey, @sessionId)'
    )
    const insertMeta = this.db.prepare(
      'INSERT OR REPLACE INTO memory_meta (key, value) VALUES (?, ?)'
    )

    const cursors = this.loadSessionCursors()

    for (const file of files) {
      const content = safeReadFile(file.path)
      const stat = fs.statSync(file.path)
      const fileHash = hashText(content)
      const sessionId = file.source === 'sessions'
        ? path.basename(file.path).replace(/\.jsonl$/i, '')
        : ''
      let workspaceKey = DEFAULT_WORKSPACE_KEY

      let sessionEntries = []
      if (file.source === 'sessions') {
        sessionEntries = this.parseSessionContent(content, workspaceKey, sessionId)
        if (sessionEntries.length > 0) {
          workspaceKey = sessionEntries[0].workspaceKey || workspaceKey
        }
      }
      insertFile.run({
        path: file.path,
        mtimeMs: stat.mtimeMs,
        size: stat.size,
        hash: fileHash,
        source: file.source,
        workspaceKey,
        sessionId,
      })
      if (file.source === 'sessions') {
        sessionEntries.forEach((entry, idx) => {
          const chunks = chunkTextByLines(entry.text)
          chunks.forEach((chunk) => {
            const embedding = embedTextLocal(chunk.content)
            insertChunk.run({
              content: chunk.content,
              path: file.path,
              source: file.source,
              workspaceKey: entry.workspaceKey || workspaceKey,
              sessionId,
              startLine: String(idx + 1),
              endLine: String(idx + 1),
              embedding: JSON.stringify(embedding),
            })
          })
        })
        cursors[file.path] = content.length
      } else {
        const chunks = chunkTextByLines(content)
        for (const chunk of chunks) {
          const embedding = embedTextLocal(chunk.content)
          insertChunk.run({
            content: chunk.content,
            path: file.path,
            source: file.source,
            workspaceKey,
            sessionId,
            startLine: String(chunk.startLine),
            endLine: String(chunk.endLine),
            embedding: JSON.stringify(embedding),
          })
        }
      }
    }
    this.saveSessionCursors(cursors)
    insertMeta.run('lastIndexedAt', new Date().toISOString())
  }

  rebuildIndexIncremental(files) {
    if (!this.db) return
    const insertChunk = this.db.prepare(
      'INSERT INTO chunks_fts (content, path, source, workspaceKey, sessionId, startLine, endLine, embedding) VALUES (@content, @path, @source, @workspaceKey, @sessionId, @startLine, @endLine, @embedding)'
    )
    const upsertFile = this.db.prepare(
      'INSERT OR REPLACE INTO memory_files (path, mtimeMs, size, hash, source, workspaceKey, sessionId) VALUES (@path, @mtimeMs, @size, @hash, @source, @workspaceKey, @sessionId)'
    )
    const insertMeta = this.db.prepare(
      'INSERT OR REPLACE INTO memory_meta (key, value) VALUES (?, ?)'
    )
    const cursors = this.loadSessionCursors()

    for (const file of files) {
      const stat = fs.statSync(file.path)
      let content = safeReadFile(file.path)
      let cursor = cursors[file.path] || 0
      const sessionId = file.source === 'sessions'
        ? path.basename(file.path).replace(/\.jsonl$/i, '')
        : ''
      let workspaceKey = DEFAULT_WORKSPACE_KEY

      if (file.source === 'sessions' && cursor > 0 && cursor < content.length) {
        content = content.slice(cursor)
      } else if (file.source === 'sessions' && cursor >= content.length) {
        cursor = 0
      }

      let sessionEntries = []
      if (file.source === 'sessions') {
        sessionEntries = this.parseSessionContent(content, workspaceKey, sessionId)
        if (sessionEntries.length > 0) {
          workspaceKey = sessionEntries[0].workspaceKey || workspaceKey
        }
      }

      const fileHash = hashText(content)
      upsertFile.run({
        path: file.path,
        mtimeMs: stat.mtimeMs,
        size: stat.size,
        hash: fileHash,
        source: file.source,
        workspaceKey,
        sessionId,
      })

      if (file.source !== 'sessions') {
        this.db.prepare('DELETE FROM chunks_fts WHERE path = ?').run(file.path)
      }

      if (file.source === 'sessions') {
        sessionEntries.forEach((entry, idx) => {
          const chunks = chunkTextByLines(entry.text)
          chunks.forEach((chunk) => {
            const embedding = embedTextLocal(chunk.content)
            insertChunk.run({
              content: chunk.content,
              path: file.path,
              source: file.source,
              workspaceKey: entry.workspaceKey || workspaceKey,
              sessionId,
              startLine: String(idx + 1),
              endLine: String(idx + 1),
              embedding: JSON.stringify(embedding),
            })
          })
        })
        cursors[file.path] = (cursors[file.path] || 0) + content.length
      } else {
        const chunks = chunkTextByLines(content)
        for (const chunk of chunks) {
          const embedding = embedTextLocal(chunk.content)
          insertChunk.run({
            content: chunk.content,
            path: file.path,
            source: file.source,
            workspaceKey,
            sessionId,
            startLine: String(chunk.startLine),
            endLine: String(chunk.endLine),
            embedding: JSON.stringify(embedding),
          })
        }
      }
    }

    this.saveSessionCursors(cursors)
    insertMeta.run('lastIndexedAt', new Date().toISOString())
  }

  ensureIndex() {
    if (!this.db) return
    const files = this.getMemoryFiles()
    const metaStmt = this.db.prepare('SELECT path, mtimeMs, size, hash FROM memory_files')
    const existingMeta = new Map(metaStmt.all().map((row) => [row.path, row]))

    let needsRebuild = this.dirty
    const changedFiles = []
    for (const file of files) {
      const stat = fs.statSync(file.path)
      const meta = existingMeta.get(file.path)
      if (!meta || meta.mtimeMs !== stat.mtimeMs || meta.size !== stat.size) {
        changedFiles.push(file)
        needsRebuild = true
      }
    }
    if (!needsRebuild && files.length === existingMeta.size) return

    if (changedFiles.length === files.length || this.dirty) {
      this.rebuildIndex(files)
    } else {
      this.rebuildIndexIncremental(changedFiles)
    }
    this.dirty = false
  }


  search({ query, topK = 5, textWeight = 0.6, vectorWeight = 0.4, workspaceKey, sources }) {
    if (!query || !this.db) {
      return { success: false, results: [] }
    }
    this.ensureIndex()
    const ftsQuery = buildFtsQuery(query)
    if (!ftsQuery) return { success: false, results: [] }
    const params = [ftsQuery]
    let whereSql = 'chunks_fts MATCH ?'
    if (workspaceKey) {
      whereSql += ' AND (workspaceKey = ? OR workspaceKey IS NULL OR workspaceKey = "")'
      params.push(workspaceKey)
    }
    const sourceList = Array.isArray(sources)
      ? sources
      : (sources ? [sources] : ['daily', 'sessions', PROFILE_SOURCE])
    if (sourceList.length) {
      const placeholders = sourceList.map(() => '?').join(',')
      whereSql += ` AND source IN (${placeholders})`
      params.push(...sourceList)
    }
    params.push(Math.max(topK * 4, 10))

    const stmt = this.db.prepare(
      `SELECT rowid, content, path, source, workspaceKey, sessionId, startLine, endLine, embedding, bm25(chunks_fts) as rank FROM chunks_fts WHERE ${whereSql} ORDER BY rank LIMIT ?`
    )
    const rows = stmt.all(...params)
    const queryEmbedding = embedTextLocal(query)
    const results = rows.map((row) => {
      const textScore = 1 / (1 + Math.abs(row.rank || 0))
      let vectorScore = 0
      if (row.embedding) {
        try {
          const vec = JSON.parse(row.embedding)
          vectorScore = cosineSimilarity(queryEmbedding, vec)
        } catch {
          vectorScore = 0
        }
      }
      const score = textWeight * textScore + vectorWeight * vectorScore
      return {
        path: row.path,
        source: row.source,
        workspaceKey: row.workspaceKey,
        sessionId: row.sessionId,
        startLine: Number(row.startLine || 0),
        endLine: Number(row.endLine || 0),
        score,
        snippet: buildSnippet(row.content),
      }
    })

    if (sourceList.includes(PROFILE_SOURCE)) {
      const profileParams = [ftsQuery, Math.max(topK * 2, 5)]
      const profileRows = this.db.prepare(
        `SELECT rowid, statement, category, factId, sourceScope, sourcePath, embedding, bm25(profile_facts_fts) as rank
         FROM profile_facts_fts
         WHERE profile_facts_fts MATCH ?
         ORDER BY rank
         LIMIT ?`,
      ).all(...profileParams)

      for (const row of profileRows) {
        let vectorScore = 0
        if (row.embedding) {
          try {
            vectorScore = cosineSimilarity(queryEmbedding, JSON.parse(row.embedding))
          } catch {
            vectorScore = 0
          }
        }
        const textScore = 1 / (1 + Math.abs(row.rank || 0))
        results.push({
          path: row.sourcePath || `profile:${row.factId}`,
          source: PROFILE_SOURCE,
          workspaceKey: '',
          sessionId: '',
          startLine: 1,
          endLine: 1,
          score: textWeight * textScore + vectorWeight * vectorScore,
          snippet: `${row.category}: ${row.statement}`,
        })
      }
    }

    results.sort((a, b) => b.score - a.score)
    return { success: true, results: results.slice(0, topK) }
  }
}

const createMemoryManager = (app) => {
  const baseDir = path.join(app.getPath('userData'), 'word-cursor', 'memory')
  return new MemoryManager({ baseDir })
}

module.exports = {
  createMemoryManager,
}
