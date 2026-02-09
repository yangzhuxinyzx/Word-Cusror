/**
 * DOCX Agent Context Generator
 *
 * 将 .docx 文件解析为可供 Agent 理解的结构化文本快照，
 * 包含页面设置、页眉页脚、正文结构与格式信息。
 *
 * 性能关键点：
 * - 解析在 Web Worker 中执行（避免阻塞 UI / 网络）
 * - 严禁把图片转成 base64 内联（会导致超大字符串、卡死与请求失败）
 */

import type { PageSettings } from './docxParser'
import { docxHtmlToElements, type FormattedElementLike } from './docxHtmlToElements'

// 配置项
interface DocxAgentContextOptions {
  /** 最大输出字符数（默认 12000，约 3000 token） */
  maxLength?: number
  /** 是否包含完整 HTML（默认 false，仅包含结构摘要） */
  includeFullHtml?: boolean
  /** 正文摘要最大段落数（默认 50） */
  maxParagraphs?: number
  /** 每段最大字符数（默认 200） */
  maxParagraphLength?: number
  /** 解析超时（ms），默认 15000 */
  timeoutMs?: number
}

// 缓存（按来源长度作为 key）
const contextCache = new Map<string, { result: string; timestamp: number }>()
const CACHE_TTL = 5 * 60 * 1000 // 5 分钟缓存

type WorkerOk = {
  html: string
  rawText: string
  pageSettings: PageSettings
  headerText?: string
  footerText?: string
}

type WorkerResp =
  | { id: string; ok: true } & WorkerOk
  | { id: string; ok: false; error: string }

let workerSingleton: Worker | null = null
const pendingWorker = new Map<
  string,
  { resolve: (v: WorkerOk) => void; reject: (e: Error) => void; timeoutId: number }
>()

function getWorker(): Worker {
  if (workerSingleton) return workerSingleton
  workerSingleton = new Worker(new URL('../workers/docxAgentContextWorker.ts', import.meta.url), {
    type: 'module',
  })
  workerSingleton.onmessage = (event: MessageEvent<WorkerResp>) => {
    const data = event.data
    const pending = pendingWorker.get(data.id)
    if (!pending) return
    pendingWorker.delete(data.id)
    clearTimeout(pending.timeoutId)
    if (data.ok) pending.resolve(data)
    else pending.reject(new Error(data.error))
  }
  workerSingleton.onerror = (err) => {
    // Fail all pending jobs
    for (const [id, pending] of pendingWorker.entries()) {
      clearTimeout(pending.timeoutId)
      pending.reject(new Error(err.message || 'Worker error'))
      pendingWorker.delete(id)
    }
  }
  return workerSingleton
}

function runWorker(arrayBuffer: ArrayBuffer, timeoutMs: number): Promise<WorkerOk> {
  const w = getWorker()
  const id = `${Date.now()}-${Math.random().toString(16).slice(2)}`
  return new Promise((resolve, reject) => {
    const timeoutId = window.setTimeout(() => {
      pendingWorker.delete(id)
      reject(new Error('DOCX 解析超时'))
    }, timeoutMs)
    pendingWorker.set(id, { resolve, reject, timeoutId })
    w.postMessage({ id, arrayBuffer }, [arrayBuffer])
  })
}

async function fetchArrayBufferFromLocalFileServer(filePath: string): Promise<ArrayBuffer> {
  // Electron main process starts a local file server (port may be dynamic if 9090 is occupied).
  // Prefer asking main process for the actual local URL.
  let url = `http://localhost:9090/file/${encodeURIComponent(filePath)}`
  try {
    if (window.electronAPI?.getLocalFileUrl) {
      url = await window.electronAPI.getLocalFileUrl(filePath)
    }
  } catch {
    // fallback to 9090
  }
  const resp = await fetch(url)
  if (!resp.ok) {
    throw new Error(`无法读取文件内容（${resp.status}）`)
  }
  return await resp.arrayBuffer()
}

function base64ToArrayBuffer(base64: string): ArrayBuffer {
  // WARNING: This can be heavy for huge docx. Prefer file server path in Electron.
  const binaryString = atob(base64)
  const len = binaryString.length
  const bytes = new Uint8Array(len)
  for (let i = 0; i < len; i++) {
    bytes[i] = binaryString.charCodeAt(i)
  }
  return bytes.buffer
}

/**
 * 从 HTML 中提取纯文本
 */
function htmlToPlainText(html: string): string {
  if (!html) return ''
  // 移除 style 和 script 标签及其内容
  let text = html.replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
  text = text.replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
  // 替换常见 HTML 实体
  text = text.replace(/&nbsp;/g, ' ')
  text = text.replace(/&amp;/g, '&')
  text = text.replace(/&lt;/g, '<')
  text = text.replace(/&gt;/g, '>')
  text = text.replace(/&quot;/g, '"')
  text = text.replace(/&#39;/g, "'")
  // 移除所有 HTML 标签
  text = text.replace(/<[^>]+>/g, ' ')
  // 清理多余空白
  text = text.replace(/\s+/g, ' ').trim()
  return text
}

const FIELD_PLACEHOLDER_REGEX = /_{3,}|（\s*）|\(\s*\)|【\s*】|\[\s*\]|\{\s*\}/

function normalizeInlineText(text: string): string {
  return (text || '').replace(/\s+/g, ' ').trim()
}

function normalizeLabelKey(label: string): string {
  return normalizeInlineText(label).toLowerCase().replace(/[：:，,。\s]/g, '')
}

function inferFieldType(label: string, value?: string): string {
  const text = `${label || ''} ${value || ''}`.toLowerCase()
  if (/日期|时间|date/.test(text)) return 'date'
  if (/金额|费用|总价|¥|\$|usd|rmb/.test(text)) return 'amount'
  if (/姓名|联系人|负责人|作者|签名|name/.test(text)) return 'person'
  if (/单位|公司|机构|部门|org|company/.test(text)) return 'organization'
  if (/电话|手机|联系方式|tel|phone/.test(text)) return 'phone'
  if (/邮箱|邮件|email/.test(text)) return 'email'
  if (/地址|住址|address/.test(text)) return 'address'
  if (/编号|编号|证件号|合同号|id|编号/.test(text)) return 'id'
  if (/%/.test(text)) return 'percent'
  return 'text'
}

type AgentFieldCandidate = {
  id: string
  label: string
  kind: 'colon' | 'blank' | 'table'
  path: string
  context: string
  currentValue?: string
  fieldType?: string
  groupKey?: string
}

function buildStructureIndexFromElements(
  elements: FormattedElementLike[],
  options: { maxIndexLines: number; maxFieldCandidates: number; maxCellText: number }
): { indexLines: string[]; fieldLines: string[] } {
  const indexLines: string[] = []
  const fieldLines: string[] = []
  const fieldCandidates: AgentFieldCandidate[] = []

  const headingStack: Array<{ level: number; index: number; text: string }> = []
  const tagCounters: Record<string, number> = { p: 0, table: 0, h1: 0, h2: 0, h3: 0, h4: 0, h5: 0, h6: 0 }
  let blockIndex = -1
  let indexLineCount = 0

  const pushIndex = (line: string) => {
    if (indexLineCount >= options.maxIndexLines) return
    indexLines.push(line)
    indexLineCount += 1
  }

  for (const el of elements) {
    if (!el) continue
    if (el.type === 'heading') {
      blockIndex += 1
      const level = Math.min(Math.max(el.level || 1, 1), 6)
      const tag = `h${level}`
      tagCounters[tag] = (tagCounters[tag] || 0) + 1
      const text = normalizeInlineText(el.content || '')
      while (headingStack.length && headingStack[headingStack.length - 1].level >= level) {
        headingStack.pop()
      }
      headingStack.push({ level, index: tagCounters[tag], text })
      const path = headingStack.map((h) => `h${h.level}[${h.index}]`).join('/')
      if (text) {
        pushIndex(`${path}: ${text.slice(0, 80)}`)
      }
      continue
    }

    if (el.type === 'paragraph') {
      blockIndex += 1
      tagCounters.p += 1
      const text = normalizeInlineText(el.content || '')
      const headingPath = headingStack.map((h) => `h${h.level}[${h.index}]`).join('/')
      const path = headingPath ? `${headingPath}/p[${tagCounters.p}]` : `p[${tagCounters.p}]`
      if (text) {
        pushIndex(`${path}: ${text.slice(0, 80)}`)
      }

      const colonIndex = text.indexOf('：') >= 0 ? text.indexOf('：') : text.indexOf(':')
      if (colonIndex > -1) {
        const label = text.slice(0, colonIndex).trim()
        const tail = text.slice(colonIndex + 1).trim()
        if (label && (!tail || FIELD_PLACEHOLDER_REGEX.test(tail))) {
          const fieldType = inferFieldType(label, tail)
          fieldCandidates.push({
            id: `p:${blockIndex}:colon`,
            label,
            kind: 'colon',
            path,
            context: text.slice(0, 80),
            currentValue: tail || '',
            fieldType,
            groupKey: normalizeLabelKey(label),
          })
        }
      }

      if (fieldCandidates.length < options.maxFieldCandidates) {
        const blankMatch = text.match(FIELD_PLACEHOLDER_REGEX)
        if (blankMatch) {
          const placeholder = blankMatch[0]
          const before = text.split(placeholder)[0].trim()
          const label = before || '未命名字段'
          const fieldType = inferFieldType(label, '')
          fieldCandidates.push({
            id: `p:${blockIndex}:blank`,
            label,
            kind: 'blank',
            path,
            context: text.slice(0, 80),
            currentValue: '',
            fieldType,
            groupKey: normalizeLabelKey(label),
          })
        }
      }
      continue
    }

    if (el.type === 'table' && el.data) {
      tagCounters.table += 1
      const tableIndex = tagCounters.table
      pushIndex(`table[${tableIndex}] (${el.rows || el.data.length}行×${el.cols || (el.data[0]?.length || 0)}列)`)
      const rows = el.data || []
      for (let r = 0; r < rows.length; r += 1) {
        for (let c = 0; c < rows[r].length - 1; c += 1) {
          if (fieldCandidates.length >= options.maxFieldCandidates) break
          const labelText = normalizeInlineText(rows[r][c] || '')
          if (!labelText) continue
          const valueText = normalizeInlineText(rows[r][c + 1] || '')
          if (!valueText || FIELD_PLACEHOLDER_REGEX.test(valueText)) {
            const fieldType = inferFieldType(labelText, valueText)
            const path = `table[${tableIndex}]/r[${r + 1}]/c[${c + 2}]`
            fieldCandidates.push({
              id: `t:${tableIndex - 1}:r:${r}:c:${c + 1}`,
              label: labelText,
              kind: 'table',
              path,
              context: `表格${tableIndex} 行${r + 1}`,
              currentValue: valueText || '',
              fieldType,
              groupKey: normalizeLabelKey(labelText),
            })
          }
        }
        if (fieldCandidates.length >= options.maxFieldCandidates) break
      }
      continue
    }
  }

  const grouped = new Map<string, number>()
  for (const field of fieldCandidates) {
    const groupKey = field.groupKey || normalizeLabelKey(field.label)
    grouped.set(groupKey, (grouped.get(groupKey) || 0) + 1)
  }

  for (const field of fieldCandidates.slice(0, options.maxFieldCandidates)) {
    const groupKey = field.groupKey || normalizeLabelKey(field.label)
    const groupCount = grouped.get(groupKey) || 1
    const currentValue = field.currentValue ? ` | 当前值: ${field.currentValue.slice(0, options.maxCellText)}` : ''
    const groupInfo = groupCount > 1 ? ` | 重复: ${groupCount}` : ''
    fieldLines.push(
      `${field.id} | ${field.label} | ${field.path}${currentValue} | 类型: ${field.fieldType || 'text'}${groupInfo}`
    )
  }

  return { indexLines, fieldLines }
}

/**
 * 从 HTML 中提取结构化大纲（标题 + 段落摘要）
 */
function extractOutline(
  html: string,
  maxParagraphs: number,
  maxParagraphLength: number
): { outline: string; stats: { headings: number; paragraphs: number; tables: number; images: number } } {
  const stats = { headings: 0, paragraphs: 0, tables: 0, images: 0 }
  const lines: string[] = []
  
  // 创建临时 DOM 解析
  const parser = new DOMParser()
  const doc = parser.parseFromString(`<div>${html}</div>`, 'text/html')
  const container = doc.body.firstElementChild
  
  if (!container) {
    return { outline: htmlToPlainText(html).slice(0, maxParagraphs * maxParagraphLength), stats }
  }
  
  let paragraphCount = 0
  
  // 遍历所有元素
  const walk = (node: Element, depth: number = 0) => {
    if (paragraphCount >= maxParagraphs) return
    
    const tagName = node.tagName.toLowerCase()
    
    // 处理标题
    if (/^h[1-6]$/.test(tagName)) {
      const level = parseInt(tagName[1])
      const text = node.textContent?.trim() || ''
      if (text) {
        const prefix = '#'.repeat(level)
        lines.push(`${prefix} ${text}`)
        stats.headings++
        paragraphCount++
      }
      return
    }
    
    // 处理表格
    if (tagName === 'table') {
      stats.tables++
      // 提取表格摘要（行数、列数、首行内容）
      const rows = node.querySelectorAll('tr')
      const firstRowCells = rows[0]?.querySelectorAll('td, th') || []
      const colCount = firstRowCells.length
      const rowCount = rows.length
      const headerText = Array.from(firstRowCells).map(c => c.textContent?.trim().slice(0, 20) || '').join(' | ')
      lines.push(`[表格: ${rowCount}行×${colCount}列${headerText ? ` - 首行: ${headerText}` : ''}]`)
      paragraphCount++
      return
    }
    
    // 处理图片
    if (tagName === 'img') {
      stats.images++
      const alt = node.getAttribute('alt') || ''
      lines.push(`[图片${alt ? `: ${alt}` : ''}]`)
      return
    }
    
    // 处理段落
    if (tagName === 'p' || tagName === 'div') {
      const text = node.textContent?.trim() || ''
      if (text && text.length > 0) {
        // 提取段落样式信息
        const style = node.getAttribute('style') || ''
        let styleHint = ''
        
        // 检测对齐方式
        if (style.includes('text-align: center')) styleHint += '[居中]'
        else if (style.includes('text-align: right')) styleHint += '[右对齐]'
        else if (style.includes('text-align: justify')) styleHint += '[两端对齐]'
        
        // 检测缩进
        if (style.includes('text-indent') || style.includes('margin-left')) styleHint += '[缩进]'
        
        // 检测字体大小
        const fontSizeMatch = style.match(/font-size:\s*(\d+(?:\.\d+)?(?:pt|px)?)/)
        if (fontSizeMatch) styleHint += `[${fontSizeMatch[1]}]`
        
        // 检测加粗
        if (node.querySelector('strong, b') || style.includes('font-weight: bold')) styleHint += '[粗体]'
        
        // 截断过长段落
        const truncatedText = text.length > maxParagraphLength 
          ? text.slice(0, maxParagraphLength) + '...' 
          : text
        
        lines.push(styleHint ? `${styleHint} ${truncatedText}` : truncatedText)
        stats.paragraphs++
        paragraphCount++
      }
      return
    }
    
    // 处理列表
    if (tagName === 'ul' || tagName === 'ol') {
      const items = node.querySelectorAll(':scope > li')
      const listType = tagName === 'ol' ? '有序列表' : '无序列表'
      lines.push(`[${listType}: ${items.length}项]`)
      items.forEach((item, idx) => {
        if (paragraphCount >= maxParagraphs) return
        const text = item.textContent?.trim().slice(0, maxParagraphLength) || ''
        if (text) {
          const prefix = tagName === 'ol' ? `${idx + 1}.` : '•'
          lines.push(`  ${prefix} ${text}`)
          paragraphCount++
        }
      })
      return
    }
    
    // 递归处理子元素
    for (const child of Array.from(node.children)) {
      if (paragraphCount >= maxParagraphs) break
      walk(child, depth + 1)
    }
  }
  
  walk(container)
  
  return { outline: lines.join('\n'), stats }
}

/**
 * 格式化页面设置为可读文本
 */
function formatPageSettings(settings: PageSettings): string {
  const lines: string[] = []
  
  // 纸张尺寸
  if (settings.width && settings.height) {
    // width/height 是 pt，1pt = 1/72 inch
    const toCm = (pt: number) => ((pt / 72) * 2.54).toFixed(1)
    const widthCm = toCm(settings.width)
    const heightCm = toCm(settings.height)
    // 判断纸张类型
    let paperType = '自定义'
    // A4 ≈ 595×842 pt
    if (Math.abs(settings.width - 595) < 5 && Math.abs(settings.height - 842) < 5) {
      paperType = 'A4'
    } else if (Math.abs(settings.width - 612) < 5 && Math.abs(settings.height - 792) < 5) {
      paperType = 'Letter'
    }
    lines.push(`纸张: ${paperType} (${widthCm}cm × ${heightCm}cm)`)
  }
  
  // 边距
  if (settings.marginTop || settings.marginBottom || settings.marginLeft || settings.marginRight) {
    const toMm = (pt?: number) => (typeof pt === 'number' ? ((pt / 72) * 25.4).toFixed(1) : '?')
    lines.push(
      `边距: 上${toMm(settings.marginTop)}mm 下${toMm(settings.marginBottom)}mm 左${toMm(settings.marginLeft)}mm 右${toMm(settings.marginRight)}mm`
    )
  }
  
  // 页眉页脚距离
  // 这里使用 headerHeight/footerHeight 近似
  if (settings.headerHeight || settings.footerHeight) {
    const toMm = (pt?: number) => (typeof pt === 'number' ? ((pt / 72) * 25.4).toFixed(1) : '?')
    lines.push(`页眉高: ${toMm(settings.headerHeight)}mm, 页脚高: ${toMm(settings.footerHeight)}mm`)
  }
  
  return lines.join('\n')
}

/**
 * 生成 DOCX 的 Agent 可读上下文（ArrayBuffer 入口，推荐）
 */
export async function generateDocxAgentContextFromArrayBuffer(
  fileName: string,
  arrayBuffer: ArrayBuffer,
  options: DocxAgentContextOptions = {}
): Promise<string> {
  const {
    maxLength = 12000,
    includeFullHtml = false,
    maxParagraphs = 50,
    maxParagraphLength = 200,
    timeoutMs = 15000,
  } = options
  
  // 检查缓存
  const cacheKey = `${fileName}:${arrayBuffer.byteLength}`
  const cached = contextCache.get(cacheKey)
  if (cached && Date.now() - cached.timestamp < CACHE_TTL) {
    console.log('[DocxAgentContext] 命中缓存:', fileName)
    return cached.result
  }
  
  console.log('[DocxAgentContext] 开始解析:', fileName)
  
  try {
    // 解析 DOCX（在 Worker 中，避免阻塞 UI；同时禁用图片 base64 内联）
    const parsed = await runWorker(arrayBuffer, timeoutMs)
    
    // 构建输出
    const sections: string[] = []
    
    // 1. 文件信息
    sections.push(`【Word 文档】${fileName}`)
    sections.push('')
    
    // 2. 页面设置
    if (parsed.pageSettings) {
      sections.push('【页面设置】')
      sections.push(formatPageSettings(parsed.pageSettings))
      sections.push('')
    }
    
    // 3. 页眉/页脚（轻量文本）
    if (parsed.headerText) {
      sections.push('【页眉】')
      sections.push(parsed.headerText.slice(0, 200))
      sections.push('')
    }
    if (parsed.footerText) {
      sections.push('【页脚】')
      sections.push(parsed.footerText.slice(0, 200))
      sections.push('')
    }
    
    // 5. 正文内容
    if (parsed.html) {
      const { outline, stats } = extractOutline(parsed.html, maxParagraphs, maxParagraphLength)
      const elements = docxHtmlToElements(parsed.html || '')
      const { indexLines, fieldLines } = buildStructureIndexFromElements(elements, {
        maxIndexLines: Math.min(maxParagraphs * 2, 120),
        maxFieldCandidates: 60,
        maxCellText: 60,
      })
      
      sections.push('【文档统计】')
      sections.push(`标题数: ${stats.headings}, 段落数: ${stats.paragraphs}, 表格数: ${stats.tables}, 图片数: ${stats.images}`)
      sections.push('')
      
      sections.push('【正文内容】')
      sections.push(outline)

      if (indexLines.length) {
        sections.push('')
        sections.push('【结构索引】')
        sections.push(indexLines.join('\n'))
      }

      if (fieldLines.length) {
        sections.push('')
        sections.push('【字段候选】')
        sections.push(fieldLines.join('\n'))
      }
      
      // 如果需要完整 HTML
      if (includeFullHtml) {
        sections.push('')
        sections.push('【原始 HTML（部分）】')
        sections.push(parsed.html.slice(0, 3000))
        if (parsed.html.length > 3000) {
          sections.push('... (已截断)')
        }
      }
    } else if (parsed.rawText) {
      sections.push('【正文文本（无格式）】')
      sections.push(parsed.rawText.slice(0, Math.min(maxLength, 6000)))
    }
    
    // 组合结果
    let result = sections.join('\n')
    
    // 长度保护
    if (result.length > maxLength) {
      result = result.slice(0, maxLength - 50) + '\n\n... (内容已截断，如需更多请指定具体章节)'
    }
    
    // 写入缓存
    contextCache.set(cacheKey, { result, timestamp: Date.now() })
    
    console.log('[DocxAgentContext] 解析完成:', fileName, '长度:', result.length)
    return result
    
  } catch (error) {
    console.error('[DocxAgentContext] 解析失败:', error)
    return `【Word 文档】${fileName}\n\n⚠️ 解析失败: ${(error as Error).message}\n\n请尝试重新打开文件或转存为新的 .docx 格式。`
  }
}

/**
 * 生成 DOCX 的 Agent 可读上下文（文件路径入口，Electron 推荐）
 */
export async function generateDocxAgentContextFromFilePath(
  fileName: string,
  filePath: string,
  options: DocxAgentContextOptions = {}
): Promise<string> {
  const arrayBuffer = await fetchArrayBufferFromLocalFileServer(filePath)
  return generateDocxAgentContextFromArrayBuffer(fileName, arrayBuffer, options)
}

/**
 * 生成 DOCX 的 Agent 可读上下文（base64 入口，兼容）
 */
export async function generateDocxAgentContext(
  fileName: string,
  base64Data: string,
  options: DocxAgentContextOptions = {}
): Promise<string> {
  // 对超大 docx（尤其含大量图片）不推荐走 base64（会触发 atob 大字符串）
  if (base64Data.length > 15_000_000) {
    return `【Word 文档】${fileName}\n\n⚠️ 文档过大（base64 长度 ${base64Data.length}），建议在桌面端使用文件路径解析模式。`
  }
  const arrayBuffer = base64ToArrayBuffer(base64Data)
  return generateDocxAgentContextFromArrayBuffer(fileName, arrayBuffer, options)
}

/**
 * 清除缓存（用于文件更新后强制刷新）
 */
export function clearDocxAgentContextCache(fileName?: string) {
  if (fileName) {
    // 清除特定文件的缓存
    for (const key of contextCache.keys()) {
      if (key.startsWith(fileName + ':')) {
        contextCache.delete(key)
      }
    }
  } else {
    // 清除所有缓存
    contextCache.clear()
  }
}

/**
 * 判断文件是否为 docx
 */
export function isDocxFile(fileName: string): boolean {
  return fileName.toLowerCase().endsWith('.docx')
}

