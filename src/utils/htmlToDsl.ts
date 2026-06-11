/**
 * HTML → DSL 转换器
 *
 * 将 Tiptap/docxParser 输出的 HTML 转换为 DocDsl 结构，
 * 用于发给 AI 模型和 DSL 工具操作。
 *
 * 支持的 HTML 模式：
 * - <h1>~<h6>: heading
 * - <p>: paragraph（含 docx-list-marker 检测 → list）
 * - <table>: table（含 colgroup、colspan、rowspan）
 * - <img>: image（base64 → 占位符）
 * - <hr class="page-break">: pageBreak
 * - <hr>: horizontalRule
 * - <blockquote>: blockquote
 * - diff/track/comment 元数据保留
 */

import type {
  DocDsl, DslBlock, DslHeading, DslParagraph, DslList, DslListItem,
  DslTable, DslTableRow, DslTableCell, DslImage, DslPageBreak,
  DslHorizontalRule, DslBlockquote, DslSectionBreak,
  DslRun, DslInline, DslParagraphFormat, DslAlignment,
  DslBorder, DslBorderSide, DslRunMeta, DslBlockMeta, DslLength,
} from '../types/docDsl'

// ─── 选项 ───

export interface HtmlToDslOptions {
  /** 去除 diff 标记（diff-old/diff-new） */
  stripDiffMarkers?: boolean
  /** 去除修订标记 */
  stripTrackChanges?: boolean
  /** 去除批注标记 */
  stripComments?: boolean
  /** 最大块数 */
  maxBlocks?: number
}

// ─── 主函数 ───

export function htmlToDsl(html: string, options?: HtmlToDslOptions): DocDsl {
  if (!html || !html.trim()) return { blocks: [] }

  const doc = new DOMParser().parseFromString(html, 'text/html')
  const body = doc.body
  const rawBlocks = convertChildren(body, options || {})

  // 合并连续的列表项为 DslList 块
  const blocks = mergeConsecutiveListItems(rawBlocks)

  // 截断
  const maxBlocks = options?.maxBlocks
  const finalBlocks = maxBlocks && maxBlocks > 0 ? blocks.slice(0, maxBlocks) : blocks

  return { blocks: finalBlocks }
}

// ─── 子节点遍历 ───

/** 临时标记：列表项（尚未合并为 DslList） */
interface PendingListItem {
  type: '__listItem'
  listType: 'bullet' | 'number'
  level: number
  content: string | DslInline[]
  _meta?: DslBlockMeta
}

type RawBlock = DslBlock | PendingListItem

function convertChildren(parent: Node, opts: HtmlToDslOptions): RawBlock[] {
  const blocks: RawBlock[] = []
  for (let i = 0; i < parent.childNodes.length; i++) {
    const node = parent.childNodes[i]
    if (node.nodeType === Node.ELEMENT_NODE) {
      const el = node as HTMLElement
      const converted = convertElement(el, opts)
      if (converted) {
        if (Array.isArray(converted)) blocks.push(...converted)
        else blocks.push(converted)
      }
    }
    // 忽略纯文本节点（块级上下文中的空白）
  }
  return blocks
}

function convertElement(el: HTMLElement, opts: HtmlToDslOptions): RawBlock | RawBlock[] | null {
  const tag = el.tagName.toLowerCase()

  // 跳过脚注/尾注区域
  if (tag === 'div' && (el.classList.contains('footnotes') || el.classList.contains('endnotes'))) {
    return null
  }

  // 跳过 script/style
  if (tag === 'script' || tag === 'style') return null

  // 标题
  if (/^h[1-6]$/.test(tag)) {
    return convertHeading(el, opts)
  }

  // 段落（可能是列表项）
  if (tag === 'p') {
    return convertParagraph(el, opts)
  }

  // 表格
  if (tag === 'table') {
    return convertTable(el, opts)
  }

  // 图片（独立块级）
  if (tag === 'img') {
    return convertImage(el)
  }

  // figure（可能包含 img + figcaption）
  if (tag === 'figure') {
    const img = el.querySelector('img')
    if (img) {
      const block = convertImage(img)
      const caption = el.querySelector('figcaption')
      if (caption && block) block.caption = caption.textContent || undefined
      return block
    }
    return null
  }

  // 水平线 / 分页符
  if (tag === 'hr') {
    if (el.classList.contains('page-break')) {
      return { type: 'pageBreak' } as DslPageBreak
    }
    return { type: 'horizontalRule' } as DslHorizontalRule
  }

  // 兼容应用内部插入的分页符块
  if (tag === 'div' && el.classList.contains('page-break')) {
    return { type: 'pageBreak' } as DslPageBreak
  }

  // 引用块
  if (tag === 'blockquote') {
    const children = convertChildren(el, opts)
    const merged = mergeConsecutiveListItems(children)
    return { type: 'blockquote', content: merged } as DslBlockquote
  }

  // ul/ol（Tiptap 可能输出原生列表）
  if (tag === 'ul' || tag === 'ol') {
    return convertNativeList(el, tag === 'ol' ? 'number' : 'bullet', opts)
  }

  // div 等容器：递归子节点
  if (tag === 'div' || tag === 'section' || tag === 'article' || tag === 'main' || tag === 'header' || tag === 'footer') {
    return convertChildren(el, opts) as RawBlock[]
  }

  // 其他未知元素：尝试当段落处理
  if (el.textContent?.trim()) {
    return {
      type: 'paragraph',
      content: extractInlines(el, opts),
    } as DslParagraph
  }

  return null
}

// ─── 标题转换 ───

function convertHeading(el: HTMLElement, opts: HtmlToDslOptions): DslHeading {
  const level = parseInt(el.tagName[1]) as 1 | 2 | 3 | 4 | 5 | 6
  const content = extractInlines(el, opts)
  const format = extractParagraphFormat(el)
  const _meta = extractBlockMeta(el, opts)

  const heading: DslHeading = { type: 'heading', level, content: simplifyInlines(content) }
  if (format) heading.format = format
  if (_meta) heading._meta = _meta
  return heading
}

// ─── 段落转换（含列表项检测） ───

function convertParagraph(el: HTMLElement, opts: HtmlToDslOptions): RawBlock | null {
  // 检测列表项：<span class="docx-list-marker">
  const listMarker = el.querySelector('.docx-list-marker')
  if (listMarker) {
    return convertListItemParagraph(el, listMarker as HTMLElement, opts)
  }

  // TOC 段落跳过
  if (el.classList.contains('docx-toc')) return null

  const content = extractInlines(el, opts)
  // 空段落保留（文档中的空行）
  if (content.length === 0) {
    return { type: 'paragraph', content: '' } as DslParagraph
  }

  const format = extractParagraphFormat(el)
  const _meta = extractBlockMeta(el, opts)

  const para: DslParagraph = { type: 'paragraph', content: simplifyInlines(content) }
  if (format) para.format = format
  if (_meta) para._meta = _meta
  return para
}

function convertListItemParagraph(
  el: HTMLElement, marker: HTMLElement, opts: HtmlToDslOptions
): PendingListItem {
  const markerText = marker.textContent || ''
  const isBullet = /^[•●○■▪▸–—-]/.test(markerText.trim())
  const listType = isBullet ? 'bullet' as const : 'number' as const

  // 推算缩进级别
  const paddingLeft = el.style.paddingLeft || ''
  const level = paddingLeft ? Math.max(0, Math.round(parseFloat(paddingLeft) / 1.5) - 1) : 0

  // 提取内容（排除 marker 本身）
  const content: DslInline[] = []
  for (let i = 0; i < el.childNodes.length; i++) {
    const child = el.childNodes[i]
    if (child === marker) continue
    if (child.nodeType === Node.ELEMENT_NODE) {
      const childEl = child as HTMLElement
      if (childEl.classList.contains('docx-list-marker')) continue
      content.push(...extractInlinesFromNode(childEl, opts))
    } else if (child.nodeType === Node.TEXT_NODE) {
      const text = child.textContent || ''
      if (text) content.push(text)
    }
  }

  return {
    type: '__listItem',
    listType,
    level,
    content: simplifyInlines(content),
    _meta: extractBlockMeta(el, opts) || undefined,
  }
}

// ─── 列表项合并 ───

function mergeConsecutiveListItems(rawBlocks: RawBlock[]): DslBlock[] {
  const result: DslBlock[] = []
  let pendingList: { listType: 'bullet' | 'number'; items: DslListItem[] } | null = null

  for (const block of rawBlocks) {
    if ((block as PendingListItem).type === '__listItem') {
      const item = block as PendingListItem
      if (pendingList && pendingList.listType === item.listType) {
        pendingList.items.push({ content: item.content })
      } else {
        if (pendingList) {
          result.push({ type: 'list', listType: pendingList.listType, items: pendingList.items } as DslList)
        }
        pendingList = { listType: item.listType, items: [{ content: item.content }] }
      }
    } else {
      if (pendingList) {
        result.push({ type: 'list', listType: pendingList.listType, items: pendingList.items } as DslList)
        pendingList = null
      }
      result.push(block as DslBlock)
    }
  }

  if (pendingList) {
    result.push({ type: 'list', listType: pendingList.listType, items: pendingList.items } as DslList)
  }

  return result
}

// ─── 原生列表转换（ul/ol） ───

function convertNativeList(el: HTMLElement, listType: 'bullet' | 'number', opts: HtmlToDslOptions): DslList {
  const items: DslListItem[] = []
  for (let i = 0; i < el.children.length; i++) {
    const li = el.children[i]
    if (li.tagName.toLowerCase() === 'li') {
      const content = extractInlines(li as HTMLElement, opts)
      const children: DslListItem[] = []
      // 检查嵌套列表
      const nestedList = li.querySelector('ul, ol')
      if (nestedList) {
        const nestedTag = nestedList.tagName.toLowerCase()
        const nested = convertNativeList(nestedList as HTMLElement, nestedTag === 'ol' ? 'number' : 'bullet', opts)
        children.push(...nested.items)
      }
      items.push({
        content: simplifyInlines(content),
        ...(children.length > 0 ? { children } : {}),
      })
    }
  }
  return { type: 'list', listType, items }
}

// ─── 表格转换 ───

function convertTable(el: HTMLElement, opts: HtmlToDslOptions): DslTable {
  const table: DslTable = { type: 'table', rows: [] }

  // 提取列宽
  const colgroup = el.querySelector('colgroup')
  if (colgroup) {
    const cols = colgroup.querySelectorAll('col')
    const widths: DslLength[] = []
    cols.forEach(col => {
      const w = (col as HTMLElement).style.width
      if (w) widths.push(w)
    })
    if (widths.length > 0) table.columnWidths = widths
  }

  // 表格宽度
  const tableWidth = el.style.width
  if (tableWidth && tableWidth !== 'auto') table.width = tableWidth

  // 遍历行
  const rows = el.querySelectorAll('tr')
  rows.forEach(tr => {
    const row = convertTableRow(tr as HTMLElement, opts)
    table.rows.push(row)
  })

  return table
}

function convertTableRow(tr: HTMLElement, opts: HtmlToDslOptions): DslTableRow {
  const row: DslTableRow = { cells: [] }

  // 行高
  const height = tr.style.height
  if (height) row.height = height

  // 是否表头行
  const isInThead = tr.parentElement?.tagName.toLowerCase() === 'thead'

  for (let i = 0; i < tr.children.length; i++) {
    const td = tr.children[i] as HTMLElement
    if (td.tagName.toLowerCase() === 'td' || td.tagName.toLowerCase() === 'th') {
      const cell = convertTableCell(td, opts)
      row.cells.push(cell)
      if (td.tagName.toLowerCase() === 'th' || isInThead) row.isHeader = true
    }
  }

  return row
}

function convertTableCell(td: HTMLElement, opts: HtmlToDslOptions): DslTableCell {
  const cell: DslTableCell = { content: '' }

  // colspan / rowspan
  const colspan = td.getAttribute('colspan')
  if (colspan && parseInt(colspan) > 1) cell.colSpan = parseInt(colspan)
  const rowspan = td.getAttribute('rowspan')
  if (rowspan && parseInt(rowspan) > 1) cell.rowSpan = parseInt(rowspan)

  // 对齐
  const textAlign = td.style.textAlign as DslAlignment
  if (textAlign && textAlign !== 'left') cell.align = textAlign
  const verticalAlign = td.style.verticalAlign
  if (verticalAlign && verticalAlign !== 'top') {
    cell.valign = verticalAlign as 'top' | 'middle' | 'bottom'
  }

  // 背景色
  const bg = td.style.backgroundColor
  if (bg) cell.backgroundColor = bg

  // 内容：检查是否有块级子元素
  const hasBlockChildren = td.querySelector('p, h1, h2, h3, h4, h5, h6, table, ul, ol')
  if (hasBlockChildren) {
    const blocks = convertChildren(td, opts)
    cell.content = mergeConsecutiveListItems(blocks)
  } else {
    const inlines = extractInlines(td, opts)
    cell.content = simplifyInlines(inlines)
  }

  return cell
}

// ─── 图片转换 ───

function convertImage(img: HTMLElement): DslImage | null {
  let src = img.getAttribute('src') || ''
  const rid = img.getAttribute('data-rid')
  const alt = img.getAttribute('alt') || ''
  const preserveInlineSrc =
    img.getAttribute('data-preserve-src') === '1' ||
    img.getAttribute('data-generated-from') === 'docx-chart' ||
    alt.startsWith('chart:')

  // base64 → 占位符
  if (src.startsWith('data:image/') && !preserveInlineSrc) {
    src = rid ? `[image:${rid}]` : '[image]'
  }

  if (!src) return null

  const block: DslImage = { type: 'image', src }

  if (alt && alt !== '文档图片') block.alt = alt

  const w = img.getAttribute('data-w') || img.style.width
  if (w) block.width = w.includes('px') || w.includes('%') ? w : `${w}px`

  const h = img.getAttribute('data-h') || img.style.height
  if (h) block.height = h.includes('px') || h.includes('%') ? h : `${h}px`

  const wrap = img.getAttribute('data-wrap')
  if (wrap) block.wrap = wrap as DslImage['wrap']

  return block
}

// ─── 内联格式提取 ───

function extractInlines(el: HTMLElement, opts: HtmlToDslOptions): DslInline[] {
  const result: DslInline[] = []
  for (let i = 0; i < el.childNodes.length; i++) {
    const node = el.childNodes[i]
    if (node.nodeType === Node.TEXT_NODE) {
      const text = node.textContent || ''
      if (text) result.push(text)
    } else if (node.nodeType === Node.ELEMENT_NODE) {
      const childEl = node as HTMLElement
      result.push(...extractInlinesFromNode(childEl, opts))
    }
  }
  return result
}

function extractInlinesFromNode(el: HTMLElement, opts: HtmlToDslOptions): DslInline[] {
  const tag = el.tagName.toLowerCase()

  // <br> → 换行
  if (tag === 'br') return ['\n']

  // 跳过列表标记
  if (el.classList.contains('docx-list-marker')) return []
  if (el.classList.contains('docx-tab')) return ['\t']

  // diff 标记
  if (el.classList.contains('diff-old') || el.classList.contains('diff-new')) {
    if (opts.stripDiffMarkers) {
      // 去除 diff 标记但保留内容（只保留 new 的内容）
      if (el.classList.contains('diff-old')) return []
      return extractInlines(el, opts)
    }
    const meta: DslRunMeta = {
      diffType: el.classList.contains('diff-old') ? 'old' : 'new',
      diffId: el.getAttribute('data-diff-id') || undefined,
    }
    const inlines = extractInlines(el, opts)
    return inlines.map(inline => {
      const run = inlineToRun(inline)
      run._meta = { ...run._meta, ...meta }
      return run
    })
  }

  // track changes
  if (el.classList.contains('docx-track')) {
    if (opts.stripTrackChanges) {
      const trackType = el.getAttribute('data-track-type')
      if (trackType === 'delete') return []
      return extractInlines(el, opts)
    }
    const meta: DslRunMeta = {
      trackType: (el.getAttribute('data-track-type') as 'insert' | 'delete') || undefined,
      trackId: el.getAttribute('data-track-id') || undefined,
      trackAuthor: el.getAttribute('data-track-author') || undefined,
      trackDate: el.getAttribute('data-track-date') || undefined,
    }
    const inlines = extractInlines(el, opts)
    return inlines.map(inline => {
      const run = inlineToRun(inline)
      run._meta = { ...run._meta, ...meta }
      return run
    })
  }

  // comments
  if (el.classList.contains('docx-comment')) {
    if (opts.stripComments) return extractInlines(el, opts)
    const ids = el.getAttribute('data-comment-ids')
    const meta: DslRunMeta = { commentIds: ids ? ids.split(',') : undefined }
    const inlines = extractInlines(el, opts)
    return inlines.map(inline => {
      const run = inlineToRun(inline)
      run._meta = { ...run._meta, ...meta }
      return run
    })
  }

  // 格式标签：strong, em, u, s, sup, sub
  if (tag === 'strong' || tag === 'b') {
    return extractInlines(el, opts).map(inline => {
      const run = inlineToRun(inline)
      run.bold = true
      return run
    })
  }
  if (tag === 'em' || tag === 'i') {
    return extractInlines(el, opts).map(inline => {
      const run = inlineToRun(inline)
      run.italic = true
      return run
    })
  }
  if (tag === 'u') {
    return extractInlines(el, opts).map(inline => {
      const run = inlineToRun(inline)
      run.underline = true
      return run
    })
  }
  if (tag === 's' || tag === 'del' || tag === 'strike') {
    return extractInlines(el, opts).map(inline => {
      const run = inlineToRun(inline)
      run.strikethrough = true
      return run
    })
  }
  if (tag === 'sup') {
    return extractInlines(el, opts).map(inline => {
      const run = inlineToRun(inline)
      run.superscript = true
      return run
    })
  }
  if (tag === 'sub') {
    return extractInlines(el, opts).map(inline => {
      const run = inlineToRun(inline)
      run.subscript = true
      return run
    })
  }

  // <span> 带样式
  if (tag === 'span') {
    const style = el.getAttribute('style') || ''
    const fontName = el.getAttribute('data-font-name')
    const inlines = extractInlines(el, opts)

    // 如果没有任何样式信息，直接返回子内容
    if (!style && !fontName) return inlines

    const runProps = extractRunPropsFromStyle(style, fontName)
    if (Object.keys(runProps).length === 0) return inlines

    return inlines.map(inline => {
      const run = inlineToRun(inline)
      Object.assign(run, runProps)
      return run
    })
  }

  // <img> 内联图片 → 占位文本
  if (tag === 'img') {
    const rid = el.getAttribute('data-rid')
    return [rid ? `[图片:${rid}]` : '[图片]']
  }

  // <a> 链接 → 保留文本
  if (tag === 'a') {
    return extractInlines(el, opts)
  }

  // 其他元素：递归提取
  return extractInlines(el, opts)
}

// ─── 样式提取辅助 ───

function extractRunPropsFromStyle(style: string, fontName?: string | null): Partial<DslRun> {
  const props: Partial<DslRun> = {}

  // 字体
  if (fontName) {
    props.fontFamily = fontName
  } else {
    const ff = style.match(/font-family:\s*"?([^",;]+)/)
    if (ff) props.fontFamily = ff[1].trim()
  }

  // 字号
  const fs = style.match(/font-size:\s*(\d+(?:\.\d+)?)\s*pt/)
  if (fs) props.fontSize = parseFloat(fs[1])

  // 颜色（排除 border-color 等前缀）
  const colorMatch = style.match(/(?:^|;\s*)color:\s*([^;]+)/)
  if (colorMatch) {
    const c = colorMatch[1].trim()
    if (c && c !== '#000000' && c !== '#000' && c !== 'rgb(0, 0, 0)') {
      props.color = c
    }
  }

  // 高亮/背景色
  const bgMatch = style.match(/background-color:\s*([^;]+)/)
  if (bgMatch) {
    const bg = bgMatch[1].trim()
    if (bg && bg !== 'transparent') props.highlight = bg
  }

  // 粗体（从 style）
  if (/font-weight:\s*(bold|[7-9]\d{2})/.test(style)) props.bold = true

  // 斜体（从 style）
  if (/font-style:\s*italic/.test(style)) props.italic = true

  // 下划线（从 style）
  if (/text-decoration[^:]*:\s*[^;]*underline/.test(style)) props.underline = true

  // 删除线（从 style）
  if (/text-decoration[^:]*:\s*[^;]*line-through/.test(style)) props.strikethrough = true

  // 字间距
  const ls = style.match(/letter-spacing:\s*(\d+(?:\.\d+)?)\s*pt/)
  if (ls) props.letterSpacing = parseFloat(ls[1])

  return props
}

function extractParagraphFormat(el: HTMLElement): DslParagraphFormat | undefined {
  const style = el.getAttribute('style') || ''
  if (!style) return undefined

  const format: DslParagraphFormat = {}

  // 对齐
  const align = style.match(/text-align:\s*(\w+)/)
  if (align && align[1] !== 'left' && align[1] !== 'start') {
    format.alignment = align[1] as DslAlignment
  }

  // 首行缩进
  const indent = style.match(/text-indent:\s*([^;]+)/)
  if (indent) format.firstLineIndent = indent[1].trim()

  // 左缩进
  const paddingLeft = style.match(/padding-left:\s*([^;]+)/)
  if (paddingLeft && paddingLeft[1].trim() !== '0em' && paddingLeft[1].trim() !== '0') {
    format.leftIndent = paddingLeft[1].trim()
  }

  // 段前
  const marginTop = style.match(/margin-top:\s*([^;]+)/)
  if (marginTop && marginTop[1].trim() !== '0' && marginTop[1].trim() !== '0px') {
    format.spaceBefore = marginTop[1].trim()
  }

  // 段后
  const marginBottom = style.match(/margin-bottom:\s*([^;]+)/)
  if (marginBottom && marginBottom[1].trim() !== '0' && marginBottom[1].trim() !== '0px') {
    format.spaceAfter = marginBottom[1].trim()
  }

  // 行距
  const lineHeight = style.match(/line-height:\s*([^;]+)/)
  if (lineHeight) {
    const lh = lineHeight[1].trim()
    if (lh !== '1.0' && lh !== 'normal') format.lineHeight = lh
  }

  // 背景色
  const bg = style.match(/background-color:\s*([^;]+)/)
  if (bg && bg[1].trim() !== 'transparent') format.backgroundColor = bg[1].trim()

  return Object.keys(format).length > 0 ? format : undefined
}

// ─── 块级元数据提取 ───

function extractBlockMeta(el: HTMLElement, opts: HtmlToDslOptions): DslBlockMeta | null {
  const diffRole = el.getAttribute('data-diff-role')
  const diffId = el.getAttribute('data-diff-id')

  if (opts.stripDiffMarkers) return null
  if (!diffRole && !diffId) return null

  return {
    diffRole: diffRole === 'new' ? 'new' : undefined,
    diffId: diffId || undefined,
  }
}

// ─── 工具函数 ───

/** 将 DslInline 转为 DslRun（如果是纯字符串则包装） */
function inlineToRun(inline: DslInline): DslRun {
  if (typeof inline === 'string') return { text: inline }
  return { ...inline }
}

/** 简化 inlines：合并相邻纯文本，单个纯文本 run 退化为 string */
function simplifyInlines(inlines: DslInline[]): string | DslInline[] {
  if (inlines.length === 0) return ''

  // 合并相邻纯文本
  const merged: DslInline[] = []
  for (const inline of inlines) {
    const last = merged[merged.length - 1]
    if (typeof inline === 'string' && typeof last === 'string') {
      merged[merged.length - 1] = last + inline
    } else if (typeof inline === 'object' && isPlainTextRun(inline) && typeof last === 'string') {
      merged[merged.length - 1] = last + inline.text
    } else if (typeof inline === 'string' && typeof last === 'object' && isPlainTextRun(last)) {
      merged[merged.length - 1] = last.text + inline
    } else {
      merged.push(inline)
    }
  }

  // 单个纯文本 → 退化为 string
  if (merged.length === 1 && typeof merged[0] === 'string') return merged[0]
  if (merged.length === 1 && typeof merged[0] === 'object' && isPlainTextRun(merged[0])) return merged[0].text

  return merged
}

/** 检查 DslRun 是否只有 text 属性（无格式） */
function isPlainTextRun(run: DslRun): boolean {
  const keys = Object.keys(run)
  return keys.length === 1 && keys[0] === 'text'
}
