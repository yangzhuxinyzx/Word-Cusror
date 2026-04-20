/**
 * DSL → AI 序列化器
 *
 * 将 DocDsl 序列化为带块索引的精简 DSL JSON，发给模型。
 * 模型读写同一种 DSL 格式，输入输出统一。
 *
 * 输出格式：每个块带 _i 索引，省略默认值，图片用占位符。
 */

import type {
  DocDsl, DslBlock, DslRun, DslInline, DslParagraphFormat,
  DslTableCell, DslTableRow,
} from '../types/docDsl'

export interface SerializeDslOptions {
  /** 最大输出字符数 */
  maxLength?: number
  /** 省略默认值 */
  omitDefaults?: boolean
}

/**
 * 将 DSL 序列化为 AI 可读的精简 JSON 字符串
 */
export function serializeDslForAI(dsl: DocDsl, options?: SerializeDslOptions): string {
  const maxLength = options?.maxLength ?? 120_000
  const omitDefaults = options?.omitDefaults ?? true

  const serializedBlocks: unknown[] = []

  for (let i = 0; i < dsl.blocks.length; i++) {
    const block = dsl.blocks[i]
    const serialized = serializeBlock(block, i, omitDefaults)
    if (serialized) serializedBlocks.push(serialized)
  }

  let result = JSON.stringify({ blocks: serializedBlocks })

  // 截断处理
  if (result.length > maxLength) {
    // 逐块减少直到满足长度限制
    while (serializedBlocks.length > 0 && JSON.stringify({ blocks: serializedBlocks }).length > maxLength) {
      serializedBlocks.pop()
    }
    result = JSON.stringify({ blocks: serializedBlocks })
    if (result.length > maxLength) {
      result = result.slice(0, maxLength - 50) + '...(truncated)'
    }
  }

  return result
}

// ─── 块序列化 ───

function serializeBlock(block: DslBlock, index: number, omitDefaults: boolean): unknown {
  const base: Record<string, unknown> = { _i: index, type: block.type }

  switch (block.type) {
    case 'heading': {
      base.level = block.level
      base.content = serializeContent(block.content)
      if (block.format) {
        const fmt = serializeFormat(block.format, omitDefaults)
        if (fmt) base.format = fmt
      }
      return base
    }

    case 'paragraph': {
      base.content = serializeContent(block.content)
      if (block.format) {
        const fmt = serializeFormat(block.format, omitDefaults)
        if (fmt) base.format = fmt
      }
      return base
    }

    case 'list': {
      base.listType = block.listType
      base.items = block.items.map(item => ({
        content: serializeContent(item.content),
        ...(item.children?.length ? { children: item.children.map(c => ({ content: serializeContent(c.content) })) } : {}),
      }))
      return base
    }

    case 'table': {
      base.rows = block.rows.map(row => serializeTableRow(row))
      if (block.columnWidths?.length) base.columnWidths = block.columnWidths
      return base
    }

    case 'image': {
      // 确保 base64 被替换为占位符
      let src = block.src
      if (src.startsWith('data:image/')) {
        src = '[image]'
      }
      base.src = src
      if (block.alt) base.alt = block.alt
      if (block.width) base.width = block.width
      if (block.height) base.height = block.height
      if (block.caption) base.caption = block.caption
      return base
    }

    case 'pageBreak':
      return base

    case 'sectionBreak':
      if (block.breakType) base.breakType = block.breakType
      return base

    case 'horizontalRule':
      return base

    case 'blockquote': {
      base.content = block.content.map((b, i) => serializeBlock(b, i, omitDefaults)).filter(Boolean)
      return base
    }

    default:
      return base
  }
}

function serializeTableRow(row: DslTableRow): unknown {
  const result: Record<string, unknown> = {
    cells: row.cells.map(cell => serializeTableCell(cell)),
  }
  if (row.isHeader) result.isHeader = true
  if (row.height) result.height = row.height
  return result
}

function serializeTableCell(cell: DslTableCell): unknown {
  const result: Record<string, unknown> = {}

  // 内容
  if (Array.isArray(cell.content) && cell.content.length > 0) {
    // 检查是否是 DslBlock[]
    const first = cell.content[0]
    if (typeof first === 'object' && 'type' in first) {
      // DslBlock[]
      result.content = (cell.content as DslBlock[]).map((b, i) => serializeBlock(b, i, true)).filter(Boolean)
    } else {
      // DslInline[]
      result.content = serializeContent(cell.content as string | DslInline[])
    }
  } else {
    result.content = serializeContent(cell.content as string | DslInline[])
  }

  if (cell.colSpan && cell.colSpan > 1) result.colSpan = cell.colSpan
  if (cell.rowSpan && cell.rowSpan > 1) result.rowSpan = cell.rowSpan
  if (cell.align && cell.align !== 'left') result.align = cell.align
  if (cell.backgroundColor) result.backgroundColor = cell.backgroundColor

  return result
}

// ─── 内容序列化 ───

function serializeContent(content: string | DslInline[]): unknown {
  if (typeof content === 'string') return content

  // 如果所有 inline 都是纯文本，合并为单个字符串
  const allPlainText = content.every(item => typeof item === 'string')
  if (allPlainText) return content.join('')

  // 否则序列化每个 run
  return content.map(item => {
    if (typeof item === 'string') return item
    return serializeRun(item)
  })
}

function serializeRun(run: DslRun): unknown {
  const result: Record<string, unknown> = { text: run.text }
  if (run.bold) result.bold = true
  if (run.italic) result.italic = true
  if (run.underline) result.underline = true
  if (run.strikethrough) result.strikethrough = true
  if (run.superscript) result.superscript = true
  if (run.subscript) result.subscript = true
  if (run.fontFamily) result.fontFamily = run.fontFamily
  if (run.fontSize) result.fontSize = run.fontSize
  if (run.color) result.color = run.color
  if (run.highlight) result.highlight = run.highlight
  if (run.letterSpacing) result.letterSpacing = run.letterSpacing
  return result
}

// ─── 格式序列化 ───

function serializeFormat(format: DslParagraphFormat, omitDefaults: boolean): unknown | null {
  const result: Record<string, unknown> = {}

  if (format.alignment && (!omitDefaults || format.alignment !== 'left')) {
    result.alignment = format.alignment
  }
  if (format.firstLineIndent) result.firstLineIndent = format.firstLineIndent
  if (format.leftIndent) result.leftIndent = format.leftIndent
  if (format.rightIndent) result.rightIndent = format.rightIndent
  if (format.spaceBefore) result.spaceBefore = format.spaceBefore
  if (format.spaceAfter) result.spaceAfter = format.spaceAfter
  if (format.lineHeight) result.lineHeight = format.lineHeight
  if (format.backgroundColor) result.backgroundColor = format.backgroundColor

  return Object.keys(result).length > 0 ? result : null
}
