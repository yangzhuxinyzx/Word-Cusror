/**
 * DocDsl - 校验与工具函数
 */

import type {
  DocDsl,
  DslBlock,
  DslHeading,
  DslParagraph,
  DslList,
  DslTable,
  DslImage,
  DslRun,
  DslInline,
  DslTableCell,
  DslTableRow,
  DslListItem,
  DslValidationResult,
  DslValidationError,
  DslLength,
  DslColor,
  DslAlignment,
  DslBorder,
  DslBorderSide,
  DslParagraphFormat,
} from '../types/docDsl'

// ============== 校验函数 ==============

/**
 * 校验 DocDsl 文档
 */
export function validateDocDsl(dsl: unknown): DslValidationResult {
  const errors: DslValidationError[] = []

  if (!dsl || typeof dsl !== 'object') {
    errors.push({ path: '', message: 'DSL must be an object', code: 'INVALID_ROOT' })
    return { valid: false, errors }
  }

  const doc = dsl as Record<string, unknown>

  // 检查 blocks 字段
  if (!Array.isArray(doc.blocks)) {
    errors.push({ path: 'blocks', message: 'blocks must be an array', code: 'MISSING_BLOCKS' })
    return { valid: false, errors }
  }

  // 校验每个块
  doc.blocks.forEach((block, index) => {
    const blockErrors = validateBlock(block, `blocks[${index}]`)
    errors.push(...blockErrors)
  })

  // 校验页面设置（可选）
  if (doc.pageSetup !== undefined) {
    const pageErrors = validatePageSetup(doc.pageSetup, 'pageSetup')
    errors.push(...pageErrors)
  }

  // 校验页眉页脚（可选）
  if (doc.headerFooter !== undefined) {
    const hfErrors = validateHeaderFooter(doc.headerFooter, 'headerFooter')
    errors.push(...hfErrors)
  }

  return { valid: errors.length === 0, errors }
}

/**
 * 校验单个块
 */
function validateBlock(block: unknown, path: string): DslValidationError[] {
  const errors: DslValidationError[] = []

  if (!block || typeof block !== 'object') {
    errors.push({ path, message: 'Block must be an object', code: 'INVALID_BLOCK' })
    return errors
  }

  const b = block as Record<string, unknown>
  const type = b.type

  if (typeof type !== 'string') {
    errors.push({ path: `${path}.type`, message: 'Block type is required', code: 'MISSING_TYPE' })
    return errors
  }

  switch (type) {
    case 'heading':
      errors.push(...validateHeading(b, path))
      break
    case 'paragraph':
      errors.push(...validateParagraph(b, path))
      break
    case 'list':
      errors.push(...validateList(b, path))
      break
    case 'table':
      errors.push(...validateTable(b, path))
      break
    case 'image':
      errors.push(...validateImage(b, path))
      break
    case 'pageBreak':
    case 'sectionBreak':
    case 'horizontalRule':
    case 'blockquote':
      // 这些类型只需要 type 字段
      break
    default:
      errors.push({ path: `${path}.type`, message: `Unknown block type: ${type}`, code: 'UNKNOWN_TYPE' })
  }

  return errors
}

function validateHeading(block: Record<string, unknown>, path: string): DslValidationError[] {
  const errors: DslValidationError[] = []

  const level = block.level
  if (typeof level !== 'number' || level < 1 || level > 6) {
    errors.push({ path: `${path}.level`, message: 'Heading level must be 1-6', code: 'INVALID_LEVEL' })
  }

  if (block.content === undefined || block.content === null) {
    errors.push({ path: `${path}.content`, message: 'Heading content is required', code: 'MISSING_CONTENT' })
  } else {
    errors.push(...validateContent(block.content, `${path}.content`))
  }

  return errors
}

function validateParagraph(block: Record<string, unknown>, path: string): DslValidationError[] {
  const errors: DslValidationError[] = []

  if (block.content === undefined || block.content === null) {
    errors.push({ path: `${path}.content`, message: 'Paragraph content is required', code: 'MISSING_CONTENT' })
  } else {
    errors.push(...validateContent(block.content, `${path}.content`))
  }

  if (block.format !== undefined) {
    errors.push(...validateParagraphFormat(block.format, `${path}.format`))
  }

  return errors
}

function validateList(block: Record<string, unknown>, path: string): DslValidationError[] {
  const errors: DslValidationError[] = []

  const listType = block.listType
  if (!['bullet', 'number', 'letter', 'roman'].includes(listType as string)) {
    errors.push({ path: `${path}.listType`, message: 'Invalid list type', code: 'INVALID_LIST_TYPE' })
  }

  if (!Array.isArray(block.items)) {
    errors.push({ path: `${path}.items`, message: 'List items must be an array', code: 'INVALID_ITEMS' })
  } else {
    block.items.forEach((item, i) => {
      errors.push(...validateListItem(item, `${path}.items[${i}]`))
    })
  }

  return errors
}

function validateListItem(item: unknown, path: string): DslValidationError[] {
  const errors: DslValidationError[] = []

  if (!item || typeof item !== 'object') {
    errors.push({ path, message: 'List item must be an object', code: 'INVALID_LIST_ITEM' })
    return errors
  }

  const i = item as Record<string, unknown>
  if (i.content === undefined) {
    errors.push({ path: `${path}.content`, message: 'List item content is required', code: 'MISSING_CONTENT' })
  } else {
    errors.push(...validateContent(i.content, `${path}.content`))
  }

  if (i.children !== undefined && Array.isArray(i.children)) {
    i.children.forEach((child, idx) => {
      errors.push(...validateListItem(child, `${path}.children[${idx}]`))
    })
  }

  return errors
}

function validateTable(block: Record<string, unknown>, path: string): DslValidationError[] {
  const errors: DslValidationError[] = []

  if (!Array.isArray(block.rows)) {
    errors.push({ path: `${path}.rows`, message: 'Table rows must be an array', code: 'INVALID_ROWS' })
    return errors
  }

  if (block.rows.length === 0) {
    errors.push({ path: `${path}.rows`, message: 'Table must have at least one row', code: 'EMPTY_TABLE' })
  }

  block.rows.forEach((row, i) => {
    errors.push(...validateTableRow(row, `${path}.rows[${i}]`))
  })

  return errors
}

function validateTableRow(row: unknown, path: string): DslValidationError[] {
  const errors: DslValidationError[] = []

  if (!row || typeof row !== 'object') {
    errors.push({ path, message: 'Table row must be an object', code: 'INVALID_ROW' })
    return errors
  }

  const r = row as Record<string, unknown>
  if (!Array.isArray(r.cells)) {
    errors.push({ path: `${path}.cells`, message: 'Row cells must be an array', code: 'INVALID_CELLS' })
    return errors
  }

  r.cells.forEach((cell, i) => {
    errors.push(...validateTableCell(cell, `${path}.cells[${i}]`))
  })

  return errors
}

function validateTableCell(cell: unknown, path: string): DslValidationError[] {
  const errors: DslValidationError[] = []

  if (!cell || typeof cell !== 'object') {
    errors.push({ path, message: 'Table cell must be an object', code: 'INVALID_CELL' })
    return errors
  }

  const c = cell as Record<string, unknown>
  if (c.content === undefined) {
    errors.push({ path: `${path}.content`, message: 'Cell content is required', code: 'MISSING_CONTENT' })
  }

  // colSpan 和 rowSpan 必须是正整数
  if (c.colSpan !== undefined && (typeof c.colSpan !== 'number' || c.colSpan < 1)) {
    errors.push({ path: `${path}.colSpan`, message: 'colSpan must be a positive integer', code: 'INVALID_COLSPAN' })
  }
  if (c.rowSpan !== undefined && (typeof c.rowSpan !== 'number' || c.rowSpan < 1)) {
    errors.push({ path: `${path}.rowSpan`, message: 'rowSpan must be a positive integer', code: 'INVALID_ROWSPAN' })
  }

  return errors
}

function validateImage(block: Record<string, unknown>, path: string): DslValidationError[] {
  const errors: DslValidationError[] = []

  if (typeof block.src !== 'string' || !block.src) {
    errors.push({ path: `${path}.src`, message: 'Image src is required', code: 'MISSING_SRC' })
  }

  return errors
}

function validateContent(content: unknown, path: string): DslValidationError[] {
  const errors: DslValidationError[] = []

  if (typeof content === 'string') {
    // 纯文本，OK
    return errors
  }

  if (Array.isArray(content)) {
    content.forEach((item, i) => {
      if (typeof item === 'string') {
        // OK
      } else if (item && typeof item === 'object') {
        const run = item as Record<string, unknown>
        if (typeof run.text !== 'string') {
          errors.push({ path: `${path}[${i}].text`, message: 'Run text is required', code: 'MISSING_TEXT' })
        }
      } else {
        errors.push({ path: `${path}[${i}]`, message: 'Invalid inline content', code: 'INVALID_INLINE' })
      }
    })
    return errors
  }

  errors.push({ path, message: 'Content must be string or array', code: 'INVALID_CONTENT' })
  return errors
}

function validateParagraphFormat(format: unknown, path: string): DslValidationError[] {
  const errors: DslValidationError[] = []

  if (!format || typeof format !== 'object') {
    errors.push({ path, message: 'Format must be an object', code: 'INVALID_FORMAT' })
    return errors
  }

  const f = format as Record<string, unknown>

  if (f.alignment !== undefined && !['left', 'center', 'right', 'justify'].includes(f.alignment as string)) {
    errors.push({ path: `${path}.alignment`, message: 'Invalid alignment value', code: 'INVALID_ALIGNMENT' })
  }

  return errors
}

function validatePageSetup(pageSetup: unknown, path: string): DslValidationError[] {
  const errors: DslValidationError[] = []

  if (!pageSetup || typeof pageSetup !== 'object') {
    errors.push({ path, message: 'pageSetup must be an object', code: 'INVALID_PAGE_SETUP' })
    return errors
  }

  const ps = pageSetup as Record<string, unknown>

  if (ps.paperSize !== undefined && !['A4', 'A3', 'Letter', 'Legal', 'custom'].includes(ps.paperSize as string)) {
    errors.push({ path: `${path}.paperSize`, message: 'Invalid paper size', code: 'INVALID_PAPER_SIZE' })
  }

  if (ps.orientation !== undefined && !['portrait', 'landscape'].includes(ps.orientation as string)) {
    errors.push({ path: `${path}.orientation`, message: 'Invalid orientation', code: 'INVALID_ORIENTATION' })
  }

  return errors
}

function validateHeaderFooter(hf: unknown, path: string): DslValidationError[] {
  const errors: DslValidationError[] = []

  if (!hf || typeof hf !== 'object') {
    errors.push({ path, message: 'headerFooter must be an object', code: 'INVALID_HEADER_FOOTER' })
  }

  return errors
}

// ============== 长度转换 ==============

/**
 * 将 DSL 长度转换为 twips (1/20 pt)
 */
export function dslLengthToTwips(length: DslLength | undefined, baseFontPt: number = 12): number | undefined {
  if (length === undefined || length === null) return undefined

  if (typeof length === 'number') {
    // 默认为 pt
    return Math.round(length * 20)
  }

  const v = length.toString().trim().toLowerCase()
  if (!v) return undefined

  // pt
  const ptMatch = v.match(/^(\d+(?:\.\d+)?)\s*pt$/)
  if (ptMatch) return Math.round(Number(ptMatch[1]) * 20)

  // px (1px ≈ 0.75pt)
  const pxMatch = v.match(/^(\d+(?:\.\d+)?)\s*px$/)
  if (pxMatch) return Math.round(Number(pxMatch[1]) * 0.75 * 20)

  // cm (1in = 2.54cm, 1in = 1440 twips)
  const cmMatch = v.match(/^(\d+(?:\.\d+)?)\s*cm$/)
  if (cmMatch) return Math.round((Number(cmMatch[1]) / 2.54) * 1440)

  // in
  const inMatch = v.match(/^(\d+(?:\.\d+)?)\s*in$/)
  if (inMatch) return Math.round(Number(inMatch[1]) * 1440)

  // em
  const emMatch = v.match(/^(\d+(?:\.\d+)?)\s*em$/)
  if (emMatch) return Math.round(Number(emMatch[1]) * baseFontPt * 20)

  // 百分比
  const percentMatch = v.match(/^(\d+(?:\.\d+)?)\s*%$/)
  if (percentMatch) {
    // 返回百分比值（调用方需特殊处理）
    return undefined
  }

  // 纯数字
  const numMatch = v.match(/^(\d+(?:\.\d+)?)$/)
  if (numMatch) return Math.round(Number(numMatch[1]) * 20)

  return undefined
}

/**
 * 将 DSL 长度转换为 pt
 */
export function dslLengthToPt(length: DslLength | undefined, baseFontPt: number = 12): number | undefined {
  const twips = dslLengthToTwips(length, baseFontPt)
  return twips !== undefined ? twips / 20 : undefined
}

/**
 * 将 DSL 长度转换为 half-points (用于 docx 字号)
 */
export function dslLengthToHalfPoints(length: DslLength | undefined): number | undefined {
  const pt = dslLengthToPt(length)
  return pt !== undefined ? Math.round(pt * 2) : undefined
}

// ============== 颜色转换 ==============

/**
 * 将 DSL 颜色转换为 6 位 hex（不含 #）
 */
export function dslColorToHex(color: DslColor | undefined): string | undefined {
  if (!color) return undefined

  const v = color.toString().trim()
  if (!v) return undefined

  // 已经是 hex
  if (v.startsWith('#')) {
    const hex = v.slice(1).trim()
    if (/^[0-9a-fA-F]{6}$/.test(hex)) return hex.toUpperCase()
    if (/^[0-9a-fA-F]{3}$/.test(hex)) {
      return hex.split('').map(c => c + c).join('').toUpperCase()
    }
    return undefined
  }

  // rgb()
  const rgbMatch = v.match(/rgb\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)/i)
  if (rgbMatch) {
    const r = Math.max(0, Math.min(255, Number(rgbMatch[1])))
    const g = Math.max(0, Math.min(255, Number(rgbMatch[2])))
    const b = Math.max(0, Math.min(255, Number(rgbMatch[3])))
    const toHex = (n: number) => n.toString(16).padStart(2, '0').toUpperCase()
    return `${toHex(r)}${toHex(g)}${toHex(b)}`
  }

  // 常用颜色名
  const colorNames: Record<string, string> = {
    black: '000000',
    white: 'FFFFFF',
    red: 'FF0000',
    green: '00FF00',
    blue: '0000FF',
    yellow: 'FFFF00',
    gray: '808080',
    grey: '808080',
    orange: 'FFA500',
    purple: '800080',
    pink: 'FFC0CB',
    brown: 'A52A2A',
    navy: '000080',
    teal: '008080',
    cyan: '00FFFF',
    magenta: 'FF00FF',
  }

  const lower = v.toLowerCase()
  if (colorNames[lower]) return colorNames[lower]

  return undefined
}

// ============== 对齐转换 ==============

/**
 * DSL 对齐方式转换为 docx AlignmentType 字符串
 */
export function dslAlignmentToDocx(alignment: DslAlignment | undefined): string | undefined {
  switch (alignment) {
    case 'left': return 'left'
    case 'center': return 'center'
    case 'right': return 'right'
    case 'justify': return 'both'
    default: return undefined
  }
}

// ============== 内容规范化 ==============

/**
 * 将 content 规范化为 DslRun 数组
 */
export function normalizeContent(content: string | DslInline[]): DslRun[] {
  if (typeof content === 'string') {
    return [{ text: content }]
  }

  return content.map(item => {
    if (typeof item === 'string') {
      return { text: item }
    }
    return item
  })
}

/**
 * 提取纯文本内容
 */
export function extractPlainText(content: string | DslInline[]): string {
  if (typeof content === 'string') {
    return content
  }

  return content.map(item => {
    if (typeof item === 'string') return item
    return item.text
  }).join('')
}

// ============== 便捷构造函数 ==============

/**
 * 创建标题块
 */
export function heading(level: 1 | 2 | 3 | 4 | 5 | 6, content: string | DslInline[], format?: DslParagraphFormat): DslHeading {
  return { type: 'heading', level, content, format }
}

/**
 * 创建段落块
 */
export function paragraph(content: string | DslInline[], format?: DslParagraphFormat): DslParagraph {
  return { type: 'paragraph', content, format }
}

/**
 * 创建文本 Run
 */
export function run(text: string, options?: Omit<DslRun, 'text'>): DslRun {
  return { text, ...options }
}

/**
 * 创建粗体文本
 */
export function bold(text: string): DslRun {
  return { text, bold: true }
}

/**
 * 创建斜体文本
 */
export function italic(text: string): DslRun {
  return { text, italic: true }
}

/**
 * 创建带颜色文本
 */
export function colored(text: string, color: DslColor): DslRun {
  return { text, color }
}

// ============== 边框处理 ==============

/**
 * 规范化边框定义
 */
export function normalizeBorder(border: DslBorder | undefined): {
  top?: DslBorderSide
  bottom?: DslBorderSide
  left?: DslBorderSide
  right?: DslBorderSide
} | undefined {
  if (!border) return undefined

  const result: {
    top?: DslBorderSide
    bottom?: DslBorderSide
    left?: DslBorderSide
    right?: DslBorderSide
  } = {}

  if (border.all) {
    result.top = border.all
    result.bottom = border.all
    result.left = border.all
    result.right = border.all
  }

  if (border.top) result.top = border.top
  if (border.bottom) result.bottom = border.bottom
  if (border.left) result.left = border.left
  if (border.right) result.right = border.right

  return result
}

// ============== 类型守卫 ==============

export function isHeading(block: DslBlock): block is DslHeading {
  return block.type === 'heading'
}

export function isParagraph(block: DslBlock): block is DslParagraph {
  return block.type === 'paragraph'
}

export function isList(block: DslBlock): block is DslList {
  return block.type === 'list'
}

export function isTable(block: DslBlock): block is DslTable {
  return block.type === 'table'
}

export function isImage(block: DslBlock): block is DslImage {
  return block.type === 'image'
}

// ============== DSL 到 HTML 渲染 ==============

/**
 * HTML 转义
 */
function escapeHtml(str: string): string {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;')
}

/**
 * 生成行内样式字符串
 */
function runToStyle(run: DslRun): string {
  const styles: string[] = []
  
  if (run.bold) styles.push('font-weight: bold')
  if (run.italic) styles.push('font-style: italic')
  if (run.underline) styles.push('text-decoration: underline')
  if (run.strikethrough) {
    if (run.underline) {
      styles.push('text-decoration: underline line-through')
    } else {
      styles.push('text-decoration: line-through')
    }
  }
  if (run.fontFamily) styles.push(`font-family: ${run.fontFamily}`)
  if (run.fontSize) styles.push(`font-size: ${run.fontSize}pt`)
  if (run.color) {
    const hex = dslColorToHex(run.color)
    if (hex) styles.push(`color: #${hex}`)
  }
  if (run.highlight) {
    const hex = dslColorToHex(run.highlight)
    if (hex) styles.push(`background-color: #${hex}`)
  }
  if (run.letterSpacing) styles.push(`letter-spacing: ${run.letterSpacing}pt`)
  
  return styles.join('; ')
}

/**
 * 渲染行内内容为 HTML
 */
function renderInlineToHtml(content: string | DslInline[]): string {
  if (typeof content === 'string') {
    return escapeHtml(content)
  }
  
  return content.map(item => {
    if (typeof item === 'string') {
      return escapeHtml(item)
    }
    
    const run = item
    let html = escapeHtml(run.text)
    
    // 上标/下标
    if (run.superscript) html = `<sup>${html}</sup>`
    if (run.subscript) html = `<sub>${html}</sub>`
    
    // 应用样式
    const style = runToStyle(run)
    if (style) {
      html = `<span style="${style}">${html}</span>`
    }
    
    return html
  }).join('')
}

/**
 * 生成段落样式
 */
function paragraphFormatToStyle(format: DslParagraphFormat | undefined): string {
  if (!format) return ''
  
  const styles: string[] = []
  
  if (format.alignment) styles.push(`text-align: ${format.alignment}`)
  if (format.firstLineIndent) styles.push(`text-indent: ${format.firstLineIndent}`)
  if (format.leftIndent) styles.push(`margin-left: ${format.leftIndent}`)
  if (format.rightIndent) styles.push(`margin-right: ${format.rightIndent}`)
  if (format.spaceBefore) styles.push(`margin-top: ${format.spaceBefore}`)
  if (format.spaceAfter) styles.push(`margin-bottom: ${format.spaceAfter}`)
  if (format.lineHeight) {
    if (typeof format.lineHeight === 'number') {
      styles.push(`line-height: ${format.lineHeight}`)
    } else {
      styles.push(`line-height: ${format.lineHeight}`)
    }
  }
  if (format.backgroundColor) {
    const hex = dslColorToHex(format.backgroundColor)
    if (hex) styles.push(`background-color: #${hex}`)
  }
  if (format.padding) styles.push(`padding: ${format.padding}`)
  
  // 边框
  if (format.border) {
    const border = normalizeBorder(format.border)
    if (border) {
      const borderSideToStyle = (side: DslBorderSide | undefined, prop: string) => {
        if (!side) return
        const width = side.width || '1px'
        const style = side.style || 'solid'
        const color = side.color ? `#${dslColorToHex(side.color) || '000000'}` : '#000000'
        styles.push(`${prop}: ${width} ${style} ${color}`)
      }
      borderSideToStyle(border.top, 'border-top')
      borderSideToStyle(border.bottom, 'border-bottom')
      borderSideToStyle(border.left, 'border-left')
      borderSideToStyle(border.right, 'border-right')
    }
  }
  
  return styles.join('; ')
}

/**
 * 渲染标题为 HTML
 */
function renderHeadingToHtml(block: DslHeading): string {
  const tag = `h${block.level}`
  const style = paragraphFormatToStyle(block.format)
  const styleAttr = style ? ` style="${style}"` : ''
  const content = renderInlineToHtml(block.content)
  return `<${tag}${styleAttr}>${content}</${tag}>`
}

/**
 * 渲染段落为 HTML
 */
function renderParagraphToHtml(block: DslParagraph): string {
  const style = paragraphFormatToStyle(block.format)
  const styleAttr = style ? ` style="${style}"` : ''
  const content = renderInlineToHtml(block.content)
  return `<p${styleAttr}>${content}</p>`
}

/**
 * 渲染列表项为 HTML
 */
function renderListItemToHtml(item: DslListItem, listType: string): string {
  const content = renderInlineToHtml(item.content)
  let html = `<li>${content}`
  
  if (item.children && item.children.length > 0) {
    const tag = listType === 'bullet' ? 'ul' : 'ol'
    const childrenHtml = item.children.map(child => renderListItemToHtml(child, listType)).join('')
    html += `<${tag}>${childrenHtml}</${tag}>`
  }
  
  html += '</li>'
  return html
}

/**
 * 渲染列表为 HTML
 */
function renderListToHtml(block: DslList): string {
  const tag = block.listType === 'bullet' ? 'ul' : 'ol'
  const items = block.items.map(item => renderListItemToHtml(item, block.listType)).join('')
  
  let attrs = ''
  if (block.listType !== 'bullet' && block.startAt && block.startAt !== 1) {
    attrs = ` start="${block.startAt}"`
  }
  if (block.listType === 'letter') {
    attrs += ' style="list-style-type: lower-alpha"'
  } else if (block.listType === 'roman') {
    attrs += ' style="list-style-type: lower-roman"'
  }
  
  return `<${tag}${attrs}>${items}</${tag}>`
}

/**
 * 渲染表格单元格为 HTML
 */
function renderTableCellToHtml(cell: DslTableCell, isHeader: boolean, tableBorder?: DslBorder): string {
  const tag = isHeader ? 'th' : 'td'
  const styles: string[] = []
  const attrs: string[] = []
  
  if (cell.colSpan && cell.colSpan > 1) attrs.push(`colspan="${cell.colSpan}"`)
  if (cell.rowSpan && cell.rowSpan > 1) attrs.push(`rowspan="${cell.rowSpan}"`)
  if (cell.align) styles.push(`text-align: ${cell.align}`)
  if (cell.valign) styles.push(`vertical-align: ${cell.valign}`)
  if (cell.backgroundColor) {
    const hex = dslColorToHex(cell.backgroundColor)
    if (hex) styles.push(`background-color: #${hex}`)
  }
  if (cell.width) styles.push(`width: ${cell.width}`)
  
  // 边框
  if (cell.border || tableBorder) {
    const border = normalizeBorder(cell.border || tableBorder)
    if (border) {
      const borderSideToStyle = (side: DslBorderSide | undefined, prop: string) => {
        if (!side) return
        const width = side.width || '1px'
        const style = side.style || 'solid'
        const color = side.color ? `#${dslColorToHex(side.color) || '000000'}` : '#000000'
        styles.push(`${prop}: ${width} ${style} ${color}`)
      }
      borderSideToStyle(border.top, 'border-top')
      borderSideToStyle(border.bottom, 'border-bottom')
      borderSideToStyle(border.left, 'border-left')
      borderSideToStyle(border.right, 'border-right')
    }
  }
  
  if (styles.length > 0) attrs.push(`style="${styles.join('; ')}"`)
  
  const attrStr = attrs.length > 0 ? ' ' + attrs.join(' ') : ''
  
  // 渲染内容
  let content: string
  if (typeof cell.content === 'string') {
    content = escapeHtml(cell.content)
  } else if (Array.isArray(cell.content)) {
    // 检查是否为块数组
    if (cell.content.length > 0 && typeof cell.content[0] === 'object' && 'type' in (cell.content[0] as object)) {
      // 块数组
      content = (cell.content as DslBlock[]).map(renderBlockToHtml).join('')
    } else {
      // 行内数组
      content = renderInlineToHtml(cell.content as DslInline[])
    }
  } else {
    content = ''
  }
  
  return `<${tag}${attrStr}>${content}</${tag}>`
}

/**
 * 渲染表格行为 HTML
 */
function renderTableRowToHtml(row: DslTableRow, tableBorder?: DslBorder): string {
  const cells = row.cells.map(cell => renderTableCellToHtml(cell, row.isHeader || false, tableBorder)).join('')
  const style = row.height ? ` style="height: ${row.height}"` : ''
  return `<tr${style}>${cells}</tr>`
}

/**
 * 渲染表格为 HTML
 */
function renderTableToHtml(block: DslTable): string {
  const styles: string[] = ['border-collapse: collapse']
  
  if (block.width) styles.push(`width: ${block.width}`)
  if (block.alignment) styles.push(`margin-left: ${block.alignment === 'center' ? 'auto' : block.alignment === 'right' ? 'auto' : '0'}`)
  if (block.alignment === 'center' || block.alignment === 'right') {
    styles.push(`margin-right: ${block.alignment === 'center' ? 'auto' : '0'}`)
  }
  
  // 默认边框
  if (block.border) {
    const border = normalizeBorder(block.border)
    const side = border?.all || border?.top || border?.left || border?.right || border?.bottom
    if (side) {
      const width = side.width || '1px'
      const style = side.style || 'solid'
      const color = side.color ? `#${dslColorToHex(side.color) || '000000'}` : '#000000'
      styles.push(`border: ${width} ${style} ${color}`)
    }
  }
  
  const styleAttr = styles.length > 0 ? ` style="${styles.join('; ')}"` : ''
  
  // 分离表头和表体
  const headerRows = block.rows.filter(r => r.isHeader)
  const bodyRows = block.rows.filter(r => !r.isHeader)
  
  let html = `<table${styleAttr}>`
  
  if (headerRows.length > 0) {
    html += '<thead>'
    html += headerRows.map(row => renderTableRowToHtml(row, block.border)).join('')
    html += '</thead>'
  }
  
  if (bodyRows.length > 0) {
    html += '<tbody>'
    html += bodyRows.map(row => renderTableRowToHtml(row, block.border)).join('')
    html += '</tbody>'
  } else if (headerRows.length === 0) {
    // 没有行
    html += '<tbody></tbody>'
  }
  
  html += '</table>'
  return html
}

/**
 * 渲染图片为 HTML
 */
function renderImageToHtml(block: DslImage): string {
  const styles: string[] = []
  const attrs: string[] = [`src="${escapeHtml(block.src)}"`]
  
  if (block.alt) attrs.push(`alt="${escapeHtml(block.alt)}"`)
  if (block.width) styles.push(`width: ${block.width}`)
  if (block.height) styles.push(`height: ${block.height}`)
  
  if (styles.length > 0) attrs.push(`style="${styles.join('; ')}"`)
  
  let img = `<img ${attrs.join(' ')} />`
  
  // 对齐包装
  if (block.alignment && block.alignment !== 'left') {
    const alignStyle = block.alignment === 'center' 
      ? 'text-align: center' 
      : block.alignment === 'right' 
        ? 'text-align: right' 
        : ''
    if (alignStyle) {
      img = `<div style="${alignStyle}">${img}</div>`
    }
  }
  
  // 标题
  if (block.caption) {
    img = `<figure style="margin: 1em 0; ${block.alignment === 'center' ? 'text-align: center' : ''}">${img}<figcaption style="font-size: 0.9em; color: #666">${escapeHtml(block.caption)}</figcaption></figure>`
  }
  
  return img
}

/**
 * 渲染单个块为 HTML
 */
function renderBlockToHtml(block: DslBlock): string {
  switch (block.type) {
    case 'heading':
      return renderHeadingToHtml(block)
    case 'paragraph':
      return renderParagraphToHtml(block)
    case 'list':
      return renderListToHtml(block)
    case 'table':
      return renderTableToHtml(block)
    case 'image':
      return renderImageToHtml(block)
    case 'pageBreak':
      return '<div style="page-break-after: always"></div>'
    case 'sectionBreak':
      return '<hr style="border: none; border-top: 1px dashed #ccc; margin: 2em 0" />'
    case 'horizontalRule':
      const hrStyles: string[] = []
      if (block.color) {
        const hex = dslColorToHex(block.color)
        if (hex) hrStyles.push(`border-color: #${hex}`)
      }
      if (block.width) hrStyles.push(`border-width: ${block.width}`)
      const hrStyle = hrStyles.length > 0 ? ` style="${hrStyles.join('; ')}"` : ''
      return `<hr${hrStyle} />`
    case 'blockquote':
      const quoteContent = block.content.map(renderBlockToHtml).join('')
      return `<blockquote style="margin: 1em 0; padding: 0.5em 1em; border-left: 4px solid #ddd; background: #f9f9f9">${quoteContent}</blockquote>`
    default:
      return ''
  }
}

/**
 * 将 DocDsl 渲染为 HTML
 */
export function dslToHtml(dsl: DocDsl): string {
  const parts: string[] = []
  
  for (const block of dsl.blocks) {
    parts.push(renderBlockToHtml(block))
  }
  
  return parts.join('\n')
}

/**
 * 将 DocDsl 渲染为 HTML（带完整样式包装）
 */
export function dslToHtmlDocument(dsl: DocDsl): string {
  const bodyHtml = dslToHtml(dsl)
  
  // 页面设置转换为 CSS
  const pageStyles: string[] = []
  if (dsl.pageSetup?.margins) {
    const m = dsl.pageSetup.margins
    if (m.top) pageStyles.push(`padding-top: ${m.top}`)
    if (m.bottom) pageStyles.push(`padding-bottom: ${m.bottom}`)
    if (m.left) pageStyles.push(`padding-left: ${m.left}`)
    if (m.right) pageStyles.push(`padding-right: ${m.right}`)
  }
  
  const pageStyle = pageStyles.length > 0 ? ` style="${pageStyles.join('; ')}"` : ''
  
  return `<div class="dsl-document"${pageStyle}>${bodyHtml}</div>`
}
