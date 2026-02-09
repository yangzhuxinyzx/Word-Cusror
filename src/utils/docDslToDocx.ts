/**
 * DocDsl 到 DOCX 转换
 * 使用 docx 库将 DSL 转换为 Word 文档
 */

import {
  Document,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  HeadingLevel,
  AlignmentType,
  WidthType,
  BorderStyle,
  LevelFormat,
  ImageRun,
  PageBreak,
  UnderlineType,
  VerticalAlign,
  Packer,
  IRunOptions,
  IParagraphOptions,
  ITableCellOptions,
  ITableRowOptions,
  ITableOptions,
  ISectionOptions,
} from 'docx'

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
  DslParagraphFormat,
  DslBorder,
  DslBorderSide,
  DslAlignment,
  DslVerticalAlign,
  DslBorderStyle,
  DslPageSetup,
} from '../types/docDsl'

import {
  dslLengthToTwips,
  dslColorToHex,
  normalizeContent,
  normalizeBorder,
} from './docDsl'

// ============== 类型映射 ==============

/**
 * 对齐方式映射
 */
function mapAlignment(align: DslAlignment | undefined): AlignmentType | undefined {
  switch (align) {
    case 'left': return AlignmentType.LEFT
    case 'center': return AlignmentType.CENTER
    case 'right': return AlignmentType.RIGHT
    case 'justify': return AlignmentType.JUSTIFIED
    default: return undefined
  }
}

/**
 * 垂直对齐映射
 */
function mapVerticalAlign(align: DslVerticalAlign | undefined): VerticalAlign | undefined {
  switch (align) {
    case 'top': return VerticalAlign.TOP
    case 'middle': return VerticalAlign.CENTER
    case 'bottom': return VerticalAlign.BOTTOM
    default: return undefined
  }
}

/**
 * 边框样式映射
 */
function mapBorderStyle(style: DslBorderStyle | undefined): BorderStyle {
  switch (style) {
    case 'none': return BorderStyle.NONE
    case 'single': return BorderStyle.SINGLE
    case 'double': return BorderStyle.DOUBLE
    case 'dashed': return BorderStyle.DASHED
    case 'dotted': return BorderStyle.DOTTED
    case 'thick': return BorderStyle.THICK
    default: return BorderStyle.SINGLE
  }
}

/**
 * 标题级别映射
 */
function mapHeadingLevel(level: number): HeadingLevel {
  switch (level) {
    case 1: return HeadingLevel.HEADING_1
    case 2: return HeadingLevel.HEADING_2
    case 3: return HeadingLevel.HEADING_3
    case 4: return HeadingLevel.HEADING_4
    case 5: return HeadingLevel.HEADING_5
    case 6: return HeadingLevel.HEADING_6
    default: return HeadingLevel.HEADING_1
  }
}

/**
 * DSL 长度转换为像素（用于图片）
 */
function dslLengthToPx(length: string | number | undefined): number | undefined {
  const twips = dslLengthToTwips(length)
  if (!twips) return undefined
  return Math.max(1, Math.round(twips / 15)) // 1px ≈ 15 twips
}

// ============== Run 转换 ==============

/**
 * 将 DslRun 转换为 TextRun 选项
 */
function dslRunToTextRunOptions(run: DslRun): IRunOptions {
  const options: IRunOptions = {
    text: run.text,
  }

  if (run.bold) options.bold = true
  if (run.italic) options.italics = true
  if (run.underline) options.underline = { type: UnderlineType.SINGLE }
  if (run.strikethrough) options.strike = true
  if (run.superscript) options.superScript = true
  if (run.subscript) options.subScript = true

  if (run.fontFamily) {
    options.font = {
      ascii: run.fontFamily,
      hAnsi: run.fontFamily,
      eastAsia: run.fontFamily,
    }
  }

  if (run.fontSize) {
    options.size = Math.round(run.fontSize * 2) // 半点
  }

  if (run.color) {
    const hex = dslColorToHex(run.color)
    if (hex) options.color = hex
  }

  if (run.highlight) {
    const hex = dslColorToHex(run.highlight)
    if (hex) options.shading = { fill: hex }
  }

  return options
}

/**
 * 将行内内容转换为 TextRun 数组
 */
function contentToTextRuns(content: string | DslInline[]): TextRun[] {
  const runs = normalizeContent(content)
  return runs.map(run => new TextRun(dslRunToTextRunOptions(run)))
}

// ============== 段落格式转换 ==============

/**
 * 将段落格式转换为 Paragraph 选项
 */
function paragraphFormatToOptions(format: DslParagraphFormat | undefined): Partial<IParagraphOptions> {
  if (!format) return {}

  const options: Partial<IParagraphOptions> = {}

  if (format.alignment) {
    options.alignment = mapAlignment(format.alignment)
  }

  // 缩进
  const indent: Record<string, number> = {}
  if (format.firstLineIndent) {
    const twips = dslLengthToTwips(format.firstLineIndent)
    if (twips) indent.firstLine = twips
  }
  if (format.leftIndent) {
    const twips = dslLengthToTwips(format.leftIndent)
    if (twips) indent.left = twips
  }
  if (format.rightIndent) {
    const twips = dslLengthToTwips(format.rightIndent)
    if (twips) indent.right = twips
  }
  if (format.hangingIndent) {
    const twips = dslLengthToTwips(format.hangingIndent)
    if (twips) indent.hanging = twips
  }
  if (Object.keys(indent).length > 0) {
    options.indent = indent
  }

  // 间距
  const spacing: Record<string, number> = {}
  if (format.spaceBefore) {
    const twips = dslLengthToTwips(format.spaceBefore)
    if (twips) spacing.before = twips
  }
  if (format.spaceAfter) {
    const twips = dslLengthToTwips(format.spaceAfter)
    if (twips) spacing.after = twips
  }
  if (format.lineHeight) {
    if (typeof format.lineHeight === 'number') {
      // 倍数行距
      spacing.line = Math.round(format.lineHeight * 240) // 240 twips = 单倍行距
      spacing.lineRule = 'auto' as any
    } else {
      const twips = dslLengthToTwips(format.lineHeight)
      if (twips) {
        spacing.line = twips
        spacing.lineRule = 'exact' as any
      }
    }
  }
  if (Object.keys(spacing).length > 0) {
    options.spacing = spacing
  }

  // 边框
  if (format.border) {
    options.border = dslBorderToParagraphBorder(format.border)
  }

  // 背景色
  if (format.backgroundColor) {
    const hex = dslColorToHex(format.backgroundColor)
    if (hex) {
      options.shading = { fill: hex }
    }
  }

  return options
}

/**
 * 将 DSL 边框转换为段落边框
 */
function dslBorderToParagraphBorder(border: DslBorder): Record<string, unknown> {
  const normalized = normalizeBorder(border)
  if (!normalized) return {}

  const result: Record<string, unknown> = {}

  const convertSide = (side: DslBorderSide | undefined) => {
    if (!side) return undefined
    const sizeTwips = dslLengthToTwips(side.width) || 12
    return {
      style: mapBorderStyle(side.style),
      size: sizeTwips / 8, // 转换为八分之一点
      color: dslColorToHex(side.color) || '000000',
    }
  }

  if (normalized.top) result.top = convertSide(normalized.top)
  if (normalized.bottom) result.bottom = convertSide(normalized.bottom)
  if (normalized.left) result.left = convertSide(normalized.left)
  if (normalized.right) result.right = convertSide(normalized.right)

  return result
}

// ============== 块转换 ==============

/**
 * 转换标题块
 */
function convertHeading(block: DslHeading): Paragraph {
  const formatOptions = paragraphFormatToOptions(block.format)
  
  return new Paragraph({
    ...formatOptions,
    heading: mapHeadingLevel(block.level),
    children: contentToTextRuns(block.content),
  })
}

/**
 * 转换段落块
 */
function convertParagraph(block: DslParagraph): Paragraph {
  const formatOptions = paragraphFormatToOptions(block.format)
  
  return new Paragraph({
    ...formatOptions,
    children: contentToTextRuns(block.content),
  })
}

/**
 * 转换列表
 */
function convertList(block: DslList): Paragraph[] {
  const paragraphs: Paragraph[] = []
  const listLevelBase = block.level || 0
  const reference =
    block.listType === 'number'
      ? 'numbered-list'
      : block.listType === 'letter'
        ? 'letter-list'
        : block.listType === 'roman'
          ? 'roman-list'
          : null
  
  const processItems = (items: DslListItem[], level: number) => {
    items.forEach((item, index) => {
      const runs = contentToTextRuns(item.content)
      
      if (block.listType === 'bullet') {
        paragraphs.push(new Paragraph({
          bullet: { level: Math.min(level, 2) },
          children: runs,
        }))
      } else {
        // 有序列表
        paragraphs.push(new Paragraph({
          numbering: {
            reference: reference || 'numbered-list',
            level: Math.min(level, 2),
          },
          children: runs,
        }))
      }
      
      // 处理嵌套
      if (item.children && item.children.length > 0) {
        processItems(item.children, level + 1)
      }
    })
  }
  
  processItems(block.items, listLevelBase)
  
  return paragraphs
}

/**
 * 转换表格单元格
 */
function convertTableCell(cell: DslTableCell): TableCell {
  const options: ITableCellOptions = {
    children: [],
  }

  const blocksToParagraphs = (blocks: DslBlock[]): Paragraph[] => {
    const paragraphs: Paragraph[] = []
    for (const block of blocks) {
      if (block.type === 'heading') {
        paragraphs.push(convertHeading(block))
      } else if (block.type === 'paragraph') {
        paragraphs.push(convertParagraph(block))
      } else if (block.type === 'list') {
        paragraphs.push(...convertList(block))
      } else if (block.type === 'image') {
        paragraphs.push(new Paragraph({ children: [new TextRun({ text: `[图片]` })] }))
      } else if (block.type === 'pageBreak') {
        paragraphs.push(convertPageBreak())
      } else if (block.type === 'horizontalRule') {
        paragraphs.push(convertHorizontalRule())
      }
    }
    return paragraphs
  }

  // 内容转换
  if (typeof cell.content === 'string') {
    const paraOptions: IParagraphOptions = { children: [new TextRun(cell.content)] }
    if (cell.align) paraOptions.alignment = mapAlignment(cell.align)
    options.children = [new Paragraph(paraOptions)]
  } else if (Array.isArray(cell.content)) {
    if (cell.content.length > 0 && typeof cell.content[0] === 'object' && 'type' in (cell.content[0] as object)) {
      // 块数组
      options.children = blocksToParagraphs(cell.content as DslBlock[])
    } else {
      // 行内数组
      const paraOptions: IParagraphOptions = { children: contentToTextRuns(cell.content as DslInline[]) }
      if (cell.align) paraOptions.alignment = mapAlignment(cell.align)
      options.children = [new Paragraph(paraOptions)]
    }
  }

  // 合并
  if (cell.colSpan && cell.colSpan > 1) {
    options.columnSpan = cell.colSpan
  }
  if (cell.rowSpan && cell.rowSpan > 1) {
    options.rowSpan = cell.rowSpan
  }

  // 对齐
  if (cell.valign) {
    options.verticalAlign = mapVerticalAlign(cell.valign)
  }

  // 背景色
  if (cell.backgroundColor) {
    const hex = dslColorToHex(cell.backgroundColor)
    if (hex) {
      options.shading = { fill: hex }
    }
  }

  // 宽度
  if (cell.width) {
    const widthTwips = dslLengthToTwips(cell.width)
    if (widthTwips) {
      options.width = { size: widthTwips, type: WidthType.DXA }
    }
  }

  // 边框
  if (cell.border) {
    options.borders = dslBorderToTableCellBorders(cell.border)
  }

  return new TableCell(options)
}

/**
 * 将 DSL 边框转换为表格单元格边框
 */
function dslBorderToTableCellBorders(border: DslBorder): Record<string, unknown> {
  const normalized = normalizeBorder(border)
  if (!normalized) return {}

  const result: Record<string, unknown> = {}

  const convertSide = (side: DslBorderSide | undefined) => {
    if (!side) return undefined
    const sizeTwips = dslLengthToTwips(side.width) || 12
    return {
      style: mapBorderStyle(side.style),
      size: sizeTwips / 8,
      color: dslColorToHex(side.color) || '000000',
    }
  }

  if (normalized.top) result.top = convertSide(normalized.top)
  if (normalized.bottom) result.bottom = convertSide(normalized.bottom)
  if (normalized.left) result.left = convertSide(normalized.left)
  if (normalized.right) result.right = convertSide(normalized.right)

  return result
}

function dslBorderToTableBorders(border: DslBorder): Record<string, unknown> {
  const normalized = normalizeBorder(border)
  if (!normalized) return {}

  const result: Record<string, unknown> = {}

  const convertSide = (side: DslBorderSide | undefined) => {
    if (!side) return undefined
    const sizeTwips = dslLengthToTwips(side.width) || 12
    return {
      style: mapBorderStyle(side.style),
      size: sizeTwips / 8,
      color: dslColorToHex(side.color) || '000000',
    }
  }

  if (normalized.top) result.top = convertSide(normalized.top)
  if (normalized.bottom) result.bottom = convertSide(normalized.bottom)
  if (normalized.left) result.left = convertSide(normalized.left)
  if (normalized.right) result.right = convertSide(normalized.right)

  return result
}

/**
 * 转换表格行
 */
function convertTableRow(row: DslTableRow): TableRow {
  const options: ITableRowOptions = {
    children: row.cells.map(convertTableCell),
  }

  if (row.height) {
    const heightTwips = dslLengthToTwips(row.height)
    if (heightTwips) {
      options.height = { value: heightTwips, rule: 'exact' as any }
    }
  }

  if (row.isHeader) {
    options.tableHeader = true
  }

  return new TableRow(options)
}

/**
 * 转换表格块
 */
function convertTable(block: DslTable): Table {
  const options: ITableOptions = {
    rows: block.rows.map(convertTableRow),
  }

  // 宽度
  if (block.width) {
    if (typeof block.width === 'string' && block.width.endsWith('%')) {
      const percent = parseFloat(block.width)
      options.width = { size: percent, type: WidthType.PERCENTAGE }
    } else {
      const widthTwips = dslLengthToTwips(block.width)
      if (widthTwips) {
        options.width = { size: widthTwips, type: WidthType.DXA }
      }
    }
  } else {
    options.width = { size: 100, type: WidthType.PERCENTAGE }
  }

  // 对齐
  if (block.alignment) {
    options.alignment = mapAlignment(block.alignment)
  }

  // 表格边框
  if (block.border) {
    options.borders = dslBorderToTableBorders(block.border)
  }

  // 列宽
  if (block.columnWidths && block.columnWidths.length > 0) {
    options.columnWidths = block.columnWidths.map(w => dslLengthToTwips(w) || 1000)
  }

  return new Table(options)
}

function decodeBase64ToBytes(base64: string): Uint8Array {
  const binary = atob(base64)
  const bytes = new Uint8Array(binary.length)
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i)
  }
  return bytes
}

async function loadImageData(src: string): Promise<Uint8Array | null> {
  if (!src) return null

  // data URL
  if (src.startsWith('data:')) {
    const base64Index = src.indexOf('base64,')
    if (base64Index !== -1) {
      const base64 = src.slice(base64Index + 7).trim()
      return decodeBase64ToBytes(base64)
    }
    return null
  }

  // 纯 base64
  if (/^[A-Za-z0-9+/=]+$/.test(src) && src.length % 4 === 0) {
    try {
      return decodeBase64ToBytes(src)
    } catch {
      // ignore
    }
  }

  // 尝试 URL 或文件路径
  let fetchUrl = src
  if (/^[A-Za-z]:\\/.test(src)) {
    fetchUrl = `file:///${src.replace(/\\/g, '/')}`
  }

  try {
    const res = await fetch(fetchUrl)
    if (!res.ok) return null
    const buffer = await res.arrayBuffer()
    return new Uint8Array(buffer)
  } catch {
    return null
  }
}

/**
 * 转换图片块
 */
async function convertImage(block: DslImage): Promise<Paragraph[]> {
  const data = await loadImageData(block.src)
  const alignment = mapAlignment(block.alignment)
  const widthPx = dslLengthToPx(block.width)
  const heightPx = dslLengthToPx(block.height)

  if (!data) {
    return [
      new Paragraph({
        children: [new TextRun({ text: `[图片加载失败: ${block.alt || block.src}]` })],
        alignment,
      })
    ]
  }

  const imageRun = new ImageRun({
    data,
    transformation: {
      width: widthPx || 300,
      height: heightPx || Math.round((widthPx || 300) * 0.75),
    },
  })

  const paragraphs: Paragraph[] = [
    new Paragraph({
      children: [imageRun],
      alignment,
    }),
  ]

  if (block.caption) {
    paragraphs.push(new Paragraph({
      children: [new TextRun({ text: block.caption, italics: true })],
      alignment,
    }))
  }

  return paragraphs
}

/**
 * 转换分页符
 */
function convertPageBreak(): Paragraph {
  return new Paragraph({
    children: [new PageBreak()],
  })
}

/**
 * 转换水平线
 */
function convertHorizontalRule(): Paragraph {
  return new Paragraph({
    border: {
      bottom: {
        style: BorderStyle.SINGLE,
        size: 6,
        color: '999999',
      },
    },
    spacing: { before: 200, after: 200 },
  })
}

/**
 * 转换块引用
 */
async function convertBlockquote(block: { content: DslBlock[] }): Promise<(Paragraph | Table)[]> {
  const converted: (Paragraph | Table)[] = []
  for (const b of block.content) {
    const items = await convertBlockAsync(b)
    converted.push(...items)
  }
  return converted
}

/**
 * 转换单个块
 */
async function convertBlockAsync(block: DslBlock): Promise<(Paragraph | Table)[]> {
  switch (block.type) {
    case 'heading':
      return [convertHeading(block)]
    case 'paragraph':
      return [convertParagraph(block)]
    case 'list':
      return convertList(block)
    case 'table':
      return [convertTable(block)]
    case 'image':
      return await convertImage(block)
    case 'pageBreak':
      return [convertPageBreak()]
    case 'sectionBreak':
      return [new Paragraph({ text: '' })] // 分节符需要特殊处理
    case 'horizontalRule':
      return [convertHorizontalRule()]
    case 'blockquote':
      return await convertBlockquote(block)
    default:
      return []
  }
}

// ============== 页面设置转换 ==============

/**
 * 转换页面设置
 */
function convertPageSetup(pageSetup: DslPageSetup | undefined): Partial<ISectionOptions['properties']> {
  if (!pageSetup) return {}

  const props: Partial<ISectionOptions['properties']> = {}

  // 页边距
  if (pageSetup.margins) {
    props.page = {
      margin: {
        top: dslLengthToTwips(pageSetup.margins.top) || 1440,
        bottom: dslLengthToTwips(pageSetup.margins.bottom) || 1440,
        left: dslLengthToTwips(pageSetup.margins.left) || 1440,
        right: dslLengthToTwips(pageSetup.margins.right) || 1440,
      },
    }
  }

  // 纸张大小和方向
  if (pageSetup.paperSize || pageSetup.orientation) {
    const page = props.page || {}
    
    // 标准纸张大小（单位：twips）
    const paperSizes: Record<string, { width: number; height: number }> = {
      A4: { width: 11906, height: 16838 },
      A3: { width: 16838, height: 23811 },
      Letter: { width: 12240, height: 15840 },
      Legal: { width: 12240, height: 20160 },
    }
    
    let width = 11906 // A4 默认
    let height = 16838
    
    if (pageSetup.paperSize && pageSetup.paperSize !== 'custom') {
      const size = paperSizes[pageSetup.paperSize]
      if (size) {
        width = size.width
        height = size.height
      }
    } else if (pageSetup.paperSize === 'custom') {
      if (pageSetup.width) width = dslLengthToTwips(pageSetup.width) || width
      if (pageSetup.height) height = dslLengthToTwips(pageSetup.height) || height
    }
    
    // 横向
    if (pageSetup.orientation === 'landscape') {
      const temp = width
      width = height
      height = temp
    }
    
    page.size = { width, height, orientation: pageSetup.orientation === 'landscape' ? 'landscape' as any : 'portrait' as any }
    props.page = page
  }

  return props
}

// ============== 主转换函数 ==============

function createNumberingLevels(format: LevelFormat) {
  return [0, 1, 2].map(level => ({
    level,
    format,
    text: `%${level + 1}.`,
    alignment: AlignmentType.LEFT,
    style: {
      paragraph: {
        indent: { left: 720 * (level + 1), hanging: 360 },
      },
    },
  }))
}

/**
 * 将 DocDsl 转换为 docx Document
 */
export async function dslToDocument(dsl: DocDsl): Promise<Document> {
  // 转换所有块
  const children: (Paragraph | Table)[] = []
  for (const block of dsl.blocks) {
    const converted = await convertBlockAsync(block)
    children.push(...converted)
  }

  // 页面设置
  const sectionProps = convertPageSetup(dsl.pageSetup)

  // 创建文档（标题颜色设为黑色，覆盖 docx 库默认的蓝色主题色）
  const doc = new Document({
    creator: 'Word-Cursor',
    title: dsl.title || '文档',
    styles: {
      default: {
        document: {
          run: {
            font: { ascii: 'Times New Roman', hAnsi: 'Times New Roman', eastAsia: '仿宋' },
            size: 24, // 小四 12pt
          },
          paragraph: {
            spacing: { before: 0, after: 0, line: 360 },
          },
        },
        heading1: {
          run: { bold: true, color: '000000', size: 32, font: { ascii: 'Arial', hAnsi: 'Arial', eastAsia: '黑体' } },
          paragraph: { spacing: { before: 240, after: 120, line: 360 } },
        },
        heading2: {
          run: { bold: true, color: '000000', size: 28, font: { ascii: 'Arial', hAnsi: 'Arial', eastAsia: '黑体' } },
          paragraph: { spacing: { before: 200, after: 100, line: 360 } },
        },
        heading3: {
          run: { bold: true, color: '000000', size: 26, font: { ascii: 'Arial', hAnsi: 'Arial', eastAsia: '黑体' } },
          paragraph: { spacing: { before: 160, after: 80, line: 340 } },
        },
      },
    },
    numbering: {
      config: [
        { reference: 'numbered-list', levels: createNumberingLevels(LevelFormat.DECIMAL) },
        { reference: 'letter-list', levels: createNumberingLevels(LevelFormat.LOWER_LETTER) },
        { reference: 'roman-list', levels: createNumberingLevels(LevelFormat.LOWER_ROMAN) },
      ],
    },
    sections: [{
      properties: sectionProps,
      children,
    }],
  })

  return doc
}

/**
 * 将 DocDsl 转换为 Blob
 */
export async function dslToDocxBlob(dsl: DocDsl): Promise<Blob> {
  const doc = await dslToDocument(dsl)
  return await Packer.toBlob(doc)
}

/**
 * 将 DocDsl 转换为 ArrayBuffer
 */
export async function dslToDocxArrayBuffer(dsl: DocDsl): Promise<ArrayBuffer> {
  const doc = await dslToDocument(dsl)
  const blob = await Packer.toBlob(doc)
  return await blob.arrayBuffer()
}

/**
 * 将 DocDsl 转换为 Base64
 */
export async function dslToDocxBase64(dsl: DocDsl): Promise<string> {
  const buffer = await dslToDocxArrayBuffer(dsl)
  const bytes = new Uint8Array(buffer)
  let binary = ''
  for (let i = 0; i < bytes.length; i++) {
    binary += String.fromCharCode(bytes[i])
  }
  return btoa(binary)
}
