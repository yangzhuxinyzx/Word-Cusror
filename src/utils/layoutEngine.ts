/**
 * 布局引擎
 * �?HTML/DocumentModel 转换为可渲染的布局信息
 * 参�?ONLYOFFICE �?Recalculate 逻辑
 */

import { 
  TextMeasurer, 
  getTextMeasurer, 
  TextStyle, 
  MeasuredText,
  A4_WIDTH_MM,
  A4_HEIGHT_MM,
  MM_TO_PX,
  PT_TO_PX
} from './textMeasurer'

const DEFAULT_LINE_HEIGHT = 1

// 布局元素类型
export type LayoutElementType = 
  | 'paragraph' 
  | 'heading' 
  | 'text' 
  | 'image' 
  | 'table' 
  | 'tableRow' 
  | 'tableCell'
  | 'list'
  | 'listItem'
  | 'pageBreak'
  | 'header'
  | 'footer'

// 布局元素基础接口
export interface LayoutElement {
  type: LayoutElementType
  x: number      // px，相对于页面左上�?
  y: number      // px，相对于页面左上�?
  width: number  // px
  height: number // px
  pageIndex: number  // 所在页码（0-based�?
  style?: ElementStyle
  children?: LayoutElement[]
}

// 文本布局元素
export interface TextLayoutElement extends LayoutElement {
  type: 'text'
  text: string
  measuredText: MeasuredText
}

// 段落布局元素
export interface ParagraphLayoutElement extends LayoutElement {
  type: 'paragraph' | 'heading'
  alignment: 'left' | 'center' | 'right' | 'justify'
  firstLineIndent: number  // 首行缩进 px
  children: TextLayoutElement[]
}

// 图片布局元素
export interface ImageLayoutElement extends LayoutElement {
  type: 'image'
  src: string
  alt?: string
}

// 表格布局元素
export interface TableLayoutElement extends LayoutElement {
  type: 'table'
  columns: number
  rows: TableRowLayoutElement[]
}

export interface TableRowLayoutElement extends LayoutElement {
  type: 'tableRow'
  cells: TableCellLayoutElement[]
}

export interface TableCellLayoutElement extends LayoutElement {
  type: 'tableCell'
  colspan?: number
  rowspan?: number
  children: LayoutElement[]
}

// 元素样式
export interface ElementStyle {
  fontFamily?: string
  fontSize?: number  // pt
  fontWeight?: 'normal' | 'bold' | number
  fontStyle?: 'normal' | 'italic'
  color?: string
  backgroundColor?: string
  textDecoration?: 'none' | 'underline' | 'line-through'
  lineHeight?: number
  marginTop?: number     // px
  marginBottom?: number  // px
  marginLeft?: number    // px
  marginRight?: number   // px
  paddingTop?: number    // px
  paddingBottom?: number // px
  paddingLeft?: number   // px
  paddingRight?: number  // px
  borderWidth?: number   // px
  borderColor?: string
  borderTopWidth?: number
  borderRightWidth?: number
  borderBottomWidth?: number
  borderLeftWidth?: number
  borderTopColor?: string
  borderRightColor?: string
  borderBottomColor?: string
  borderLeftColor?: string
  borderStyle?: 'solid' | 'dashed' | 'dotted'
}

// 页面配置
export interface PageConfig {
  width: number      // px
  height: number     // px
  marginTop: number  // px
  marginBottom: number
  marginLeft: number
  marginRight: number
  headerHeight: number
  footerHeight: number
}

export interface PageSettingsLike {
  width: number
  height: number
  marginTop: number
  marginBottom: number
  marginLeft: number
  marginRight: number
  headerHeight: number
  footerHeight: number
}

// 布局结果
export interface LayoutResult {
  pages: PageLayout[]
  totalHeight: number
}

export interface PageLayout {
  pageIndex: number
  elements: LayoutElement[]
  header?: LayoutElement
  footer?: LayoutElement
}

// 默认页面配置（A4�?
export function getDefaultPageConfig(scale: number = 1): PageConfig {
  return {
    width: A4_WIDTH_MM * MM_TO_PX * scale,
    height: A4_HEIGHT_MM * MM_TO_PX * scale,
    marginTop: 25.4 * MM_TO_PX * scale,    // 1 inch = 25.4mm
    marginBottom: 25.4 * MM_TO_PX * scale,
    marginLeft: 31.7 * MM_TO_PX * scale,   // Word 默认左边�?
    marginRight: 31.7 * MM_TO_PX * scale,
    headerHeight: 12.7 * MM_TO_PX * scale, // 0.5 inch
    footerHeight: 12.7 * MM_TO_PX * scale
  }
}

export function pageSettingsToPageConfig(settings: PageSettingsLike, scale: number = 1): PageConfig {
  return {
    width: settings.width * PT_TO_PX * scale,
    height: settings.height * PT_TO_PX * scale,
    marginTop: settings.marginTop * PT_TO_PX * scale,
    marginBottom: settings.marginBottom * PT_TO_PX * scale,
    marginLeft: settings.marginLeft * PT_TO_PX * scale,
    marginRight: settings.marginRight * PT_TO_PX * scale,
    headerHeight: settings.headerHeight * PT_TO_PX * scale,
    footerHeight: settings.footerHeight * PT_TO_PX * scale,
  }
}

/**
 * 布局引擎�?
 */
export class LayoutEngine {
  private measurer: TextMeasurer
  private pageConfig: PageConfig
  private scale: number
  
  constructor(pageConfig?: PageConfig, scale: number = 1) {
    this.measurer = getTextMeasurer()
    this.scale = scale
    this.pageConfig = pageConfig || getDefaultPageConfig(scale)
  }
  
  /**
   * 获取内容区域尺寸
   */
  getContentArea(): { x: number; y: number; width: number; height: number } {
    const { width, height, marginTop, marginBottom, marginLeft, marginRight, headerHeight, footerHeight } = this.pageConfig
    return {
      x: marginLeft,
      y: marginTop,
      width: width - marginLeft - marginRight,
      height: height - marginTop - marginBottom - headerHeight - footerHeight
    }
  }
  
  /**
   * �?HTML 解析并布局
   */
  layoutFromHtml(html: string, headerHtml?: string, footerHtml?: string): LayoutResult {
    // 解析 HTML �?DOM
    const parser = new DOMParser()
    const doc = parser.parseFromString(`<div>${html}</div>`, 'text/html')
    const container = doc.body.firstChild as HTMLElement
    
    if (!container) {
      return { pages: [], totalHeight: 0 }
    }
    
    // 解析元素
    const elements = this.parseElements(container)
    
    // 进行布局和分�?
    const pages = this.paginateElements(elements)
    
    // 添加页眉页脚
    if (headerHtml || footerHtml) {
      this.addHeaderFooter(pages, headerHtml, footerHtml)
    }
    
    // 计算总高�?
    const totalHeight = pages.length * this.pageConfig.height
    
    return { pages, totalHeight }
  }
  
  /**
   * 解析 HTML 元素
   */
  private parseElements(container: HTMLElement): LayoutElement[] {
    const elements: LayoutElement[] = []
    
    for (const child of Array.from(container.children)) {
      const element = this.parseElement(child as HTMLElement)
      if (element) {
        elements.push(element)
      }
    }
    
    return elements
  }
  
  /**
   * 解析单个 HTML 元素
   */
  private parseElement(el: HTMLElement): LayoutElement | null {
    const tagName = el.tagName.toLowerCase()
    
    // 处理分页�?
    if (el.classList.contains('page-break')) {
      return {
        type: 'pageBreak',
        x: 0,
        y: 0,
        width: 0,
        height: 0,
        pageIndex: 0
      }
    }
    
    switch (tagName) {
      case 'p':
        return this.parseParagraph(el)
      case 'h1':
      case 'h2':
      case 'h3':
      case 'h4':
      case 'h5':
      case 'h6':
        return this.parseHeading(el, tagName)
      case 'img':
        return this.parseImage(el as HTMLImageElement)
      case 'table':
        return this.parseTable(el as HTMLTableElement)
      case 'ul':
      case 'ol':
        return this.parseList(el)
      case 'div':
      case 'span':
        // 递归处理容器
        if (el.children.length > 0) {
          const children: LayoutElement[] = []
          for (const child of Array.from(el.children)) {
            const childEl = this.parseElement(child as HTMLElement)
            if (childEl) children.push(childEl)
          }
          if (children.length === 1) return children[0]
          return {
            type: 'paragraph',
            x: 0,
            y: 0,
            width: 0,
            height: 0,
            pageIndex: 0,
            children
          } as LayoutElement
        }
        // 作为文本处理
        return this.parseParagraph(el)
      default:
        // 默认作为段落处理
        return this.parseParagraph(el)
    }
  }
  
  /**
   * 解析段落
   */
  private parseParagraph(el: HTMLElement): ParagraphLayoutElement {
    const style = this.extractStyle(el)
    const isToc = el.classList.contains('docx-toc')
    let text = el.textContent || ''
    const contentArea = this.getContentArea()
    
    // 获取对齐方式
    const computedStyle = el.getAttribute('style') || ''
    let alignment: 'left' | 'center' | 'right' | 'justify' = 'left'
    if (computedStyle.includes('text-align: center') || computedStyle.includes('text-align:center')) {
      alignment = 'center'
    } else if (computedStyle.includes('text-align: right') || computedStyle.includes('text-align:right')) {
      alignment = 'right'
    } else if (computedStyle.includes('text-align: justify')) {
      alignment = 'justify'
    }
    
    // 首行缩进
    let firstLineIndent = 0
    if (computedStyle.includes('text-indent')) {
      const match = computedStyle.match(/text-indent:\s*(\d+(?:\.\d+)?)(px|pt|em)/)
      if (match) {
        const value = parseFloat(match[1])
        const unit = match[2]
        if (unit === 'px') firstLineIndent = value * this.scale
        else if (unit === 'pt') firstLineIndent = value * PT_TO_PX * this.scale
        else if (unit === 'em') firstLineIndent = value * (style.fontSize || 10.5) * PT_TO_PX * this.scale
      }
    }
    
    // 测量文本
    const textStyle: TextStyle = {
      fontFamily: style.fontFamily || 'DengXian, "Microsoft YaHei", "SimSun", serif',
      fontSize: style.fontSize || 10.5,
      fontWeight: style.fontWeight,
      fontStyle: style.fontStyle,
      lineHeight: style.lineHeight || DEFAULT_LINE_HEIGHT
    }
    
    if (isToc) {
      const leftText = el.querySelector('.docx-toc-left')?.textContent || ''
      const rightText = el.querySelector('.docx-toc-right')?.textContent || ''
      if (leftText || rightText) {
        const dotWidth = this.measurer.measureText('.', textStyle)
        const leftWidth = this.measurer.measureText(leftText, textStyle)
        const rightWidth = this.measurer.measureText(rightText, textStyle)
        const leaderWidth = Math.max(0, contentArea.width - firstLineIndent - leftWidth - rightWidth)
        const dotCount = dotWidth > 0 ? Math.max(2, Math.floor(leaderWidth / dotWidth)) : 2
        const dots = '.'.repeat(dotCount)
        text = `${leftText}${dots}${rightText}`
      }
    }

    const measured = this.measurer.measureParagraph(
      text,
      contentArea.width - firstLineIndent,
      textStyle
    )
    
    // 计算段落间距
    const marginTop = style.marginTop || 0
    const marginBottom = style.marginBottom || (8 * PT_TO_PX * this.scale)
    
    return {
      type: 'paragraph',
      x: 0,
      y: 0,
      width: contentArea.width,
      height: measured.height + marginTop + marginBottom,
      pageIndex: 0,
      alignment,
      firstLineIndent,
      style,
      children: [{
        type: 'text',
        x: firstLineIndent,
        y: marginTop,
        width: measured.width,
        height: measured.height,
        pageIndex: 0,
        text,
        measuredText: measured,
        style
      }]
    }
  }
  
  /**
   * 解析标题
   */
  private parseHeading(el: HTMLElement, tagName: string): ParagraphLayoutElement {
    const paragraph = this.parseParagraph(el)
    paragraph.type = 'heading'
    
    // 根据标题级别设置字体大小
    const level = parseInt(tagName.substring(1))
    const baseFontSize = 10.5
    const fontSizes: Record<number, number> = {
      1: 20,
      2: 16,
      3: 14,
      4: 12,
      5: 10.5,
      6: 10
    }
    
    const fontSize = fontSizes[level] || baseFontSize
    
    const spacingBefore = level === 1 ? 18 : 8
    const spacingAfter = 4

    if (paragraph.style) {
      paragraph.style.fontSize = fontSize
      paragraph.style.fontWeight = 'bold'
      paragraph.style.marginTop = spacingBefore * PT_TO_PX * this.scale
      paragraph.style.marginBottom = spacingAfter * PT_TO_PX * this.scale
    }
    
    // 重新测量
    const contentArea = this.getContentArea()
    const textStyle: TextStyle = {
      fontFamily: paragraph.style?.fontFamily || 'DengXian, "Microsoft YaHei", "SimSun", serif',
      fontSize,
      fontWeight: 'bold',
      lineHeight: DEFAULT_LINE_HEIGHT
    }
    
    const text = el.textContent || ''
    const measured = this.measurer.measureParagraph(text, contentArea.width, textStyle)
    
    paragraph.height = measured.height + (paragraph.style?.marginTop || 0) + (paragraph.style?.marginBottom || 0)
    
    if (paragraph.children[0]) {
      paragraph.children[0].height = measured.height
      paragraph.children[0].measuredText = measured
    }
    
    return paragraph
  }
  
  /**
   * 解析图片
   */
  private parseImage(el: HTMLImageElement): ImageLayoutElement {
    const src = el.src
    const alt = el.alt
    
    // 获取图片尺寸
    let width = el.width || 200
    let height = el.height || 150
    
    // �?style 获取尺寸
    const styleWidth = el.style.width
    const styleHeight = el.style.height
    
    if (styleWidth) {
      const match = styleWidth.match(/(\d+(?:\.\d+)?)(px|pt|%)/)
      if (match) {
        const value = parseFloat(match[1])
        const unit = match[2]
        if (unit === 'px') width = value * this.scale
        else if (unit === 'pt') width = value * PT_TO_PX * this.scale
        else if (unit === '%') {
          const contentArea = this.getContentArea()
          width = contentArea.width * value / 100
        }
      }
    }
    
    if (styleHeight) {
      const match = styleHeight.match(/(\d+(?:\.\d+)?)(px|pt|%)/)
      if (match) {
        const value = parseFloat(match[1])
        const unit = match[2]
        if (unit === 'px') height = value * this.scale
        else if (unit === 'pt') height = value * PT_TO_PX * this.scale
      }
    }
    
    // 限制图片最大宽�?
    const contentArea = this.getContentArea()
    if (width > contentArea.width) {
      const ratio = contentArea.width / width
      width = contentArea.width
      height = height * ratio
    }
    
    return {
      type: 'image',
      x: 0,
      y: 0,
      width,
      height: height + 10 * this.scale, // 添加图片间距
      pageIndex: 0,
      src,
      alt
    }
  }
  
  /**
   * 解析表格
   */
  private parseTable(el: HTMLTableElement): TableLayoutElement {
    const contentArea = this.getContentArea()
    const rows: TableRowLayoutElement[] = []
    let tableHeight = 0
    const scale = this.scale
    const styleAttr = el.getAttribute('style') || ''
    const parseSize = (raw: string | undefined, base: number): number | undefined => {
      if (!raw) return undefined
      const trimmed = raw.trim()
      if (trimmed.endsWith('%')) {
        const pct = Number.parseFloat(trimmed)
        return Number.isFinite(pct) ? (base * pct) / 100 : undefined
      }
      if (trimmed.endsWith('pt')) {
        const pt = Number.parseFloat(trimmed)
        return Number.isFinite(pt) ? pt * PT_TO_PX * scale : undefined
      }
      if (trimmed.endsWith('px')) {
        const px = Number.parseFloat(trimmed)
        return Number.isFinite(px) ? px * scale : undefined
      }
      const num = Number.parseFloat(trimmed)
      return Number.isFinite(num) ? num : undefined
    }
    const widthMatch = styleAttr.match(/width:\s*([^;]+)/)
    const tableWidth = parseSize(widthMatch?.[1], contentArea.width) || contentArea.width
    const gridAttr = el.getAttribute('data-tbl-grid')
    const gridTotalAttr = el.getAttribute('data-tbl-grid-total')
    const gridTwips = gridAttr
      ? gridAttr
          .split(',')
          .map((v) => Number.parseFloat(v))
          .filter((v) => Number.isFinite(v) && v > 0)
      : []
    const gridTotal = gridTotalAttr ? Number.parseFloat(gridTotalAttr) : gridTwips.reduce((sum, v) => sum + v, 0)
    let columnWidths: number[] = []
    if (gridTwips.length && gridTotal > 0) {
      columnWidths = gridTwips.map((twips) => (twips / gridTotal) * tableWidth)
    } else {
      const cols = el.querySelectorAll('colgroup col')
      if (cols.length) {
        columnWidths = Array.from(cols).map((col) => {
          const colStyle = (col as HTMLElement).getAttribute('style') || ''
          const colWidthMatch = colStyle.match(/width:\s*([^;]+)/)
          return parseSize(colWidthMatch?.[1], tableWidth) || 0
        })
      }
    }

    const tableRows = el.querySelectorAll('tr')

    for (const tr of Array.from(tableRows)) {
      const cells: TableCellLayoutElement[] = []
      let rowHeight = 0
      const trStyle = (tr as HTMLElement).getAttribute('style') || ''
      const heightMatch = trStyle.match(/height:\s*([^;]+)/)
      const explicitRowHeight = parseSize(heightMatch?.[1], contentArea.height)
      const tds = tr.querySelectorAll('td, th')
      const fallbackCellWidth = tableWidth / Math.max(tds.length, 1)
      let colIndex = 0

      for (const td of Array.from(tds)) {
        const text = td.textContent || ''
        const cellStyle = this.extractStyle(td as HTMLElement)
        const textStyle: TextStyle = {
          fontFamily: cellStyle.fontFamily || 'DengXian, "Microsoft YaHei", "SimSun", serif',
          fontSize: cellStyle.fontSize || 10.5,
          fontWeight: cellStyle.fontWeight,
          fontStyle: cellStyle.fontStyle,
          lineHeight: cellStyle.lineHeight || DEFAULT_LINE_HEIGHT
        }

        const parsePadding = (attr: string) => {
          const match = (td as HTMLElement).getAttribute('style')?.match(new RegExp(`${attr}:\\s*([^;]+)`, 'i'))
          return parseSize(match?.[1], tableWidth)
        }
        const paddingLeft = parsePadding('padding-left')
        const paddingRight = parsePadding('padding-right')
        const paddingTop = parsePadding('padding-top')
        const paddingBottom = parsePadding('padding-bottom')
        const paddingX = (paddingLeft ?? 5 * PT_TO_PX * scale) + (paddingRight ?? 5 * PT_TO_PX * scale)
        const paddingY = (paddingTop ?? 2 * PT_TO_PX * scale) + (paddingBottom ?? 2 * PT_TO_PX * scale)

        const colspan = parseInt((td as HTMLTableCellElement).getAttribute('colspan') || '1')
        const colSpan = Number.isFinite(colspan) && colspan > 0 ? colspan : 1
        const cellWidth = columnWidths.length
          ? columnWidths.slice(colIndex, colIndex + colSpan).reduce((sum, w) => sum + w, 0) || fallbackCellWidth
          : fallbackCellWidth

        const measured = this.measurer.measureParagraph(text, Math.max(0, cellWidth - paddingX), textStyle)
        const cellHeight = measured.height + paddingY * 2

        if (cellHeight > rowHeight) rowHeight = cellHeight

        cells.push({
          type: 'tableCell',
          x: 0,
          y: 0,
          width: cellWidth,
          height: cellHeight,
          pageIndex: 0,
          colspan: colSpan,
          rowspan: parseInt((td as HTMLTableCellElement).getAttribute('rowspan') || '1'),
          style: cellStyle,
          children: [{
            type: 'text',
            x: paddingLeft ?? 5 * PT_TO_PX * scale,
            y: paddingTop ?? 2 * PT_TO_PX * scale,
            width: measured.width,
            height: measured.height,
            pageIndex: 0,
            text,
            measuredText: measured,
            style: cellStyle
          } as TextLayoutElement]
        })
        colIndex += colSpan
      }

      if (explicitRowHeight && explicitRowHeight > rowHeight) {
        rowHeight = explicitRowHeight
      }

      for (const cell of cells) {
        cell.height = rowHeight
      }

      rows.push({
        type: 'tableRow',
        x: 0,
        y: tableHeight,
        width: tableWidth,
        height: rowHeight,
        pageIndex: 0,
        cells
      })

      tableHeight += rowHeight
    }

    return {
      type: 'table',
      x: 0,
      y: 0,
      width: tableWidth,
      height: tableHeight + 20 * this.scale, // 表格间距
      pageIndex: 0,
      columns: rows[0]?.cells.length || 0,
      rows
    }
  }
  
  /**
   * 解析列表
   */
  private parseList(el: HTMLElement): LayoutElement {
    const items = el.querySelectorAll('li')
    const children: LayoutElement[] = []
    let totalHeight = 0
    const isOrdered = el.tagName.toLowerCase() === 'ol'
    
    let index = 1
    for (const li of Array.from(items)) {
      const prefix = isOrdered ? `${index}. ` : '�?'
      const text = prefix + (li.textContent || '')
      
      const para = this.parseParagraph(li)
      if (para.children[0]) {
        para.children[0].text = text
        // 重新测量
        const contentArea = this.getContentArea()
        const textStyle: TextStyle = {
          fontFamily: para.style?.fontFamily || 'DengXian, "Microsoft YaHei", "SimSun", serif',
          fontSize: para.style?.fontSize || 10.5,
          lineHeight: para.style?.lineHeight || DEFAULT_LINE_HEIGHT
        }
        para.children[0].measuredText = this.measurer.measureParagraph(text, contentArea.width, textStyle)
        para.height = para.children[0].measuredText.height + 8 * this.scale
      }
      
      para.y = totalHeight
      totalHeight += para.height
      children.push(para)
      index++
    }
    
    return {
      type: 'list',
      x: 0,
      y: 0,
      width: this.getContentArea().width,
      height: totalHeight,
      pageIndex: 0,
      children
    }
  }
  
  /**
   * 提取元素样式
   */
  private extractStyle(el: HTMLElement): ElementStyle {
    const style: ElementStyle = {}
    const styleAttr = el.getAttribute('style') || ''
    
    // 解析 font-family
    const fontFamilyMatch = styleAttr.match(/font-family:\s*([^;]+)/)
    if (fontFamilyMatch) {
      style.fontFamily = fontFamilyMatch[1].trim().replace(/['"]/g, '')
    }
    
    // 解析 font-size
    const fontSizeMatch = styleAttr.match(/font-size:\s*(\d+(?:\.\d+)?)(px|pt|em)/)
    if (fontSizeMatch) {
      const value = parseFloat(fontSizeMatch[1])
      const unit = fontSizeMatch[2]
      if (unit === 'pt') style.fontSize = value
      else if (unit === 'px') style.fontSize = value / PT_TO_PX
      else if (unit === 'em') style.fontSize = value * 12
    }
    
    // 解析 font-weight
    if (styleAttr.includes('font-weight: bold') || styleAttr.includes('font-weight:bold')) {
      style.fontWeight = 'bold'
    }
    
    // 解析 font-style
    if (styleAttr.includes('font-style: italic') || styleAttr.includes('font-style:italic')) {
      style.fontStyle = 'italic'
    }
    
    // 解析 color
    const colorMatch = styleAttr.match(/(?:^|;)\s*color:\s*([^;]+)/)
    if (colorMatch) {
      style.color = colorMatch[1].trim()
    }
    
    // 解析 background-color
    const bgColorMatch = styleAttr.match(/background-color:\s*([^;]+)/)
    if (bgColorMatch) {
      style.backgroundColor = bgColorMatch[1].trim()
    }
    
    // 解析 text-decoration
    if (styleAttr.includes('text-decoration: underline') || styleAttr.includes('text-decoration:underline')) {
      style.textDecoration = 'underline'
    } else if (styleAttr.includes('text-decoration: line-through')) {
      style.textDecoration = 'line-through'
    }

    const parseBorderSide = (prop: string) => {
      const match = styleAttr.match(new RegExp(`${prop}:\\s*([\\d.]+)(px|pt)\\s+[^;]*\\s+([^;]+)`, 'i'))
      if (!match) return undefined
      const value = parseFloat(match[1])
      const unit = match[2]
      const color = match[3]
      const widthPx = unit === 'pt' ? value * PT_TO_PX * this.scale : value * this.scale
      return {
        width: Number.isFinite(widthPx) ? widthPx : undefined,
        color: color ? color.trim() : undefined,
      }
    }

    const topBorder = parseBorderSide('border-top')
    const rightBorder = parseBorderSide('border-right')
    const bottomBorder = parseBorderSide('border-bottom')
    const leftBorder = parseBorderSide('border-left')
    if (topBorder) {
      style.borderTopWidth = topBorder.width
      style.borderTopColor = topBorder.color
    }
    if (rightBorder) {
      style.borderRightWidth = rightBorder.width
      style.borderRightColor = rightBorder.color
    }
    if (bottomBorder) {
      style.borderBottomWidth = bottomBorder.width
      style.borderBottomColor = bottomBorder.color
    }
    if (leftBorder) {
      style.borderLeftWidth = leftBorder.width
      style.borderLeftColor = leftBorder.color
    }

    // 解析 border
    const borderMatch = styleAttr.match(/border:\s*([\d.]+)(px|pt)\s+[^;]*\s+([^;]+)/i)
    if (borderMatch) {
      const value = parseFloat(borderMatch[1])
      const unit = borderMatch[2]
      const color = borderMatch[3]
      const widthPx = unit === 'pt' ? value * PT_TO_PX * this.scale : value * this.scale
      if (Number.isFinite(widthPx)) style.borderWidth = widthPx
      if (color) style.borderColor = color.trim()
      if (style.borderTopWidth == null) style.borderTopWidth = style.borderWidth
      if (style.borderRightWidth == null) style.borderRightWidth = style.borderWidth
      if (style.borderBottomWidth == null) style.borderBottomWidth = style.borderWidth
      if (style.borderLeftWidth == null) style.borderLeftWidth = style.borderWidth
      if (style.borderTopColor == null) style.borderTopColor = style.borderColor
      if (style.borderRightColor == null) style.borderRightColor = style.borderColor
      if (style.borderBottomColor == null) style.borderBottomColor = style.borderColor
      if (style.borderLeftColor == null) style.borderLeftColor = style.borderColor
    }
    
    // 解析 line-height
    const lineHeightMatch = styleAttr.match(/line-height:\s*(\d+(?:\.\d+)?)(px|pt)?/)
    if (lineHeightMatch) {
      const value = parseFloat(lineHeightMatch[1])
      const unit = lineHeightMatch[2]
      const fontSize = style.fontSize || 10.5
      if (unit === 'pt') {
        style.lineHeight = fontSize > 0 ? value / fontSize : DEFAULT_LINE_HEIGHT
      } else if (unit === 'px') {
        const pt = value / PT_TO_PX
        style.lineHeight = fontSize > 0 ? pt / fontSize : DEFAULT_LINE_HEIGHT
      } else {
        style.lineHeight = value
      }
    }
    
    return style
  }
  
  /**
   * 分页算法
   */
  private paginateElements(elements: LayoutElement[]): PageLayout[] {
    const pages: PageLayout[] = []
    const contentArea = this.getContentArea()
    const maxHeight = contentArea.height
    
    let currentPage: PageLayout = {
      pageIndex: 0,
      elements: []
    }
    let currentY = 0
    
    for (const element of elements) {
      // 处理分页�?
      if (element.type === 'pageBreak') {
        pages.push(currentPage)
        currentPage = {
          pageIndex: pages.length,
          elements: []
        }
        currentY = 0
        continue
      }
      
      // 检查是否需要换�?
      if (currentY + element.height > maxHeight && currentPage.elements.length > 0) {
        // 当前页放不下，换�?
        pages.push(currentPage)
        currentPage = {
          pageIndex: pages.length,
          elements: []
        }
        currentY = 0
      }
      
      // 设置元素位置
      element.x = contentArea.x
      element.y = contentArea.y + currentY
      element.pageIndex = currentPage.pageIndex
      
      // 更新子元素位�?
      this.updateChildPositions(element, element.x, element.y, currentPage.pageIndex)
      
      currentPage.elements.push(element)
      currentY += element.height
    }
    
    // 保存最后一�?
    if (currentPage.elements.length > 0 || pages.length === 0) {
      pages.push(currentPage)
    }
    
    return pages
  }
  
  /**
   * 更新子元素位�?
   */
  private updateChildPositions(element: LayoutElement, parentX: number, parentY: number, pageIndex: number): void {
    if (!element.children) return
    
    for (const child of element.children) {
      child.x += parentX
      child.y += parentY
      child.pageIndex = pageIndex
      
      this.updateChildPositions(child, child.x, child.y, pageIndex)
    }
  }
  
  /**
   * 添加页眉页脚
   */
  private addHeaderFooter(pages: PageLayout[], headerHtml?: string, footerHtml?: string): void {
    const { marginLeft, marginTop, marginBottom, width, height, headerHeight, footerHeight } = this.pageConfig
    const contentWidth = width - marginLeft * 2
    
    for (let i = 0; i < pages.length; i++) {
      const page = pages[i]
      const pageNumber = i + 1
      const totalPages = pages.length
      
      // 处理页眉
      if (headerHtml) {
        const headerText = headerHtml
          .replace(/<[^>]+>/g, '')  // 移除 HTML 标签
          .replace(/\{PAGE\}/g, String(pageNumber))
          .replace(/\{NUMPAGES\}/g, String(totalPages))
        
        const textStyle: TextStyle = {
          fontFamily: 'DengXian, "Microsoft YaHei", "SimSun", serif',
          fontSize: 9,
          color: 'var(--word-ink-muted)',
          lineHeight: 1
        }
        
        const measured = this.measurer.measureParagraph(headerText, contentWidth, textStyle)
        
        page.header = {
          type: 'header',
          x: marginLeft,
          y: marginTop - headerHeight,
          width: contentWidth,
          height: headerHeight,
          pageIndex: i,
          children: [{
            type: 'text',
            x: marginLeft,
            y: marginTop - headerHeight,
            width: measured.width,
            height: measured.height,
            pageIndex: i,
            text: headerText,
            measuredText: measured,
            style: {
              fontSize: 9,
              color: 'var(--word-ink-muted)',
              borderStyle: 'solid',
              borderWidth: 0.5,
              borderColor: 'var(--word-rule)'
            }
          } as TextLayoutElement]
        }
      }
      
      // 处理页脚
      if (footerHtml) {
        const footerText = footerHtml
          .replace(/<[^>]+>/g, '')
          .replace(/\{PAGE\}/g, String(pageNumber))
          .replace(/\{NUMPAGES\}/g, String(totalPages))
        
        const textStyle: TextStyle = {
          fontFamily: 'DengXian, "Microsoft YaHei", "SimSun", serif',
          fontSize: 9,
          color: 'var(--word-ink-muted)',
          lineHeight: 1
        }
        
        const measured = this.measurer.measureParagraph(footerText, contentWidth, textStyle)
        
        page.footer = {
          type: 'footer',
          x: marginLeft,
          y: height - marginBottom,
          width: contentWidth,
          height: footerHeight,
          pageIndex: i,
          children: [{
            type: 'text',
            x: marginLeft,
            y: height - marginBottom,
            width: measured.width,
            height: measured.height,
            pageIndex: i,
            text: footerText,
            measuredText: measured,
            style: {
              fontSize: 9,
              color: 'var(--word-ink-muted)'
            }
          } as TextLayoutElement]
        }
      } else {
        // 默认页码
        const footerText = `${pageNumber}`
        const textStyle: TextStyle = {
          fontFamily: 'DengXian, "Microsoft YaHei", "SimSun", serif',
          fontSize: 9,
          color: 'var(--word-ink-muted)',
          lineHeight: 1
        }
        
        const measured = this.measurer.measureParagraph(footerText, contentWidth, textStyle)
        
        page.footer = {
          type: 'footer',
          x: marginLeft,
          y: height - marginBottom,
          width: contentWidth,
          height: footerHeight,
          pageIndex: i,
          children: [{
            type: 'text',
            x: marginLeft + (contentWidth - measured.width) / 2,  // 居中
            y: height - marginBottom,
            width: measured.width,
            height: measured.height,
            pageIndex: i,
            text: footerText,
            measuredText: measured,
            style: {
              fontSize: 9,
              color: 'var(--word-ink-muted)'
            }
          } as TextLayoutElement]
        }
      }
    }
  }
  
  /**
   * 获取页面配置
   */
  getPageConfig(): PageConfig {
    return this.pageConfig
  }
  
  /**
   * 更新页面配置
   */
  setPageConfig(config: Partial<PageConfig>): void {
    this.pageConfig = { ...this.pageConfig, ...config }
  }
}

/**
 * 创建布局引擎实例
 */
export function createLayoutEngine(pageConfig?: PageConfig, scale: number = 1): LayoutEngine {
  return new LayoutEngine(pageConfig, scale)
}














