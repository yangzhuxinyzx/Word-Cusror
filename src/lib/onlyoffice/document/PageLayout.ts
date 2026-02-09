/**
 * 页面布局引擎
 * 
 * 参考 ONLYOFFICE 的设计:
 * - DocumentServer/sdkjs-src/word/Editor/Paragraph_Recalculate.js
 * - DocumentServer/sdkjs-src/word/Editor/Document.js (Recalculate_Page)
 * 
 * 功能:
 * - 计算元素位置和尺寸
 * - 自动分页
 * - 页眉页脚定位
 * - 处理分页符
 */

import { TextMeasurer, getTextMeasurer } from '../core/TextMeasurer'
import { MM_TO_PX, PT_TO_PX } from '../core/Constants'
import {
  DocumentElement,
  PageSettings,
  HeaderFooter,
  LayoutElement,
  LayoutPage,
  ParagraphElement,
  ImageElement,
  TableElement,
  ListElement,
  TextRun,
  getDefaultPageSettings,
  getDefaultTextStyle
} from './DocumentModel'

// ============== 类型定义 ==============

export interface LayoutOptions {
  /** 页面设置 */
  pageSettings?: PageSettings
  /** 是否启用分页 */
  enablePagination?: boolean
  /** 缩放比例 */
  scale?: number
}

export interface LayoutResult {
  pages: LayoutPage[]
  totalHeight: number  // mm
}

// ============== PageLayoutEngine 类 ==============

export class PageLayoutEngine {
  private textMeasurer: TextMeasurer
  private pageSettings: PageSettings
  private scale: number
  
  constructor(options: LayoutOptions = {}) {
    this.textMeasurer = getTextMeasurer()
    this.pageSettings = options.pageSettings || getDefaultPageSettings()
    this.scale = options.scale || 1
  }
  
  // ============== 主布局方法 ==============
  
  /**
   * 对文档内容进行布局
   * 参考: CDocument.prototype.Recalculate_Page
   */
  layout(
    content: DocumentElement[],
    header?: HeaderFooter,
    footer?: HeaderFooter
  ): LayoutResult {
    const pages: LayoutPage[] = []
    const contentArea = this.getContentArea()
    
    let currentPage = this.createNewPage(0, header, footer)
    let currentY = 0  // 当前 Y 位置（相对于内容区域）
    
    for (let i = 0; i < content.length; i++) {
      const element = content[i]
      
      // 处理分页符
      if (element.type === 'pageBreak') {
        pages.push(currentPage)
        currentPage = this.createNewPage(pages.length, header, footer)
        currentY = 0
        continue
      }
      
      // 计算元素尺寸
      const elementSize = this.measureElement(element, contentArea.width)
      
      // 检查是否需要分页
      if (currentY + elementSize.height > contentArea.height) {
        // 当前页放不下，需要分页
        pages.push(currentPage)
        currentPage = this.createNewPage(pages.length, header, footer)
        currentY = 0
      }
      
      // 添加元素到当前页
      const layoutElement: LayoutElement = {
        element,
        pageIndex: pages.length,
        x: contentArea.x + (elementSize.marginLeft || 0),
        y: contentArea.y + currentY + (elementSize.marginTop || 0),
        width: elementSize.width,
        height: elementSize.height
      }
      
      currentPage.elements.push(layoutElement)
      currentY += elementSize.height + (elementSize.marginTop || 0) + (elementSize.marginBottom || 0)
    }
    
    // 添加最后一页
    if (currentPage.elements.length > 0 || pages.length === 0) {
      pages.push(currentPage)
    }
    
    // 计算总高度
    const totalHeight = pages.length * this.pageSettings.height
    
    return { pages, totalHeight }
  }
  
  /**
   * 从 HTML 解析并布局
   */
  layoutFromHtml(
    html: string,
    header?: HeaderFooter,
    footer?: HeaderFooter
  ): LayoutResult {
    const content = this.parseHtmlToElements(html)
    return this.layout(content, header, footer)
  }
  
  // ============== 辅助方法 ==============
  
  /**
   * 获取内容区域
   */
  private getContentArea(): { x: number; y: number; width: number; height: number } {
    const { width, height, margins, headerHeight, footerHeight } = this.pageSettings
    
    return {
      x: margins.left,
      y: margins.top + headerHeight,
      width: width - margins.left - margins.right,
      height: height - margins.top - margins.bottom - headerHeight - footerHeight
    }
  }
  
  /**
   * 创建新页面
   */
  private createNewPage(pageIndex: number, header?: HeaderFooter, footer?: HeaderFooter): LayoutPage {
    return {
      pageIndex,
      width: this.pageSettings.width,
      height: this.pageSettings.height,
      elements: [],
      header,
      footer
    }
  }
  
  /**
   * 测量元素尺寸
   */
  private measureElement(
    element: DocumentElement,
    maxWidth: number
  ): { width: number; height: number; marginTop?: number; marginBottom?: number; marginLeft?: number } {
    switch (element.type) {
      case 'paragraph':
      case 'heading':
        return this.measureParagraph(element as ParagraphElement, maxWidth)
      
      case 'image':
        return this.measureImage(element as ImageElement, maxWidth)
      
      case 'table':
        return this.measureTable(element as TableElement, maxWidth)
      
      case 'list':
        return this.measureList(element as ListElement, maxWidth)
      
      case 'lineBreak':
        return { width: maxWidth, height: 5 }  // 换行符高度
      
      default:
        return { width: maxWidth, height: 10 }
    }
  }
  
  /**
   * 测量段落
   */
  private measureParagraph(
    paragraph: ParagraphElement,
    maxWidth: number
  ): { width: number; height: number; marginTop?: number; marginBottom?: number; marginLeft?: number } {
    const style = paragraph.style || {}
    const textStyle = paragraph.textStyle || getDefaultTextStyle()
    
    // 计算可用宽度
    const availableWidth = maxWidth - (style.marginLeft || 0) - (style.marginRight || 0)
    const availableWidthPx = availableWidth * MM_TO_PX * this.scale
    
    // 测量所有文本运行
    let totalHeight = 0
    let currentLineWidth = style.firstLineIndent ? (style.firstLineIndent * MM_TO_PX * this.scale) : 0
    let lineCount = 1
    
    const lineHeightPx = (textStyle.fontSize || 10.5) * PT_TO_PX * (textStyle.lineHeight || 1.15) * this.scale
    
    for (const run of paragraph.children) {
      if (run.type === 'text') {
        const runStyle = run.style || textStyle
        
        this.textMeasurer.setFont({
          fontFamily: runStyle.fontFamily,
          fontSize: runStyle.fontSize,
          fontWeight: runStyle.fontWeight,
          fontStyle: runStyle.fontStyle,
          lineHeight: runStyle.lineHeight
        })
        
        // 测量文本
        const measured = this.textMeasurer.measureText(run.text, {
          maxWidth: availableWidthPx,
          wordWrap: true
        })
        
        // 累计行数
        lineCount += measured.lines.length - 1
        
        if (measured.lines.length > 0) {
          currentLineWidth = measured.lines[measured.lines.length - 1].width
        }
      }
    }
    
    totalHeight = lineCount * (lineHeightPx / MM_TO_PX / this.scale)
    
    return {
      width: availableWidth,
      height: totalHeight,
      marginTop: style.marginTop || 0,
      marginBottom: style.marginBottom || 0,
      marginLeft: style.marginLeft || 0
    }
  }
  
  /**
   * 测量图片
   */
  private measureImage(
    image: ImageElement,
    maxWidth: number
  ): { width: number; height: number } {
    let width = image.width || 100
    let height = image.height || 75
    
    // 如果图片太宽，按比例缩小
    if (width > maxWidth) {
      const ratio = maxWidth / width
      width = maxWidth
      height = height * ratio
    }
    
    return { width, height }
  }
  
  /**
   * 测量表格
   */
  private measureTable(
    table: TableElement,
    maxWidth: number
  ): { width: number; height: number } {
    const tableWidth = table.width || maxWidth
    let tableHeight = 0
    
    for (const row of table.rows) {
      const rowHeight = row.height || 10  // 默认行高 10mm
      tableHeight += rowHeight
    }
    
    return { width: Math.min(tableWidth, maxWidth), height: tableHeight }
  }
  
  /**
   * 测量列表
   */
  private measureList(
    list: ListElement,
    maxWidth: number
  ): { width: number; height: number } {
    let totalHeight = 0
    
    for (const item of list.items) {
      // 递归测量列表项内容
      for (const child of item.children) {
        const size = this.measureElement(child, maxWidth - 10)  // 减去缩进
        totalHeight += size.height + (size.marginTop || 0) + (size.marginBottom || 0)
      }
    }
    
    return { width: maxWidth, height: totalHeight }
  }
  
  // ============== HTML 解析 ==============
  
  /**
   * 将 HTML 解析为文档元素
   * 简化版解析器，用于演示
   */
  private parseHtmlToElements(html: string): DocumentElement[] {
    const elements: DocumentElement[] = []
    
    // 创建临时 DOM 解析
    const parser = new DOMParser()
    const doc = parser.parseFromString(html, 'text/html')
    const body = doc.body
    
    // 递归解析子节点
    this.parseNodeChildren(body, elements)
    
    return elements
  }
  
  /**
   * 解析节点的子元素
   */
  private parseNodeChildren(parent: Element, elements: DocumentElement[]): void {
    for (let i = 0; i < parent.children.length; i++) {
      const child = parent.children[i]
      const element = this.parseNode(child)
      if (element) {
        elements.push(element)
      }
    }
  }
  
  /**
   * 解析单个 DOM 节点
   */
  private parseNode(node: Element): DocumentElement | null {
    const tagName = node.tagName.toLowerCase()
    
    switch (tagName) {
      case 'p':
        return this.parseParagraph(node)
      
      case 'h1':
      case 'h2':
      case 'h3':
      case 'h4':
      case 'h5':
      case 'h6':
        return this.parseHeading(node, parseInt(tagName[1]))
      
      case 'img':
        return this.parseImage(node as HTMLImageElement)
      
      case 'table':
        return this.parseTable(node)
      
      case 'ul':
      case 'ol':
        return this.parseList(node, tagName === 'ol')
      
      case 'hr':
        return { type: 'pageBreak' }
      
      case 'br':
        return { type: 'lineBreak' }
      
      case 'div':
      case 'section':
      case 'article':
        // 对于容器元素，递归解析其内容
        // 但我们需要返回单个元素，所以转换为段落
        return this.parseParagraph(node)
      
      default:
        return null
    }
  }
  
  /**
   * 解析段落
   */
  private parseParagraph(node: Element): ParagraphElement {
    const children: TextRun[] = []
    this.parseTextContent(node, children)
    
    // 解析样式
    const style = this.parseElementStyle(node)
    
    return {
      type: 'paragraph',
      children,
      style: style.paragraph,
      textStyle: style.text
    }
  }
  
  /**
   * 解析标题
   */
  private parseHeading(node: Element, level: number): ParagraphElement {
    const children: TextRun[] = []
    this.parseTextContent(node, children)
    
    // 标题默认加粗
    const defaultTextStyle = getDefaultTextStyle()
    defaultTextStyle.fontWeight = 'bold'
    defaultTextStyle.fontSize = this.getHeadingFontSize(level)
    
    const style = this.parseElementStyle(node)
    
    return {
      type: 'heading',
      level,
      children,
      style: style.paragraph,
      textStyle: { ...defaultTextStyle, ...style.text }
    }
  }
  
  /**
   * 获取标题字号
   */
  private getHeadingFontSize(level: number): number {
    const sizes = {
      1: 22,
      2: 18,
      3: 16,
      4: 14,
      5: 12,
      6: 10.5
    }
    return sizes[level as keyof typeof sizes] || 10.5
  }
  
  /**
   * 解析文本内容
   */
  private parseTextContent(node: Element, runs: TextRun[]): void {
    for (let i = 0; i < node.childNodes.length; i++) {
      const child = node.childNodes[i]
      
      if (child.nodeType === Node.TEXT_NODE) {
        const text = child.textContent?.trim()
        if (text) {
          runs.push({ type: 'text', text })
        }
      } else if (child.nodeType === Node.ELEMENT_NODE) {
        const element = child as Element
        const tagName = element.tagName.toLowerCase()
        
        // 处理内联格式
        const style = this.parseInlineStyle(element)
        const text = element.textContent?.trim()
        
        if (text) {
          runs.push({ type: 'text', text, style })
        }
      }
    }
  }
  
  /**
   * 解析内联样式
   */
  private parseInlineStyle(element: Element): any {
    const tagName = element.tagName.toLowerCase()
    const style: any = {}
    
    if (tagName === 'strong' || tagName === 'b') {
      style.fontWeight = 'bold'
    }
    if (tagName === 'em' || tagName === 'i') {
      style.fontStyle = 'italic'
    }
    if (tagName === 'u') {
      style.textDecoration = 'underline'
    }
    if (tagName === 's' || tagName === 'strike' || tagName === 'del') {
      style.textDecoration = 'line-through'
    }
    
    // 解析 style 属性
    const inlineStyle = element.getAttribute('style')
    if (inlineStyle) {
      const styleProps = this.parseStyleAttribute(inlineStyle)
      Object.assign(style, styleProps)
    }
    
    return style
  }
  
  /**
   * 解析元素样式
   */
  private parseElementStyle(element: Element): { paragraph: any; text: any } {
    const paragraph: any = {}
    const text: any = {}
    
    const inlineStyle = element.getAttribute('style')
    if (inlineStyle) {
      const props = this.parseStyleAttribute(inlineStyle)
      
      // 分配到段落或文本样式
      if (props.textAlign) paragraph.alignment = props.textAlign
      if (props.textIndent) paragraph.firstLineIndent = parseFloat(props.textIndent) / MM_TO_PX
      if (props.marginTop) paragraph.marginTop = parseFloat(props.marginTop) / MM_TO_PX
      if (props.marginBottom) paragraph.marginBottom = parseFloat(props.marginBottom) / MM_TO_PX
      if (props.fontFamily) text.fontFamily = props.fontFamily
      if (props.fontSize) text.fontSize = parseFloat(props.fontSize)
      if (props.fontWeight) text.fontWeight = props.fontWeight
      if (props.color) text.color = props.color
    }
    
    return { paragraph, text }
  }
  
  /**
   * 解析 style 属性
   */
  private parseStyleAttribute(styleStr: string): Record<string, string> {
    const result: Record<string, string> = {}
    
    const parts = styleStr.split(';')
    for (const part of parts) {
      const [key, value] = part.split(':').map(s => s.trim())
      if (key && value) {
        // 转换为 camelCase
        const camelKey = key.replace(/-([a-z])/g, (_, c) => c.toUpperCase())
        result[camelKey] = value
      }
    }
    
    return result
  }
  
  /**
   * 解析图片
   */
  private parseImage(img: HTMLImageElement): ImageElement {
    return {
      type: 'image',
      src: img.src,
      alt: img.alt,
      width: img.width ? img.width / MM_TO_PX : undefined,
      height: img.height ? img.height / MM_TO_PX : undefined
    }
  }
  
  /**
   * 解析表格
   */
  private parseTable(table: Element): TableElement {
    const rows: any[] = []
    const tableRows = table.querySelectorAll('tr')
    
    tableRows.forEach(tr => {
      const cells: any[] = []
      const tableCells = tr.querySelectorAll('td, th')
      
      tableCells.forEach(td => {
        const cellChildren: DocumentElement[] = []
        this.parseNodeChildren(td, cellChildren)
        
        cells.push({
          type: 'tableCell',
          children: cellChildren.length ? cellChildren : [{ type: 'paragraph', children: [{ type: 'text', text: td.textContent || '' }] }]
        })
      })
      
      rows.push({ type: 'tableRow', cells })
    })
    
    return { type: 'table', rows }
  }
  
  /**
   * 解析列表
   */
  private parseList(list: Element, ordered: boolean): ListElement {
    const items: any[] = []
    const listItems = list.querySelectorAll(':scope > li')
    
    listItems.forEach(li => {
      const children: DocumentElement[] = []
      this.parseNodeChildren(li, children)
      
      // 如果没有子元素，创建一个文本段落
      if (children.length === 0) {
        children.push({
          type: 'paragraph',
          children: [{ type: 'text', text: li.textContent || '' }]
        })
      }
      
      items.push({ type: 'listItem', children })
    })
    
    return { type: 'list', ordered, items }
  }
}

// ============== 工厂函数 ==============

export function createPageLayoutEngine(options?: LayoutOptions): PageLayoutEngine {
  return new PageLayoutEngine(options)
}

export default PageLayoutEngine














