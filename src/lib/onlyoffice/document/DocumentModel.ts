/**
 * 文档数据模型
 * 
 * 参考 ONLYOFFICE 的设计:
 * - DocumentServer/sdkjs-src/word/Editor/Document.js
 * - DocumentServer/sdkjs-src/word/Editor/Paragraph.js
 * 
 * 定义文档的数据结构，用于渲染和布局
 */

import {
  A4_WIDTH_MM,
  A4_HEIGHT_MM,
  DEFAULT_MARGIN_TOP_MM,
  DEFAULT_MARGIN_BOTTOM_MM,
  DEFAULT_MARGIN_LEFT_MM,
  DEFAULT_MARGIN_RIGHT_MM,
  HEADER_HEIGHT_MM,
  FOOTER_HEIGHT_MM,
  DEFAULT_FONT_SIZE_PT,
  DEFAULT_FONT_FAMILY_CN,
  DEFAULT_LINE_HEIGHT
} from '../core/Constants'

// ============== 基础类型 ==============

export type Alignment = 'left' | 'center' | 'right' | 'justify'

export interface TextStyle {
  fontFamily?: string
  fontSize?: number  // pt
  fontWeight?: 'normal' | 'bold'
  fontStyle?: 'normal' | 'italic'
  textDecoration?: 'none' | 'underline' | 'line-through'
  color?: string
  backgroundColor?: string
  lineHeight?: number
}

export interface ParagraphStyle {
  alignment?: Alignment
  firstLineIndent?: number  // mm
  marginTop?: number        // mm
  marginBottom?: number     // mm
  marginLeft?: number       // mm
  marginRight?: number      // mm
  lineSpacing?: number      // 倍数
}

// ============== 文档元素 ==============

export type ElementType = 
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
  | 'lineBreak'

export interface BaseElement {
  type: ElementType
  id?: string
}

/**
 * 文本运行（Run）- 具有相同样式的一段文本
 */
export interface TextRun extends BaseElement {
  type: 'text'
  text: string
  style?: TextStyle
}

/**
 * 段落元素
 */
export interface ParagraphElement extends BaseElement {
  type: 'paragraph' | 'heading'
  level?: number  // 标题级别 (1-6)
  children: TextRun[]
  style?: ParagraphStyle
  textStyle?: TextStyle
}

/**
 * 图片元素
 */
export interface ImageElement extends BaseElement {
  type: 'image'
  src: string
  alt?: string
  width?: number   // mm
  height?: number  // mm
  alignment?: Alignment
}

/**
 * 表格单元格
 */
export interface TableCellElement extends BaseElement {
  type: 'tableCell'
  children: DocumentElement[]
  colSpan?: number
  rowSpan?: number
  width?: number  // mm
  style?: {
    backgroundColor?: string
    borderColor?: string
    borderWidth?: number
    padding?: number
    verticalAlign?: 'top' | 'middle' | 'bottom'
  }
}

/**
 * 表格行
 */
export interface TableRowElement extends BaseElement {
  type: 'tableRow'
  cells: TableCellElement[]
  height?: number  // mm
}

/**
 * 表格元素
 */
export interface TableElement extends BaseElement {
  type: 'table'
  rows: TableRowElement[]
  width?: number  // mm (总宽度)
  columnWidths?: number[]  // mm (各列宽度)
  style?: {
    borderCollapse?: boolean
    borderColor?: string
    borderWidth?: number
  }
}

/**
 * 列表项
 */
export interface ListItemElement extends BaseElement {
  type: 'listItem'
  children: DocumentElement[]
  level?: number
}

/**
 * 列表元素
 */
export interface ListElement extends BaseElement {
  type: 'list'
  ordered: boolean
  items: ListItemElement[]
  style?: {
    bulletStyle?: string
    startNumber?: number
  }
}

/**
 * 分页符
 */
export interface PageBreakElement extends BaseElement {
  type: 'pageBreak'
}

/**
 * 换行符
 */
export interface LineBreakElement extends BaseElement {
  type: 'lineBreak'
}

/**
 * 所有文档元素的联合类型
 */
export type DocumentElement = 
  | ParagraphElement
  | ImageElement
  | TableElement
  | ListElement
  | PageBreakElement
  | LineBreakElement

// ============== 页眉页脚 ==============

export interface HeaderFooter {
  content: DocumentElement[]
  style?: TextStyle
  alignment?: Alignment
}

// ============== 页面设置 ==============

export interface PageSettings {
  /** 页面宽度 (mm) */
  width: number
  /** 页面高度 (mm) */
  height: number
  /** 页面方向 */
  orientation: 'portrait' | 'landscape'
  /** 边距 (mm) */
  margins: {
    top: number
    bottom: number
    left: number
    right: number
  }
  /** 页眉高度 (mm) */
  headerHeight: number
  /** 页脚高度 (mm) */
  footerHeight: number
}

// ============== 布局后的元素 ==============

export interface LayoutRect {
  x: number      // mm
  y: number      // mm
  width: number  // mm
  height: number // mm
}

export interface LayoutElement extends LayoutRect {
  element: DocumentElement
  pageIndex: number
}

export interface LayoutPage {
  pageIndex: number
  width: number   // mm
  height: number  // mm
  elements: LayoutElement[]
  header?: HeaderFooter
  footer?: HeaderFooter
}

// ============== 文档模型 ==============

export interface DocumentModelData {
  /** 页面设置 */
  pageSettings: PageSettings
  /** 文档内容 */
  content: DocumentElement[]
  /** 页眉 */
  header?: HeaderFooter
  /** 页脚 */
  footer?: HeaderFooter
  /** 元数据 */
  metadata?: {
    title?: string
    author?: string
    createdAt?: Date
    modifiedAt?: Date
  }
}

/**
 * 文档模型类
 */
export class DocumentModel {
  private data: DocumentModelData
  
  constructor(data?: Partial<DocumentModelData>) {
    this.data = {
      pageSettings: data?.pageSettings || getDefaultPageSettings(),
      content: data?.content || [],
      header: data?.header,
      footer: data?.footer,
      metadata: data?.metadata
    }
  }
  
  // ============== 获取器 ==============
  
  getPageSettings(): PageSettings {
    return this.data.pageSettings
  }
  
  getContent(): DocumentElement[] {
    return this.data.content
  }
  
  getHeader(): HeaderFooter | undefined {
    return this.data.header
  }
  
  getFooter(): HeaderFooter | undefined {
    return this.data.footer
  }
  
  getMetadata() {
    return this.data.metadata
  }
  
  // ============== 设置器 ==============
  
  setPageSettings(settings: Partial<PageSettings>): void {
    this.data.pageSettings = { ...this.data.pageSettings, ...settings }
  }
  
  setContent(content: DocumentElement[]): void {
    this.data.content = content
  }
  
  setHeader(header: HeaderFooter | undefined): void {
    this.data.header = header
  }
  
  setFooter(footer: HeaderFooter | undefined): void {
    this.data.footer = footer
  }
  
  // ============== 内容操作 ==============
  
  addElement(element: DocumentElement): void {
    this.data.content.push(element)
  }
  
  insertElement(index: number, element: DocumentElement): void {
    this.data.content.splice(index, 0, element)
  }
  
  removeElement(index: number): DocumentElement | undefined {
    return this.data.content.splice(index, 1)[0]
  }
  
  clearContent(): void {
    this.data.content = []
  }
  
  // ============== 计算方法 ==============
  
  /**
   * 获取内容区域尺寸
   */
  getContentArea(): { x: number; y: number; width: number; height: number } {
    const { width, height, margins, headerHeight, footerHeight } = this.data.pageSettings
    
    return {
      x: margins.left,
      y: margins.top + headerHeight,
      width: width - margins.left - margins.right,
      height: height - margins.top - margins.bottom - headerHeight - footerHeight
    }
  }
  
  /**
   * 导出为 JSON
   */
  toJSON(): DocumentModelData {
    return JSON.parse(JSON.stringify(this.data))
  }
  
  /**
   * 从 JSON 导入
   */
  static fromJSON(json: DocumentModelData): DocumentModel {
    return new DocumentModel(json)
  }
}

// ============== 默认值 ==============

/**
 * 获取默认页面设置
 */
export function getDefaultPageSettings(): PageSettings {
  return {
    width: A4_WIDTH_MM,
    height: A4_HEIGHT_MM,
    orientation: 'portrait',
    margins: {
      top: DEFAULT_MARGIN_TOP_MM,
      bottom: DEFAULT_MARGIN_BOTTOM_MM,
      left: DEFAULT_MARGIN_LEFT_MM,
      right: DEFAULT_MARGIN_RIGHT_MM
    },
    headerHeight: HEADER_HEIGHT_MM,
    footerHeight: FOOTER_HEIGHT_MM
  }
}

/**
 * 获取默认文本样式
 */
export function getDefaultTextStyle(): TextStyle {
  return {
    fontFamily: DEFAULT_FONT_FAMILY_CN,
    fontSize: DEFAULT_FONT_SIZE_PT,
    fontWeight: 'normal',
    fontStyle: 'normal',
    textDecoration: 'none',
    color: '#000000',
    lineHeight: DEFAULT_LINE_HEIGHT
  }
}

/**
 * 获取默认段落样式
 */
export function getDefaultParagraphStyle(): ParagraphStyle {
  return {
    alignment: 'left',
    firstLineIndent: 0,
    marginTop: 0,
    marginBottom: 0,
    marginLeft: 0,
    marginRight: 0,
    lineSpacing: DEFAULT_LINE_HEIGHT
  }
}

export default DocumentModel














