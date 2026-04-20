/**
 * 元素渲染器
 * 
 * 参考 ONLYOFFICE 的设计:
 * - DocumentServer/sdkjs-src/word/Drawing/DrawingDocument.js
 * - DocumentServer/sdkjs-src/word/Editor/Paragraph.js (Draw 方法)
 * 
 * 功能:
 * - 将布局后的元素渲染到 Canvas
 * - 支持各种文档元素类型
 * - 处理样式和格式
 */

import { Graphics } from '../core/Graphics'
import { getTextMeasurer, TextMeasurer } from '../core/TextMeasurer'
import { MM_TO_PX, PT_TO_PX, DEFAULT_FONT_SIZE_PT, DEFAULT_FONT_FAMILY_CN } from '../core/Constants'
import {
  DocumentElement,
  LayoutElement,
  LayoutPage,
  ParagraphElement,
  ImageElement,
  TableElement,
  ListElement,
  HeaderFooter,
  TextRun,
  PageSettings,
  getDefaultTextStyle
} from './DocumentModel'

// ============== 类型定义 ==============

export interface RenderOptions {
  /** 缩放比例 */
  scale?: number
  /** 是否显示页码 */
  showPageNumber?: boolean
  /** 总页数 */
  totalPages?: number
  /** 图片缓存 */
  imageCache?: Map<string, HTMLImageElement>
}

// ============== ElementRenderer 类 ==============

export class ElementRenderer {
  private graphics: Graphics
  private textMeasurer: TextMeasurer
  private scale: number
  private imageCache: Map<string, HTMLImageElement>
  
  constructor(graphics: Graphics, options: RenderOptions = {}) {
    this.graphics = graphics
    this.textMeasurer = getTextMeasurer()
    this.scale = options.scale || 1
    this.imageCache = options.imageCache || new Map()
  }
  
  // ============== 页面渲染 ==============
  
  /**
   * 渲染单个页面
   * 参考: CDrawingDocument.prototype.OnPaint
   */
  renderPage(page: LayoutPage, pageSettings: PageSettings, options: RenderOptions = {}): void {
    const g = this.graphics
    
    // 清除画布
    g.clear()
    
    // 绘制页面背景（白色）
    g.setBrushColor(255, 255, 255, 1)
    g.fillRect(0, 0, page.width, page.height)
    
    // 绘制页眉
    if (page.header) {
      this.renderHeaderFooter(page.header, pageSettings, 'header', page.pageIndex + 1, options.totalPages)
    }
    
    // 绘制页面内容
    for (const layoutElement of page.elements) {
      this.renderElement(layoutElement)
    }
    
    // 绘制页脚
    if (page.footer) {
      this.renderHeaderFooter(page.footer, pageSettings, 'footer', page.pageIndex + 1, options.totalPages)
    }
    
    // 绘制页码（如果页脚没有内容但需要显示页码）
    if (options.showPageNumber && !page.footer) {
      this.renderPageNumber(page.pageIndex + 1, options.totalPages || 1, pageSettings)
    }
  }
  
  /**
   * 渲染页眉或页脚
   */
  private renderHeaderFooter(
    headerFooter: HeaderFooter,
    pageSettings: PageSettings,
    type: 'header' | 'footer',
    pageNumber: number,
    totalPages?: number
  ): void {
    const g = this.graphics
    const { margins, headerHeight, footerHeight, width } = pageSettings
    
    // 计算位置
    let y: number
    if (type === 'header') {
      y = margins.top
    } else {
      y = pageSettings.height - margins.bottom - footerHeight
    }
    
    const x = margins.left
    const contentWidth = width - margins.left - margins.right
    
    // 设置样式
    const style = headerFooter.style || getDefaultTextStyle()
    g.setFont({
      family: style.fontFamily,
      size: style.fontSize || 10,
      bold: style.fontWeight === 'bold',
      italic: style.fontStyle === 'italic'
    })
    g.setBrushColor(102, 102, 102, 1)  // 灰色文字
    
    // 渲染内容
    for (const element of headerFooter.content) {
      if (element.type === 'paragraph') {
        const para = element as ParagraphElement
        let text = ''
        for (const run of para.children) {
          if (run.type === 'text') {
            text += run.text
          }
        }
        
        // 替换页码占位符
        text = text.replace('{PAGE}', String(pageNumber))
        text = text.replace('{NUMPAGES}', String(totalPages || 1))
        
        // 根据对齐方式计算 X 坐标
        const textWidth = this.textMeasurer.measureWidth(text, {
          fontFamily: style.fontFamily,
          fontSize: style.fontSize || 10
        }) / MM_TO_PX / this.scale
        
        let textX = x
        const alignment = headerFooter.alignment || para.style?.alignment || 'left'
        
        if (alignment === 'center') {
          textX = x + (contentWidth - textWidth) / 2
        } else if (alignment === 'right') {
          textX = x + contentWidth - textWidth
        }
        
        g.fillText(text, textX, y)
      }
    }
    
    // 绘制分隔线
    g.setPenColor(200, 200, 200, 1)
    g.setLineWidthPx(0.5)
    g.beginPath()
    
    if (type === 'header') {
      g.moveTo(x, y + headerHeight - 2)
      g.lineTo(x + contentWidth, y + headerHeight - 2)
    } else {
      g.moveTo(x, y + 2)
      g.lineTo(x + contentWidth, y + 2)
    }
    
    g.stroke()
  }
  
  /**
   * 渲染页码
   */
  private renderPageNumber(pageNumber: number, totalPages: number, pageSettings: PageSettings): void {
    const g = this.graphics
    const { margins, footerHeight, width, height } = pageSettings
    
    const text = `${pageNumber} / ${totalPages}`
    const y = height - margins.bottom - footerHeight / 2
    
    g.setFont({
      family: DEFAULT_FONT_FAMILY_CN,
      size: 10,
      bold: false,
      italic: false
    })
    g.setBrushColor(102, 102, 102, 1)
    
    const textWidth = this.textMeasurer.measureWidth(text, { fontSize: 10 }) / MM_TO_PX / this.scale
    const x = margins.left + (width - margins.left - margins.right - textWidth) / 2
    
    g.fillText(text, x, y)
  }
  
  // ============== 元素渲染 ==============
  
  /**
   * 渲染单个布局元素
   */
  renderElement(layoutElement: LayoutElement): void {
    const { element, x, y, width, height } = layoutElement
    
    switch (element.type) {
      case 'paragraph':
      case 'heading':
        this.renderParagraph(element as ParagraphElement, x, y, width)
        break
      
      case 'image':
        this.renderImage(element as ImageElement, x, y, width, height)
        break
      
      case 'table':
        this.renderTable(element as TableElement, x, y, width)
        break
      
      case 'list':
        this.renderList(element as ListElement, x, y, width)
        break
      
      case 'pageBreak':
      case 'lineBreak':
        // 不渲染，只用于布局
        break
    }
  }
  
  /**
   * 渲染段落
   * 参考: Paragraph.prototype.Draw
   */
  private renderParagraph(paragraph: ParagraphElement, x: number, y: number, maxWidth: number): void {
    const g = this.graphics
    const style = paragraph.style || {}
    const textStyle = paragraph.textStyle || getDefaultTextStyle()
    
    // 应用段落缩进
    let currentX = x + (style.firstLineIndent || 0)
    let currentY = y
    
    const lineHeight = ((textStyle.fontSize || DEFAULT_FONT_SIZE_PT) * PT_TO_PX * (textStyle.lineHeight || 1.15)) / MM_TO_PX
    
    for (const run of paragraph.children) {
      if (run.type === 'text') {
        this.renderTextRun(run, currentX, currentY, maxWidth - (currentX - x), textStyle)
        
        // 简化处理：假设文本在一行内
        const textWidth = this.textMeasurer.measureWidth(run.text, {
          fontFamily: run.style?.fontFamily || textStyle.fontFamily,
          fontSize: run.style?.fontSize || textStyle.fontSize
        }) / MM_TO_PX / this.scale
        
        currentX += textWidth
        
        // 如果超出宽度，换行
        if (currentX > x + maxWidth) {
          currentX = x
          currentY += lineHeight
        }
      }
    }
  }
  
  /**
   * 渲染文本运行
   */
  private renderTextRun(run: TextRun, x: number, y: number, maxWidth: number, defaultStyle: any): void {
    const g = this.graphics
    const style = run.style || defaultStyle
    
    // 设置字体
    g.setFont({
      family: style.fontFamily || DEFAULT_FONT_FAMILY_CN,
      size: style.fontSize || DEFAULT_FONT_SIZE_PT,
      bold: style.fontWeight === 'bold',
      italic: style.fontStyle === 'italic',
      underline: style.textDecoration === 'underline',
      strikethrough: style.textDecoration === 'line-through'
    })
    
    // 设置颜色
    const color = this.parseColor(style.color || '#000000')
    g.setBrushColor(color.r, color.g, color.b, color.a)
    
    // 使用文本测量器进行自动换行
    const maxWidthPx = maxWidth * MM_TO_PX * this.scale
    const measured = this.textMeasurer.measureText(run.text, {
      maxWidth: maxWidthPx,
      wordWrap: true,
      style: {
        fontFamily: style.fontFamily,
        fontSize: style.fontSize,
        fontWeight: style.fontWeight,
        fontStyle: style.fontStyle,
        lineHeight: style.lineHeight
      }
    })
    
    // 渲染每一行
    let lineY = y
    for (const line of measured.lines) {
      g.fillText(line.text, x, lineY)
      lineY += measured.lineHeight / MM_TO_PX / this.scale
    }
  }
  
  /**
   * 渲染图片
   */
  private renderImage(image: ImageElement, x: number, y: number, width: number, height: number): void {
    const g = this.graphics
    
    // 尝试从缓存获取图片
    let img = this.imageCache.get(image.src)
    
    if (img && img.complete) {
      g.drawImage(img, x, y, width, height)
    } else {
      // 图片未加载，绘制占位符
      g.setPenColor(200, 200, 200, 1)
      g.setLineWidthPx(1)
      g.strokeRect(x, y, width, height)
      
      // 绘制 X 形状
      g.beginPath()
      g.moveTo(x, y)
      g.lineTo(x + width, y + height)
      g.moveTo(x + width, y)
      g.lineTo(x, y + height)
      g.stroke()
      
      // 如果图片未缓存，开始加载
      if (!img) {
        img = new Image()
        img.crossOrigin = 'anonymous'
        img.onload = () => {
          // 图片加载完成后，需要重新渲染（由调用方处理）
        }
        img.src = image.src
        this.imageCache.set(image.src, img)
      }
    }
  }
  
  /**
   * 渲染表格
   */
  private renderTable(table: TableElement, x: number, y: number, maxWidth: number): void {
    const g = this.graphics
    
    // 计算列宽
    const columnCount = table.rows[0]?.cells.length || 0
    const columnWidths = table.columnWidths || Array(columnCount).fill(maxWidth / columnCount)
    const tableWidth = table.width || maxWidth
    
    // 绘制表格
    let currentY = y
    
    g.setPenColor(0, 0, 0, 1)
    g.setLineWidthPx(0.5)
    
    for (const row of table.rows) {
      let currentX = x
      const rowHeight = row.height || 10
      
      for (let i = 0; i < row.cells.length; i++) {
        const cell = row.cells[i]
        const cellWidth = columnWidths[i] || (tableWidth / columnCount)
        
        // 绘制单元格边框
        g.strokeRect(currentX, currentY, cellWidth, rowHeight)
        
        // 绘制单元格内容
        const padding = 2  // 2mm 内边距
        for (const child of cell.children) {
          if (child.type === 'paragraph') {
            this.renderParagraph(child as ParagraphElement, currentX + padding, currentY + padding, cellWidth - padding * 2)
          }
        }
        
        currentX += cellWidth
      }
      
      currentY += rowHeight
    }
  }
  
  /**
   * 渲染列表
   */
  private renderList(list: ListElement, x: number, y: number, maxWidth: number): void {
    const g = this.graphics
    
    const indent = 10  // 10mm 缩进
    const bulletWidth = 5
    let currentY = y
    
    for (let i = 0; i < list.items.length; i++) {
      const item = list.items[i]
      
      // 绘制项目符号或数字
      g.setFont({
        family: DEFAULT_FONT_FAMILY_CN,
        size: DEFAULT_FONT_SIZE_PT,
        bold: false,
        italic: false
      })
      g.setBrushColor(0, 0, 0, 1)
      
      const bullet = list.ordered ? `${i + 1}.` : '•'
      g.fillText(bullet, x, currentY)
      
      // 绘制列表项内容
      for (const child of item.children) {
        if (child.type === 'paragraph') {
          this.renderParagraph(child as ParagraphElement, x + indent, currentY, maxWidth - indent)
          
          // 估算高度并更新 Y
          const lineHeight = DEFAULT_FONT_SIZE_PT * PT_TO_PX * 1.15 / MM_TO_PX
          currentY += lineHeight
        }
      }
    }
  }
  
  // ============== 工具方法 ==============
  
  /**
   * 解析颜色字符串
   */
  private parseColor(colorStr: string): { r: number; g: number; b: number; a: number } {
    // 处理 #RRGGBB 格式
    if (colorStr.startsWith('#')) {
      const hex = colorStr.slice(1)
      if (hex.length === 6) {
        return {
          r: parseInt(hex.slice(0, 2), 16),
          g: parseInt(hex.slice(2, 4), 16),
          b: parseInt(hex.slice(4, 6), 16),
          a: 1
        }
      } else if (hex.length === 3) {
        return {
          r: parseInt(hex[0] + hex[0], 16),
          g: parseInt(hex[1] + hex[1], 16),
          b: parseInt(hex[2] + hex[2], 16),
          a: 1
        }
      }
    }
    
    // 处理 rgb(r, g, b) 格式
    const rgbMatch = colorStr.match(/rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)/)
    if (rgbMatch) {
      return {
        r: parseInt(rgbMatch[1]),
        g: parseInt(rgbMatch[2]),
        b: parseInt(rgbMatch[3]),
        a: 1
      }
    }
    
    // 处理 rgba(r, g, b, a) 格式
    const rgbaMatch = colorStr.match(/rgba\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*,\s*([\d.]+)\s*\)/)
    if (rgbaMatch) {
      return {
        r: parseInt(rgbaMatch[1]),
        g: parseInt(rgbaMatch[2]),
        b: parseInt(rgbaMatch[3]),
        a: parseFloat(rgbaMatch[4])
      }
    }
    
    // 默认黑色
    return { r: 0, g: 0, b: 0, a: 1 }
  }
  
  /**
   * 设置图片缓存
   */
  setImageCache(cache: Map<string, HTMLImageElement>): void {
    this.imageCache = cache
  }
  
  /**
   * 预加载图片
   */
  async preloadImage(src: string): Promise<HTMLImageElement> {
    return new Promise((resolve, reject) => {
      let img = this.imageCache.get(src)
      
      if (img && img.complete) {
        resolve(img)
        return
      }
      
      img = new Image()
      img.crossOrigin = 'anonymous'
      img.onload = () => {
        this.imageCache.set(src, img!)
        resolve(img!)
      }
      img.onerror = reject
      img.src = src
    })
  }
}

// ============== 工厂函数 ==============

export function createElementRenderer(graphics: Graphics, options?: RenderOptions): ElementRenderer {
  return new ElementRenderer(graphics, options)
}

export default ElementRenderer














