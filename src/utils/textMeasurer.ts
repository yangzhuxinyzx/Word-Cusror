/**
 * 文本测量工具
 * 使用 Canvas API 精确测量文本尺寸
 * 参考 ONLYOFFICE 的 CGraphics 设计
 */

// A4 尺寸常量 (mm)
export const A4_WIDTH_MM = 210
export const A4_HEIGHT_MM = 297

// 屏幕 DPI (96 为标准 Windows DPI)
export const SCREEN_DPI = 96
export const MM_TO_PX = SCREEN_DPI / 25.4  // 1mm ≈ 3.78px
export const PT_TO_PX = SCREEN_DPI / 72    // 1pt ≈ 1.33px

// 默认样式
export interface TextStyle {
  fontFamily: string
  fontSize: number  // pt
  fontWeight?: 'normal' | 'bold' | number
  fontStyle?: 'normal' | 'italic'
  color?: string
  lineHeight?: number  // 行高倍数，默认 1.0
  letterSpacing?: number  // 字间距 pt
}

export interface MeasuredText {
  width: number   // px
  height: number  // px
  lines: MeasuredLine[]
}

export interface MeasuredLine {
  text: string
  width: number   // px
  height: number  // px
  y: number       // 相对于段落顶部的 y 坐标
  words: MeasuredWord[]
}

export interface MeasuredWord {
  text: string
  width: number
  x: number  // 相对于行开头的 x 坐标
}

/**
 * 文本测量器类
 * 使用离屏 Canvas 进行精确文本测量
 */
export class TextMeasurer {
  private canvas: HTMLCanvasElement
  private ctx: CanvasRenderingContext2D
  private cache: Map<string, number> = new Map()
  
  constructor() {
    // 创建离屏 Canvas 用于测量
    this.canvas = document.createElement('canvas')
    this.canvas.width = 1
    this.canvas.height = 1
    const ctx = this.canvas.getContext('2d')
    if (!ctx) {
      throw new Error('无法创建 Canvas 2D 上下文')
    }
    this.ctx = ctx
  }
  
  /**
   * 构建 CSS font 字符串
   */
  private buildFontString(style: TextStyle): string {
    const fontStyle = style.fontStyle || 'normal'
    const fontWeight = style.fontWeight || 'normal'
    const fontSize = style.fontSize * PT_TO_PX  // 转换为 px
    const fontFamily = style.fontFamily || 'DengXian, "等线", "Microsoft YaHei", "SimSun", serif'
    
    return `${fontStyle} ${fontWeight} ${fontSize}px ${fontFamily}`
  }
  
  /**
   * 测量单个文本片段的宽度
   */
  measureText(text: string, style: TextStyle): number {
    const fontString = this.buildFontString(style)
    const cacheKey = `${text}|${fontString}`
    
    // 检查缓存
    if (this.cache.has(cacheKey)) {
      return this.cache.get(cacheKey)!
    }
    
    this.ctx.font = fontString
    const metrics = this.ctx.measureText(text)
    const width = metrics.width
    
    // 缓存结果
    this.cache.set(cacheKey, width)
    
    return width
  }
  
  /**
   * 测量单个字符的宽度
   */
  measureChar(char: string, style: TextStyle): number {
    return this.measureText(char, style)
  }
  
  /**
   * 获取行高（px）
   */
  getLineHeight(style: TextStyle): number {
    const fontSize = style.fontSize * PT_TO_PX
    const lineHeight = style.lineHeight || 1
    return fontSize * lineHeight
  }
  
  /**
   * 判断字符是否是中文
   */
  private isChinese(char: string): boolean {
    const code = char.charCodeAt(0)
    return (code >= 0x4E00 && code <= 0x9FFF) ||  // CJK 统一汉字
           (code >= 0x3400 && code <= 0x4DBF) ||  // CJK 扩展 A
           (code >= 0x20000 && code <= 0x2A6DF)   // CJK 扩展 B
  }
  
  /**
   * 判断字符是否可以作为行尾
   */
  private canBreakAfter(char: string): boolean {
    // 中文字符后可以断行
    if (this.isChinese(char)) return true
    // 空格后可以断行
    if (char === ' ' || char === '\t') return true
    // 标点符号后可以断行
    if (',.!?;:。，！？；：、）】》'.includes(char)) return true
    return false
  }
  
  /**
   * 判断字符是否不能作为行首
   */
  private cannotStartLine(char: string): boolean {
    return ',.!?;:。，！？；：、）】》"\''.includes(char)
  }
  
  /**
   * 判断字符是否不能作为行尾
   */
  private cannotEndLine(char: string): boolean {
    return '（【《"\''.includes(char)
  }
  
  /**
   * 将文本按照最大宽度分割成多行
   * 支持中英文混排和换行规则
   */
  measureParagraph(text: string, maxWidth: number, style: TextStyle): MeasuredText {
    const lines: MeasuredLine[] = []
    const lineHeight = this.getLineHeight(style)
    
    // 处理空文本
    if (!text || text.trim() === '') {
      return {
        width: 0,
        height: lineHeight,
        lines: [{
          text: '',
          width: 0,
          height: lineHeight,
          y: 0,
          words: []
        }]
      }
    }
    
    let currentLine = ''
    let currentWidth = 0
    let currentY = 0
    
    // 按字符遍历
    const chars = Array.from(text)
    
    for (let i = 0; i < chars.length; i++) {
      const char = chars[i]
      
      // 处理换行符
      if (char === '\n') {
        lines.push({
          text: currentLine,
          width: currentWidth,
          height: lineHeight,
          y: currentY,
          words: this.splitLineToWords(currentLine, style)
        })
        currentY += lineHeight
        currentLine = ''
        currentWidth = 0
        continue
      }
      
      const charWidth = this.measureChar(char, style)
      
      // 检查是否需要换行
      if (currentWidth + charWidth > maxWidth && currentLine.length > 0) {
        // 检查换行规则
        let breakPoint = currentLine.length
        
        // 如果当前字符不能作为行首，尝试找到更好的断点
        if (this.cannotStartLine(char)) {
          // 将上一个字符也移到下一行
          breakPoint = currentLine.length - 1
          if (breakPoint < 0) breakPoint = 0
        }
        
        // 保存当前行
        const lineText = currentLine.substring(0, breakPoint)
        const lineWidth = this.measureText(lineText, style)
        
        lines.push({
          text: lineText,
          width: lineWidth,
          height: lineHeight,
          y: currentY,
          words: this.splitLineToWords(lineText, style)
        })
        
        currentY += lineHeight
        
        // 开始新行
        const remaining = currentLine.substring(breakPoint)
        currentLine = remaining + char
        currentWidth = this.measureText(currentLine, style)
      } else {
        currentLine += char
        currentWidth += charWidth
      }
    }
    
    // 保存最后一行
    if (currentLine.length > 0) {
      lines.push({
        text: currentLine,
        width: currentWidth,
        height: lineHeight,
        y: currentY,
        words: this.splitLineToWords(currentLine, style)
      })
    }
    
    // 计算总高度和最大宽度
    const totalHeight = lines.length * lineHeight
    const totalWidth = Math.max(...lines.map(l => l.width), 0)
    
    return {
      width: totalWidth,
      height: totalHeight,
      lines
    }
  }
  
  /**
   * 将行分割成单词（用于精确渲染）
   */
  private splitLineToWords(line: string, style: TextStyle): MeasuredWord[] {
    const words: MeasuredWord[] = []
    let currentWord = ''
    let currentX = 0
    
    for (const char of line) {
      if (this.isChinese(char)) {
        // 中文字符单独作为一个"词"
        if (currentWord) {
          const wordWidth = this.measureText(currentWord, style)
          words.push({ text: currentWord, width: wordWidth, x: currentX })
          currentX += wordWidth
          currentWord = ''
        }
        const charWidth = this.measureChar(char, style)
        words.push({ text: char, width: charWidth, x: currentX })
        currentX += charWidth
      } else if (char === ' ') {
        // 空格结束当前词
        if (currentWord) {
          const wordWidth = this.measureText(currentWord, style)
          words.push({ text: currentWord, width: wordWidth, x: currentX })
          currentX += wordWidth
          currentWord = ''
        }
        const spaceWidth = this.measureChar(' ', style)
        words.push({ text: ' ', width: spaceWidth, x: currentX })
        currentX += spaceWidth
      } else {
        currentWord += char
      }
    }
    
    // 保存最后一个词
    if (currentWord) {
      const wordWidth = this.measureText(currentWord, style)
      words.push({ text: currentWord, width: wordWidth, x: currentX })
    }
    
    return words
  }
  
  /**
   * 清除测量缓存
   */
  clearCache(): void {
    this.cache.clear()
  }
}

export async function ensureFontsReady(): Promise<void> {
  if (typeof document === 'undefined') return
  const fonts = (document as any).fonts as FontFaceSet | undefined
  if (!fonts?.ready) return
  try {
    await fonts.ready
  } catch {
    // ignore font loading errors, fallback fonts will be used
  }
}

// 创建全局单例
let globalMeasurer: TextMeasurer | null = null

export function getTextMeasurer(): TextMeasurer {
  if (!globalMeasurer) {
    globalMeasurer = new TextMeasurer()
  }
  return globalMeasurer
}

/**
 * 单位转换工具函数
 */
export function mmToPx(mm: number): number {
  return mm * MM_TO_PX
}

export function pxToMm(px: number): number {
  return px / MM_TO_PX
}

export function ptToPx(pt: number): number {
  return pt * PT_TO_PX
}

export function pxToPt(px: number): number {
  return px / PT_TO_PX
}

/**
 * A4 页面尺寸（px）
 */
export function getA4SizePx(scale: number = 1): { width: number; height: number } {
  return {
    width: A4_WIDTH_MM * MM_TO_PX * scale,
    height: A4_HEIGHT_MM * MM_TO_PX * scale
  }
}














