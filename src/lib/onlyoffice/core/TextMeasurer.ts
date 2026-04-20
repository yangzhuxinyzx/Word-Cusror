/**
 * 文本测量器
 * 
 * 参考 ONLYOFFICE 的设计:
 * - DocumentServer/sdkjs-src/common/libfont/textmeasurer.js
 * - DocumentServer/sdkjs-src/common/libfont/manager.js
 * 
 * 功能:
 * - 精确测量文本宽度和高度
 * - 支持多字体、多字号
 * - 行高计算
 * - 缓存优化
 */

import {
  PT_TO_PX,
  DEFAULT_FONT_FAMILY_CN,
  DEFAULT_FONT_SIZE_PT,
  DEFAULT_LINE_HEIGHT
} from './Constants'

// ============== 类型定义 ==============

export interface TextStyle {
  fontFamily?: string
  fontSize?: number  // pt
  fontWeight?: 'normal' | 'bold' | number
  fontStyle?: 'normal' | 'italic'
  lineHeight?: number  // 倍数
}

export interface MeasuredLine {
  text: string
  width: number   // px
  height: number  // px
  x: number       // px
  y: number       // px (相对于段落起始)
}

export interface MeasuredText {
  lines: MeasuredLine[]
  totalWidth: number   // px (最宽行)
  totalHeight: number  // px
  lineHeight: number   // px (单行高度)
}

export interface MeasureOptions {
  maxWidth?: number     // px，最大宽度（用于自动换行）
  style?: TextStyle
  wordWrap?: boolean    // 是否自动换行
}

// ============== 缓存 ==============

interface CacheKey {
  text: string
  font: string
}

interface CacheEntry {
  width: number
  height: number
  ascent: number
  descent: number
}

// ============== TextMeasurer 类 ==============

export class TextMeasurer {
  private canvas: HTMLCanvasElement
  private ctx: CanvasRenderingContext2D
  
  // 缓存（使用 Map 以 font+text 为键）
  private cache: Map<string, CacheEntry> = new Map()
  private maxCacheSize: number = 10000
  
  // 当前字体设置
  private currentFont: string = ''
  
  constructor() {
    // 创建离屏 Canvas 用于测量
    this.canvas = document.createElement('canvas')
    this.canvas.width = 1
    this.canvas.height = 1
    this.ctx = this.canvas.getContext('2d')!
    this.ctx.textBaseline = 'top'
  }
  
  // ============== 字体设置 ==============
  
  /**
   * 构建 CSS 字体字符串
   */
  private buildFontString(style: TextStyle): string {
    const fontStyle = style.fontStyle || 'normal'
    const fontWeight = style.fontWeight || 'normal'
    const fontSize = (style.fontSize || DEFAULT_FONT_SIZE_PT) * PT_TO_PX
    const fontFamily = style.fontFamily || DEFAULT_FONT_FAMILY_CN
    
    return `${fontStyle} ${fontWeight} ${fontSize}px ${fontFamily}`
  }
  
  /**
   * 设置当前字体
   */
  setFont(style: TextStyle): void {
    const fontString = this.buildFontString(style)
    if (fontString !== this.currentFont) {
      this.currentFont = fontString
      this.ctx.font = fontString
    }
  }
  
  // ============== 测量方法 ==============
  
  /**
   * 测量单个文本的宽度
   */
  measureWidth(text: string, style?: TextStyle): number {
    if (style) {
      this.setFont(style)
    }
    
    const cacheKey = `${this.currentFont}|${text}`
    const cached = this.cache.get(cacheKey)
    
    if (cached) {
      return cached.width
    }
    
    const metrics = this.ctx.measureText(text)
    const width = metrics.width
    
    // 计算高度
    const height = this.getLineHeight(style)
    
    // 缓存结果
    this.addToCache(cacheKey, {
      width,
      height,
      ascent: metrics.actualBoundingBoxAscent || height * 0.8,
      descent: metrics.actualBoundingBoxDescent || height * 0.2
    })
    
    return width
  }
  
  /**
   * 测量文本高度
   */
  measureHeight(text: string, style?: TextStyle): number {
    if (style) {
      this.setFont(style)
    }
    
    const cacheKey = `${this.currentFont}|${text}`
    const cached = this.cache.get(cacheKey)
    
    if (cached) {
      return cached.height
    }
    
    const metrics = this.ctx.measureText(text)
    let height: number
    
    // 使用 actualBoundingBox 计算高度（如果支持）
    if (metrics.actualBoundingBoxAscent !== undefined && metrics.actualBoundingBoxDescent !== undefined) {
      height = metrics.actualBoundingBoxAscent + metrics.actualBoundingBoxDescent
    } else {
      // 回退到基于字号的估算
      height = this.getLineHeight(style)
    }
    
    return height
  }
  
  /**
   * 获取行高
   */
  getLineHeight(style?: TextStyle): number {
    const fontSize = (style?.fontSize || DEFAULT_FONT_SIZE_PT) * PT_TO_PX
    const lineHeightMultiplier = style?.lineHeight || DEFAULT_LINE_HEIGHT
    return fontSize * lineHeightMultiplier
  }
  
  /**
   * 测量文本（支持自动换行）
   */
  measureText(text: string, options: MeasureOptions = {}): MeasuredText {
    const style = options.style || {}
    this.setFont(style)
    
    const lineHeight = this.getLineHeight(style)
    const maxWidth = options.maxWidth
    const wordWrap = options.wordWrap !== false && maxWidth !== undefined
    
    if (!wordWrap || !maxWidth) {
      // 不换行，直接测量
      const width = this.measureWidth(text, style)
      return {
        lines: [{
          text,
          width,
          height: lineHeight,
          x: 0,
          y: 0
        }],
        totalWidth: width,
        totalHeight: lineHeight,
        lineHeight
      }
    }
    
    // 自动换行
    const lines = this.wrapText(text, maxWidth, style)
    let totalWidth = 0
    let y = 0
    
    const measuredLines: MeasuredLine[] = lines.map(lineText => {
      const width = this.measureWidth(lineText, style)
      totalWidth = Math.max(totalWidth, width)
      
      const line: MeasuredLine = {
        text: lineText,
        width,
        height: lineHeight,
        x: 0,
        y
      }
      
      y += lineHeight
      return line
    })
    
    return {
      lines: measuredLines,
      totalWidth,
      totalHeight: y,
      lineHeight
    }
  }
  
  /**
   * 文本换行
   */
  private wrapText(text: string, maxWidth: number, style?: TextStyle): string[] {
    if (!text) return ['']
    
    const lines: string[] = []
    
    // 按自然换行符分割
    const paragraphs = text.split('\n')
    
    for (const paragraph of paragraphs) {
      if (!paragraph) {
        lines.push('')
        continue
      }
      
      // 对中英文混合文本进行处理
      const words = this.splitIntoWords(paragraph)
      let currentLine = ''
      let currentWidth = 0
      
      for (const word of words) {
        const wordWidth = this.measureWidth(word, style)
        
        if (currentWidth + wordWidth <= maxWidth) {
          currentLine += word
          currentWidth += wordWidth
        } else {
          if (currentLine) {
            lines.push(currentLine)
          }
          
          // 如果单个词超过最大宽度，需要强制断开
          if (wordWidth > maxWidth) {
            const chars = word.split('')
            currentLine = ''
            currentWidth = 0
            
            for (const char of chars) {
              const charWidth = this.measureWidth(char, style)
              if (currentWidth + charWidth <= maxWidth) {
                currentLine += char
                currentWidth += charWidth
              } else {
                if (currentLine) {
                  lines.push(currentLine)
                }
                currentLine = char
                currentWidth = charWidth
              }
            }
          } else {
            currentLine = word
            currentWidth = wordWidth
          }
        }
      }
      
      if (currentLine) {
        lines.push(currentLine)
      }
    }
    
    return lines.length ? lines : ['']
  }
  
  /**
   * 将文本分割为"词"（考虑中英文混合）
   * 中文每个字符独立，英文按空格分词
   */
  private splitIntoWords(text: string): string[] {
    const words: string[] = []
    let currentWord = ''
    let isInEnglishWord = false
    
    for (let i = 0; i < text.length; i++) {
      const char = text[i]
      const code = char.charCodeAt(0)
      
      // 判断是否为 CJK 字符
      const isCJK = (code >= 0x4E00 && code <= 0x9FFF) ||  // CJK Unified Ideographs
                    (code >= 0x3400 && code <= 0x4DBF) ||  // CJK Unified Ideographs Extension A
                    (code >= 0x3000 && code <= 0x303F) ||  // CJK Symbols and Punctuation
                    (code >= 0xFF00 && code <= 0xFFEF)     // Halfwidth and Fullwidth Forms
      
      if (isCJK) {
        // CJK 字符：结束当前英文单词，单独作为一个"词"
        if (currentWord) {
          words.push(currentWord)
          currentWord = ''
        }
        words.push(char)
        isInEnglishWord = false
      } else if (char === ' ') {
        // 空格：结束当前单词，空格作为独立"词"
        if (currentWord) {
          words.push(currentWord)
          currentWord = ''
        }
        words.push(char)
        isInEnglishWord = false
      } else {
        // 英文/数字/其他：累积到当前单词
        currentWord += char
        isInEnglishWord = true
      }
    }
    
    if (currentWord) {
      words.push(currentWord)
    }
    
    return words
  }
  
  // ============== 缓存管理 ==============
  
  private addToCache(key: string, entry: CacheEntry): void {
    // LRU 简化实现：超过最大缓存时清空一半
    if (this.cache.size >= this.maxCacheSize) {
      const keysToDelete = Array.from(this.cache.keys()).slice(0, this.maxCacheSize / 2)
      keysToDelete.forEach(k => this.cache.delete(k))
    }
    
    this.cache.set(key, entry)
  }
  
  /**
   * 清空缓存
   */
  clearCache(): void {
    this.cache.clear()
  }
  
  /**
   * 获取缓存大小
   */
  getCacheSize(): number {
    return this.cache.size
  }
}

// ============== 单例 ==============

let globalTextMeasurer: TextMeasurer | null = null

/**
 * 获取全局文本测量器实例
 */
export function getTextMeasurer(): TextMeasurer {
  if (!globalTextMeasurer) {
    globalTextMeasurer = new TextMeasurer()
  }
  return globalTextMeasurer
}

export default TextMeasurer














