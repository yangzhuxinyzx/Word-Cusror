/**
 * Canvas 渲染引擎
 * 
 * 参考 ONLYOFFICE 的设计:
 * - DocumentServer/sdkjs-src/common/Drawings/GraphicsBase.js
 * - DocumentServer/sdkjs-src/word/Drawing/Graphics.js
 * 
 * 核心功能:
 * - 封装 Canvas 2D Context
 * - mm 到 px 精确转换
 * - 文本绘制（支持中英文混排）
 * - 图形绘制（矩形、线条、路径）
 * - 图片绘制
 * - 变换矩阵支持
 */

import {
  MM_TO_PX,
  PT_TO_PX,
  DEFAULT_FONT_FAMILY_CN,
  DEFAULT_FONT_SIZE_PT,
  getDevicePixelRatio
} from './Constants'

// ============== 类型定义 ==============

export interface GraphicsConfig {
  /** Canvas 元素 */
  canvas: HTMLCanvasElement
  /** 页面宽度 (mm) */
  widthMm: number
  /** 页面高度 (mm) */
  heightMm: number
  /** 缩放比例 */
  scale?: number
}

export interface FontStyle {
  family?: string
  size?: number  // pt
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strikethrough?: boolean
}

export interface Color {
  r: number
  g: number
  b: number
  a?: number
}

// ============== 渲染器类型（来自 ONLYOFFICE） ==============

export enum RendererType {
  Base = 0,
  Drawer = 1,
  PDF = 2,
  BoundsChecker = 3,
  Track = 4,
  TextDrawer = 5,
  NativeDrawer = 6
}

// ============== Graphics 类 ==============

export class Graphics {
  // Canvas 相关
  private canvas: HTMLCanvasElement
  private ctx: CanvasRenderingContext2D
  private devicePixelRatio: number
  
  // 尺寸（mm 单位）
  private widthMm: number = 210
  private heightMm: number = 297
  
  // 尺寸（px 单位，考虑设备像素比）
  private widthPx: number = 0
  private heightPx: number = 0
  
  // 缩放比例
  private scale: number = 1
  
  // 当前状态
  private currentFont: FontStyle = {
    family: DEFAULT_FONT_FAMILY_CN,
    size: DEFAULT_FONT_SIZE_PT,
    bold: false,
    italic: false
  }
  
  private penColor: Color = { r: 0, g: 0, b: 0, a: 1 }
  private brushColor: Color = { r: 0, g: 0, b: 0, a: 1 }
  
  // 渲染器类型
  private type: RendererType = RendererType.Drawer
  
  // 标志
  public isDarkMode: boolean = false
  public isPrintMode: boolean = false
  public isPrintPreview: boolean = false
  
  constructor() {
    this.devicePixelRatio = getDevicePixelRatio()
    this.canvas = document.createElement('canvas')
    this.ctx = this.canvas.getContext('2d')!
  }
  
  // ============== 初始化 ==============
  
  /**
   * 初始化渲染器
   * 参考: CGraphics.prototype.init
   */
  init(config: GraphicsConfig): void {
    this.canvas = config.canvas
    this.ctx = this.canvas.getContext('2d')!
    this.widthMm = config.widthMm
    this.heightMm = config.heightMm
    this.scale = config.scale || 1
    this.devicePixelRatio = getDevicePixelRatio()
    
    // 计算像素尺寸
    this.widthPx = this.widthMm * MM_TO_PX * this.scale
    this.heightPx = this.heightMm * MM_TO_PX * this.scale
    
    // 设置 Canvas 尺寸（考虑设备像素比）
    this.canvas.width = this.widthPx * this.devicePixelRatio
    this.canvas.height = this.heightPx * this.devicePixelRatio
    this.canvas.style.width = `${this.widthPx}px`
    this.canvas.style.height = `${this.heightPx}px`
    
    // 缩放上下文以适应设备像素比
    this.ctx.scale(this.devicePixelRatio, this.devicePixelRatio)
    
    // 设置默认样式
    this.ctx.textBaseline = 'top'
    this.ctx.imageSmoothingEnabled = true
    this.ctx.imageSmoothingQuality = 'high'
  }
  
  /**
   * 使用已有的 Context 初始化
   */
  initFromContext(
    ctx: CanvasRenderingContext2D,
    width: number,
    height: number,
    widthMm: number,
    heightMm: number
  ): void {
    this.ctx = ctx
    this.widthPx = width
    this.heightPx = height
    this.widthMm = widthMm
    this.heightMm = heightMm
    
    this.ctx.textBaseline = 'top'
  }
  
  // ============== 坐标转换 ==============
  
  /**
   * mm 转 px（考虑缩放）
   */
  mmToPx(mm: number): number {
    return mm * MM_TO_PX * this.scale
  }
  
  /**
   * pt 转 px（考虑缩放）
   */
  ptToPx(pt: number): number {
    return pt * PT_TO_PX * this.scale
  }
  
  // ============== 变换 ==============
  
  /**
   * 设置变换矩阵
   * 参考: CGraphics.prototype.transform
   */
  transform(a: number, b: number, c: number, d: number, tx: number, ty: number): void {
    this.ctx.transform(a, b, c, d, tx * this.scale, ty * this.scale)
  }
  
  /**
   * 重置变换
   */
  resetTransform(): void {
    this.ctx.setTransform(this.devicePixelRatio, 0, 0, this.devicePixelRatio, 0, 0)
  }
  
  /**
   * 保存状态
   */
  save(): void {
    this.ctx.save()
  }
  
  /**
   * 恢复状态
   */
  restore(): void {
    this.ctx.restore()
  }
  
  // ============== 颜色设置 ==============
  
  /**
   * 设置笔触颜色
   * 参考: CGraphics.prototype.p_color
   */
  setPenColor(r: number, g: number, b: number, a: number = 1): void {
    this.penColor = { r, g, b, a }
    this.ctx.strokeStyle = `rgba(${r}, ${g}, ${b}, ${a})`
  }
  
  /**
   * 设置填充颜色
   * 参考: CGraphics.prototype.b_color1
   */
  setBrushColor(r: number, g: number, b: number, a: number = 1): void {
    this.brushColor = { r, g, b, a }
    this.ctx.fillStyle = `rgba(${r}, ${g}, ${b}, ${a})`
  }
  
  /**
   * 从 CSS 颜色字符串设置填充色
   */
  setFillStyle(color: string): void {
    this.ctx.fillStyle = color
  }
  
  /**
   * 从 CSS 颜色字符串设置描边色
   */
  setStrokeStyle(color: string): void {
    this.ctx.strokeStyle = color
  }
  
  // ============== 字体设置 ==============
  
  /**
   * 设置字体
   * 参考: CGraphics.prototype.SetFont
   */
  setFont(style: FontStyle): void {
    this.currentFont = { ...this.currentFont, ...style }
    
    const fontStyle = this.currentFont.italic ? 'italic' : 'normal'
    const fontWeight = this.currentFont.bold ? 'bold' : 'normal'
    const fontSize = this.ptToPx(this.currentFont.size || DEFAULT_FONT_SIZE_PT)
    const fontFamily = this.currentFont.family || DEFAULT_FONT_FAMILY_CN
    
    this.ctx.font = `${fontStyle} ${fontWeight} ${fontSize}px ${fontFamily}`
  }
  
  /**
   * 设置字体（简化版）
   */
  setFontSimple(family: string, sizePt: number, bold: boolean = false, italic: boolean = false): void {
    this.setFont({ family, size: sizePt, bold, italic })
  }
  
  // ============== 文本绘制 ==============
  
  /**
   * 绘制文本
   * 参考: CGraphics.prototype.FillText
   */
  fillText(text: string, x: number, y: number): void {
    const px = this.mmToPx(x)
    const py = this.mmToPx(y)
    
    this.ctx.fillText(text, px, py)
    
    // 绘制下划线
    if (this.currentFont.underline) {
      const metrics = this.ctx.measureText(text)
      const fontSize = this.ptToPx(this.currentFont.size || DEFAULT_FONT_SIZE_PT)
      const underlineY = py + fontSize + 2
      
      this.ctx.beginPath()
      this.ctx.moveTo(px, underlineY)
      this.ctx.lineTo(px + metrics.width, underlineY)
      this.ctx.strokeStyle = this.ctx.fillStyle as string
      this.ctx.lineWidth = 1
      this.ctx.stroke()
    }
    
    // 绘制删除线
    if (this.currentFont.strikethrough) {
      const metrics = this.ctx.measureText(text)
      const fontSize = this.ptToPx(this.currentFont.size || DEFAULT_FONT_SIZE_PT)
      const strikeY = py + fontSize / 2
      
      this.ctx.beginPath()
      this.ctx.moveTo(px, strikeY)
      this.ctx.lineTo(px + metrics.width, strikeY)
      this.ctx.strokeStyle = this.ctx.fillStyle as string
      this.ctx.lineWidth = 1
      this.ctx.stroke()
    }
  }
  
  /**
   * 绘制文本（像素坐标）
   */
  fillTextPx(text: string, x: number, y: number): void {
    this.ctx.fillText(text, x * this.scale, y * this.scale)
  }
  
  /**
   * 测量文本宽度
   */
  measureText(text: string): TextMetrics {
    return this.ctx.measureText(text)
  }
  
  /**
   * 测量文本宽度（返回 mm）
   */
  measureTextMm(text: string): number {
    const metrics = this.ctx.measureText(text)
    return metrics.width / MM_TO_PX / this.scale
  }
  
  // ============== 图形绘制 ==============
  
  /**
   * 绘制矩形路径
   * 参考: CGraphics.prototype.rect
   */
  rect(x: number, y: number, w: number, h: number): void {
    this.ctx.rect(
      this.mmToPx(x),
      this.mmToPx(y),
      this.mmToPx(w),
      this.mmToPx(h)
    )
  }
  
  /**
   * 绘制矩形路径（像素坐标）
   */
  rectPx(x: number, y: number, w: number, h: number): void {
    this.ctx.rect(
      x * this.scale,
      y * this.scale,
      w * this.scale,
      h * this.scale
    )
  }
  
  /**
   * 填充矩形
   */
  fillRect(x: number, y: number, w: number, h: number): void {
    this.ctx.fillRect(
      this.mmToPx(x),
      this.mmToPx(y),
      this.mmToPx(w),
      this.mmToPx(h)
    )
  }
  
  /**
   * 填充矩形（像素坐标）
   */
  fillRectPx(x: number, y: number, w: number, h: number): void {
    this.ctx.fillRect(
      x * this.scale,
      y * this.scale,
      w * this.scale,
      h * this.scale
    )
  }
  
  /**
   * 描边矩形
   */
  strokeRect(x: number, y: number, w: number, h: number): void {
    this.ctx.strokeRect(
      this.mmToPx(x),
      this.mmToPx(y),
      this.mmToPx(w),
      this.mmToPx(h)
    )
  }
  
  /**
   * 描边矩形（像素坐标）
   */
  strokeRectPx(x: number, y: number, w: number, h: number): void {
    this.ctx.strokeRect(
      x * this.scale,
      y * this.scale,
      w * this.scale,
      h * this.scale
    )
  }
  
  /**
   * 清除矩形区域
   */
  clearRect(x: number, y: number, w: number, h: number): void {
    this.ctx.clearRect(
      this.mmToPx(x),
      this.mmToPx(y),
      this.mmToPx(w),
      this.mmToPx(h)
    )
  }
  
  /**
   * 清除整个画布
   */
  clear(): void {
    this.ctx.clearRect(0, 0, this.widthPx, this.heightPx)
  }
  
  // ============== 路径操作 ==============
  
  /**
   * 开始新路径
   */
  beginPath(): void {
    this.ctx.beginPath()
  }
  
  /**
   * 关闭路径
   */
  closePath(): void {
    this.ctx.closePath()
  }
  
  /**
   * 移动到
   */
  moveTo(x: number, y: number): void {
    this.ctx.moveTo(this.mmToPx(x), this.mmToPx(y))
  }
  
  /**
   * 移动到（像素坐标）
   */
  moveToPx(x: number, y: number): void {
    this.ctx.moveTo(x * this.scale, y * this.scale)
  }
  
  /**
   * 画线到
   */
  lineTo(x: number, y: number): void {
    this.ctx.lineTo(this.mmToPx(x), this.mmToPx(y))
  }
  
  /**
   * 画线到（像素坐标）
   */
  lineToPx(x: number, y: number): void {
    this.ctx.lineTo(x * this.scale, y * this.scale)
  }
  
  /**
   * 填充路径
   */
  fill(): void {
    this.ctx.fill()
  }
  
  /**
   * 描边路径
   */
  stroke(): void {
    this.ctx.stroke()
  }
  
  /**
   * 设置线宽
   */
  setLineWidth(width: number): void {
    this.ctx.lineWidth = this.mmToPx(width)
  }
  
  /**
   * 设置线宽（像素）
   */
  setLineWidthPx(width: number): void {
    this.ctx.lineWidth = width * this.scale
  }
  
  // ============== 图片绘制 ==============
  
  /**
   * 绘制图片
   * 参考: CGraphics.prototype.drawImage
   */
  drawImage(
    image: HTMLImageElement | HTMLCanvasElement,
    x: number,
    y: number,
    w: number,
    h: number
  ): void {
    this.ctx.drawImage(
      image,
      this.mmToPx(x),
      this.mmToPx(y),
      this.mmToPx(w),
      this.mmToPx(h)
    )
  }
  
  /**
   * 绘制图片（像素坐标）
   */
  drawImagePx(
    image: HTMLImageElement | HTMLCanvasElement,
    x: number,
    y: number,
    w: number,
    h: number
  ): void {
    this.ctx.drawImage(
      image,
      x * this.scale,
      y * this.scale,
      w * this.scale,
      h * this.scale
    )
  }
  
  /**
   * 从 Base64 绘制图片
   */
  async drawImageFromBase64(
    base64: string,
    x: number,
    y: number,
    w: number,
    h: number
  ): Promise<void> {
    return new Promise((resolve, reject) => {
      const img = new Image()
      img.onload = () => {
        this.drawImage(img, x, y, w, h)
        resolve()
      }
      img.onerror = reject
      img.src = base64
    })
  }
  
  // ============== 裁剪 ==============
  
  /**
   * 设置裁剪区域
   */
  clip(): void {
    this.ctx.clip()
  }
  
  // ============== 全局透明度 ==============
  
  /**
   * 设置全局透明度
   */
  setGlobalAlpha(alpha: number): void {
    this.ctx.globalAlpha = alpha
  }
  
  // ============== 获取器 ==============
  
  getCanvas(): HTMLCanvasElement {
    return this.canvas
  }
  
  getContext(): CanvasRenderingContext2D {
    return this.ctx
  }
  
  getWidthMm(): number {
    return this.widthMm
  }
  
  getHeightMm(): number {
    return this.heightMm
  }
  
  getWidthPx(): number {
    return this.widthPx
  }
  
  getHeightPx(): number {
    return this.heightPx
  }
  
  getScale(): number {
    return this.scale
  }
}

// ============== 导出单例创建函数 ==============

export function createGraphics(): Graphics {
  return new Graphics()
}

export default Graphics














