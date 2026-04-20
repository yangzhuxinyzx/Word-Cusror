/**
 * Canvas 页面渲染器
 * 使用 Canvas 2D API 精确渲染单个 A4 页面
 * 参考 ONLYOFFICE 的 CGraphics 和 printpreview.js
 */

import React, { useEffect, useRef, useCallback, useMemo } from 'react'
import type { 
  PageLayout, 
  LayoutElement, 
  TextLayoutElement, 
  ImageLayoutElement,
  TableLayoutElement,
  TableRowLayoutElement,
  ParagraphLayoutElement,
  PageConfig,
  ElementStyle
} from '../utils/layoutEngine'
import { PT_TO_PX } from '../utils/textMeasurer'
import { getNativeStyleMetrics } from '../utils/nativeTextMeasurement'

interface CanvasPageRendererProps {
  page: PageLayout
  pageIndex?: number
  pageConfig: PageConfig
  scale?: number
  showPageNumber?: boolean
  totalPages?: number
  // 图片缓存 Map
  imageCache?: Map<string, HTMLImageElement>
}

/**
 * 单页 Canvas 渲染组件
 */
export const CanvasPageRenderer: React.FC<CanvasPageRendererProps> = ({
  page,
  pageIndex = 1,
  pageConfig,
  scale = 1,
  showPageNumber = true,
  totalPages = 1,
  imageCache = new Map()
}) => {
  const canvasRef = useRef<HTMLCanvasElement>(null)
  
  // 计算 Canvas 实际尺寸（考虑设备像素比）
  const devicePixelRatio = typeof window !== 'undefined' ? window.devicePixelRatio || 1 : 1
  const renderScale = devicePixelRatio * 2
  
  const canvasWidth = useMemo(() => pageConfig.width * renderScale, [pageConfig.width, renderScale])
  const canvasHeight = useMemo(() => pageConfig.height * renderScale, [pageConfig.height, renderScale])
  
  const getWordPalette = useCallback(() => {
    if (typeof window === 'undefined') {
      return {
        pageBg: '#ffffff',
        ink: '#000000',
        inkMuted: '#666666',
        rule: '#999999'
      }
    }
    const styles = getComputedStyle(document.documentElement)
    const read = (name: string, fallback: string) => styles.getPropertyValue(name).trim() || fallback
    return {
      pageBg: read('--word-page-bg', '#ffffff'),
      ink: read('--word-ink', '#000000'),
      inkMuted: read('--word-ink-muted', '#666666'),
      rule: read('--word-rule', '#999999')
    }
  }, [])

  const resolveColor = useCallback((
    color: string | undefined,
    palette: { ink: string; inkMuted: string; rule: string }
  ) => {
    if (!color) return undefined
    const trimmed = color.trim()
    if (trimmed === 'var(--word-ink)') return palette.ink
    if (trimmed === 'var(--word-ink-muted)') return palette.inkMuted
    if (trimmed === 'var(--word-rule)') return palette.rule
    return color
  }, [])

  /**
   * 设置 Canvas 上下文样式
   */
  const setContextStyle = useCallback((
    ctx: CanvasRenderingContext2D, 
    style?: ElementStyle
  ) => {
    if (!style) return
    
    // 设置字体
    const fontSize = (style.fontSize || 12) * PT_TO_PX * scale * renderScale
    const fontWeight = style.fontWeight || 'normal'
    const fontStyle = style.fontStyle || 'normal'
    const fontFamily = style.fontFamily || 'DengXian, "Microsoft YaHei", "SimSun", serif'
    
    ctx.font = `${fontStyle} ${fontWeight} ${fontSize}px ${fontFamily}`
    
    // 设置颜色
    const palette = getWordPalette()
    ctx.fillStyle = resolveColor(style.color, palette) || palette.ink
  }, [scale, renderScale, getWordPalette, resolveColor])
  
  /**
   * 绘制文本元素
   */
  const drawText = useCallback((
    ctx: CanvasRenderingContext2D,
    element: TextLayoutElement
  ) => {
    const { measuredText, style } = element
    if (!measuredText || !measuredText.lines) return
    
    setContextStyle(ctx, style)
    const palette = getWordPalette()
    const resolvedColor = resolveColor(style?.color, palette) || palette.ink
    
    const x = element.x * renderScale
    let y = element.y * renderScale
    const letterSpacingPx = ((style?.letterSpacing || 0) * PT_TO_PX * scale) * renderScale
    const fontSizePx = (style?.fontSize || 12) * PT_TO_PX * scale * renderScale
    const nativeStyleMetrics = style
      ? getNativeStyleMetrics({
          fontFamily: style.fontFamily || 'DengXian, "Microsoft YaHei", "SimSun", serif',
          fontSize: style.fontSize || 12,
          fontWeight: style.fontWeight,
          fontStyle: style.fontStyle,
          lineHeight: style.lineHeight,
          letterSpacing: style.letterSpacing,
          scaleFactor: scale,
        })
      : null
    const baselineOffset = (nativeStyleMetrics?.baseline || (fontSizePx * 0.82) / renderScale) * renderScale
    
    // 绘制每一行
    for (const line of measuredText.lines) {
      const lineY = y + line.y * renderScale + baselineOffset
      
      // 绘制文本
      if (letterSpacingPx) {
        let cursorX = x
        for (const char of Array.from(line.text)) {
          ctx.fillText(char, cursorX, lineY)
          cursorX += ctx.measureText(char).width + letterSpacingPx
        }
      } else {
        ctx.fillText(line.text, x, lineY)
      }
      
      // 绘制下划线
      if (style?.textDecoration === 'underline') {
        const underlineY = lineY + 2 * renderScale
        ctx.beginPath()
        ctx.moveTo(x, underlineY)
        ctx.lineTo(x + line.width * renderScale, underlineY)
        ctx.strokeStyle = resolvedColor
        ctx.lineWidth = 1 * renderScale
        ctx.stroke()
      }
      
      // 绘制删除线
      if (style?.textDecoration === 'line-through') {
        const strikeY = lineY - baselineOffset * 0.35
        ctx.beginPath()
        ctx.moveTo(x, strikeY)
        ctx.lineTo(x + line.width * renderScale, strikeY)
        ctx.strokeStyle = resolvedColor
        ctx.lineWidth = 1 * renderScale
        ctx.stroke()
      }
    }
  }, [setContextStyle, scale, renderScale, getWordPalette, resolveColor])
  
  /**
   * 绘制段落
   */
  const drawParagraph = useCallback((
    ctx: CanvasRenderingContext2D,
    element: ParagraphLayoutElement
  ) => {
    // 处理对齐
    const contentWidth = pageConfig.width - pageConfig.marginLeft - pageConfig.marginRight
    
    if (element.children) {
      for (const child of element.children) {
        if (child.type === 'text') {
          const textChild = child as TextLayoutElement
          
          // 根据对齐方式调整 x 坐标
          let adjustedX = textChild.x
          
          if (element.alignment === 'center') {
            adjustedX = pageConfig.marginLeft + (contentWidth - textChild.width) / 2
          } else if (element.alignment === 'right') {
            adjustedX = pageConfig.marginLeft + contentWidth - textChild.width
          }
          
          drawText(ctx, {
            ...textChild,
            x: adjustedX
          })
        }
      }
    }
  }, [drawText, pageConfig])
  
  /**
   * 绘制图片
   */
  const drawImage = useCallback((
    ctx: CanvasRenderingContext2D,
    element: ImageLayoutElement
  ) => {
    const { src, x, y, width, height } = element
    
    // 尝试从缓存获取图片
    let img = imageCache.get(src)
    
    if (img && img.complete) {
      // 图片已加载，直接绘制
      ctx.drawImage(
        img,
        x * renderScale,
        y * renderScale,
        width * renderScale,
        (height - 10 * scale) * renderScale  // 减去间距
      )
    } else {
      // 绘制占位符
      const palette = getWordPalette()
      ctx.fillStyle = palette.pageBg
      ctx.fillRect(
        x * renderScale,
        y * renderScale,
        width * renderScale,
        (height - 10 * scale) * renderScale
      )
      
      // 绘制加载提示
      ctx.fillStyle = palette.inkMuted
        ctx.font = `${12 * renderScale}px sans-serif`
      ctx.textAlign = 'center'
      ctx.fillText(
        '图片加载中...',
          (x + width / 2) * renderScale,
          (y + height / 2) * renderScale
      )
      ctx.textAlign = 'left'
      
      // 如果图片未缓存，创建并加载
      if (!img) {
        img = new Image()
        img.crossOrigin = 'anonymous'
        imageCache.set(src, img)
        
        img.onload = () => {
          // 图片加载完成后重新渲染
          const canvas = canvasRef.current
          if (canvas) {
            const ctx = canvas.getContext('2d')
            if (ctx) {
              ctx.drawImage(
                img!,
                x * renderScale,
                y * renderScale,
                width * renderScale,
                (height - 10 * scale) * renderScale
              )
            }
          }
        }
        
        img.src = src
      }
    }
  }, [imageCache, renderScale, scale, getWordPalette])
  
  /**
   * 绘制表格
   */
  const drawTable = useCallback((
    ctx: CanvasRenderingContext2D,
    element: TableLayoutElement
  ) => {
    const { rows, x, y, width } = element
    
    const palette = getWordPalette()
    
    let currentY = y
    
    for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
      const row = rows[rowIndex]
      
      for (const cell of row.cells) {
        const cellX = (x + (cell.x || 0)) * renderScale
        const cellY = currentY * renderScale
        const cellWidth = cell.width * renderScale
        const rowSpan = Math.max(1, cell.rowspan || 1)
        const spannedHeight = rows
          .slice(rowIndex, Math.min(rows.length, rowIndex + rowSpan))
          .reduce((sum, tableRow) => sum + tableRow.height, 0)
        const cellHeight = spannedHeight * renderScale
        
        // 绘制单元格边框（支持每边不同样式）
        const defaultBorderWidth = 0.5 * PT_TO_PX * scale
        const topWidth = cell.style?.borderTopWidth ?? cell.style?.borderWidth ?? defaultBorderWidth
        const rightWidth = cell.style?.borderRightWidth ?? cell.style?.borderWidth ?? defaultBorderWidth
        const bottomWidth = cell.style?.borderBottomWidth ?? cell.style?.borderWidth ?? defaultBorderWidth
        const leftWidth = cell.style?.borderLeftWidth ?? cell.style?.borderWidth ?? defaultBorderWidth
        const topColor = resolveColor(cell.style?.borderTopColor, palette) || resolveColor(cell.style?.borderColor, palette) || palette.rule
        const rightColor = resolveColor(cell.style?.borderRightColor, palette) || resolveColor(cell.style?.borderColor, palette) || palette.rule
        const bottomColor = resolveColor(cell.style?.borderBottomColor, palette) || resolveColor(cell.style?.borderColor, palette) || palette.rule
        const leftColor = resolveColor(cell.style?.borderLeftColor, palette) || resolveColor(cell.style?.borderColor, palette) || palette.rule

        const drawLine = (x1: number, y1: number, x2: number, y2: number, width: number, color: string) => {
          if (!width || width <= 0) return
          ctx.strokeStyle = color
            ctx.lineWidth = width * renderScale
          ctx.beginPath()
          ctx.moveTo(x1, y1)
          ctx.lineTo(x2, y2)
          ctx.stroke()
        }

        drawLine(cellX, cellY, cellX + cellWidth, cellY, topWidth, topColor)
        drawLine(cellX + cellWidth, cellY, cellX + cellWidth, cellY + cellHeight, rightWidth, rightColor)
        drawLine(cellX, cellY + cellHeight, cellX + cellWidth, cellY + cellHeight, bottomWidth, bottomColor)
        drawLine(cellX, cellY, cellX, cellY + cellHeight, leftWidth, leftColor)
        
        // 绘制单元格内容
        if (cell.children) {
          for (const child of cell.children) {
            if (child.type === 'text') {
              const textChild = child as TextLayoutElement
              const paddingLeft = textChild.x || 0
              const paddingTop = textChild.y || 0
              const paddingRight = cell.style?.paddingRight || paddingLeft
              const paddingBottom = cell.style?.paddingBottom || paddingTop
              const usableWidth = Math.max(0, cell.width - paddingLeft - paddingRight)
              const usableHeight = Math.max(0, (spannedHeight) - paddingTop - paddingBottom)
              let textX = x + (cell.x || 0) + paddingLeft
              let textY = currentY + paddingTop

              if (cell.style?.textAlign === 'center') {
                textX += Math.max(0, (usableWidth - textChild.width) / 2)
              } else if (cell.style?.textAlign === 'right') {
                textX += Math.max(0, usableWidth - textChild.width)
              }

              if (cell.style?.verticalAlign === 'middle') {
                textY += Math.max(0, (usableHeight - textChild.height) / 2)
              } else if (cell.style?.verticalAlign === 'bottom') {
                textY += Math.max(0, usableHeight - textChild.height)
              }

              drawText(ctx, {
                ...textChild,
                x: textX,
                y: textY
              })
            }
          }
        }
      }
      
      currentY += row.height
    }
  }, [drawText, renderScale, scale, getWordPalette])
  
  /**
   * 绘制页眉
   */
  const drawHeader = useCallback((
    ctx: CanvasRenderingContext2D,
    header: LayoutElement
  ) => {
    const x = header.x * devicePixelRatio
    const y = header.y * devicePixelRatio
    const width = header.width * devicePixelRatio
    const height = header.height * devicePixelRatio
    
    // 绘制页眉文本
    if (header.children) {
      for (const child of header.children) {
        if (child.type === 'text') {
          const textChild = child as TextLayoutElement
          
          // 页眉靠右对齐
          const textX = x + width - (textChild.measuredText?.width || 0) * devicePixelRatio
          
          drawText(ctx, {
            ...textChild,
            x: textX / devicePixelRatio
          })
        }
      }
    }
    
  }, [drawText, devicePixelRatio, getWordPalette])
  
  /**
   * 绘制页脚
   */
  const drawFooter = useCallback((
    ctx: CanvasRenderingContext2D,
    footer: LayoutElement
  ) => {
    // 绘制页脚文本（居中）
    if (footer.children) {
      for (const child of footer.children) {
        if (child.type === 'text') {
          const textChild = child as TextLayoutElement
          
          // 页脚居中
          const contentWidth = pageConfig.width - pageConfig.marginLeft - pageConfig.marginRight
          const textX = pageConfig.marginLeft + (contentWidth - (textChild.measuredText?.width || 0)) / 2
          
          drawText(ctx, {
            ...textChild,
            x: textX
          })
        }
      }
    }
  }, [drawText, pageConfig])
  
  /**
   * 绘制单个元素
   */
  const drawElement = useCallback((
    ctx: CanvasRenderingContext2D,
    element: LayoutElement
  ) => {
    switch (element.type) {
      case 'text':
        drawText(ctx, element as TextLayoutElement)
        break
      case 'paragraph':
      case 'heading':
        drawParagraph(ctx, element as ParagraphLayoutElement)
        break
      case 'image':
        drawImage(ctx, element as ImageLayoutElement)
        break
      case 'table':
        drawTable(ctx, element as TableLayoutElement)
        break
      case 'list':
        // 列表：绘制子元素
        if (element.children) {
          for (const child of element.children) {
            drawElement(ctx, child)
          }
        }
        break
      case 'header':
        drawHeader(ctx, element)
        break
      case 'footer':
        drawFooter(ctx, element)
        break
      default:
        // 递归处理子元素
        if (element.children) {
          for (const child of element.children) {
            drawElement(ctx, child)
          }
        }
    }
  }, [drawText, drawParagraph, drawImage, drawTable, drawHeader, drawFooter])
  
  /**
   * 渲染整个页面
   */
  const renderPage = useCallback(() => {
    const canvas = canvasRef.current
    if (!canvas) return
    
    const ctx = canvas.getContext('2d')
    if (!ctx) return
    
    const palette = getWordPalette()
    // 清空画布
    ctx.fillStyle = palette.pageBg
    ctx.fillRect(0, 0, canvasWidth, canvasHeight)
    
    // 设置默认样式
    ctx.textBaseline = 'top'
    ctx.fillStyle = palette.ink
    
    // 绘制页眉
    if (page.header) {
      drawElement(ctx, page.header)
    }
    
    // 绘制所有元素
    for (const element of page.elements) {
      drawElement(ctx, element)
    }
    
    // 绘制页脚
    if (page.footer) {
      drawElement(ctx, page.footer)
    }
    
  }, [page, canvasWidth, canvasHeight, drawElement, getWordPalette])
  
  // 组件挂载和更新时渲染
  useEffect(() => {
    renderPage()
  }, [renderPage])
  
  return (
    <canvas
      ref={canvasRef}
      data-testid={`word-render-canvas-page-${pageIndex}`}
      data-word-page-index={pageIndex}
      width={canvasWidth}
      height={canvasHeight}
      style={{
        width: pageConfig.width,
        height: pageConfig.height,
        display: 'block',
        backgroundColor: 'var(--word-page-bg)',
        boxShadow: 'var(--word-page-shadow)',
        borderRadius: '1px',
        border: '1px solid var(--word-page-border)'
      }}
    />
  )
}

export default CanvasPageRenderer














