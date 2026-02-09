import React, { useRef, useEffect, useState, useCallback, useMemo } from 'react'

/**
 * WordPageContainer - 精确的 A4 页面容器组件
 * 
 * 基于 ONLYOFFICE 的设计思想：
 * - 使用 mm 作为基本单位
 * - 精确的 A4 尺寸（210mm x 297mm）
 * - 可配置的边距和页眉/页脚区域
 * - 真实的分页显示效果
 */

// ============== 精确单位转换（来自 ONLYOFFICE） ==============
const SCREEN_DPI = 96 // 标准屏幕 DPI
const MM_PER_INCH = 25.4

// mm 到 px 的精确转换系数
export const MM_TO_PX = SCREEN_DPI / MM_PER_INCH // ≈ 3.7795

// A4 纸张尺寸（mm）
export const A4_WIDTH_MM = 210
export const A4_HEIGHT_MM = 297

// Word 默认边距（mm）- 来自 ONLYOFFICE
export const DEFAULT_MARGINS = {
  top: 25.4,    // 1 inch = 25.4mm
  bottom: 25.4, // 1 inch
  left: 31.7,   // 1.25 inch ≈ 31.7mm (Word 默认)
  right: 31.7,  // 1.25 inch
}

// 页眉页脚区域高度（mm）
export const HEADER_HEIGHT_MM = 12.7  // 0.5 inch
export const FOOTER_HEIGHT_MM = 12.7  // 0.5 inch

// 页间距（px）
export const PAGE_GAP = 20

// ============== 工具函数 ==============
export function mmToPx(mm: number): number {
  return mm * MM_TO_PX
}

export function pxToMm(px: number): number {
  return px / MM_TO_PX
}

export function ptToPx(pt: number): number {
  // 1pt = 1/72 inch, 1 inch = 96px (at 96 DPI)
  return pt * (SCREEN_DPI / 72)
}

export function pxToPt(px: number): number {
  return px * (72 / SCREEN_DPI)
}

// ============== 类型定义 ==============
export interface PageMargins {
  top: number    // mm
  bottom: number // mm
  left: number   // mm
  right: number  // mm
}

export interface PageSize {
  width: number  // mm
  height: number // mm
}

export interface HeaderFooterConfig {
  content?: React.ReactNode
  height?: number  // mm
  style?: React.CSSProperties
}

export interface PageSettings {
  size: PageSize
  margins: PageMargins
  orientation: 'portrait' | 'landscape'
  header?: HeaderFooterConfig
  footer?: HeaderFooterConfig
}

export interface WordPageContainerProps {
  children: React.ReactNode
  settings?: Partial<PageSettings>
  scale?: number  // 缩放比例，默认 1
  showPageBorder?: boolean
  showShadow?: boolean
  pageNumber?: number
  totalPages?: number
  className?: string
  style?: React.CSSProperties
  onContentOverflow?: (overflow: boolean, contentHeight: number) => void
}

// ============== 默认设置 ==============
const DEFAULT_SETTINGS: PageSettings = {
  size: {
    width: A4_WIDTH_MM,
    height: A4_HEIGHT_MM,
  },
  margins: DEFAULT_MARGINS,
  orientation: 'portrait',
}

// ============== 主组件 ==============
export function WordPageContainer({
  children,
  settings: customSettings,
  scale = 1,
  showPageBorder = true,
  showShadow = true,
  pageNumber,
  totalPages,
  className = '',
  style,
  onContentOverflow,
}: WordPageContainerProps) {
  const contentRef = useRef<HTMLDivElement>(null)
  const [contentOverflow, setContentOverflow] = useState(false)
  
  // 合并设置
  const settings = useMemo(() => ({
    ...DEFAULT_SETTINGS,
    ...customSettings,
    size: { ...DEFAULT_SETTINGS.size, ...customSettings?.size },
    margins: { ...DEFAULT_MARGINS, ...customSettings?.margins },
    header: customSettings?.header,
    footer: customSettings?.footer,
  }), [customSettings])

  // 计算页面尺寸（考虑横向/纵向）
  const pageSize = useMemo(() => {
    if (settings.orientation === 'landscape') {
      return {
        width: mmToPx(settings.size.height),
        height: mmToPx(settings.size.width),
      }
    }
    return {
      width: mmToPx(settings.size.width),
      height: mmToPx(settings.size.height),
    }
  }, [settings.orientation, settings.size])

  // 计算边距（px）
  const margins = useMemo(() => ({
    top: mmToPx(settings.margins.top),
    bottom: mmToPx(settings.margins.bottom),
    left: mmToPx(settings.margins.left),
    right: mmToPx(settings.margins.right),
  }), [settings.margins])

  // 计算内容区域尺寸
  const contentArea = useMemo(() => {
    const headerHeight = settings.header?.height ? mmToPx(settings.header.height) : 0
    const footerHeight = settings.footer?.height ? mmToPx(settings.footer.height) : 0
    
    return {
      width: pageSize.width - margins.left - margins.right,
      height: pageSize.height - margins.top - margins.bottom - headerHeight - footerHeight,
    }
  }, [pageSize, margins, settings.header, settings.footer])

  // 监测内容溢出
  useEffect(() => {
    if (!contentRef.current || !onContentOverflow) return

    const observer = new ResizeObserver((entries) => {
      for (const entry of entries) {
        const contentHeight = entry.contentRect.height
        const isOverflow = contentHeight > contentArea.height
        setContentOverflow(isOverflow)
        onContentOverflow(isOverflow, contentHeight)
      }
    })

    observer.observe(contentRef.current)
    return () => observer.disconnect()
  }, [contentArea.height, onContentOverflow])

  // 页眉渲染
  const renderHeader = () => {
    if (!settings.header?.content) return null
    
    const headerHeight = settings.header.height ? mmToPx(settings.header.height) : mmToPx(HEADER_HEIGHT_MM)
    
    return (
      <div 
        className="word-page-header"
        style={{
          position: 'absolute',
          top: margins.top - headerHeight,
          left: margins.left,
          right: margins.right,
          height: headerHeight,
          display: 'flex',
          alignItems: 'flex-end',
          paddingBottom: 4,
          borderBottom: '1px solid var(--word-rule)',
          color: 'var(--word-ink-muted)',
          ...settings.header.style,
        }}
      >
        {settings.header.content}
      </div>
    )
  }

  // 页脚渲染
  const renderFooter = () => {
    const hasFooterContent = settings.footer?.content
    const showPageNumber = pageNumber !== undefined
    
    if (!hasFooterContent && !showPageNumber) return null
    
    const footerHeight = settings.footer?.height ? mmToPx(settings.footer.height) : mmToPx(FOOTER_HEIGHT_MM)
    
    return (
      <div 
        className="word-page-footer"
        style={{
          position: 'absolute',
          bottom: margins.bottom - footerHeight,
          left: margins.left,
          right: margins.right,
          height: footerHeight,
          display: 'flex',
          alignItems: 'flex-start',
          justifyContent: 'center',
          paddingTop: 4,
          borderTop: '1px solid var(--word-rule)',
          color: 'var(--word-ink-muted)',
          ...settings.footer?.style,
        }}
      >
        {hasFooterContent ? (
          settings.footer!.content
        ) : showPageNumber ? (
          <span className="page-number-text" style={{ fontSize: '10pt', color: 'var(--word-ink-muted)' }}>
            {totalPages ? `${pageNumber} / ${totalPages}` : pageNumber}
          </span>
        ) : null}
      </div>
    )
  }

  return (
    <div
      className={`word-page-container ${className}`}
      style={{
        position: 'relative',
        width: pageSize.width * scale,
        height: pageSize.height * scale,
        backgroundColor: 'var(--word-page-bg)',
        boxShadow: showShadow ? 'var(--word-page-shadow)' : undefined,
        border: showPageBorder ? '1px solid var(--word-page-border)' : undefined,
        overflow: 'hidden',
        transform: scale !== 1 ? `scale(${scale})` : undefined,
        transformOrigin: 'top left',
        ...style,
      }}
    >
      {/* 页面内容层 - 模拟 ONLYOFFICE 的 CEditorPage */}
      <div
        style={{
          position: 'absolute',
          top: 0,
          left: 0,
          right: 0,
          bottom: 0,
          transform: scale !== 1 ? `scale(${1/scale})` : undefined,
          transformOrigin: 'top left',
          width: pageSize.width,
          height: pageSize.height,
        }}
      >
        {/* 页眉 */}
        {renderHeader()}
        
        {/* 正文内容区域 */}
        <div
          ref={contentRef}
          className="word-page-content"
          style={{
            position: 'absolute',
            top: margins.top,
            left: margins.left,
            width: contentArea.width,
            minHeight: contentArea.height,
            // Word 默认样式
            fontFamily: 'var(--word-font-family-cn)',
            fontSize: 'var(--word-font-size)',  // 五号字
            lineHeight: 'var(--word-line-height)',     // Word 默认行距
            color: 'var(--word-ink)',
            overflow: 'visible', // 允许溢出以便检测
          }}
        >
          {children}
        </div>
        
        {/* 页脚 */}
        {renderFooter()}
        
        {/* 内容溢出警告（仅开发调试用） */}
        {contentOverflow && process.env.NODE_ENV === 'development' && (
          <div
            style={{
              position: 'absolute',
              bottom: 0,
              left: 0,
              right: 0,
              height: 4,
              background: 'linear-gradient(to right, transparent, rgba(255, 0, 0, 0.3), transparent)',
            }}
          />
        )}
      </div>
    </div>
  )
}

// ============== 多页容器组件 ==============
export interface MultiPageContainerProps {
  pages: React.ReactNode[]
  settings?: Partial<PageSettings>
  scale?: number
  showPageNumbers?: boolean
  gap?: number  // 页间距（px）
  className?: string
}

export function MultiPageContainer({
  pages,
  settings,
  scale = 1,
  showPageNumbers = true,
  gap = PAGE_GAP,
  className = '',
}: MultiPageContainerProps) {
  return (
    <div
      className={`multi-page-container ${className}`}
      style={{
        display: 'flex',
        flexDirection: 'column',
        alignItems: 'center',
        gap: gap,
        padding: gap,
        backgroundColor: 'var(--word-canvas-bg)',  // 类似 Word 的灰色背景
        minHeight: '100%',
      }}
    >
      {pages.map((pageContent, index) => (
        <WordPageContainer
          key={index}
          settings={settings}
          scale={scale}
          pageNumber={showPageNumbers ? index + 1 : undefined}
          totalPages={showPageNumbers ? pages.length : undefined}
        >
          {pageContent}
        </WordPageContainer>
      ))}
    </div>
  )
}

// ============== 页面计算工具 ==============

/**
 * 根据内容高度计算需要多少页
 * 参考 ONLYOFFICE 的 Recalculate_Page 逻辑
 */
export function calculatePageCount(
  contentHeight: number, 
  pageSettings: PageSettings = DEFAULT_SETTINGS
): number {
  const { size, margins, header, footer } = pageSettings
  
  const headerHeight = header?.height ? mmToPx(header.height) : 0
  const footerHeight = footer?.height ? mmToPx(footer.height) : 0
  
  const availableHeight = mmToPx(size.height) - mmToPx(margins.top) - mmToPx(margins.bottom) - headerHeight - footerHeight
  
  if (contentHeight <= availableHeight) return 1
  
  return Math.ceil(contentHeight / availableHeight)
}

/**
 * 获取 A4 页面的精确像素尺寸
 */
export function getA4SizePx(orientation: 'portrait' | 'landscape' = 'portrait'): { width: number; height: number } {
  if (orientation === 'landscape') {
    return {
      width: mmToPx(A4_HEIGHT_MM),
      height: mmToPx(A4_WIDTH_MM),
    }
  }
  return {
    width: mmToPx(A4_WIDTH_MM),
    height: mmToPx(A4_HEIGHT_MM),
  }
}

export default WordPageContainer














