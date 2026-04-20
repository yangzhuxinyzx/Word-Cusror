/**
 * Canvas 分页预览容器
 * 管理多页 Canvas 渲染，支持滚动和缩放
 * 
 * 基于 ONLYOFFICE HtmlPage.js 的设计思想：
 * - 精确的 A4 尺寸和边距
 * - 使用 mm 作为基本单位
 * - 高质量 Canvas 渲染
 * 
 * ONLYOFFICE 关键参数 @ 96 DPI：
 * - A4 尺寸: 210mm x 297mm = 793.7px x 1122.5px
 * - 默认边距: 上下 25.4mm (96px), 左右 31.7mm (120px)
 * - 页眉页脚高度: 12.7mm (48px)
 */

import React, { useState, useEffect, useMemo, useCallback, useRef } from 'react'
import CanvasPageRenderer from './CanvasPageRenderer'
import { 
  LayoutEngine, 
  createLayoutEngine, 
  getDefaultPageConfig,
  LayoutResult,
  PageConfig,
  pageSettingsToPageConfig
} from '../utils/layoutEngine'
import { A4_WIDTH_MM, A4_HEIGHT_MM, MM_TO_PX } from '../utils/textMeasurer'
import type { DocxParseResult, PageSettings } from '../utils/docxParser'
import { ensureFontsReady } from '../utils/textMeasurer'
import { prewarmNativeTextMeasurementsFromHtml } from '../utils/nativeTextMeasurement'
import type { DocxInspectionReport } from '../types'

// ONLYOFFICE 精确参数 @ 96 DPI
const ONLYOFFICE_PARAMS = {
  DPI: 96,
  MM_PER_INCH: 25.4,
  MM_TO_PX: 96 / 25.4,  // ≈ 3.7795
  A4_WIDTH_MM: 210,
  A4_HEIGHT_MM: 297,
  MARGIN_TOP_MM: 25.4,    // 1 inch
  MARGIN_BOTTOM_MM: 25.4, // 1 inch
  MARGIN_LEFT_MM: 31.7,   // 1.25 inch
  MARGIN_RIGHT_MM: 31.7,  // 1.25 inch
  HEADER_HEIGHT_MM: 12.7, // 0.5 inch
  FOOTER_HEIGHT_MM: 12.7, // 0.5 inch
  PAGE_GAP_PX: 20,        // 页间距
}

interface CanvasPagePreviewProps {
  // 可以传入 HTML 或 DocxParseResult
  html?: string
  docxData?: DocxParseResult
  // 页眉页脚
  headerHtml?: string
  footerHtml?: string
  // 缩放比例
  scale?: number
  // 是否显示页码
  showPageNumbers?: boolean
  // 页面设置（可选，否则使用默认 A4）
  pageSettings?: PageSettings
  // 布局完成后回传当前页和总页数
  onPageChange?: (currentPage: number, totalPages: number) => void
  docxInspectionReport?: DocxInspectionReport | null
}

/**
 * Canvas 预览容器组件
 */
export const CanvasPagePreview: React.FC<CanvasPagePreviewProps> = ({
  html,
  docxData,
  headerHtml,
  footerHtml,
  scale = 0.8,
  showPageNumbers = true,
  pageSettings,
  onPageChange,
  docxInspectionReport,
}) => {
  const [layoutResult, setLayoutResult] = useState<LayoutResult | null>(null)
  const [pageConfig, setPageConfig] = useState<PageConfig>(getDefaultPageConfig(scale))
  const [isLoading, setIsLoading] = useState(true)
  const containerRef = useRef<HTMLDivElement>(null)
  const imageCache = useRef<Map<string, HTMLImageElement>>(new Map())
  const resolvedContentHtml = docxData?.bodyHtml || html || ''
  const resolvedHeaderHtml = headerHtml || docxData?.headerHtml || ''
  const resolvedFooterHtml = footerHtml || docxData?.footerHtml || ''
  
  // 从 DocxParseResult 或 pageSettings 更新页面配置
  useEffect(() => {
    let config: PageConfig
    
    if (pageSettings) {
      // 使用传入的页面设置
      config = pageSettingsToPageConfig(pageSettings, scale)
    } else if (docxData?.pageSettings) {
      // 从 docxData 获取页面设置
      config = pageSettingsToPageConfig(docxData.pageSettings, scale)
    } else {
      // 默认 A4 配置
      config = getDefaultPageConfig(scale)
    }
    
    setPageConfig(config)
  }, [pageSettings, docxData?.pageSettings, scale])
  
  // 执行布局计算
  useEffect(() => {
    const performLayout = async () => {
      setIsLoading(true)
      
      try {
        await ensureFontsReady()
        await prewarmNativeTextMeasurementsFromHtml(resolvedContentHtml, scale, docxInspectionReport)
        const contentHtml = resolvedContentHtml
        const header = resolvedHeaderHtml
        const footer = resolvedFooterHtml

        if (!contentHtml) {
          setLayoutResult({ pages: [], totalHeight: 0 })
          setIsLoading(false)
          return
        }
        
        // 创建布局引擎
        const engine = createLayoutEngine(pageConfig, scale)
        
        // 执行布局
        const result = engine.layoutFromHtml(contentHtml, header, footer)
        
        console.log('[CanvasPagePreview] 布局完成:', {
          pagesCount: result.pages.length,
          totalHeight: result.totalHeight
        })
        
        setLayoutResult(result)
      } catch (error) {
        console.error('[CanvasPagePreview] 布局失败:', error)
        setLayoutResult({ pages: [], totalHeight: 0 })
      } finally {
        setIsLoading(false)
      }
    }
    
    performLayout()
  }, [docxInspectionReport, pageConfig, resolvedContentHtml, resolvedFooterHtml, resolvedHeaderHtml, scale])
  
  // 预加载图片
  useEffect(() => {
    if (!resolvedContentHtml) return
    
    // 从 HTML 中提取所有图片 src
    const imgRegex = /<img[^>]+src=["']([^"']+)["']/gi
    let match
    const srcs: string[] = []
    
    while ((match = imgRegex.exec(resolvedContentHtml)) !== null) {
      srcs.push(match[1])
    }
    
    // 预加载图片
    for (const src of srcs) {
      if (!imageCache.current.has(src)) {
        const img = new Image()
        img.crossOrigin = 'anonymous'
        img.src = src
        imageCache.current.set(src, img)
      }
    }
  }, [resolvedContentHtml])

  useEffect(() => {
    if (typeof window === 'undefined') return

    const pageCount = layoutResult?.pages?.length || 0
    const debugPayload = {
      status: isLoading ? 'loading' : 'ready',
      layoutResult,
      pageConfig,
      scale,
      pageCount,
    }

    ;(window as any).__wordCursorLayoutDebug = debugPayload
    return () => {
      if ((window as any).__wordCursorLayoutDebug === debugPayload) {
        delete (window as any).__wordCursorLayoutDebug
      }
    }
  }, [isLoading, layoutResult, pageConfig, scale])

  useEffect(() => {
    if (!onPageChange || isLoading) return
    const totalPages = layoutResult?.pages?.length || 0
    onPageChange(totalPages > 0 ? 1 : 0, totalPages)
  }, [isLoading, layoutResult, onPageChange])
  
  // 渲染页面列表
  const renderPages = useMemo(() => {
    if (!layoutResult || layoutResult.pages.length === 0) {
      return (
        <div 
          style={{
            width: pageConfig.width,
            height: pageConfig.height,
            backgroundColor: 'var(--word-page-bg)',
            boxShadow: 'var(--word-page-shadow)',
            border: '1px solid var(--word-page-border)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            color: 'var(--word-ink-muted)',
            fontSize: 14
          }}
        >
          暂无内容
        </div>
      )
    }
    
    return layoutResult.pages.map((page, index) => (
      <div 
        key={index}
        data-word-page-index={index + 1}
        data-testid={`word-render-page-${index + 1}`}
        style={{
          marginBottom: ONLYOFFICE_PARAMS.PAGE_GAP_PX,
          position: 'relative'
        }}
      >
        <CanvasPageRenderer
          page={page}
          pageIndex={index + 1}
          pageConfig={pageConfig}
          scale={scale}
          showPageNumber={showPageNumbers}
          totalPages={layoutResult.pages.length}
          imageCache={imageCache.current}
        />
        
        {/* 页码标签（ONLYOFFICE 风格） */}
        {showPageNumbers && (
          <div
            style={{
              position: 'absolute',
              bottom: -18,
              left: '50%',
              transform: 'translateX(-50%)',
              fontSize: 11,
              color: 'var(--word-ink-muted)',
              whiteSpace: 'nowrap',
              fontFamily: '"Segoe UI", -apple-system, BlinkMacSystemFont, sans-serif'
            }}
          >
            第 {index + 1} 页 / 共 {layoutResult.pages.length} 页
          </div>
        )}
      </div>
    ))
  }, [layoutResult, pageConfig, scale, showPageNumbers])
  
  return (
    <div
      ref={containerRef}
      className="canvas-page-preview"
      data-testid="word-render-canvas-preview"
      data-render-status={isLoading ? 'loading' : 'ready'}
      data-page-count={layoutResult?.pages?.length || 0}
      style={{
        display: 'flex',
        flexDirection: 'column',
        alignItems: 'center',
        padding: `${ONLYOFFICE_PARAMS.PAGE_GAP_PX * 2}px ${ONLYOFFICE_PARAMS.PAGE_GAP_PX}px`,
        // ONLYOFFICE 风格的工作区背景色
        backgroundColor: 'var(--word-canvas-bg)',
        minHeight: '100%',
        overflow: 'auto'
      }}
    >
      {isLoading ? (
        <div
          style={{
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            width: pageConfig.width,
            height: pageConfig.height,
            backgroundColor: 'var(--word-page-bg)',
            boxShadow: '0 1px 3px rgba(0,0,0,0.12)',
            color: 'var(--word-ink-muted)',
            fontSize: 14
          }}
        >
          <div style={{ textAlign: 'center' }}>
            <div 
              style={{
                width: 32,
                height: 32,
                border: '3px solid var(--word-page-border)',
                borderTopColor: 'var(--word-ink-muted)',
                borderRadius: '50%',
                animation: 'spin 1s linear infinite',
                margin: '0 auto 12px'
              }}
            />
            正在渲染页面...
          </div>
        </div>
      ) : (
        renderPages
      )}
      
      {/* 添加旋转动画 */}
      <style>{`
        @keyframes spin {
          to { transform: rotate(360deg); }
        }
      `}</style>
    </div>
  )
}

export default CanvasPagePreview

