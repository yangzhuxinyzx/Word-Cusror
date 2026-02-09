/**
 * ONLYOFFICE 风格预览组件 (优化版)
 * 
 * 复用现有的 CanvasPageRenderer 和 layoutEngine 渲染逻辑
 * 提供与 CanvasPagePreview 相同的渲染质量
 * 
 * 此组件的目的是作为一个更模块化的入口点，
 * 方便未来逐步替换为我们自己的渲染引擎
 */

import React, { useState, useEffect, useMemo, useRef } from 'react'
// 复用现有的渲染组件
import CanvasPageRenderer from '../../../components/CanvasPageRenderer'
import { 
  createLayoutEngine, 
  getDefaultPageConfig,
  LayoutResult,
  PageConfig
} from '../../../utils/layoutEngine'
import type { DocxParseResult, PageSettings } from '../../../utils/docxParser'
import { HeaderFooter as HeaderFooterType, PageSettings as OOPageSettings } from '../document/DocumentModel'

// ============== 类型定义 ==============

export interface OnlyOfficePreviewProps {
  /** HTML 内容 */
  html?: string
  /** 完整的 docx 解析结果（优先使用） */
  docxData?: DocxParseResult
  /** 页眉（简化版，用于直接传入） */
  header?: HeaderFooterType
  /** 页脚（简化版，用于直接传入） */
  footer?: HeaderFooterType
  /** 页眉 HTML（用于直接传入） */
  headerHtml?: string
  /** 页脚 HTML（用于直接传入） */
  footerHtml?: string
  /** 缩放比例 (0.5 - 2.0) */
  scale?: number
  /** 是否显示页码 */
  showPageNumbers?: boolean
  /** 页面设置 */
  pageSettings?: OOPageSettings | PageSettings
  /** 容器类名 */
  className?: string
  /** 容器样式 */
  style?: React.CSSProperties
  /** 页面变化回调 */
  onPageChange?: (currentPage: number, totalPages: number) => void
}

// ONLYOFFICE 风格参数
const ONLYOFFICE_STYLE = {
  PAGE_GAP_PX: 20,
  BACKGROUND_COLOR: 'var(--word-canvas-bg)',
  PAGE_SHADOW: 'var(--word-page-shadow)',
}

// ============== 主预览组件 ==============

export const OnlyOfficePreview: React.FC<OnlyOfficePreviewProps> = ({
  html,
  docxData,
  header,
  footer,
  headerHtml,
  footerHtml,
  scale = 0.8,
  showPageNumbers = true,
  pageSettings,
  className = '',
  style,
  onPageChange
}) => {
  const containerRef = useRef<HTMLDivElement>(null)
  const [layoutResult, setLayoutResult] = useState<LayoutResult | null>(null)
  const [pageConfig, setPageConfig] = useState<PageConfig>(getDefaultPageConfig(scale))
  const [isLoading, setIsLoading] = useState(true)
  const [currentPage, setCurrentPage] = useState(1)
  const imageCacheRef = useRef<Map<string, HTMLImageElement>>(new Map())
  
  // 更新页面配置
  useEffect(() => {
    let config: PageConfig
    
    if (pageSettings && 'width' in pageSettings) {
      // 使用传入的页面设置
      const ps = pageSettings as PageSettings
      config = {
        width: ps.width * (96/72) * scale,
        height: ps.height * (96/72) * scale,
        marginTop: ps.marginTop * (96/72) * scale,
        marginBottom: ps.marginBottom * (96/72) * scale,
        marginLeft: ps.marginLeft * (96/72) * scale,
        marginRight: ps.marginRight * (96/72) * scale,
        headerHeight: ps.headerHeight * (96/72) * scale,
        footerHeight: ps.footerHeight * (96/72) * scale
      }
    } else if (docxData?.pageSettings) {
      // 从 docxData 获取页面设置
      const ps = docxData.pageSettings
      config = {
        width: ps.width * (96/72) * scale,
        height: ps.height * (96/72) * scale,
        marginTop: ps.marginTop * (96/72) * scale,
        marginBottom: ps.marginBottom * (96/72) * scale,
        marginLeft: ps.marginLeft * (96/72) * scale,
        marginRight: ps.marginRight * (96/72) * scale,
        headerHeight: ps.headerHeight * (96/72) * scale,
        footerHeight: ps.footerHeight * (96/72) * scale
      }
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
        // 获取要布局的 HTML 内容
        let contentHtml = html
        let hHtml = headerHtml
        let fHtml = footerHtml
        
        if (docxData) {
          contentHtml = docxData.bodyHtml
          hHtml = hHtml || docxData.headerHtml
          fHtml = fHtml || docxData.footerHtml
        }
        
        // 如果有 HeaderFooter 对象，转换为 HTML
        if (header && !hHtml) {
          hHtml = headerFooterToHtml(header)
        }
        if (footer && !fHtml) {
          fHtml = headerFooterToHtml(footer)
        }
        
        if (!contentHtml) {
          setLayoutResult({ pages: [], totalHeight: 0 })
          setIsLoading(false)
          return
        }
        
        // 创建布局引擎（复用现有的）
        const engine = createLayoutEngine(pageConfig, scale)
        
        // 执行布局
        const result = engine.layoutFromHtml(contentHtml, hHtml, fHtml)
        
        console.log('[OnlyOfficePreview] 布局完成:', {
          pagesCount: result.pages.length,
          totalHeight: result.totalHeight
        })
        
        setLayoutResult(result)
        
        // 预加载图片
        await preloadImages(contentHtml)
        
      } catch (error) {
        console.error('[OnlyOfficePreview] 布局失败:', error)
        setLayoutResult({ pages: [], totalHeight: 0 })
      } finally {
        setIsLoading(false)
      }
    }
    
    performLayout()
  }, [html, docxData, header, footer, headerHtml, footerHtml, pageConfig, scale])
  
  // 预加载图片
  const preloadImages = async (htmlContent: string) => {
    const imgRegex = /<img[^>]+src=["']([^"']+)["']/gi
    let match
    const promises: Promise<void>[] = []
    
    while ((match = imgRegex.exec(htmlContent)) !== null) {
      const src = match[1]
      if (!imageCacheRef.current.has(src)) {
        const promise = new Promise<void>((resolve) => {
          const img = new Image()
          img.crossOrigin = 'anonymous'
          img.onload = () => {
            imageCacheRef.current.set(src, img)
            resolve()
          }
          img.onerror = () => resolve()
          img.src = src
        })
        promises.push(promise)
      }
    }
    
    await Promise.all(promises)
  }
  
  // 监听滚动以更新当前页码
  useEffect(() => {
    const container = containerRef.current
    if (!container || !layoutResult) return
    
    const handleScroll = () => {
      const scrollTop = container.scrollTop
      const pageHeight = pageConfig.height + ONLYOFFICE_STYLE.PAGE_GAP_PX
      
      const newCurrentPage = Math.floor(scrollTop / pageHeight) + 1
      const clampedPage = Math.max(1, Math.min(newCurrentPage, layoutResult.pages.length))
      
      if (clampedPage !== currentPage) {
        setCurrentPage(clampedPage)
        onPageChange?.(clampedPage, layoutResult.pages.length)
      }
    }
    
    container.addEventListener('scroll', handleScroll)
    return () => container.removeEventListener('scroll', handleScroll)
  }, [layoutResult, pageConfig, currentPage, onPageChange])
  
  // 渲染页面列表
  const renderPages = useMemo(() => {
    if (!layoutResult || layoutResult.pages.length === 0) {
      return (
        <div
          style={{
            width: pageConfig.width,
            height: pageConfig.height,
            backgroundColor: 'var(--word-page-bg)',
            boxShadow: ONLYOFFICE_STYLE.PAGE_SHADOW,
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
        style={{
          marginBottom: ONLYOFFICE_STYLE.PAGE_GAP_PX,
          position: 'relative'
        }}
      >
        {/* 复用现有的 CanvasPageRenderer */}
        <CanvasPageRenderer
          page={page}
          pageConfig={pageConfig}
          scale={scale}
          showPageNumber={showPageNumbers}
          totalPages={layoutResult.pages.length}
          imageCache={imageCacheRef.current}
        />
        
        {/* 页码标签 */}
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
      className={`onlyoffice-preview ${className}`}
      style={{
        display: 'flex',
        flexDirection: 'column',
        alignItems: 'center',
        gap: ONLYOFFICE_STYLE.PAGE_GAP_PX,
        padding: `${ONLYOFFICE_STYLE.PAGE_GAP_PX * 2}px ${ONLYOFFICE_STYLE.PAGE_GAP_PX}px`,
        backgroundColor: ONLYOFFICE_STYLE.BACKGROUND_COLOR,
        minHeight: '100%',
        overflow: 'auto',
        ...style
      }}
    >
      {isLoading ? (
        <div
          style={{
            width: pageConfig.width,
            height: pageConfig.height,
            backgroundColor: 'var(--word-page-bg)',
            boxShadow: ONLYOFFICE_STYLE.PAGE_SHADOW,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
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
                animation: 'onlyoffice-spin 1s linear infinite',
                margin: '0 auto 12px'
              }}
            />
            正在渲染页面...
          </div>
        </div>
      ) : (
        renderPages
      )}
      
      <style>{`
        @keyframes onlyoffice-spin {
          to { transform: rotate(360deg); }
        }
        
        .onlyoffice-preview::-webkit-scrollbar {
          width: 8px;
        }
        
        .onlyoffice-preview::-webkit-scrollbar-track {
          background: #d0d0d0;
        }
        
        .onlyoffice-preview::-webkit-scrollbar-thumb {
          background: #888;
          border-radius: 4px;
        }
        
        .onlyoffice-preview::-webkit-scrollbar-thumb:hover {
          background: #666;
        }
      `}</style>
    </div>
  )
}

// ============== 工具函数 ==============

/**
 * 将 HeaderFooter 对象转换为 HTML
 */
function headerFooterToHtml(hf: HeaderFooterType): string {
  if (!hf.content || hf.content.length === 0) return ''
  
  let html = ''
  for (const element of hf.content) {
    if (element.type === 'paragraph') {
      html += '<p>'
      for (const child of (element as any).children || []) {
        if (child.type === 'text') {
          html += child.text || ''
        }
      }
      html += '</p>'
    }
  }
  
  return html
}

export default OnlyOfficePreview
