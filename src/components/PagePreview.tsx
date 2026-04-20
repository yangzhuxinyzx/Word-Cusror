import React, { useState, useEffect, useMemo, useRef, useCallback } from 'react'
import type { DocxParseResult, SectionConfig, HeaderFooterContent, HeaderFooterStyle } from '../utils/docxParser'

interface PagePreviewProps {
  docxData: DocxParseResult
  scale?: number  // 缩放比例，默认 1
  showPageNumbers?: boolean
}

// A4 纸尺寸常量 (pt 转 px，96 DPI)
const PT_TO_PX = 96 / 72  // 1pt = 1.333px

export const PagePreview: React.FC<PagePreviewProps> = ({
  docxData,
  scale = 0.8,
  showPageNumbers = true
}) => {
  const [pages, setPages] = useState<string[]>([])
  const contentRef = useRef<HTMLDivElement>(null)
  
  // 获取节配置
  const sectionConfig = useMemo<SectionConfig | undefined>(() => {
    const config = docxData.documentModel?.sections?.[0]
    console.log('[PagePreview] 获取节配置:', {
      hasDocumentModel: !!docxData.documentModel,
      sectionsCount: docxData.documentModel?.sections?.length,
      hasConfig: !!config,
      titlePage: config?.titlePage,
      hasHeaderDefault: !!config?.headerDefault,
      hasFooterDefault: !!config?.footerDefault
    })
    return config
  }, [docxData.documentModel])
  
  // 根据页码获取正确的页眉/页脚（基于 ONLYOFFICE 逻辑）
  const getHeaderFooterForPage = useCallback((
    pageNumber: number, 
    totalPages: number,
    isHeader: boolean
  ): { content: string; style: HeaderFooterStyle } => {
    const section = sectionConfig
    
    // 默认样式
    const defaultStyle: HeaderFooterStyle = {
      fontFamily: 'var(--word-font-family-cn)',
      fontSize: '9pt',
      color: 'var(--word-ink-muted)',
      alignment: isHeader ? 'right' : 'center'
    }
    
    // 如果没有节配置，使用旧版数据
    if (!section) {
      const html = isHeader ? docxData.headerHtml : docxData.footerHtml
      const style = isHeader ? docxData.headerStyle : docxData.footerStyle
      return { 
        content: html
          .replace(/\{PAGE\}/g, String(pageNumber))
          .replace(/\{NUMPAGES\}/g, String(totalPages)),
        style: style || defaultStyle 
      }
    }
    
    let hfContent: HeaderFooterContent | undefined
    
    // 根据 ONLYOFFICE 逻辑选择正确的页眉/页脚
    // 1. 首页：如果 titlePage 启用，使用 First
    // 2. 偶数页：如果 evenAndOddHeaders 启用，使用 Even
    // 3. 其他：使用 Default
    
    if (pageNumber === 1 && section.titlePage) {
      // 首页
      hfContent = isHeader ? section.headerFirst : section.footerFirst
    } else if (pageNumber % 2 === 0 && section.evenAndOddHeaders) {
      // 偶数页
      hfContent = isHeader ? section.headerEven : section.footerEven
    }
    
    // 如果没有特殊配置，使用默认
    if (!hfContent) {
      hfContent = isHeader ? section.headerDefault : section.footerDefault
    }
    
    if (!hfContent) {
      // 没有页眉/页脚内容
      return { 
        content: !isHeader && showPageNumbers ? `<p class="header-footer-para">${pageNumber}</p>` : '',
        style: defaultStyle 
      }
    }
    
    // 替换页码占位符
    const content = hfContent.html
      .replace(/\{PAGE\}/g, String(pageNumber))
      .replace(/\{NUMPAGES\}/g, String(totalPages))
    
    return { content, style: hfContent.style || defaultStyle }
  }, [sectionConfig, docxData, showPageNumbers])
  
  // 计算页面尺寸 (px)
  const pageSize = useMemo(() => {
    const { pageSettings } = docxData
    return {
      width: pageSettings.width * PT_TO_PX * scale,
      height: pageSettings.height * PT_TO_PX * scale,
      marginTop: pageSettings.marginTop * PT_TO_PX * scale,
      marginBottom: pageSettings.marginBottom * PT_TO_PX * scale,
      marginLeft: pageSettings.marginLeft * PT_TO_PX * scale,
      marginRight: pageSettings.marginRight * PT_TO_PX * scale,
      headerHeight: pageSettings.headerHeight * PT_TO_PX * scale,
      footerHeight: pageSettings.footerHeight * PT_TO_PX * scale,
      // 可用内容高度
      contentHeight: (pageSettings.height - pageSettings.marginTop - pageSettings.marginBottom) * PT_TO_PX * scale
    }
  }, [docxData.pageSettings, scale])
  
  // 改进的分页算法（参考 ONLYOFFICE Recalculate 逻辑）
  useEffect(() => {
    if (!contentRef.current) return
    
    // 临时创建一个隐藏的测量容器
    const measureContainer = document.createElement('div')
    measureContainer.style.cssText = `
      position: absolute;
      visibility: hidden;
      width: ${pageSize.width - pageSize.marginLeft - pageSize.marginRight}px;
      font-size: ${14 * scale}px;
      font-family: var(--word-font-family-cn);
      line-height: var(--word-line-height);
    `
    measureContainer.innerHTML = docxData.bodyHtml
    document.body.appendChild(measureContainer)
    
    const pageContents: string[] = []
    let currentPageContent = ''
    let currentHeight = 0
    const maxHeight = pageSize.contentHeight
    
    // 获取所有顶层元素
    const elements = measureContainer.children
    
    // 检查元素是否是分页符
    const isPageBreak = (el: HTMLElement): boolean => {
      return el.classList.contains('page-break') || 
             el.tagName === 'HR' && el.classList.contains('page-break') ||
             el.style.pageBreakBefore === 'always' ||
             el.style.pageBreakAfter === 'always'
    }
    
    // 检查元素是否不可分割（表格、带有 keep-together 的元素）
    const isNonBreakable = (el: HTMLElement): boolean => {
      const tagName = el.tagName.toLowerCase()
      // 表格尽量不在中间断开
      if (tagName === 'table') return true
      // 带有特定样式的元素
      if (el.style.pageBreakInside === 'avoid') return true
      return false
    }
    
    for (let i = 0; i < elements.length; i++) {
      const element = elements[i] as HTMLElement
      
      // 检查是否是分页符
      if (isPageBreak(element)) {
        // 遇到分页符，强制换页
        if (currentPageContent) {
          pageContents.push(currentPageContent)
          currentPageContent = ''
          currentHeight = 0
        }
        continue
      }
      
      const elementHeight = element.offsetHeight
      const willExceed = currentHeight + elementHeight > maxHeight
      
      // 如果当前元素加上已有内容超过页面高度
      if (willExceed && currentPageContent) {
        // 检查是否是不可分割元素
        if (isNonBreakable(element)) {
          // 不可分割元素：如果元素本身就超过页面高度，还是要放
          if (elementHeight > maxHeight) {
            // 元素太大，无法放入单页，直接放
            currentPageContent += element.outerHTML
            currentHeight += elementHeight
          } else {
            // 元素可以放入单页，先换页再放
            pageContents.push(currentPageContent)
            currentPageContent = element.outerHTML
            currentHeight = elementHeight
          }
        } else {
          // 可分割元素：换页后放
          pageContents.push(currentPageContent)
          currentPageContent = element.outerHTML
          currentHeight = elementHeight
        }
      } else {
        // 添加元素到当前页
        currentPageContent += element.outerHTML
        currentHeight += elementHeight
      }
    }
    
    // 保存最后一页
    if (currentPageContent) {
      pageContents.push(currentPageContent)
    }
    
    // 至少保证有一页
    if (pageContents.length === 0) {
      pageContents.push(docxData.bodyHtml || '<p></p>')
    }
    
    document.body.removeChild(measureContainer)
    setPages(pageContents)
  }, [docxData.bodyHtml, pageSize, scale])
  
  // 将 HeaderFooterStyle 转换为 React.CSSProperties
  // 参考 ONLYOFFICE 的页眉/页脚样式：
  // - 页眉有下划线分隔
  // - 页眉内容靠右对齐
  // - 页脚居中
  // - 小字体、灰色文字
  const styleToCSS = useCallback((style: HeaderFooterStyle, isHeader: boolean): React.CSSProperties => {
    return {
      fontFamily: style.fontFamily || 'var(--word-font-family-cn)',
      fontSize: style.fontSize || '9pt',
      color: style.color || 'var(--word-ink-muted)',
      textAlign: style.alignment || (isHeader ? 'right' : 'center'),
      // 页眉始终有下划线（参考 ONLYOFFICE/Word 默认样式）
      borderBottom: isHeader ? '0.5pt solid var(--word-rule)' : 'none',
      borderTop: !isHeader ? 'none' : 'none',
      paddingBottom: isHeader ? '4px' : '0',
      paddingTop: !isHeader ? '4px' : '0',
      marginBottom: '0',
      marginTop: '0',
      lineHeight: style.lineHeight || '1'
    }
  }, [])
  
  // 渲染单个页面
  const renderPage = (content: string, pageNumber: number, totalPages: number) => {
    // 根据页码获取正确的页眉/页脚
    const header = getHeaderFooterForPage(pageNumber, totalPages, true)
    const footer = getHeaderFooterForPage(pageNumber, totalPages, false)
    
    const headerCSSStyle = styleToCSS(header.style, true)
    const footerCSSStyle = styleToCSS(footer.style, false)
    
    return (
      <div 
        key={pageNumber}
        className="page-preview-page"
        style={{
          width: pageSize.width,
          minHeight: pageSize.height,
          backgroundColor: 'var(--word-page-bg)',
          // 参考 Word/ONLYOFFICE 的页面阴影效果 - 更真实的纸张感
          boxShadow: 'var(--word-page-shadow)',
          marginBottom: '24px',
          position: 'relative',
          display: 'flex',
          flexDirection: 'column',
          // 纸张圆角（非常轻微）
          borderRadius: '1px',
          // 纸张边缘
          border: '1px solid var(--word-page-border)'
        }}
      >
        {/* 页眉区域 */}
        <div 
          className="page-header"
          style={{
            padding: `${Math.max(pageSize.marginTop - pageSize.headerHeight, 0)}px ${pageSize.marginRight}px 0 ${pageSize.marginLeft}px`,
            minHeight: pageSize.marginTop,
            ...headerCSSStyle
          }}
        >
          <div dangerouslySetInnerHTML={{ __html: header.content }} />
        </div>
        
        {/* 正文区域 */}
        <div 
          className="page-content"
          style={{
            flex: 1,
            padding: `0 ${pageSize.marginRight}px 0 ${pageSize.marginLeft}px`,
            overflow: 'hidden',
            minHeight: pageSize.contentHeight,
            fontSize: `${14 * scale}px`,
            lineHeight: 'var(--word-line-height)'
          }}
        >
          <div 
            className="word-editor-content"
            dangerouslySetInnerHTML={{ __html: content }} 
          />
        </div>
        
        {/* 页脚区域 */}
        <div 
          className="page-footer"
          style={{
            padding: `0 ${pageSize.marginRight}px ${Math.max(pageSize.marginBottom - pageSize.footerHeight, 0)}px ${pageSize.marginLeft}px`,
            minHeight: pageSize.marginBottom,
            ...footerCSSStyle
          }}
        >
          <div dangerouslySetInnerHTML={{ __html: footer.content }} />
        </div>
        
        {/* 页码标签 */}
        {showPageNumbers && (
          <div 
            style={{
              position: 'absolute',
              bottom: '-25px',
              left: '50%',
              transform: 'translateX(-50%)',
              fontSize: '12px',
              color: 'var(--word-ink-muted)'
            }}
          >
            第 {pageNumber} 页 / 共 {totalPages} 页
          </div>
        )}
      </div>
    )
  }
  
  return (
    <div 
      className="page-preview-container"
      style={{
        display: 'flex',
        flexDirection: 'column',
        alignItems: 'center',
        // 参考 ONLYOFFICE 的画布背景 - 中性灰色
        padding: '40px 20px',
        backgroundColor: 'var(--word-canvas-bg)',
        minHeight: '100%',
        overflow: 'auto',
        // 添加渐变背景增加深度感
        background: 'var(--word-canvas-bg)'
      }}
    >
      {/* 隐藏的测量容器 */}
      <div ref={contentRef} style={{ display: 'none' }} />
      
      {/* 渲染所有页面 */}
      {pages.map((content, index) => 
        renderPage(content, index + 1, pages.length)
      )}
      
      {/* 如果没有内容，显示空白页 */}
      {pages.length === 0 && renderPage('<p></p>', 1, 1)}
    </div>
  )
}

export default PagePreview

