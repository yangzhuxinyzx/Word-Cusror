/**
 * ONLYOFFICE 风格渲染引擎
 * 
 * 基于 ONLYOFFICE Document Server 的设计思想实现的独立预览渲染系统
 * 无需安装任何外部服务，完全在前端运行
 */

// Core
export { Graphics, createGraphics } from './core/Graphics'
export { TextMeasurer, getTextMeasurer } from './core/TextMeasurer'
export * from './core/Constants'

// Document
export { 
  DocumentModel, 
  getDefaultPageSettings, 
  getDefaultTextStyle, 
  getDefaultParagraphStyle 
} from './document/DocumentModel'
export type {
  DocumentElement,
  ParagraphElement,
  ImageElement,
  TableElement,
  ListElement,
  TextRun,
  HeaderFooter,
  PageSettings,
  LayoutElement,
  LayoutPage
} from './document/DocumentModel'

export { PageLayoutEngine, createPageLayoutEngine } from './document/PageLayout'
export type { LayoutOptions, LayoutResult } from './document/PageLayout'

export { ElementRenderer, createElementRenderer } from './document/ElementRenderer'
export type { RenderOptions } from './document/ElementRenderer'

// Components
export { OnlyOfficePreview } from './components/OnlyOfficePreview'
export type { OnlyOfficePreviewProps } from './components/OnlyOfficePreview'














