/**
 * DocDsl - 文档结构化 DSL 类型定义
 * 用于模型输出的结构化文档描述，替代 HTML/Markdown
 */

// ============== 基础类型 ==============

/** 对齐方式 */
export type DslAlignment = 'left' | 'center' | 'right' | 'justify'

/** 垂直对齐 */
export type DslVerticalAlign = 'top' | 'middle' | 'bottom'

/** 边框样式 */
export type DslBorderStyle = 'none' | 'single' | 'double' | 'dashed' | 'dotted' | 'thick'

/** 列表类型 */
export type DslListType = 'bullet' | 'number' | 'letter' | 'roman'

/** 图片环绕方式 */
export type DslImageWrap = 'inline' | 'square' | 'tight' | 'behind' | 'front'

/** 长度单位：支持 pt, px, cm, in, em */
export type DslLength = string | number // e.g. "12pt", "2em", 24 (默认pt)

/** 颜色：支持 hex (#RRGGBB), rgb(), 或颜色名 */
export type DslColor = string // e.g. "#FF0000", "rgb(255,0,0)", "red"

// ============== 行内格式（Run） ==============

/** 运行时元数据（diff/track/comments，不持久化到 DOCX） */
export interface DslRunMeta {
  /** diff 类型 */
  diffType?: 'old' | 'new'
  /** diff ID */
  diffId?: string
  /** 修订类型 */
  trackType?: 'insert' | 'delete'
  /** 修订 ID */
  trackId?: string
  /** 修订作者 */
  trackAuthor?: string
  /** 修订日期 */
  trackDate?: string
  /** 批注 ID 列表 */
  commentIds?: string[]
}

/** 块级元数据 */
export interface DslBlockMeta {
  /** diff 角色（新增块） */
  diffRole?: 'new'
  /** diff ID */
  diffId?: string
}

/** 行内文本片段 */
export interface DslRun {
  /** 文本内容 */
  text: string
  /** 粗体 */
  bold?: boolean
  /** 斜体 */
  italic?: boolean
  /** 下划线 */
  underline?: boolean
  /** 删除线 */
  strikethrough?: boolean
  /** 上标 */
  superscript?: boolean
  /** 下标 */
  subscript?: boolean
  /** 字体 */
  fontFamily?: string
  /** 字号 (pt) */
  fontSize?: number
  /** 文字颜色 */
  color?: DslColor
  /** 高亮/背景色 */
  highlight?: DslColor
  /** 字符间距 (pt) */
  letterSpacing?: number
  /** 运行时元数据（diff/track/comments） */
  _meta?: DslRunMeta
}

/** 行内元素：可以是纯文本或带格式的 Run */
export type DslInline = string | DslRun

// ============== 段落格式 ==============

/** 段落格式属性 */
export interface DslParagraphFormat {
  /** 对齐方式 */
  alignment?: DslAlignment
  /** 首行缩进 */
  firstLineIndent?: DslLength
  /** 左缩进 */
  leftIndent?: DslLength
  /** 右缩进 */
  rightIndent?: DslLength
  /** 悬挂缩进 */
  hangingIndent?: DslLength
  /** 段前间距 */
  spaceBefore?: DslLength
  /** 段后间距 */
  spaceAfter?: DslLength
  /** 行距（倍数或固定值） */
  lineHeight?: number | DslLength
  /** 边框 */
  border?: DslBorder
  /** 背景色 */
  backgroundColor?: DslColor
  /** 内边距 */
  padding?: DslLength
}

// ============== 边框 ==============

/** 单边边框 */
export interface DslBorderSide {
  style?: DslBorderStyle
  width?: DslLength
  color?: DslColor
}

/** 四边边框 */
export interface DslBorder {
  top?: DslBorderSide
  bottom?: DslBorderSide
  left?: DslBorderSide
  right?: DslBorderSide
  /** 简写：同时设置四边 */
  all?: DslBorderSide
}

// ============== 块级元素 ==============

/** 标题块 */
export interface DslHeading {
  type: 'heading'
  /** 标题级别 1-6 */
  level: 1 | 2 | 3 | 4 | 5 | 6
  /** 内容：可以是纯文本或带格式的行内元素数组 */
  content: string | DslInline[]
  /** 覆盖格式 */
  format?: DslParagraphFormat
  /** 块级元数据 */
  _meta?: DslBlockMeta
}

/** 段落块 */
export interface DslParagraph {
  type: 'paragraph'
  /** 内容：可以是纯文本或带格式的行内元素数组 */
  content: string | DslInline[]
  /** 段落格式 */
  format?: DslParagraphFormat
  /** 块级元数据 */
  _meta?: DslBlockMeta
}

/** 列表项 */
export interface DslListItem {
  /** 内容 */
  content: string | DslInline[]
  /** 嵌套子列表 */
  children?: DslListItem[]
}

/** 列表块 */
export interface DslList {
  type: 'list'
  /** 列表类型 */
  listType: DslListType
  /** 起始编号（仅对 number/letter/roman 有效） */
  startAt?: number
  /** 列表项 */
  items: DslListItem[]
  /** 缩进级别 */
  level?: number
}

/** 表格单元格 */
export interface DslTableCell {
  /** 内容：可以是纯文本、行内元素数组、或嵌套块 */
  content: string | DslInline[] | DslBlock[]
  /** 跨列数 */
  colSpan?: number
  /** 跨行数 */
  rowSpan?: number
  /** 水平对齐 */
  align?: DslAlignment
  /** 垂直对齐 */
  valign?: DslVerticalAlign
  /** 背景色 */
  backgroundColor?: DslColor
  /** 边框 */
  border?: DslBorder
  /** 宽度 */
  width?: DslLength
}

/** 表格行 */
export interface DslTableRow {
  /** 单元格 */
  cells: DslTableCell[]
  /** 行高 */
  height?: DslLength
  /** 是否为表头行 */
  isHeader?: boolean
}

/** 表格块 */
export interface DslTable {
  type: 'table'
  /** 行 */
  rows: DslTableRow[]
  /** 列宽数组 */
  columnWidths?: DslLength[]
  /** 表格宽度（百分比或固定值） */
  width?: DslLength | string
  /** 表格对齐 */
  alignment?: DslAlignment
  /** 默认边框样式 */
  border?: DslBorder
  /** 表头行重复（跨页） */
  repeatHeader?: boolean
}

/** 图片块 */
export interface DslImage {
  type: 'image'
  /** 图片源：URL 或 base64 */
  src: string
  /** 替代文本 */
  alt?: string
  /** 宽度 */
  width?: DslLength
  /** 高度 */
  height?: DslLength
  /** 对齐方式 */
  alignment?: DslAlignment
  /** 环绕方式 */
  wrap?: DslImageWrap
  /** 标题/说明 */
  caption?: string
}

/** 分页符 */
export interface DslPageBreak {
  type: 'pageBreak'
}

/** 分节符 */
export interface DslSectionBreak {
  type: 'sectionBreak'
  /** 分节类型 */
  breakType?: 'nextPage' | 'continuous' | 'evenPage' | 'oddPage'
}

/** 水平线 */
export interface DslHorizontalRule {
  type: 'horizontalRule'
  /** 样式 */
  style?: DslBorderStyle
  /** 颜色 */
  color?: DslColor
  /** 粗细 */
  width?: DslLength
}

/** 引用块 */
export interface DslBlockquote {
  type: 'blockquote'
  /** 内容块 */
  content: DslBlock[]
}

/** 所有块类型的联合 */
export type DslBlock =
  | DslHeading
  | DslParagraph
  | DslList
  | DslTable
  | DslImage
  | DslPageBreak
  | DslSectionBreak
  | DslHorizontalRule
  | DslBlockquote

// ============== 页面设置 ==============

/** 页边距 */
export interface DslMargins {
  top?: DslLength
  bottom?: DslLength
  left?: DslLength
  right?: DslLength
  gutter?: DslLength
}

/** 纸张大小 */
export type DslPaperSize = 'A4' | 'A3' | 'Letter' | 'Legal' | 'custom'

/** 页面设置 */
export interface DslPageSetup {
  /** 纸张大小 */
  paperSize?: DslPaperSize
  /** 自定义宽度（paperSize 为 custom 时） */
  width?: DslLength
  /** 自定义高度（paperSize 为 custom 时） */
  height?: DslLength
  /** 页面方向 */
  orientation?: 'portrait' | 'landscape'
  /** 页边距 */
  margins?: DslMargins
}

/** 页眉/页脚 */
export interface DslHeaderFooter {
  /** 页眉内容 */
  header?: {
    content: string | DslInline[]
    alignment?: DslAlignment
    showOnFirstPage?: boolean
  }
  /** 页脚内容 */
  footer?: {
    content: string | DslInline[]
    alignment?: DslAlignment
    showOnFirstPage?: boolean
  }
  /** 页码设置 */
  pageNumber?: {
    enabled: boolean
    position: 'header' | 'footer'
    alignment?: DslAlignment
    format?: 'arabic' | 'roman' | 'letter'
    startFrom?: number
  }
}

// ============== 样式定义 ==============

/** 样式定义 */
export interface DslStyleDef {
  /** 样式名称 */
  name: string
  /** 基于哪个样式 */
  basedOn?: string
  /** 段落格式 */
  paragraph?: DslParagraphFormat
  /** 字符格式 */
  run?: Omit<DslRun, 'text'>
}

// ============== 文档根节点 ==============

/** 文档 DSL 根节点 */
export interface DocDsl {
  /** DSL 版本 */
  version?: '1.0'
  /** 文档标题 */
  title?: string
  /** 页面设置 */
  pageSetup?: DslPageSetup
  /** 页眉页脚 */
  headerFooter?: DslHeaderFooter
  /** 自定义样式 */
  styles?: DslStyleDef[]
  /** 文档内容块 */
  blocks: DslBlock[]
}

// ============== 编辑操作 ==============

/** DSL 编辑操作类型 */
export type DslEditOpType =
  | 'insert_block'      // 插入块
  | 'delete_block'      // 删除块
  | 'replace_text'      // 替换文本
  | 'format_text'       // 格式化文本
  | 'format_paragraph'  // 格式化段落
  | 'insert_table'      // 插入表格
  | 'modify_table'      // 修改表格（合并、边框等）
  | 'insert_image'      // 插入图片
  | 'modify_image'      // 修改图片
  | 'set_page_setup'    // 设置页面
  | 'set_header_footer' // 设置页眉页脚

/** 目标定位 */
export interface DslEditTarget {
  /** 定位方式 */
  scope: 'document' | 'selection' | 'anchor' | 'index'
  /** 锚点文本（scope 为 anchor 时） */
  anchorText?: string
  /** 块索引（scope 为 index 时） */
  blockIndex?: number
  /** 位置：之前/之后（用于插入） */
  position?: 'before' | 'after' | 'start' | 'end'
}

/** 编辑操作 */
export interface DslEditOp {
  /** 操作类型 */
  type: DslEditOpType
  /** 目标定位 */
  target: DslEditTarget
  /** 操作参数 */
  params: Record<string, unknown>
  /** 预览模式（不实际执行） */
  dryRun?: boolean
}

// ============== 工具类型 ==============

/** DSL 校验结果 */
export interface DslValidationResult {
  valid: boolean
  errors: DslValidationError[]
}

/** DSL 校验错误 */
export interface DslValidationError {
  path: string
  message: string
  code: string
}

/** DSL 渲染结果 */
export interface DslRenderResult {
  html: string
  warnings?: string[]
}
