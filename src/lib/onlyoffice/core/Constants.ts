/**
 * 常量定义 - 基于 ONLYOFFICE sdkjs-src
 * 
 * 参考文件:
 * - DocumentServer/sdkjs-src/word/Drawing/HtmlPage.js
 * - DocumentServer/sdkjs-src/word/Editor/Document.js
 */

// ============== 屏幕和单位转换 ==============

/** 标准屏幕 DPI */
export const SCREEN_DPI = 96

/** 每英寸毫米数 */
export const MM_PER_INCH = 25.4

/** mm 到 px 转换系数 @ 96 DPI */
export const MM_TO_PX = SCREEN_DPI / MM_PER_INCH  // ≈ 3.7795

/** px 到 mm 转换系数 @ 96 DPI */
export const PX_TO_MM = MM_PER_INCH / SCREEN_DPI  // ≈ 0.2646

/** pt 到 px 转换系数 @ 96 DPI (1pt = 1/72 inch) */
export const PT_TO_PX = SCREEN_DPI / 72  // ≈ 1.333

/** px 到 pt 转换系数 */
export const PX_TO_PT = 72 / SCREEN_DPI  // ≈ 0.75

/** twip 到 px 转换 (1 twip = 1/20 pt) */
export const TWIP_TO_PX = SCREEN_DPI / 72 / 20  // ≈ 0.0667

/** EMU 到 px 转换 (1 inch = 914400 EMU) */
export const EMU_TO_PX = SCREEN_DPI / 914400

// ============== A4 页面尺寸 ==============

/** A4 宽度 (mm) */
export const A4_WIDTH_MM = 210

/** A4 高度 (mm) */
export const A4_HEIGHT_MM = 297

/** A4 宽度 (px @ 96 DPI) */
export const A4_WIDTH_PX = A4_WIDTH_MM * MM_TO_PX  // ≈ 793.7

/** A4 高度 (px @ 96 DPI) */
export const A4_HEIGHT_PX = A4_HEIGHT_MM * MM_TO_PX  // ≈ 1122.5

// ============== Word 默认边距 ==============

/** 上边距 (mm) - 1 inch */
export const DEFAULT_MARGIN_TOP_MM = 25.4

/** 下边距 (mm) - 1 inch */
export const DEFAULT_MARGIN_BOTTOM_MM = 25.4

/** 左边距 (mm) - 1.25 inch (Word 默认) */
export const DEFAULT_MARGIN_LEFT_MM = 31.7

/** 右边距 (mm) - 1.25 inch (Word 默认) */
export const DEFAULT_MARGIN_RIGHT_MM = 31.7

/** 上边距 (px) */
export const DEFAULT_MARGIN_TOP_PX = DEFAULT_MARGIN_TOP_MM * MM_TO_PX  // ≈ 96

/** 下边距 (px) */
export const DEFAULT_MARGIN_BOTTOM_PX = DEFAULT_MARGIN_BOTTOM_MM * MM_TO_PX  // ≈ 96

/** 左边距 (px) */
export const DEFAULT_MARGIN_LEFT_PX = DEFAULT_MARGIN_LEFT_MM * MM_TO_PX  // ≈ 120

/** 右边距 (px) */
export const DEFAULT_MARGIN_RIGHT_PX = DEFAULT_MARGIN_RIGHT_MM * MM_TO_PX  // ≈ 120

// ============== 页眉页脚 ==============

/** 页眉高度 (mm) - 0.5 inch */
export const HEADER_HEIGHT_MM = 12.7

/** 页脚高度 (mm) - 0.5 inch */
export const FOOTER_HEIGHT_MM = 12.7

/** 页眉高度 (px) */
export const HEADER_HEIGHT_PX = HEADER_HEIGHT_MM * MM_TO_PX  // ≈ 48

/** 页脚高度 (px) */
export const FOOTER_HEIGHT_PX = FOOTER_HEIGHT_MM * MM_TO_PX  // ≈ 48

// ============== 字体默认值 ==============

/** 中文默认字体 */
export const DEFAULT_FONT_FAMILY_CN = '"等线", "DengXian", "宋体", "SimSun", serif'

/** 英文默认字体 */
export const DEFAULT_FONT_FAMILY_EN = '"Calibri", "Times New Roman", serif'

/** 默认字号 (pt) - 五号字 */
export const DEFAULT_FONT_SIZE_PT = 10.5

/** 默认字号 (px) */
export const DEFAULT_FONT_SIZE_PX = DEFAULT_FONT_SIZE_PT * PT_TO_PX  // ≈ 14

/** 默认行高倍数 */
export const DEFAULT_LINE_HEIGHT = 1.15

// ============== 页面间距 ==============

/** 页间距 (px) */
export const PAGE_GAP_PX = 20

// ============== 工具函数 ==============

/**
 * 毫米转像素
 */
export function mmToPx(mm: number): number {
  return mm * MM_TO_PX
}

/**
 * 像素转毫米
 */
export function pxToMm(px: number): number {
  return px * PX_TO_MM
}

/**
 * 点转像素
 */
export function ptToPx(pt: number): number {
  return pt * PT_TO_PX
}

/**
 * 像素转点
 */
export function pxToPt(px: number): number {
  return px * PX_TO_PT
}

/**
 * Twip 转像素
 */
export function twipToPx(twip: number): number {
  return twip * TWIP_TO_PX
}

/**
 * EMU 转像素
 */
export function emuToPx(emu: number): number {
  return emu * EMU_TO_PX
}

/**
 * 获取设备像素比
 */
export function getDevicePixelRatio(): number {
  return typeof window !== 'undefined' ? window.devicePixelRatio || 1 : 1
}














