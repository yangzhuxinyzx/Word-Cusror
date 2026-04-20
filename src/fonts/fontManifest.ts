export type FontSourceFormat = 'woff2' | 'woff' | 'truetype' | 'opentype'

export type BundledFontSource = {
  /**
   * URL 必须是“相对当前页面”的相对路径，避免 Electron 生产环境 file:// 下绝对路径失效。
   * 推荐：`./fonts/onlyoffice/<file>.woff2`
   */
  url: string
  format: FontSourceFormat
}

export type BundledFontEntry = {
  /** CSS font-family 名称（用于 font-family: 'xxx'） */
  family: string
  /** 400/500/700 或 'normal'/'bold' */
  weight?: number | 'normal' | 'bold'
  style?: 'normal' | 'italic'
  sources: BundledFontSource[]
}

/**
 * 内置字体清单（可选）
 *
 * 你也可以不改这里，直接把 `public/fonts/onlyoffice/manifest.json` 填好；
 * 应用启动时会优先读取 manifest.json。
 */
export const bundledFonts: BundledFontEntry[] = []

/**
 * WordEditor 字体下拉默认会加入这些“常见系统字体”，避免没有任何字体可选。
 * 注意：是否真的存在取决于用户系统/环境。
 */
export const commonSystemFonts: string[] = [
  '等线',
  'DengXian',
  '宋体',
  'SimSun',
  '微软雅黑',
  'Microsoft YaHei',
  '黑体',
  'SimHei',
  '仿宋',
  'FangSong',
  '楷体',
  'KaiTi',
  'Times New Roman',
  'Arial',
  'Calibri',
]

export function normalizeFontFamilyForDisplay(raw: string | null | undefined): string {
  const v = (raw || '').trim()
  if (!v) return ''
  // var(--xxx) 直接返回空，让上层用“文档默认字体”兜底
  if (/^var\(/i.test(v)) return ''
  // 取逗号前的第一个字体名
  const first = v.split(',')[0]?.trim() || ''
  return first.replace(/^['"]+|['"]+$/g, '').trim()
}

export function toChineseDefaultFallbackStack(primary?: string): string {
  const p = (primary || '').trim()
  const quoted = p ? (/[^\w-]/.test(p) ? `"${p.replace(/"/g, '\\"')}"` : `"${p}"`) : ''
  const lower = p.toLowerCase()

  const serifFallback = `"STSong","Songti SC","SimSun",serif`
  const fangsongFallback = `"STFangsong","FangSong","Fangsong SC","STSong","Songti SC",serif`
  const kaitiFallback = `"STKaiti","Kaiti SC","KaiTi","STSong","Songti SC",serif`
  const sansFallback = `"DengXian","PingFang SC","Microsoft YaHei","Heiti SC",sans-serif`

  let fallback = serifFallback
  if (/仿宋|fangsong/.test(lower)) {
    fallback = fangsongFallback
  } else if (/楷体|kaiti/.test(lower)) {
    fallback = kaitiFallback
  } else if (/等线|dengxian|微软雅黑|yahei|黑体|simhei/.test(lower)) {
    fallback = sansFallback
  }

  return quoted ? `${quoted},${fallback}` : serifFallback
}













