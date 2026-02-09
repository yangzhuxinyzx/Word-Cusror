import JSZip from 'jszip'
import mammoth from 'mammoth'

// 图片数据映射类型
interface ImageMap {
  [rId: string]: string  // rId -> base64 data URL
}

// NOTE:
// We intentionally keep images as `data:` URLs for docx HTML.
// Reason: DocumentContext caches/restores HTML across reloads, and `blob:` URLs are NOT stable across reloads,
// leading to `net::ERR_FILE_NOT_FOUND` when restoring cached content.

// 给 Agent 用的图片元信息（不传二进制）
export interface DocxImageMeta {
  rId: string
  target?: string
  widthPx?: number
  heightPx?: number
  alt?: string
  floating?: boolean
}

type ImageParseContext = {
  embedImages: boolean
  relsMap?: Record<string, string>
  images?: DocxImageMeta[]
}

// 脚注/尾注数据类型
interface FootnoteData {
  id: string
  content: string  // HTML 内容
}

interface FootnoteMap {
  footnotes: Map<string, FootnoteData>
  endnotes: Map<string, FootnoteData>
}

// MIME 类型映射
const IMAGE_MIME_TYPES: Record<string, string> = {
  'png': 'image/png',
  'jpg': 'image/jpeg',
  'jpeg': 'image/jpeg',
  'gif': 'image/gif',
  'bmp': 'image/bmp',
  'webp': 'image/webp',
  'svg': 'image/svg+xml',
  'tiff': 'image/tiff',
  'tif': 'image/tiff',
  'emf': 'image/emf',
  'wmf': 'image/wmf',
}

// EMU 到像素的转换（96 DPI）
const EMU_PER_PIXEL = 9525

function emuToPixels(emu: number): number {
  return Math.round(emu / EMU_PER_PIXEL)
}

// Word 字号到 pt 的映射
const WORD_FONT_SIZE_MAP: Record<number, string> = {
  // 半点值 (half-points) 到 pt
  20: '10pt',   // 五号
  21: '10.5pt', // 五号半
  24: '12pt',   // 小四
  28: '14pt',   // 四号
  30: '15pt',   // 小三
  32: '16pt',   // 三号
  36: '18pt',   // 小二
  44: '22pt',   // 二号
  52: '26pt',   // 小一
  72: '36pt',   // 一号
}

// 将 Word 的 half-points 转换为 pt
function halfPointsToPt(halfPoints: number): string {
  if (WORD_FONT_SIZE_MAP[halfPoints]) {
    return WORD_FONT_SIZE_MAP[halfPoints]
  }
  return `${halfPoints / 2}pt`
}

interface RunStyle {
  bold?: boolean
  italic?: boolean
  underline?: boolean
  underlineStyle?: 'solid' | 'dotted' | 'dashed' | 'double' | 'wavy'
  strike?: boolean
  fontSize?: string
  fontFamily?: string  // CSS font-family 值（包含回退字体栈）
  fontName?: string    // 原始字体名（用于 UI 显示）
  color?: string
  highlight?: string
}

interface ParagraphStyle {
  alignment?: 'left' | 'center' | 'right' | 'justify'
  indent?: number
  heading?: number
  fontSize?: string
  fontFamily?: string
  color?: string
  lineHeight?: string
  marginTop?: string
  marginBottom?: string
}

// 字体名标准化映射：将常见的字体别名统一为首选名称
// 确保 DOCX 中的字体名能匹配到 @font-face 定义
const FONT_NAME_NORMALIZE_MAP: Record<string, string> = {
  // 等线系列 - 统一为中文名（与 manifest.json 对应）
  'DengXian': '等线',
  'DengXian Light': '等线',
  '等线 Light': '等线',
  
  // Word 主题字体格式（带括号的格式）
  // Word 显示 "宋体 (中文正文)" 表示使用主题的正文字体
  // 有些情况下可能直接出现在 XML 中
  '宋体 (中文正文)': '宋体',
  '宋体(中文正文)': '宋体',
  '宋体 (中文标题)': '宋体',
  '宋体(中文标题)': '宋体',
  '等线 (中文正文)': '等线',
  '等线(中文正文)': '等线',
  '等线 (中文标题)': '等线',
  '等线(中文标题)': '等线',
  '微软雅黑 (中文正文)': '微软雅黑',
  '微软雅黑(中文正文)': '微软雅黑',
  '微软雅黑 (中文标题)': '微软雅黑',
  '微软雅黑(中文标题)': '微软雅黑',
  '黑体 (中文正文)': '黑体',
  '黑体(中文正文)': '黑体',
  '黑体 (中文标题)': '黑体',
  '黑体(中文标题)': '黑体',
  '楷体 (中文正文)': '楷体',
  '楷体(中文正文)': '楷体',
  '仿宋 (中文正文)': '仿宋',
  '仿宋(中文正文)': '仿宋',
  
  // 宋体系列
  'SimSun': '宋体',
  'NSimSun': '新宋体',
  'Songti SC': '宋体',
  '宋体-简': '宋体',
  
  // 黑体系列
  'SimHei': '黑体',
  'Heiti SC': '黑体',
  '黑体-简': '黑体',
  
  // 微软雅黑系列
  'Microsoft YaHei': '微软雅黑',
  'Microsoft YaHei Light': '微软雅黑',
  'Microsoft YaHei UI': '微软雅黑',
  
  // 仿宋系列
  'FangSong': '仿宋',
  'FangSong_GB2312': '仿宋_GB2312',
  'Fangsong SC': '仿宋',
  
  // 楷体系列
  'KaiTi': '楷体',
  'KaiTi_GB2312': '楷体_GB2312',
  'Kaiti SC': '楷体',
  
  // 华文系列
  'STSong': '华文宋体',
  'STFangsong': '华文仿宋',
  'STFANGSO': '华文仿宋',
  'STKaiti': '华文楷体',
  'STKAITI': '华文楷体',
  'STXihei': '华文细黑',
  'STXIHEI': '华文细黑',
  'STHeiti': '华文细黑',
  'STHEITI': '华文细黑',
  'STZhongsong': '华文中宋',
  'STXinwei': '华文新魏',
  'STXingkai': '华文行楷',
  'STHupo': '华文琥珀',
  'STLiti': '华文隶书',
  'STCaiyun': '华文彩云',
  
  // 其他中文字体
  'LiSu': '隶书',
  'YouYuan': '幼圆',
  'FZYaoti': '方正姚体',
  'FZShuTi': '方正舒体',
}

/**
 * 标准化字体名称，确保与 @font-face 定义匹配
 * @param fontName 原始字体名
 * @returns 标准化后的字体名
 */
function normalizeFontName(fontName: string | null | undefined): string {
  if (!fontName) return ''
  let trimmed = fontName.trim()
  if (!trimmed) return ''
  
  // 检查是否有直接映射
  if (FONT_NAME_NORMALIZE_MAP[trimmed]) {
    return FONT_NAME_NORMALIZE_MAP[trimmed]
  }
  
  // 处理 Word 主题字体格式：去掉括号后缀如 "(中文正文)"、"(中文标题)"、"(西文正文)" 等
  // 这些是 Word UI 显示的标记，不是实际字体名的一部分
  const themeMarkerMatch = trimmed.match(/^(.+?)\s*[\(（].*?[\)）]\s*$/)
  if (themeMarkerMatch) {
    trimmed = themeMarkerMatch[1].trim()
    // 去掉括号后再次检查映射
    if (FONT_NAME_NORMALIZE_MAP[trimmed]) {
      return FONT_NAME_NORMALIZE_MAP[trimmed]
    }
  }
  
  // 检查小写映射（兼容大小写不一致）
  const lower = trimmed.toLowerCase()
  for (const [key, value] of Object.entries(FONT_NAME_NORMALIZE_MAP)) {
    if (key.toLowerCase() === lower) {
      return value
    }
  }
  
  return trimmed
}

// 字体映射 - 将 Word 字体映射到系统可用字体
const FONT_FALLBACK_MAP: Record<string, string> = {
  // 宋体系列
  '宋体': '"宋体", "SimSun", "Songti SC", serif',
  'SimSun': '"宋体", "SimSun", "Songti SC", serif',
  '新宋体': '"新宋体", "NSimSun", "宋体", serif',
  'NSimSun': '"新宋体", "NSimSun", "宋体", serif',
  '华文宋体': '"华文宋体", "STSong", "宋体", serif',
  'STSong': '"华文宋体", "STSong", "宋体", serif',
  '方正小标宋简体': '"方正小标宋简体", "宋体", serif',
  '方正小标宋_GBK': '"方正小标宋_GBK", "宋体", serif',
  
  // 黑体系列
  '黑体': '"黑体", "SimHei", "Microsoft YaHei", "微软雅黑", "Heiti SC", sans-serif',
  'SimHei': '"黑体", "SimHei", "Microsoft YaHei", "微软雅黑", "Heiti SC", sans-serif',
  '华文黑体': '"华文黑体", "STHeiti", "黑体", sans-serif',
  'STHeiti': '"华文黑体", "STHeiti", "黑体", sans-serif',
  '微软雅黑': '"微软雅黑", "Microsoft YaHei", "黑体", sans-serif',
  'Microsoft YaHei': '"微软雅黑", "Microsoft YaHei", "黑体", sans-serif',
  
  // 楷体系列
  '楷体': '"楷体", "STKAITI", "KaiTi", "Kaiti SC", serif',
  'KaiTi': '"楷体", "KaiTi", "Kaiti SC", serif',
  '楷体_GB2312': '"楷体_GB2312", "楷体", "KaiTi", serif',
  '华文楷体': '"华文楷体", "STKAITI", "STKaiti", "楷体", serif',
  'STKaiti': '"华文楷体", "STKAITI", "STKaiti", "楷体", serif',
  'STKAITI': '"华文楷体", "STKAITI", "STKaiti", "楷体", serif',
  
  // 仿宋系列
  '仿宋': '"仿宋", "STFANGSO", "FangSong", "Fangsong SC", serif',
  'FangSong': '"仿宋", "FangSong", "Fangsong SC", serif',
  '仿宋_GB2312': '"仿宋_GB2312", "仿宋", "FangSong", serif',
  '华文仿宋': '"华文仿宋", "STFANGSO", "STFangsong", "仿宋", serif',
  'STFangsong': '"华文仿宋", "STFANGSO", "STFangsong", "仿宋", serif',
  'STFANGSO': '"华文仿宋", "STFANGSO", "STFangsong", "仿宋", serif',
  
  // 其他常用字体
  '华文中宋': '"华文中宋", "STZhongsong", "宋体", serif',
  'STZhongsong': '"华文中宋", "STZhongsong", "宋体", serif',
  '华文细黑': '"华文细黑", "STXIHEI", "STXihei", "黑体", sans-serif',
  '等线': '"等线", "DengXian", "微软雅黑", sans-serif',
  'DengXian': '"等线", "DengXian", "微软雅黑", sans-serif',
  
  // 英文字体
  'Times New Roman': '"Times New Roman", "宋体", serif',
  'Arial': '"Arial", "黑体", sans-serif',
  'Calibri': '"Calibri", "等线", sans-serif',
}

type ThemeFontMap = Record<string, string>

let currentThemeFontMap: ThemeFontMap | null = null

function getWAttr(el: Element, localName: string): string | null {
  return el.getAttribute(`w:${localName}`) || el.getAttribute(localName)
}

/**
 * 判断颜色是否应该在深色模式下被忽略
 * 黑色或接近黑色的颜色在深色模式下应该忽略，让 CSS 自动处理
 * 但保留明显的强调色（如红色、蓝色等）
 */
function shouldIgnoreColorInDarkMode(colorHex: string): boolean {
  // 忽略黑色/接近黑色的颜色（亮度 < 30），让 CSS 变量控制文字色
  // 保留鲜艳的强调色（红、蓝、绿等）
  const hex = colorHex.replace('#', '').toLowerCase()
  if (hex === '000000' || hex === '000' || hex === 'auto') return true
  if (hex.length < 6) return false
  const r = parseInt(hex.slice(0, 2), 16)
  const g = parseInt(hex.slice(2, 4), 16)
  const b = parseInt(hex.slice(4, 6), 16)
  if (isNaN(r) || isNaN(g) || isNaN(b)) return false
  // 亮度低于 30（接近纯黑）→ 忽略
  const brightness = (r * 299 + g * 587 + b * 114) / 1000
  return brightness < 30
}

function resolveThemeFont(themeKey: string | null | undefined): string {
  if (!themeKey) return ''
  return currentThemeFontMap?.[themeKey] || ''
}

function resolveFontNameFromRFonts(rFonts: Element): string {
  // Prefer EastAsia (Chinese) then ASCII/HAnsi (Latin), then CS (complex scripts)
  // 所有返回值都经过 normalizeFontName 标准化，确保与 @font-face 匹配
  const eastAsia = getWAttr(rFonts, 'eastAsia')
  const eastAsiaTheme = getWAttr(rFonts, 'eastAsiaTheme')
  const ascii = getWAttr(rFonts, 'ascii')
  const asciiTheme = getWAttr(rFonts, 'asciiTheme')
  if (eastAsia) return normalizeFontName(eastAsia)
  const eastAsiaFromTheme = resolveThemeFont(eastAsiaTheme)
  if (eastAsiaFromTheme) return normalizeFontName(eastAsiaFromTheme)

  if (ascii) return normalizeFontName(ascii)
  const asciiFromTheme = resolveThemeFont(asciiTheme)
  if (asciiFromTheme) return normalizeFontName(asciiFromTheme)

  const hAnsi = getWAttr(rFonts, 'hAnsi')
  if (hAnsi) return normalizeFontName(hAnsi)
  const hAnsiTheme = getWAttr(rFonts, 'hAnsiTheme')
  const hAnsiFromTheme = resolveThemeFont(hAnsiTheme)
  if (hAnsiFromTheme) return normalizeFontName(hAnsiFromTheme)

  const cs = getWAttr(rFonts, 'cs')
  if (cs) return normalizeFontName(cs)
  const csTheme = getWAttr(rFonts, 'csTheme')
  const csFromTheme = resolveThemeFont(csTheme)
  if (csFromTheme) return normalizeFontName(csFromTheme)

  return ''
}

async function loadThemeFontMap(zip: JSZip): Promise<ThemeFontMap | null> {
  try {
    const themeXml =
      (await zip.file('word/theme/theme1.xml')?.async('string')) ||
      (await zip.file('word/theme/theme.xml')?.async('string'))
    if (!themeXml) return null

    const parser = new DOMParser()
    const doc = parser.parseFromString(themeXml, 'application/xml')
    const parseError = doc.querySelector('parsererror')
    if (parseError) return null

    const root = doc.documentElement
    const majorFont = findElementByLocalName(root as any, 'majorFont')
    const minorFont = findElementByLocalName(root as any, 'minorFont')

    const getTypeface = (fontEl: Element | null, localName: string) => {
      if (!fontEl) return ''
      const el = findElementByLocalName(fontEl, localName)
      return el?.getAttribute('typeface') || ''
    }

    const normalizeThemeTypeface = (v: string) => {
      const t = (v || '').trim()
      // In many Office themes, typeface may be a placeholder like "+mn-ea", "+mj-lt"
      if (!t) return ''
      if (t.startsWith('+')) return ''
      return t
    }

    const getTypefaceByScript = (fontEl: Element | null, scripts: string[]) => {
      if (!fontEl) return ''
      const fontNodes = findAllElementsByLocalName(fontEl, 'font')
      for (const s of scripts) {
        for (const n of fontNodes) {
          const script = n.getAttribute('script') || n.getAttribute('w:script') || ''
          if (script === s) {
            const tf = normalizeThemeTypeface(n.getAttribute('typeface') || '')
            if (tf) return tf
          }
        }
      }
      // fallback: first non-empty typeface in any font node
      for (const n of fontNodes) {
        const tf = normalizeThemeTypeface(n.getAttribute('typeface') || '')
        if (tf) return tf
      }
      return ''
    }

    const getEastAsiaTypeface = (fontEl: Element | null) => {
      // Prefer <a:ea typeface="..."> if it's a real name; otherwise read script-specific mapping (Hans/Hant).
      const ea = normalizeThemeTypeface(getTypeface(fontEl, 'ea'))
      if (ea) return ea
      return getTypefaceByScript(fontEl, ['Hans', 'Hant', 'Jpan', 'Hang'])
    }

    const getLatinTypeface = (fontEl: Element | null) => {
      const latin = normalizeThemeTypeface(getTypeface(fontEl, 'latin'))
      if (latin) return latin
      return getTypefaceByScript(fontEl, ['Latn'])
    }

    const getCsTypeface = (fontEl: Element | null) => {
      const cs = normalizeThemeTypeface(getTypeface(fontEl, 'cs'))
      if (cs) return cs
      return getTypefaceByScript(fontEl, ['Arab', 'Hebr', 'Thaa'])
    }

    // Word theme keys commonly used in w:rFonts:
    // minorHAnsi/majorHAnsi, minorEastAsia/majorEastAsia, minorBidi/majorBidi
    const map: ThemeFontMap = {}
    map.majorHAnsi = getLatinTypeface(majorFont) || getCsTypeface(majorFont)
    map.minorHAnsi = getLatinTypeface(minorFont) || getCsTypeface(minorFont)
    map.majorEastAsia = getEastAsiaTypeface(majorFont)
    map.minorEastAsia = getEastAsiaTypeface(minorFont)
    map.majorBidi = getCsTypeface(majorFont)
    map.minorBidi = getCsTypeface(minorFont)

    // Some docs may use majorAscii/minorAscii (non-standard but seen in the wild)
    if (map.majorHAnsi) map.majorAscii = map.majorHAnsi
    if (map.minorHAnsi) map.minorAscii = map.minorHAnsi

    // 为空值的主题键添加常用默认值，确保主题字体引用不会返回空
    // 这解决了某些 DOCX 文件主题中没有定义中文字体的问题
    const THEME_FONT_DEFAULTS: ThemeFontMap = {
      // 拉丁字体默认值（Office 2016+ 默认）
      majorHAnsi: 'Calibri Light',
      minorHAnsi: 'Calibri',
      majorAscii: 'Calibri Light',
      minorAscii: 'Calibri',
      // 东亚字体默认值（中文环境常用）
      majorEastAsia: '等线',
      minorEastAsia: '等线',
      // 复杂文本脚本默认值
      majorBidi: 'Times New Roman',
      minorBidi: 'Times New Roman',
    }

    // 只为空值填充默认值，保留已解析的值
    for (const [key, defaultValue] of Object.entries(THEME_FONT_DEFAULTS)) {
      if (!map[key]) {
        map[key] = defaultValue
      }
    }

    return map
  } catch {
    // 解析失败时返回默认的主题字体映射，而不是 null
    // 这确保即使 theme.xml 解析失败，主题字体引用仍能正常工作
    return {
      majorHAnsi: 'Calibri Light',
      minorHAnsi: 'Calibri',
      majorAscii: 'Calibri Light',
      minorAscii: 'Calibri',
      majorEastAsia: '等线',
      minorEastAsia: '等线',
      majorBidi: 'Times New Roman',
      minorBidi: 'Times New Roman',
    }
  }
}

// 主题颜色映射类型
type ThemeColorMap = Record<string, string>

// 当前解析的主题颜色映射
let currentThemeColorMap: ThemeColorMap | null = null

/**
 * 从 theme1.xml 加载主题颜色映射
 * 支持的主题颜色：dk1, lt1, dk2, lt2, accent1-accent6, hlink, folHlink
 */
async function loadThemeColorMap(zip: JSZip): Promise<ThemeColorMap | null> {
  try {
    const themeXml =
      (await zip.file('word/theme/theme1.xml')?.async('string')) ||
      (await zip.file('word/theme/theme.xml')?.async('string'))
    if (!themeXml) return null

    const parser = new DOMParser()
    const doc = parser.parseFromString(themeXml, 'application/xml')
    const parseError = doc.querySelector('parsererror')
    if (parseError) return null

    const root = doc.documentElement
    const clrScheme = findElementByLocalName(root as any, 'clrScheme')
    if (!clrScheme) return null

    const map: ThemeColorMap = {}
    
    // 主题颜色名称列表
    const colorNames = [
      'dk1', 'lt1', 'dk2', 'lt2',
      'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6',
      'hlink', 'folHlink'
    ]
    
    for (const colorName of colorNames) {
      const colorEl = findElementByLocalName(clrScheme, colorName)
      if (colorEl) {
        // 尝试获取 srgbClr
        const srgbClr = findElementByLocalName(colorEl, 'srgbClr')
        if (srgbClr) {
          const val = srgbClr.getAttribute('val')
          if (val) {
            map[colorName] = val.toUpperCase()
          }
        } else {
          // 尝试获取 sysClr（系统颜色）
          const sysClr = findElementByLocalName(colorEl, 'sysClr')
          if (sysClr) {
            const lastClr = sysClr.getAttribute('lastClr')
            if (lastClr) {
              map[colorName] = lastClr.toUpperCase()
            }
          }
        }
      }
    }
    
    // 默认 Office 主题颜色（如果没有解析到）
    const defaultColors: ThemeColorMap = {
      'dk1': '000000',
      'lt1': 'FFFFFF',
      'dk2': '44546A',
      'lt2': 'E7E6E6',
      'accent1': '5B9BD5',
      'accent2': 'ED7D31',
      'accent3': 'A5A5A5',
      'accent4': 'FFC000',
      'accent5': '4472C4',
      'accent6': '70AD47',
      'hlink': '0563C1',
      'folHlink': '954F72',
    }
    
    // 填充默认值
    for (const [key, defaultValue] of Object.entries(defaultColors)) {
      if (!map[key]) {
        map[key] = defaultValue
      }
    }
    
    return map
  } catch {
    // 解析失败时返回默认的主题颜色映射
    return {
      'dk1': '000000',
      'lt1': 'FFFFFF',
      'dk2': '44546A',
      'lt2': 'E7E6E6',
      'accent1': '5B9BD5',
      'accent2': 'ED7D31',
      'accent3': 'A5A5A5',
      'accent4': 'FFC000',
      'accent5': '4472C4',
      'accent6': '70AD47',
      'hlink': '0563C1',
      'folHlink': '954F72',
    }
  }
}

/**
 * 解析主题颜色引用
 * @param themeColor 主题颜色名称（如 'accent1', 'dk1' 等）
 * @returns 颜色的 HEX 值（不带 #）
 */
function resolveThemeColor(themeColor: string | null | undefined): string | null {
  if (!themeColor) return null
  if (!currentThemeColorMap) return null
  return currentThemeColorMap[themeColor] || null
}

// 编号定义映射类型
interface NumberingInfo {
  numFmt: string       // 编号格式: decimal, bullet, lowerLetter, upperLetter, lowerRoman, upperRoman, chineseCounting 等
  lvlText: string      // 级别文本模板
  start: number        // 起始编号
  suff?: string        // 编号后缀: tab, space, nothing
  indLeftTwips?: number
  indHangingTwips?: number
  indFirstLineTwips?: number
}

type NumberingMap = Record<string, Record<number, NumberingInfo>>  // numId -> ilvl -> NumberingInfo
type NumberingState = Record<string, number[]>

let currentNumberingMap: NumberingMap | null = null

/**
 * 从 numbering.xml 加载编号定义映射
 */
async function loadNumberingMap(zip: JSZip): Promise<NumberingMap | null> {
  try {
    const numberingXml = await zip.file('word/numbering.xml')?.async('string')
    if (!numberingXml) return null

    const parser = new DOMParser()
    const doc = parser.parseFromString(numberingXml, 'application/xml')
    const parseError = doc.querySelector('parsererror')
    if (parseError) return null

    const map: NumberingMap = {}
    
    // 首先解析抽象编号定义
    const abstractNums: Record<string, Record<number, NumberingInfo>> = {}
    const abstractNumEls = doc.getElementsByTagName('w:abstractNum')
    
    for (let i = 0; i < abstractNumEls.length; i++) {
      const abstractNum = abstractNumEls[i]
      const abstractNumId = abstractNum.getAttribute('w:abstractNumId')
      if (!abstractNumId) continue
      
      abstractNums[abstractNumId] = {}
      
      const lvlEls = abstractNum.getElementsByTagName('w:lvl')
      for (let j = 0; j < lvlEls.length; j++) {
        const lvl = lvlEls[j]
        const ilvl = parseInt(lvl.getAttribute('w:ilvl') || '0')
        
        const numFmtEl = lvl.getElementsByTagName('w:numFmt')[0]
        const numFmt = numFmtEl?.getAttribute('w:val') || 'decimal'
        
        const lvlTextEl = lvl.getElementsByTagName('w:lvlText')[0]
        const lvlText = lvlTextEl?.getAttribute('w:val') || '%1.'
        
        const startEl = lvl.getElementsByTagName('w:start')[0]
        const start = parseInt(startEl?.getAttribute('w:val') || '1')

        const suffEl = lvl.getElementsByTagName('w:suff')[0]
        const suff = suffEl?.getAttribute('w:val') || undefined

        const ind = lvl.getElementsByTagName('w:ind')[0]
        const left = ind?.getAttribute('w:left') || ind?.getAttribute('w:start')
        const hanging = ind?.getAttribute('w:hanging')
        const firstLine = ind?.getAttribute('w:firstLine')

        abstractNums[abstractNumId][ilvl] = { 
          numFmt, 
          lvlText, 
          start,
          suff,
          indLeftTwips: left ? parseInt(left) : undefined,
          indHangingTwips: hanging ? parseInt(hanging) : undefined,
          indFirstLineTwips: firstLine ? parseInt(firstLine) : undefined,
        }
      }
    }
    
    // 然后解析编号实例（numId -> abstractNumId 的映射）
    const numEls = doc.getElementsByTagName('w:num')
    for (let i = 0; i < numEls.length; i++) {
      const num = numEls[i]
      const numId = num.getAttribute('w:numId')
      if (!numId) continue
      
      const abstractNumIdEl = num.getElementsByTagName('w:abstractNumId')[0]
      const abstractNumId = abstractNumIdEl?.getAttribute('w:val')
      if (abstractNumId && abstractNums[abstractNumId]) {
        map[numId] = abstractNums[abstractNumId]
      }
    }
    
    return map
  } catch {
    return null
  }
}

/**
 * 获取编号信息
 * @param numId 编号 ID
 * @param ilvl 级别
 * @returns 编号信息
 */
function getNumberingInfo(numId: string, ilvl: number): NumberingInfo | null {
  if (!currentNumberingMap) return null
  const numInfo = currentNumberingMap[numId]
  if (!numInfo) return null
  return numInfo[ilvl] || numInfo[0] || null
}

/**
 * 将 Word 编号格式转换为 CSS list-style-type
 */
function numFmtToCssListStyle(numFmt: string): string {
  switch (numFmt) {
    case 'bullet':
      return 'disc'
    case 'decimal':
    case 'chineseCounting':
    case 'japaneseCounting':
      return 'decimal'
    case 'lowerLetter':
      return 'lower-alpha'
    case 'upperLetter':
      return 'upper-alpha'
    case 'lowerRoman':
      return 'lower-roman'
    case 'upperRoman':
      return 'upper-roman'
    case 'decimalEnclosedCircle':
      return 'decimal'
    default:
      return 'decimal'
  }
}

function toRoman(num: number, upper = true): string {
  if (num <= 0) return ''
  const map: Array<[number, string]> = [
    [1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'],
    [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'],
    [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']
  ]
  let n = num
  let result = ''
  for (const [val, sym] of map) {
    while (n >= val) {
      result += sym
      n -= val
    }
  }
  return upper ? result : result.toLowerCase()
}

function toAlpha(num: number, upper = true): string {
  if (num <= 0) return ''
  let n = num
  let result = ''
  while (n > 0) {
    n -= 1
    const charCode = (n % 26) + 65
    result = String.fromCharCode(charCode) + result
    n = Math.floor(n / 26)
  }
  return upper ? result : result.toLowerCase()
}

function toChineseNumber(num: number): string {
  if (num <= 0) return ''
  const digits = ['零', '一', '二', '三', '四', '五', '六', '七', '八', '九']
  const units = ['', '十', '百', '千', '万']
  if (num < 10) return digits[num]
  if (num < 20) return `十${num % 10 ? digits[num % 10] : ''}`
  if (num < 100) {
    const tens = Math.floor(num / 10)
    const ones = num % 10
    return `${digits[tens]}十${ones ? digits[ones] : ''}`
  }
  let n = num
  let unitIndex = 0
  let result = ''
  let zero = false
  while (n > 0 && unitIndex < units.length) {
    const digit = n % 10
    if (digit === 0) {
      if (!zero && result) {
        result = digits[0] + result
        zero = true
      }
    } else {
      result = `${digits[digit]}${units[unitIndex]}${result}`
      zero = false
    }
    n = Math.floor(n / 10)
    unitIndex += 1
  }
  result = result.replace(/^一十/, '十')
  return result
}

function formatNumberingValue(numFmt: string, value: number): string {
  switch (numFmt) {
    case 'lowerLetter':
      return toAlpha(value, false)
    case 'upperLetter':
      return toAlpha(value, true)
    case 'lowerRoman':
      return toRoman(value, false)
    case 'upperRoman':
      return toRoman(value, true)
    case 'chineseCounting':
    case 'japaneseCounting':
      return toChineseNumber(value)
    default:
      return String(value)
  }
}

function buildListMarkerText(
  numId: string,
  ilvl: number,
  numberingState: NumberingState
): string | null {
  if (!currentNumberingMap) return null
  const numInfo = currentNumberingMap[numId]
  if (!numInfo) return null
  const levelInfo = numInfo[ilvl] || numInfo[0]
  if (!levelInfo) return null

  const counters = numberingState[numId] || []
  const start = levelInfo.start || 1
  const current = (counters[ilvl] || (start - 1)) + 1
  counters[ilvl] = current
  for (let i = ilvl + 1; i < counters.length; i++) {
    counters[i] = 0
  }
  numberingState[numId] = counters

  const template = levelInfo.lvlText || '%1.'
  const marker = template.replace(/%(\d)/g, (_match, digit) => {
    const idx = parseInt(digit, 10) - 1
    if (idx < 0) return ''
    const levelCount = counters[idx] || (numInfo[idx]?.start || 1)
    const fmt = numInfo[idx]?.numFmt || levelInfo.numFmt
    return formatNumberingValue(fmt, levelCount)
  })

  const suff = levelInfo.suff
  if (suff === 'space') return `${marker} `
  if (suff === 'tab') return `${marker} `
  return marker
}

// 获取安全的字体族
function getSafeFontFamily(fontName: string | null | undefined): string {
  if (!fontName) return ''
  
  if (FONT_FALLBACK_MAP[fontName]) {
    return FONT_FALLBACK_MAP[fontName]
  }
  
  const isChinese = /[\u4e00-\u9fa5]/.test(fontName) || 
                    fontName.includes('Song') || 
                    fontName.includes('Hei') || 
                    fontName.includes('Kai') ||
                    fontName.includes('Fang')
  
  if (isChinese) {
    return `"${fontName}", "宋体", "SimSun", serif`
  }
  
  return `"${fontName}", "Arial", sans-serif`
}

// 检查文件是否是有效的 ZIP/DOCX 格式
function isValidZip(bytes: Uint8Array): boolean {
  return bytes.length >= 4 && bytes[0] === 0x50 && bytes[1] === 0x4B
}

// 检查是否是旧版 .doc 格式 (OLE Compound Document)
function isOldDocFormat(bytes: Uint8Array): boolean {
  return bytes.length >= 4 && 
         bytes[0] === 0xD0 && bytes[1] === 0xCF && 
         bytes[2] === 0x11 && bytes[3] === 0xE0
}

// 使用 mammoth 解析 Word 文档（支持 .doc 和 .docx）
async function parseWithMammoth(arrayBuffer: ArrayBuffer): Promise<string> {
  try {
    const result = await mammoth.convertToHtml({ arrayBuffer })
    
    if (result.messages.length > 0) {
      console.log('Mammoth 解析消息:', result.messages)
    }
    
    let html = result.value
    
    // 如果没有内容，返回空段落
    if (!html || html.trim() === '') {
      return '<p></p>'
    }
    
    // 添加一些基本样式处理
    // 将 mammoth 生成的简单 HTML 转换为更适合显示的格式
    html = html
      // 段落添加缩进样式
      .replace(/<p>/g, '<p>')
      // 表格添加边框
      .replace(/<table>/g, '<table style="border-collapse: collapse; width: auto; margin: 0;">')
      .replace(/<td>/g, '<td style="border: 0.5pt solid var(--word-rule); padding: 2pt 5pt;">')
      .replace(/<th>/g, '<th style="border: 0.5pt solid var(--word-rule); padding: 2pt 5pt; background: var(--word-page-bg);">')
    
    return html
  } catch (error) {
    console.error('Mammoth 解析失败:', error)
    throw error
  }
}

// 解析 docx 文件并转换为 HTML
export async function parseDocxToHtml(base64Data: string): Promise<string> {
  try {
    console.log('开始解析 Word 文档，数据长度:', base64Data.length)
    
    // 解码 base64
    const binaryString = atob(base64Data)
    const bytes = new Uint8Array(binaryString.length)
    for (let i = 0; i < binaryString.length; i++) {
      bytes[i] = binaryString.charCodeAt(i)
    }
    
    console.log('文件头字节:', bytes[0], bytes[1], bytes[2], bytes[3])
    
    // 获取 ArrayBuffer 用于 mammoth
    const arrayBuffer = bytes.buffer

    // 检查文件格式并选择解析方式
    if (isOldDocFormat(bytes)) {
      console.log('检测到旧版 .doc 格式，使用 mammoth 解析')
      return await parseWithMammoth(arrayBuffer)
    }

    if (isValidZip(bytes)) {
      console.log('检测到 .docx 格式')
      
      // 先尝试用我们的自定义解析器（保留更多样式）
      try {
        const customResult = await parseDocxCustom(bytes)
        if (customResult && customResult.trim()) {
          // 原先用 “□” 来判定失败太激进：很多 docx（项目符号/特殊符号/域代码）也可能包含该字符，
          // 但自定义解析器仍然能正确保留字体等关键信息。这里改成：
          // - 只要产出了 font-family（或段落字体下沉标记），就优先使用自定义解析结果。
          // - 否则再按旧规则（不包含 □）作为保守判断。
          const hasFontInfo =
            customResult.includes('font-family:') || customResult.includes('data-para-font="1"')
          if (hasFontInfo || !customResult.includes('□')) {
            console.log('自定义解析器成功')
            return customResult
          }
        }
      } catch (e) {
        console.log('自定义解析器失败，回退到 mammoth:', e)
      }
      
      // 回退到 mammoth
      console.log('使用 mammoth 解析 .docx')
      return await parseWithMammoth(arrayBuffer)
    }

    // 未知格式，尝试用 mammoth
    console.log('未知格式，尝试用 mammoth 解析')
    try {
      return await parseWithMammoth(arrayBuffer)
    } catch (e) {
      console.error('mammoth 也无法解析:', e)
      return `<div style="padding: 40px; text-align: center; color: #888;">
        <p style="font-size: 18px; margin-bottom: 10px;">⚠️ 无法识别的文件格式</p>
        <p style="font-size: 14px;">请确保文件是有效的 Word 文档 (.doc 或 .docx)</p>
      </div>`
    }
  } catch (error) {
    console.error('Word 文档解析错误:', error)
    return `<div style="padding: 40px; text-align: center; color: #888;">
      <p style="font-size: 18px; margin-bottom: 10px;">⚠️ 文档解析失败</p>
      <p style="font-size: 14px;">${(error as Error).message}</p>
    </div>`
  }
}

export interface DocxHtmlForAgentResult {
  html: string
  images: DocxImageMeta[]
  pageSettings?: PageSettings
}

async function parseWithMammothNoInlineImages(arrayBuffer: ArrayBuffer): Promise<string> {
  try {
    const result = await mammoth.convertToHtml(
      { arrayBuffer },
      {
        // 不要把图片内联成 base64（避免超大字符串/卡死）
        convertImage: mammoth.images.imgElement(async () => ({ src: 'about:blank' })),
      } as any
    )

    let html = result.value || ''
    if (!html || html.trim() === '') return '<p></p>'

    html = html
      .replace(/<table>/g, '<table style="border-collapse: collapse; width: auto; margin: 0;">')
      .replace(/<td>/g, '<td style="border: 0.5pt solid var(--word-rule); padding: 2pt 5pt;">')
      .replace(/<th>/g, '<th style="border: 0.5pt solid var(--word-rule); padding: 2pt 5pt; background: var(--word-page-bg);">')

    return html
  } catch (e) {
    // 最差情况也别 throw，让上层给出可读错误
    return '<p></p>'
  }
}

async function parseDocxCustomForAgent(bytes: Uint8Array): Promise<DocxHtmlForAgentResult> {
  const ab = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer
  const zip = await JSZip.loadAsync(ab)

  const documentXml = await zip.file('word/document.xml')?.async('string')
  if (!documentXml) {
    throw new Error('找不到 document.xml')
  }

  // 加载主题颜色、字体和编号定义
  currentThemeColorMap = await loadThemeColorMap(zip)
  currentNumberingMap = await loadNumberingMap(zip)
  
  const stylesXml = await zip.file('word/styles.xml')?.async('string')
  const styles = stylesXml ? parseStyles(stylesXml, currentThemeColorMap || undefined) : {}

  // relationships（用于定位图片 target / header/footer 等）
  const relsMap = await parseRelationships(zip)

  // Agent 模式：不提取图片 base64
  const imageMap: ImageMap = {}
  const imageCtx: ImageParseContext = { embedImages: false, relsMap, images: [] }

  // 脚注/尾注
  const footnoteMap = await parseFootnotesAndEndnotes(zip, styles)

  // 页眉/页脚（图片只保留占位 + 元信息）
  const headerFooterData = await parseHeadersAndFooters(zip, relsMap, styles, imageMap, imageCtx)

  const parser = new DOMParser()
  const doc = parser.parseFromString(documentXml, 'application/xml')
  const parseError = doc.querySelector('parsererror')
  if (parseError) {
    throw new Error('XML 解析失败')
  }

  const pageSettings = parsePageSettings(doc)

  const body = doc.getElementsByTagName('w:body')[0]
  if (!body) {
    throw new Error('找不到文档主体')
  }

  let html = ''
  const numberingState: NumberingState = {}

  // 注意：页眉不再嵌入正文 HTML，由前端单独渲染
  html += '<div class="docx-body">'

  const children = body.childNodes
  for (let i = 0; i < children.length; i++) {
    const child = children[i] as Element
    if (child.nodeName === 'w:p') {
      html += parseParagraph(child, styles, imageMap, footnoteMap, imageCtx, numberingState)
    } else if (child.nodeName === 'w:tbl') {
      html += parseTable(child, styles, imageMap, footnoteMap, imageCtx)
    } else if (child.nodeName === 'w:sdt') {
      const sdtContent = child.getElementsByTagName('w:sdtContent')[0]
      if (sdtContent) {
        const sdtChildren = sdtContent.childNodes
        for (let j = 0; j < sdtChildren.length; j++) {
          const sdtChild = sdtChildren[j] as Element
          if (sdtChild.nodeName === 'w:p') {
            html += parseParagraph(sdtChild, styles, imageMap, footnoteMap, imageCtx, numberingState)
          } else if (sdtChild.nodeName === 'w:tbl') {
            html += parseTable(sdtChild, styles, imageMap, footnoteMap, imageCtx)
          }
        }
      }
    }
  }

  html += '</div>'
  html += generateFootnotesHtml(footnoteMap)
  // 注意：页脚不再嵌入正文 HTML，由前端单独渲染

  return { html: html || '<p></p>', images: imageCtx.images || [], pageSettings }
}

// 给 Agent 使用：返回“全文 HTML（尽力）+ 图片元信息（不内联 base64）”
export async function parseDocxToHtmlForAgent(base64Data: string): Promise<DocxHtmlForAgentResult> {
  const images: DocxImageMeta[] = []
  try {
    const binaryString = atob(base64Data)
    const bytes = new Uint8Array(binaryString.length)
    for (let i = 0; i < binaryString.length; i++) {
      bytes[i] = binaryString.charCodeAt(i)
    }

    const arrayBuffer = bytes.buffer

    // .doc（旧格式）只能走 mammoth（不内联图片）
    if (isOldDocFormat(bytes)) {
      const html = await parseWithMammothNoInlineImages(arrayBuffer)
      return { html, images }
    }

    if (isValidZip(bytes)) {
      try {
        return await parseDocxCustomForAgent(bytes)
      } catch (e) {
        const html = await parseWithMammothNoInlineImages(arrayBuffer)
        return { html, images }
      }
    }

    const html = await parseWithMammothNoInlineImages(arrayBuffer)
    return { html, images }
  } catch (error) {
    return {
      html: `<div style="padding: 16px; color: #888;">DOCX 解析失败: ${(error as Error).message}</div>`,
      images,
    }
  }
}

// 解析 relationships 文件，构建 rId -> 路径映射（包括图片、页眉、页脚等）
async function parseRelationships(zip: JSZip): Promise<Record<string, string>> {
  const relsMap: Record<string, string> = {}
  
  const relsXml = await zip.file('word/_rels/document.xml.rels')?.async('string')
  if (!relsXml) {
    // 尝试其他可能的路径
    const altRelsXml = await zip.file('word/_rels/document2.xml.rels')?.async('string')
    if (!altRelsXml) {
      return relsMap
    }
  }
  
  const parser = new DOMParser()
  const doc = parser.parseFromString(relsXml!, 'application/xml')
  
  const relationships = doc.getElementsByTagName('Relationship')
  
  for (let i = 0; i < relationships.length; i++) {
    const rel = relationships[i]
    const id = rel.getAttribute('Id')
    const target = rel.getAttribute('Target')
    const type = rel.getAttribute('Type')
    
    if (id && target) {
      // 保存所有关系的映射
      let path = target
      if (target.startsWith('/')) {
        path = target.slice(1)
      } else if (!target.startsWith('word/') && !target.startsWith('../')) {
        path = `word/${target}`
      }
      relsMap[id] = path
    }
  }
  
  return relsMap
}

// 从 zip 中提取图片
async function extractImages(zip: JSZip, relsMap: Record<string, string>): Promise<ImageMap> {
  const imageMap: ImageMap = {}
  
  // 列出 word/media 目录下的所有文件
  const mediaFiles: string[] = []
  zip.forEach((relativePath, file) => {
    if (relativePath.includes('media/') && !file.dir) {
      mediaFiles.push(relativePath)
    }
  })
  
  for (const [rId, imagePath] of Object.entries(relsMap)) {
    try {
      let imageFile = zip.file(imagePath)
      
      // 如果找不到，尝试其他路径格式
      if (!imageFile) {
        // 尝试不带 word/ 前缀
        const altPath1 = imagePath.replace('word/', '')
        imageFile = zip.file(altPath1)
        if (imageFile) {
        }
      }
      
      if (!imageFile) {
        // 尝试直接在 media 目录下查找
        const fileName = imagePath.split('/').pop()
        const altPath2 = `word/media/${fileName}`
        imageFile = zip.file(altPath2)
        if (imageFile) {
        }
      }
      
      if (!imageFile) {
        continue
      }
      
      // 获取图片的 base64 数据（JSZip 直接提供）
      const imageBase64 = await imageFile.async('base64')
      
      // 确定 MIME 类型
      const ext = imagePath.split('.').pop()?.toLowerCase() || 'png'
      const mimeType = IMAGE_MIME_TYPES[ext] || 'image/png'
      
      imageMap[rId] = `data:${mimeType};base64,${imageBase64}`
    } catch (err) {
      console.error(`提取图片失败: ${imagePath}`, err)
    }
  }
  
  return imageMap
}

// 解析脚注和尾注
async function parseFootnotesAndEndnotes(zip: JSZip, styles: Record<string, any>): Promise<FootnoteMap> {
  const result: FootnoteMap = {
    footnotes: new Map(),
    endnotes: new Map()
  }
  
  // 解析脚注
  const footnotesXml = await zip.file('word/footnotes.xml')?.async('string')
  if (footnotesXml) {
    const parser = new DOMParser()
    const doc = parser.parseFromString(footnotesXml, 'application/xml')
    const footnotes = doc.getElementsByTagName('w:footnote')
    
    for (let i = 0; i < footnotes.length; i++) {
      const footnote = footnotes[i]
      const id = footnote.getAttribute('w:id')
      const type = footnote.getAttribute('w:type')
      
      // 跳过分隔符和延续分隔符（id=0 和 id=-1）
      if (!id || type === 'separator' || type === 'continuationSeparator') continue
      if (id === '0' || id === '-1') continue
      
      // 解析脚注内容
      let content = ''
      const paras = footnote.getElementsByTagName('w:p')
      for (let j = 0; j < paras.length; j++) {
        content += parseSimpleParagraph(paras[j], styles)
      }
      
      if (content.trim()) {
        result.footnotes.set(id, { id, content })
      }
    }
  }
  
  // 解析尾注
  const endnotesXml = await zip.file('word/endnotes.xml')?.async('string')
  if (endnotesXml) {
    const parser = new DOMParser()
    const doc = parser.parseFromString(endnotesXml, 'application/xml')
    const endnotes = doc.getElementsByTagName('w:endnote')
    
    for (let i = 0; i < endnotes.length; i++) {
      const endnote = endnotes[i]
      const id = endnote.getAttribute('w:id')
      const type = endnote.getAttribute('w:type')
      
      // 跳过分隔符和延续分隔符
      if (!id || type === 'separator' || type === 'continuationSeparator') continue
      if (id === '0' || id === '-1') continue
      
      // 解析尾注内容
      let content = ''
      const paras = endnote.getElementsByTagName('w:p')
      for (let j = 0; j < paras.length; j++) {
        content += parseSimpleParagraph(paras[j], styles)
      }
      
      if (content.trim()) {
        result.endnotes.set(id, { id, content })
      }
    }
  }
  
  return result
}

// 简化的段落解析（用于脚注/尾注/页眉/页脚）
function parseSimpleParagraph(para: Element, styles: Record<string, any>): string {
  let content = ''
  const runs = para.getElementsByTagName('w:r')
  
  for (let i = 0; i < runs.length; i++) {
    const run = runs[i]
    const texts = run.getElementsByTagName('w:t')
    for (let j = 0; j < texts.length; j++) {
      content += texts[j].textContent || ''
    }
  }
  
  return content ? `<span class="note-content">${escapeHtml(content)}</span>` : ''
}

// 页眉/页脚样式
// 页眉/页脚样式
export interface HeaderFooterStyle {
  fontFamily?: string
  fontSize?: string
  color?: string
  lineHeight?: string
  alignment?: 'left' | 'center' | 'right'
  borderBottom?: boolean  // 页眉下划线
  borderTop?: boolean     // 页脚上划线
}

// 页眉/页脚内容（单个）
export interface HeaderFooterContent {
  html: string
  style: HeaderFooterStyle
}

// 页眉/页脚数据类型（旧版兼容）
interface HeaderFooterData {
  headerHtml: string
  footerHtml: string
  headerStyle?: HeaderFooterStyle
  footerStyle?: HeaderFooterStyle
}

// 页面设置
export interface PageSettings {
  width: number       // 页面宽度 (pt)
  height: number      // 页面高度 (pt)
  marginTop: number   // 上边距 (pt)
  marginBottom: number // 下边距 (pt)
  marginLeft: number  // 左边距 (pt)
  marginRight: number // 右边距 (pt)
  headerHeight: number // 页眉高度 (pt)
  footerHeight: number // 页脚高度 (pt)
  orientation?: 'portrait' | 'landscape'
}

// 节配置（基于 ONLYOFFICE SectPr 模型）
export interface SectionConfig {
  pageSettings: PageSettings
  // 首页页眉/页脚
  headerFirst?: HeaderFooterContent
  footerFirst?: HeaderFooterContent
  // 偶数页页眉/页脚
  headerEven?: HeaderFooterContent
  footerEven?: HeaderFooterContent
  // 默认/奇数页页眉/页脚
  headerDefault?: HeaderFooterContent
  footerDefault?: HeaderFooterContent
  // 设置标志
  titlePage: boolean        // 首页不同
  evenAndOddHeaders: boolean  // 奇偶页不同
}

// 完整文档模型（新版）
export interface DocumentModel {
  sections: SectionConfig[]
  bodyHtml: string
  footnotes: FootnoteMap
  totalPages?: number
}

// 完整的文档解析结果（旧版兼容）
export interface DocxParseResult {
  bodyHtml: string           // 正文 HTML
  headerHtml: string         // 页眉 HTML
  footerHtml: string         // 页脚 HTML
  headerStyle: HeaderFooterStyle
  footerStyle: HeaderFooterStyle
  pageSettings: PageSettings
  totalPages?: number        // 估算的总页数
  // 新增：完整文档模型
  documentModel?: DocumentModel
}

// 解析页眉和页脚
async function parseHeadersAndFooters(
  zip: JSZip,
  relsMap: Record<string, string>,
  styles: Record<string, any>,
  imageMap: ImageMap,
  imageCtx?: ImageParseContext
): Promise<HeaderFooterData> {
  const result: HeaderFooterData = {
    headerHtml: '',
    footerHtml: ''
  }
  
  // 从 relationships 中找到页眉和页脚文件
  const headerFiles: string[] = []
  const footerFiles: string[] = []
  
  for (const [rId, target] of Object.entries(relsMap)) {
    if (target.includes('header') && target.endsWith('.xml')) {
      const path = target.startsWith('word/') ? target : `word/${target}`
      headerFiles.push(path)
    } else if (target.includes('footer') && target.endsWith('.xml')) {
      const path = target.startsWith('word/') ? target : `word/${target}`
      footerFiles.push(path)
    }
  }
  
  // 解析第一个页眉（通常是默认页眉）
  if (headerFiles.length > 0) {
    const defaultHeader = headerFiles.find(f => f.includes('header2.xml')) || 
                          headerFiles.find(f => f.includes('header1.xml')) || 
                          headerFiles[0]
    
    const headerXml = await zip.file(defaultHeader)?.async('string')
    if (headerXml) {
      const parser = new DOMParser()
      const doc = parser.parseFromString(headerXml, 'application/xml')
      const body = doc.documentElement
      result.headerStyle = parseHeaderFooterStyleFromXml(doc, styles, 'right')
      
      const paras = body.getElementsByTagName('w:p')
      const headerParts = collectHeaderFooterParts(paras, styles, imageMap, imageCtx)

      if (headerParts.length > 0) {
        result.headerHtml = headerParts.join('')
      }
    }
  }

  if (footerFiles.length > 0) {
    const defaultFooter = footerFiles.find(f => f.includes('footer2.xml')) || 
                          footerFiles.find(f => f.includes('footer1.xml')) || 
                          footerFiles[0]
    
    const footerXml = await zip.file(defaultFooter)?.async('string')
    if (footerXml) {
      const parser = new DOMParser()
      const doc = parser.parseFromString(footerXml, 'application/xml')
      const body = doc.documentElement
      result.footerStyle = parseHeaderFooterStyleFromXml(doc, styles, 'center')
      
      const paras = body.getElementsByTagName('w:p')
      const footerParts = collectHeaderFooterParts(paras, styles, imageMap, imageCtx)
      
      if (footerParts.length > 0) {
        result.footerHtml = footerParts.join('')
      } else {
        result.footerHtml = '<p class="header-footer-para">{PAGE}</p>'
      }
    }
  }
  return result
}

// 增强版简化段落解析（支持图片、域代码等）
// skipPageFields: 是否跳过页码相关字段（在页眉中通常跳过）
// usePagePlaceholder: 是否为页码使用占位符（用于页脚）
type HeaderFooterTabAlign = 'left' | 'center' | 'right'
interface HeaderFooterTabStop {
  posPt: number
  align: HeaderFooterTabAlign
}

interface HeaderFooterPara {
  html: string
  tabStops: HeaderFooterTabStop[]
}

function extractHeaderFooterInner(paraHtml: string): string {
  const match = paraHtml.match(/<p[^>]*>([\s\S]*)<\/p>/)
  return match ? match[1] : paraHtml
}

function isHeaderFooterPageOnly(innerHtml: string): boolean {
  const text = innerHtml
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, '')
    .replace(/\s+/g, '')
  // 识别页码占位符或纯数字（页码值）
  return text === '{PAGE}' || text === '{NUMPAGES}' || /^\d+$/.test(text)
}

function parseTabStops(para: Element): HeaderFooterTabStop[] {
  const pPr = para.getElementsByTagName('w:pPr')[0]
  if (!pPr) return []
  const tabs = pPr.getElementsByTagName('w:tabs')[0]
  if (!tabs) return []
  const tabEls = tabs.getElementsByTagName('w:tab')
  const stops: HeaderFooterTabStop[] = []
  for (let i = 0; i < tabEls.length; i++) {
    const tab = tabEls[i]
    const pos = tab.getAttribute('w:pos')
    if (!pos) continue
    const val = (tab.getAttribute('w:val') || 'left').toLowerCase()
    if (val === 'clear' || val === 'bar') continue
    let align: HeaderFooterTabAlign = 'left'
    if (val === 'center') align = 'center'
    else if (val === 'right' || val === 'decimal') align = 'right'
    const posPt = parseInt(pos, 10) / 20
    if (!posPt) continue
    stops.push({ posPt, align })
  }
  return stops.sort((a, b) => a.posPt - b.posPt)
}

function mergeHeaderFooterPageFields(paras: HeaderFooterPara[]): HeaderFooterPara[] {
  if (paras.length < 2) return paras

  const inners = paras.map(item => extractHeaderFooterInner(item.html))
  const pageOnly = inners.map(isHeaderFooterPageOnly)
  const baseIndex = pageOnly.findIndex(isPage => !isPage)
  if (baseIndex < 0) return paras

  const rightParts = inners.filter((_, idx) => idx !== baseIndex && pageOnly[idx])
  if (rightParts.length === 0) return paras

  const mergedInner = inners[baseIndex] + HEADER_FOOTER_TAB_TOKEN + rightParts.join(' ')
  const mergedPara: HeaderFooterPara = {
    html: `<p class=\"header-footer-para\">${mergedInner}</p>`,
    tabStops: paras[baseIndex].tabStops
  }

  const merged: HeaderFooterPara[] = []
  for (let i = 0; i < paras.length; i++) {
    if (i === baseIndex) {
      merged.push(mergedPara)
    } else if (pageOnly[i]) {
      continue
    } else {
      merged.push(paras[i])
    }
  }
  return merged
}

function collectHeaderFooterParts(
  paras: HTMLCollectionOf<Element>,
  styles: Record<string, any>,
  imageMap: ImageMap,
  imageCtx?: ImageParseContext
): string[] {
  const raw: HeaderFooterPara[] = []
  for (let i = 0; i < paras.length; i++) {
    const para = paras[i]
    const paraContent = parseSimpleParagraphWithImages(para, styles, imageMap, false, true, imageCtx)
    const plain = paraContent.replace(/<[^>]+>/g, '').trim()
    if (plain || paraContent.includes('{PAGE}') || paraContent.includes('{NUMPAGES}')) {
      raw.push({ html: paraContent, tabStops: parseTabStops(para) })
    }

    // 处理页眉/页脚中的文本框（wps:txbx / v:textbox）
    // Word 常把页码放在文本框里，这里提取出来参与合并与对齐
    const textboxParas = extractHeaderFooterTextBoxParas(para, styles, imageMap, imageCtx)
    for (const item of textboxParas) {
      raw.push(item)
    }
  }
  const merged = mergeHeaderFooterPageFields(raw)
  return merged.map(item => formatHeaderFooterTabs(item.html, item.tabStops))
}

function extractHeaderFooterTextBoxParas(
  para: Element,
  styles: Record<string, any>,
  imageMap: ImageMap,
  imageCtx?: ImageParseContext
): HeaderFooterPara[] {
  const results: HeaderFooterPara[] = []
  const seen = new Set<string>()
  const textBoxes = findAllElementsByLocalName(para, 'txbxContent')
  if (!textBoxes.length) return results

  for (const box of textBoxes) {
    const innerParas = findAllElementsByLocalName(box, 'p')
    for (const innerPara of innerParas) {
      const html = parseSimpleParagraphWithImages(innerPara, styles, imageMap, false, true, imageCtx)
      if (!html) continue
      const plain = html.replace(/<[^>]+>/g, '').replace(/\s+/g, '')
      if (!plain) continue
      if (seen.has(plain)) continue
      seen.add(plain)
      results.push({ html, tabStops: parseTabStops(innerPara) })
    }
  }

  return results
}

const DOCX_TAB_TOKEN = '[[TAB]]'
const HEADER_FOOTER_TAB_TOKEN = DOCX_TAB_TOKEN
function parseSimpleParagraphWithImages(
  para: Element,
  styles: Record<string, any>,
  imageMap: ImageMap,
  skipPageFields: boolean = false,
  usePagePlaceholder: boolean = false,
  imageCtx?: ImageParseContext
): string {
  let content = ''
  let inComplexField = false  // 标记是否在复杂域内
  let fieldResult = ''        // 复杂域的结果
  let currentFieldType = ''   // 当前域类型
  let hasPageField = false    // 是否有页码字段
  
  const childNodes = para.childNodes
  
  for (let i = 0; i < childNodes.length; i++) {
    const child = childNodes[i] as Element
    if (child.nodeName === 'w:r') {
      // 检查是否是域开始/结束标记
      const fldChar = child.getElementsByTagName('w:fldChar')[0]
      if (fldChar) {
        const fldType = fldChar.getAttribute('w:fldCharType')
        if (fldType === 'begin') {
          inComplexField = true
          fieldResult = ''
          currentFieldType = ''
        } else if (fldType === 'separate') {
          // 分隔符后面是域结果
        } else if (fldType === 'end') {
          // 域结束，添加结果（除非是页码字段且需要跳过）
          const isPageField = currentFieldType.includes('PAGE') && !currentFieldType.includes('NUMPAGES')
          const isNumPagesField = currentFieldType.includes('NUMPAGES')
          
          if (isPageField || isNumPagesField) {
            hasPageField = true
            if (!skipPageFields) {
              // 页眉/页脚中统一使用占位符，方便前端替换为实际页码
              if (usePagePlaceholder) {
                content += isPageField ? '{PAGE}' : '{NUMPAGES}'
              } else if (fieldResult.trim()) {
                content += fieldResult
              }
            }
          } else if (fieldResult) {
            content += fieldResult
          }
          inComplexField = false
          fieldResult = ''
          currentFieldType = ''
        }
        continue
      }
      
      // 如果在复杂域中，检查指令文本
      const instrText = child.getElementsByTagName('w:instrText')[0]
      if (instrText) {
        // 记录域类型
        currentFieldType += (instrText.textContent || '').toUpperCase()
        continue
      }
      
      // 解析 run 中的文本/制表符
      const runChildren = Array.from(child.childNodes) as Element[]
      for (let j = 0; j < runChildren.length; j++) {
        const runChild = runChildren[j]
        if (runChild.nodeName === 'w:t') {
          const text = escapeHtml(runChild.textContent || '')
          if (inComplexField) {
            fieldResult += text
          } else {
            content += text
          }
        } else if (runChild.nodeName === 'w:tab' || runChild.nodeName === 'w:ptab') {
          if (inComplexField) {
            fieldResult += HEADER_FOOTER_TAB_TOKEN
          } else {
            content += HEADER_FOOTER_TAB_TOKEN
          }
        }
      }
      
      // 检查是否有图片
      const drawing = child.getElementsByTagName('w:drawing')[0]
      if (drawing) {
        content += parseDrawing(drawing, imageMap, imageCtx)
      }
    } else if (child.nodeName === 'w:fldSimple') {
      // 简单域（如页码）
      const instr = (child.getAttribute('w:instr') || '').toUpperCase()
      const isPageField = instr.includes('PAGE') && !instr.includes('NUMPAGES')
      const isNumPagesField = instr.includes('NUMPAGES')
      
      if (isPageField || isNumPagesField) {
        hasPageField = true
        // 如果是页码字段且需要跳过，则跳过
        if (skipPageFields) {
          continue
        }
      }
      
      // 获取域的值
      let fieldValue = ''
      const innerRuns = child.getElementsByTagName('w:r')
      for (let j = 0; j < innerRuns.length; j++) {
        const texts = innerRuns[j].getElementsByTagName('w:t')
        for (let k = 0; k < texts.length; k++) {
          fieldValue += escapeHtml(texts[k].textContent || '')
        }
      }
      
      if (fieldValue.trim()) {
        content += fieldValue
      } else if (usePagePlaceholder && (isPageField || isNumPagesField)) {
        content += isPageField ? '{PAGE}' : '{NUMPAGES}'
      }
    } else if (child.nodeName === 'w:hyperlink') {
      // 超链接
      const innerRuns = child.getElementsByTagName('w:r')
      for (let j = 0; j < innerRuns.length; j++) {
        const texts = innerRuns[j].getElementsByTagName('w:t')
        for (let k = 0; k < texts.length; k++) {
          content += escapeHtml(texts[k].textContent || '')
        }
      }
    } else if (child.nodeName === 'w:sdt') {
      // 结构化文档标签（可能包含日期、页码等）
      const sdtContent = child.getElementsByTagName('w:sdtContent')[0]
      if (sdtContent) {
        const innerRuns = sdtContent.getElementsByTagName('w:r')
        for (let j = 0; j < innerRuns.length; j++) {
          const texts = innerRuns[j].getElementsByTagName('w:t')
          for (let k = 0; k < texts.length; k++) {
            content += escapeHtml(texts[k].textContent || '')
          }
        }
      }
    }
  }
  
  // 如果没有内容但有页码字段，返回占位符
  if (!content && hasPageField && usePagePlaceholder && !skipPageFields) {
    return `<p class="header-footer-para">{PAGE}</p>`
  }
  
  return content ? `<p class="header-footer-para">${content}</p>` : ''
}

function formatHeaderFooterTabs(paraHtml: string, tabStops: HeaderFooterTabStop[] = []): string {
  if (!paraHtml || !paraHtml.includes(HEADER_FOOTER_TAB_TOKEN)) return paraHtml
  const match = paraHtml.match(/<p[^>]*>([\s\S]*)<\/p>/)
  const inner = match ? match[1] : paraHtml
  const segments = inner.split(HEADER_FOOTER_TAB_TOKEN)

  if (tabStops.length > 0) {
    const spans: string[] = []
    for (let i = 0; i < segments.length; i++) {
      const seg = segments[i] || ''
      if (i === 0) {
        spans.push(`<span class=\"hf-seg hf-left\">${seg}</span>`)
        continue
      }
      const stop = tabStops[Math.min(i - 1, tabStops.length - 1)]
      const align = stop?.align || 'left'
      const pos = stop?.posPt || 0
      const transforms: string[] = []
      if (align === 'right') transforms.push('translateX(-100%)')
      if (align === 'center') transforms.push('translateX(-50%)')
      const styleParts: string[] = []
      if (pos) styleParts.push(`left:${pos}pt`)
      if (transforms.length) styleParts.push(`transform:${transforms.join(' ')}`)
      const styleAttr = styleParts.length ? ` style=\"${styleParts.join(';')}\"` : ''
      spans.push(`<span class=\"hf-seg ${align}\"${styleAttr}>${seg}</span>`)
    }
    return `<p class=\"header-footer-para hf-tabs-abs\">${spans.join('')}</p>`
  }

  const left = segments[0] || ''
  const center = segments.length > 2 ? (segments[1] || '') : ''
  const right = segments.length > 2 ? segments.slice(2).join('') : (segments[1] || '')
  return `<p class=\"header-footer-para hf-tabs\"><span class=\"hf-left\">${left}</span><span class=\"hf-center\">${center}</span><span class=\"hf-right\">${right}</span></p>`
}

// 生成页眉/页脚区域的 HTML
function generateHeaderFooterHtml(data: HeaderFooterData, isHeader: boolean): string {
  let content = isHeader ? data.headerHtml : data.footerHtml
  
  if (!content) {
    return ''
  }
  
  // 在普通编辑模式下，将页码占位符替换为 "1"（因为不分页）
  content = content.replace(/\{PAGE\}/g, '1')
  content = content.replace(/\{NUMPAGES\}/g, '1')
  
  const className = isHeader ? 'docx-header' : 'docx-footer'
  return `<div class="${className}">${content}</div>`
}

// 生成脚注/尾注区域的 HTML
function generateFootnotesHtml(footnoteMap: FootnoteMap): string {
  let html = ''
  
  // 脚注区域
  if (footnoteMap.footnotes.size > 0) {
    html += '<div class="docx-footnotes"><hr class="footnote-separator" />'
    html += '<div class="footnotes-title">脚注</div>'
    
    footnoteMap.footnotes.forEach((note, id) => {
      html += `<p class="footnote-item"><sup class="footnote-number">[${id}]</sup> ${note.content}</p>`
    })
    
    html += '</div>'
  }
  
  // 尾注区域
  if (footnoteMap.endnotes.size > 0) {
    html += '<div class="docx-endnotes"><hr class="endnote-separator" />'
    html += '<div class="endnotes-title">尾注</div>'
    
    footnoteMap.endnotes.forEach((note, id) => {
      html += `<p class="endnote-item"><sup class="endnote-number">[${id}]</sup> ${note.content}</p>`
    })
    
    html += '</div>'
  }
  
  return html
}

// 自定义 docx 解析器（保留更多样式信息）
async function parseDocxCustom(bytes: Uint8Array): Promise<string> {
  // JSZip.loadAsync 需要 ArrayBuffer；Uint8Array.buffer 可能是 SharedArrayBuffer
  const ab = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer
  const zip = await JSZip.loadAsync(ab)
  
  const documentXml = await zip.file('word/document.xml')?.async('string')
  if (!documentXml) {
    throw new Error('找不到 document.xml')
  }

  const previousTheme = currentThemeFontMap
  const previousThemeColors = currentThemeColorMap
  const previousNumbering = currentNumberingMap
  currentThemeFontMap = await loadThemeFontMap(zip)
  currentThemeColorMap = await loadThemeColorMap(zip)
  currentNumberingMap = await loadNumberingMap(zip)

  try {
    const stylesXml = await zip.file('word/styles.xml')?.async('string')
    const styles = stylesXml ? parseStyles(stylesXml, currentThemeColorMap || undefined) : {}

    // 解析图片关系映射并提取图片
    const relsMap = await parseRelationships(zip)
    const imageMap = await extractImages(zip, relsMap)

    // 解析脚注和尾注
    const footnoteMap = await parseFootnotesAndEndnotes(zip, styles)
    
    // 解析页眉和页脚
    const headerFooterData = await parseHeadersAndFooters(zip, relsMap, styles, imageMap)

    const parser = new DOMParser()
    const doc = parser.parseFromString(documentXml, 'application/xml')

    const parseError = doc.querySelector('parsererror')
    if (parseError) {
      throw new Error('XML 解析失败')
    }

    // 获取 body 元素
    const body = doc.getElementsByTagName('w:body')[0]
    if (!body) {
      throw new Error('找不到文档主体')
    }

    let html = ''
    const numberingState: NumberingState = {}
    
    // 注意：页眉/页脚不再嵌入正文 HTML，而是由前端单独渲染
    // 页眉页脚数据通过 parseDocxComplete 函数返回
    
    // 包装正文内容
    html += '<div class="docx-body">'
    
    // 遍历 body 的直接子元素，处理段落和表格
    const children = body.childNodes
    for (let i = 0; i < children.length; i++) {
      const child = children[i] as Element
      if (child.nodeName === 'w:p') {
        html += parseParagraph(child, styles, imageMap, footnoteMap, undefined, numberingState)
      } else if (child.nodeName === 'w:tbl') {
        html += parseTable(child, styles, imageMap, footnoteMap)
      } else if (child.nodeName === 'w:sdt') {
        // 结构化文档标签，递归处理其内容
        const sdtContent = child.getElementsByTagName('w:sdtContent')[0]
        if (sdtContent) {
          const sdtChildren = sdtContent.childNodes
          for (let j = 0; j < sdtChildren.length; j++) {
            const sdtChild = sdtChildren[j] as Element
            if (sdtChild.nodeName === 'w:p') {
              html += parseParagraph(sdtChild, styles, imageMap, footnoteMap, undefined, numberingState)
            } else if (sdtChild.nodeName === 'w:tbl') {
              html += parseTable(sdtChild, styles, imageMap, footnoteMap)
            }
          }
        }
      }
    }
    
    html += '</div>'

    // 添加脚注/尾注区域
    html += generateFootnotesHtml(footnoteMap)
    
    // 注意：页脚不再嵌入正文 HTML，由前端单独渲染

    return html || '<p></p>'
  } finally {
    currentThemeFontMap = previousTheme
  }
}

// 解析页面设置 (从 sectPr 元素获取)
function parsePageSettings(doc: Document): PageSettings {
  // 默认 A4 纸设置 (pt)
  const defaults: PageSettings = {
    width: 595,        // A4 宽度 (210mm ≈ 595pt)
    height: 842,       // A4 高度 (297mm ≈ 842pt)
    marginTop: 72,     // 1 英寸 = 72pt
    marginBottom: 72,
    marginLeft: 90,    // 1.25 英寸 ≈ 90pt
    marginRight: 90,
    headerHeight: 36,  // 0.5 英寸
    footerHeight: 36,
    orientation: 'portrait'
  }
  
  // 查找 sectPr 元素（通常在 body 末尾或 sectPr 标签中）
  const sectPr = doc.getElementsByTagName('w:sectPr')[0]
  if (!sectPr) return defaults
  
  // 解析页面大小 (w:pgSz)
  const pgSz = sectPr.getElementsByTagName('w:pgSz')[0]
  if (pgSz) {
    const w = pgSz.getAttribute('w:w')
    const h = pgSz.getAttribute('w:h')
    const orient = pgSz.getAttribute('w:orient')
    // Word 使用 twips (1/20 pt)
    if (w) defaults.width = Math.round(parseInt(w) / 20)
    if (h) defaults.height = Math.round(parseInt(h) / 20)
    if (orient === 'landscape') defaults.orientation = 'landscape'
  }
  
  // 解析页边距 (w:pgMar)
  const pgMar = sectPr.getElementsByTagName('w:pgMar')[0]
  if (pgMar) {
    const top = pgMar.getAttribute('w:top')
    const bottom = pgMar.getAttribute('w:bottom')
    const left = pgMar.getAttribute('w:left')
    const right = pgMar.getAttribute('w:right')
    const header = pgMar.getAttribute('w:header')
    const footer = pgMar.getAttribute('w:footer')
    
    if (top) defaults.marginTop = Math.round(parseInt(top) / 20)
    if (bottom) defaults.marginBottom = Math.round(parseInt(bottom) / 20)
    if (left) defaults.marginLeft = Math.round(parseInt(left) / 20)
    if (right) defaults.marginRight = Math.round(parseInt(right) / 20)
    if (header) defaults.headerHeight = Math.round(parseInt(header) / 20)
    if (footer) defaults.footerHeight = Math.round(parseInt(footer) / 20)
  }
  
  return defaults
}

// 页眉/页脚引用类型映射
interface HeaderFooterRefs {
  headerFirst?: string    // rId
  headerEven?: string
  headerDefault?: string
  footerFirst?: string
  footerEven?: string
  footerDefault?: string
}

// 解析节配置（基于 ONLYOFFICE SectPr 模型）
async function parseSectionConfig(
  sectPr: Element, 
  zip: JSZip, 
  relsMap: Record<string, string>, 
  styles: Record<string, any>, 
  imageMap: ImageMap
): Promise<SectionConfig> {
  console.log('[SectionConfig] 开始解析节配置...')
  
  // 解析页面设置
  const pageSettings = parseSectPrPageSettings(sectPr)
  console.log('[SectionConfig] 页面设置:', pageSettings)
  
  // 检查 titlePage 设置（首页不同）
  const titlePageEl = sectPr.getElementsByTagName('w:titlePg')[0]
  const titlePage = !!titlePageEl
  console.log('[SectionConfig] 首页不同 (titlePage):', titlePage)
  
  // 检查 evenAndOddHeaders 设置（奇偶页不同）
  // 注意：这个设置通常在 settings.xml 中，这里先默认 false
  const evenAndOddHeaders = false
  
  // 解析页眉/页脚引用
  const refs: HeaderFooterRefs = {}
  
  // 解析 w:headerReference
  const headerRefs = sectPr.getElementsByTagName('w:headerReference')
  console.log('[SectionConfig] 找到', headerRefs.length, '个页眉引用')
  for (let i = 0; i < headerRefs.length; i++) {
    const ref = headerRefs[i]
    const type = ref.getAttribute('w:type') || 'default'
    const rId = ref.getAttribute('r:id')
    console.log('[SectionConfig] 页眉引用:', { type, rId, target: rId ? relsMap[rId] : 'N/A' })
    if (rId) {
      if (type === 'first') refs.headerFirst = rId
      else if (type === 'even') refs.headerEven = rId
      else refs.headerDefault = rId  // default 或其他
    }
  }
  
  // 解析 w:footerReference
  const footerRefs = sectPr.getElementsByTagName('w:footerReference')
  console.log('[SectionConfig] 找到', footerRefs.length, '个页脚引用')
  for (let i = 0; i < footerRefs.length; i++) {
    const ref = footerRefs[i]
    const type = ref.getAttribute('w:type') || 'default'
    const rId = ref.getAttribute('r:id')
    console.log('[SectionConfig] 页脚引用:', { type, rId, target: rId ? relsMap[rId] : 'N/A' })
    if (rId) {
      if (type === 'first') refs.footerFirst = rId
      else if (type === 'even') refs.footerEven = rId
      else refs.footerDefault = rId
    }
  }
  
  console.log('[SectionConfig] 解析的引用:', refs)
  
  // 解析各个页眉/页脚文件
  const config: SectionConfig = {
    pageSettings,
    titlePage,
    evenAndOddHeaders
  }
  
  // 解析页眉
  if (refs.headerFirst) {
    console.log('[SectionConfig] 解析首页页眉:', relsMap[refs.headerFirst])
    config.headerFirst = await parseHeaderFooterFile(zip, relsMap[refs.headerFirst], styles, imageMap, true)
  }
  if (refs.headerEven) {
    console.log('[SectionConfig] 解析偶数页页眉:', relsMap[refs.headerEven])
    config.headerEven = await parseHeaderFooterFile(zip, relsMap[refs.headerEven], styles, imageMap, true)
  }
  if (refs.headerDefault) {
    console.log('[SectionConfig] 解析默认页眉:', relsMap[refs.headerDefault])
    config.headerDefault = await parseHeaderFooterFile(zip, relsMap[refs.headerDefault], styles, imageMap, true)
  }
  
  // 解析页脚
  if (refs.footerFirst) {
    console.log('[SectionConfig] 解析首页页脚:', relsMap[refs.footerFirst])
    config.footerFirst = await parseHeaderFooterFile(zip, relsMap[refs.footerFirst], styles, imageMap, false)
  }
  if (refs.footerEven) {
    console.log('[SectionConfig] 解析偶数页页脚:', relsMap[refs.footerEven])
    config.footerEven = await parseHeaderFooterFile(zip, relsMap[refs.footerEven], styles, imageMap, false)
  }
  if (refs.footerDefault) {
    console.log('[SectionConfig] 解析默认页脚:', relsMap[refs.footerDefault])
    config.footerDefault = await parseHeaderFooterFile(zip, relsMap[refs.footerDefault], styles, imageMap, false)
  }
  
  console.log('[SectionConfig] 最终配置:', {
    titlePage: config.titlePage,
    evenAndOddHeaders: config.evenAndOddHeaders,
    hasHeaderFirst: !!config.headerFirst,
    hasHeaderEven: !!config.headerEven,
    hasHeaderDefault: !!config.headerDefault,
    hasFooterFirst: !!config.footerFirst,
    hasFooterEven: !!config.footerEven,
    hasFooterDefault: !!config.footerDefault,
    headerDefaultContent: config.headerDefault?.html?.substring(0, 100),
    footerDefaultContent: config.footerDefault?.html?.substring(0, 100)
  })
  
  return config
}

// 从 sectPr 元素解析页面设置
function parseSectPrPageSettings(sectPr: Element): PageSettings {
  const settings: PageSettings = {
    width: 595,
    height: 842,
    marginTop: 72,
    marginBottom: 72,
    marginLeft: 90,
    marginRight: 90,
    headerHeight: 36,
    footerHeight: 36,
    orientation: 'portrait'
  }
  
  const pgSz = sectPr.getElementsByTagName('w:pgSz')[0]
  if (pgSz) {
    const w = pgSz.getAttribute('w:w')
    const h = pgSz.getAttribute('w:h')
    const orient = pgSz.getAttribute('w:orient')
    if (w) settings.width = Math.round(parseInt(w) / 20)
    if (h) settings.height = Math.round(parseInt(h) / 20)
    if (orient === 'landscape') settings.orientation = 'landscape'
  }
  
  const pgMar = sectPr.getElementsByTagName('w:pgMar')[0]
  if (pgMar) {
    const top = pgMar.getAttribute('w:top')
    const bottom = pgMar.getAttribute('w:bottom')
    const left = pgMar.getAttribute('w:left')
    const right = pgMar.getAttribute('w:right')
    const header = pgMar.getAttribute('w:header')
    const footer = pgMar.getAttribute('w:footer')
    
    if (top) settings.marginTop = Math.round(parseInt(top) / 20)
    if (bottom) settings.marginBottom = Math.round(parseInt(bottom) / 20)
    if (left) settings.marginLeft = Math.round(parseInt(left) / 20)
    if (right) settings.marginRight = Math.round(parseInt(right) / 20)
    if (header) settings.headerHeight = Math.round(parseInt(header) / 20)
    if (footer) settings.footerHeight = Math.round(parseInt(footer) / 20)
  }
  
  return settings
}

// 解析单个页眉/页脚文件
async function parseHeaderFooterFile(
  zip: JSZip, 
  target: string | undefined, 
  styles: Record<string, any>, 
  imageMap: ImageMap,
  isHeader: boolean,
  imageCtx?: ImageParseContext
): Promise<HeaderFooterContent | undefined> {
  if (!target) return undefined
  
  const path = target.startsWith('word/') ? target : `word/${target}`
  const xml = await zip.file(path)?.async('string')
  if (!xml) return undefined
  
  const parser = new DOMParser()
  const doc = parser.parseFromString(xml, 'application/xml')
  
  const style = parseHeaderFooterStyleFromXml(doc, styles, isHeader ? 'right' : 'center')
  
  const body = doc.documentElement
  const paras = body.getElementsByTagName('w:p')
  const parts = collectHeaderFooterParts(paras, styles, imageMap, imageCtx)
  
  let html = ''
  if (parts.length > 0) {
    html = parts.join('')
  } else if (!isHeader) {
    html = '<p class="header-footer-para">{PAGE}</p>'
  }
  
  return { html, style }
}

function parseHeaderFooterStyleFromXml(
  xmlDoc: Document, 
  styles: Record<string, any> = {}, 
  defaultAlignment: 'left' | 'center' | 'right' = 'right'
): HeaderFooterStyle {
  const style: HeaderFooterStyle = {
    fontFamily: 'var(--word-font-family-cn)',
    fontSize: '9pt',
    color: 'var(--word-ink-muted)',
    alignment: defaultAlignment
  }

  const firstPara = xmlDoc.getElementsByTagName('w:p')[0]
  if (!firstPara) return style

  const applyStyle = (source: any) => {
    if (!source) return
    if (source.fontFamily) style.fontFamily = source.fontFamily
    if (source.fontSize) style.fontSize = source.fontSize
    // 在深色模式下，忽略黑色或接近黑色的颜色，保持使用 CSS 变量
    if (source.color && !shouldIgnoreColorInDarkMode(source.color)) {
      style.color = source.color
    }
    if (source.alignment) style.alignment = source.alignment
    if (source.lineHeight) style.lineHeight = source.lineHeight
  }

  const applyRunProps = (rPr?: Element) => {
    if (!rPr) return
    const sz = rPr.getElementsByTagName('w:sz')[0]
    if (sz) {
      const val = sz.getAttribute('w:val')
      if (val) {
        style.fontSize = `${parseInt(val) / 2}pt`
      }
    }

    const rFonts = rPr.getElementsByTagName('w:rFonts')[0]
    if (rFonts) {
      const fontName = resolveFontNameFromRFonts(rFonts)
      if (fontName) {
        style.fontFamily = getSafeFontFamily(fontName) || `"${fontName}", serif`
      }
    }

    const color = rPr.getElementsByTagName('w:color')[0]
    if (color) {
      const val = color.getAttribute('w:val')
      if (val && val !== 'auto') {
        const colorHex = `#${val}`
        // 在深色模式下，忽略黑色或接近黑色的颜色，保持使用 CSS 变量
        if (!shouldIgnoreColorInDarkMode(colorHex)) {
          style.color = colorHex
        }
      }
    }
  }

  const applyLineHeight = (spacing?: Element) => {
    if (!spacing) return
    const line = spacing.getAttribute('w:line')
    if (!line) return
    const lineVal = parseInt(line)
    if (!lineVal) return
    const lineRule = spacing.getAttribute('w:lineRule')
    if (lineRule === 'exact' || lineRule === 'atLeast') {
      style.lineHeight = `${(lineVal / 20).toFixed(1)}pt`
    } else {
      style.lineHeight = (lineVal / 240).toFixed(2)
    }
  }

  const pPr = firstPara.getElementsByTagName('w:pPr')[0]
  if (pPr) {
    const pStyle = pPr.getElementsByTagName('w:pStyle')[0]
    if (pStyle) {
      const styleId = pStyle.getAttribute('w:val')
      if (styleId && styles[styleId]) {
        applyStyle(styles[styleId])
      }
    }

    const jc = pPr.getElementsByTagName('w:jc')[0]
    if (jc) {
      const align = jc.getAttribute('w:val')
      if (align === 'center') style.alignment = 'center'
      else if (align === 'left' || align === 'start') style.alignment = 'left'
      else if (align === 'right' || align === 'end') style.alignment = 'right'
    }

    const spacing = pPr.getElementsByTagName('w:spacing')[0]
    applyLineHeight(spacing)

    const pBdr = pPr.getElementsByTagName('w:pBdr')[0]
    if (pBdr) {
      const bottom = pBdr.getElementsByTagName('w:bottom')[0]
      if (bottom) {
        const val = bottom.getAttribute('w:val')
        if (val && val !== 'none' && val !== 'nil') {
          style.borderBottom = true
        }
      }
      const top = pBdr.getElementsByTagName('w:top')[0]
      if (top) {
        const val = top.getAttribute('w:val')
        if (val && val !== 'none' && val !== 'nil') {
          style.borderTop = true
        }
      }
    }

    const pRpr = pPr.getElementsByTagName('w:rPr')[0]
    applyRunProps(pRpr)
  }

  const firstRun = firstPara.getElementsByTagName('w:r')[0]
  if (firstRun) {
    const rPr = firstRun.getElementsByTagName('w:rPr')[0]
    applyRunProps(rPr)
  }

  return style
}
async function parseDocxCustomComplete(bytes: Uint8Array): Promise<DocxParseResult> {
  const ab = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer
  const zip = await JSZip.loadAsync(ab)
  
  const documentXml = await zip.file('word/document.xml')?.async('string')
  if (!documentXml) {
    throw new Error('找不到 document.xml')
  }

  const previousTheme = currentThemeFontMap
  const previousThemeColors = currentThemeColorMap
  const previousNumbering = currentNumberingMap
  currentThemeFontMap = await loadThemeFontMap(zip)
  currentThemeColorMap = await loadThemeColorMap(zip)
  currentNumberingMap = await loadNumberingMap(zip)

  try {
    const stylesXml = await zip.file('word/styles.xml')?.async('string')
    const styles = stylesXml ? parseStyles(stylesXml, currentThemeColorMap || undefined) : {}

    // 解析图片关系映射并提取图片
    const relsMap = await parseRelationships(zip)
    const imageMap = await extractImages(zip, relsMap)

    // 解析脚注和尾注
    const footnoteMap = await parseFootnotesAndEndnotes(zip, styles)

    const parser = new DOMParser()
    const doc = parser.parseFromString(documentXml, 'application/xml')

    const parseError = doc.querySelector('parsererror')
    if (parseError) {
      throw new Error('XML 解析失败')
    }
    
    // 解析页面设置
    const pageSettings = parsePageSettings(doc)

    // 获取 body 元素
    const body = doc.getElementsByTagName('w:body')[0]
    if (!body) {
      throw new Error('找不到文档主体')
    }

    // 解析正文内容（不包含页眉页脚）
    let bodyHtml = ''
    const numberingState: NumberingState = {}
    const children = body.childNodes
    for (let i = 0; i < children.length; i++) {
      const child = children[i] as Element
      if (child.nodeName === 'w:p') {
        bodyHtml += parseParagraph(child, styles, imageMap, footnoteMap, undefined, numberingState)
      } else if (child.nodeName === 'w:tbl') {
        bodyHtml += parseTable(child, styles, imageMap, footnoteMap)
      } else if (child.nodeName === 'w:sdt') {
        const sdtContent = child.getElementsByTagName('w:sdtContent')[0]
        if (sdtContent) {
          const sdtChildren = sdtContent.childNodes
          for (let j = 0; j < sdtChildren.length; j++) {
            const sdtChild = sdtChildren[j] as Element
            if (sdtChild.nodeName === 'w:p') {
              bodyHtml += parseParagraph(sdtChild, styles, imageMap, footnoteMap, undefined, numberingState)
            } else if (sdtChild.nodeName === 'w:tbl') {
              bodyHtml += parseTable(sdtChild, styles, imageMap, footnoteMap)
            }
          }
        }
      }
    }

    // 添加脚注/尾注区域
    bodyHtml += generateFootnotesHtml(footnoteMap)

    // 解析页眉页脚
    let headerHtml = ''
    let footerHtml = ''
    let headerStyle: HeaderFooterStyle = { fontSize: '9pt', alignment: 'right', color: '#666' }
    let footerStyle: HeaderFooterStyle = { fontSize: '9pt', alignment: 'center', color: '#666' }
    
    // 找到页眉页脚文件
    const headerFiles: string[] = []
    const footerFiles: string[] = []
    
    for (const [, target] of Object.entries(relsMap)) {
      if (target.includes('header') && target.endsWith('.xml')) {
        headerFiles.push(target.startsWith('word/') ? target : `word/${target}`)
      } else if (target.includes('footer') && target.endsWith('.xml')) {
        footerFiles.push(target.startsWith('word/') ? target : `word/${target}`)
      }
    }
    
    // 解析页眉
    if (headerFiles.length > 0) {
      const defaultHeader = headerFiles.find(f => f.includes('header2.xml')) || 
                            headerFiles.find(f => f.includes('header1.xml')) || 
                            headerFiles[0]
      const headerXml = await zip.file(defaultHeader)?.async('string')
      if (headerXml) {
        const headerDoc = parser.parseFromString(headerXml, 'application/xml')
        headerStyle = parseHeaderFooterStyleFromXml(headerDoc, styles, 'right')
        
        const paras = headerDoc.getElementsByTagName('w:p')
        const headerParts = collectHeaderFooterParts(paras, styles, imageMap)
        
        if (headerParts.length > 0) {
          headerHtml = headerParts.join('')
        }
      }
    }
    
    if (footerFiles.length > 0) {
      const defaultFooter = footerFiles.find(f => f.includes('footer2.xml')) || 
                            footerFiles.find(f => f.includes('footer1.xml')) || 
                            footerFiles[0]
      const footerXml = await zip.file(defaultFooter)?.async('string')
      if (footerXml) {
        const footerDoc = parser.parseFromString(footerXml, 'application/xml')
        footerStyle = parseHeaderFooterStyleFromXml(footerDoc, styles, 'center')
        
        const paras = footerDoc.getElementsByTagName('w:p')
        const footerParts = collectHeaderFooterParts(paras, styles, imageMap)
        
        if (footerParts.length > 0) {
          footerHtml = footerParts.join('')
        } else {
          footerHtml = '<p class="header-footer-para">{PAGE}</p>'
        }
      }
    }

    const sectPr = doc.getElementsByTagName('w:sectPr')[0]
    let sectionConfig: SectionConfig | undefined
    
    if (sectPr) {
      sectionConfig = await parseSectionConfig(sectPr, zip, relsMap, styles, imageMap)
    }
    
    // 构建完整的文档模型
    const documentModel: DocumentModel = {
      sections: sectionConfig ? [sectionConfig] : [{
        pageSettings,
        titlePage: false,
        evenAndOddHeaders: false,
        headerDefault: headerHtml ? { html: headerHtml, style: headerStyle } : undefined,
        footerDefault: footerHtml ? { html: footerHtml, style: footerStyle } : undefined
      }],
      bodyHtml,
      footnotes: footnoteMap
    }

    console.log('[DocumentModel] 创建完成:', {
      sectionsCount: documentModel.sections.length,
      hasSectionConfig: !!sectionConfig,
      section0: documentModel.sections[0] ? {
        titlePage: documentModel.sections[0].titlePage,
        hasHeaderDefault: !!documentModel.sections[0].headerDefault,
        hasFooterDefault: !!documentModel.sections[0].footerDefault
      } : null
    })

    return {
      bodyHtml,
      headerHtml,
      footerHtml,
      headerStyle,
      footerStyle,
      pageSettings,
      documentModel
    }
  } finally {
    currentThemeFontMap = previousTheme
  }
}

// 导出完整解析函数
export async function parseDocxComplete(base64Data: string): Promise<DocxParseResult> {
  const binaryString = atob(base64Data)
  const bytes = new Uint8Array(binaryString.length)
  for (let i = 0; i < binaryString.length; i++) {
    bytes[i] = binaryString.charCodeAt(i)
  }
  
  if (!isValidZip(bytes)) {
    throw new Error('不是有效的 .docx 文件')
  }
  
  return parseDocxCustomComplete(bytes)
}

// 表格条件格式选项
interface TableLookOptions {
  firstRow: boolean      // 首行特殊样式
  lastRow: boolean       // 末行特殊样式
  firstColumn: boolean   // 首列特殊样式
  lastColumn: boolean    // 末列特殊样式
  noHBand: boolean       // 禁用横向条带
  noVBand: boolean       // 禁用纵向条带
}

// 解析 tblLook 属性
function parseTblLook(tblLook: Element | null): TableLookOptions {
  const defaults: TableLookOptions = {
    firstRow: true,
    lastRow: false,
    firstColumn: false,
    lastColumn: false,
    noHBand: true,
    noVBand: true,
  }
  
  if (!tblLook) return defaults
  
  // tblLook 可以使用 val 属性（位掩码）或单独的属性
  const val = tblLook.getAttribute('w:val')
  if (val) {
    // val 是一个十六进制位掩码
    const mask = parseInt(val, 16)
    return {
      firstRow: !!(mask & 0x0020),
      lastRow: !!(mask & 0x0040),
      firstColumn: !!(mask & 0x0080),
      lastColumn: !!(mask & 0x0100),
      noHBand: !!(mask & 0x0200),
      noVBand: !!(mask & 0x0400),
    }
  }
  
  // 也可能使用单独的属性
  return {
    firstRow: tblLook.getAttribute('w:firstRow') !== '0',
    lastRow: tblLook.getAttribute('w:lastRow') === '1',
    firstColumn: tblLook.getAttribute('w:firstColumn') === '1',
    lastColumn: tblLook.getAttribute('w:lastColumn') === '1',
    noHBand: tblLook.getAttribute('w:noHBand') !== '0',
    noVBand: tblLook.getAttribute('w:noVBand') !== '0',
  }
}

interface TableBorderStyle {
  widthPt: number
  color: string
}

interface TableBorders {
  top?: TableBorderStyle
  right?: TableBorderStyle
  bottom?: TableBorderStyle
  left?: TableBorderStyle
  insideH?: TableBorderStyle
  insideV?: TableBorderStyle
}

interface TableStylePr {
  backgroundColor?: string
  textColor?: string
  bold?: boolean
}

function parseBorderStyle(borderEl?: Element | null): TableBorderStyle | undefined {
  if (!borderEl) return undefined
  const val = (borderEl.getAttribute('w:val') || '').toLowerCase()
  if (!val || val === 'none' || val === 'nil') return undefined
  const sz = borderEl.getAttribute('w:sz')
  const color = borderEl.getAttribute('w:color')
  const themeColor = borderEl.getAttribute('w:themeColor')
  const widthPt = sz ? parseInt(sz, 10) / 8 : 0.5
  let resolvedColor = color && color !== 'auto' ? color : null
  if (themeColor) {
    const themeResolved = resolveThemeColor(themeColor)
    if (themeResolved) resolvedColor = themeResolved
  }
  return {
    widthPt,
    color: resolvedColor ? `#${resolvedColor}` : 'var(--word-rule)',
  }
}

function parseTableBorders(tblBorders?: Element | null): TableBorders | null {
  if (!tblBorders) return null
  return {
    top: parseBorderStyle(tblBorders.getElementsByTagName('w:top')[0]),
    right: parseBorderStyle(tblBorders.getElementsByTagName('w:right')[0]),
    bottom: parseBorderStyle(tblBorders.getElementsByTagName('w:bottom')[0]),
    left: parseBorderStyle(tblBorders.getElementsByTagName('w:left')[0]),
    insideH: parseBorderStyle(tblBorders.getElementsByTagName('w:insideH')[0]),
    insideV: parseBorderStyle(tblBorders.getElementsByTagName('w:insideV')[0]),
  }
}

function mergeTableBorders(base?: TableBorders | null, override?: TableBorders | null): TableBorders | null {
  if (!base && !override) return null
  return {
    top: override?.top || base?.top,
    right: override?.right || base?.right,
    bottom: override?.bottom || base?.bottom,
    left: override?.left || base?.left,
    insideH: override?.insideH || base?.insideH,
    insideV: override?.insideV || base?.insideV,
  }
}

function mergeCellMargin(
  base?: { top?: number; right?: number; bottom?: number; left?: number } | null,
  override?: { top?: number; right?: number; bottom?: number; left?: number } | null
) {
  if (!base && !override) return undefined
  return {
    top: override?.top ?? base?.top,
    right: override?.right ?? base?.right,
    bottom: override?.bottom ?? base?.bottom,
    left: override?.left ?? base?.left,
  }
}

// 解析表格
function parseTable(
  tbl: Element,
  styles: Record<string, any>,
  imageMap: ImageMap = {},
  footnoteMap?: FootnoteMap,
  imageCtx?: ImageParseContext
): string {
  let tableStyle = 'border-collapse: collapse; margin: 0; table-layout: fixed;'
  let tableStyleId: string | null = null
  let tableBorders: TableBorders | null = null
  let tableCellMargin: { top?: number; right?: number; bottom?: number; left?: number } | undefined
  let headerRowStyle: TableStylePr | null = null
  let tableLayoutType: string | undefined
  
  // 获取表格网格定义（列宽）
  const tblGrid = tbl.getElementsByTagName('w:tblGrid')[0]
  const columnWidths: number[] = []
  let totalWidth = 0
  
  if (tblGrid) {
    const gridCols = tblGrid.getElementsByTagName('w:gridCol')
    for (let i = 0; i < gridCols.length; i++) {
      const w = gridCols[i].getAttribute('w:w')
      if (w) {
        const width = parseInt(w)
        columnWidths.push(width)
        totalWidth += width
      }
    }
  }
  
  // 解析表格属性
  const tblPr = tbl.getElementsByTagName('w:tblPr')[0]
  let tblLookOptions: TableLookOptions = {
    firstRow: true,
    lastRow: false,
    firstColumn: false,
    lastColumn: false,
    noHBand: true,
    noVBand: true,
  }
  if (tblPr) {
    // 解析表格样式引用
    const tblStyle = tblPr.getElementsByTagName('w:tblStyle')[0]
    if (tblStyle) {
      tableStyleId = tblStyle.getAttribute('w:val')
    }

    if (tableStyleId && styles[tableStyleId]) {
      const tableStyleData = styles[tableStyleId]
      tableBorders = tableStyleData.tblBorders || null
      tableCellMargin = tableStyleData.tblCellMargin
      headerRowStyle = tableStyleData.tblStylePr?.firstRow || null
    }
    
    // 解析 tblLook 条件格式
    const tblLook = tblPr.getElementsByTagName('w:tblLook')[0]
    tblLookOptions = parseTblLook(tblLook)
    
    // 表格对齐
    const jc = tblPr.getElementsByTagName('w:jc')[0]
    if (jc) {
      const val = jc.getAttribute('w:val')
      if (val === 'center') {
        tableStyle += ' margin-left: auto; margin-right: auto;'
      }
    }

    // 浮动表格定位（tblpPr）
    const tblpPr = tblPr.getElementsByTagName('w:tblpPr')[0]
    if (tblpPr) {
      const xSpec = tblpPr.getAttribute('w:tblpXSpec')
      if (xSpec === 'right') {
        tableStyle += ' margin-left: auto; margin-right: 0;'
      } else if (xSpec === 'center') {
        tableStyle += ' margin-left: auto; margin-right: auto;'
      } else if (xSpec === 'left') {
        tableStyle += ' margin-left: 0; margin-right: auto;'
      }
    }
    
    // 表格宽度
    const tblW = tblPr.getElementsByTagName('w:tblW')[0]
    let tableWidthCss: string | null = null
    if (tblW) {
      const w = tblW.getAttribute('w:w')
      const type = tblW.getAttribute('w:type')
      if (w && type === 'pct') {
        tableWidthCss = `${parseInt(w) / 50}%`
      } else if (w && (!type || type === 'dxa')) {
        tableWidthCss = `${parseInt(w) / 20}pt`
      } else if (type === 'auto') {
        tableWidthCss = totalWidth > 0 ? `${totalWidth / 20}pt` : '100%'
      }
    } else if (totalWidth > 0) {
      // 使用计算出的总宽度
      tableWidthCss = `${totalWidth / 20}pt`
    } else {
      tableWidthCss = '100%'
    }
    if (tableWidthCss) {
      tableStyle += ` width: ${tableWidthCss}; --table-width: ${tableWidthCss};`
    }

    const tblBordersEl = tblPr.getElementsByTagName('w:tblBorders')[0]
    const tblCellMarEl = tblPr.getElementsByTagName('w:tblCellMar')[0]
    const tblBordersFromPr = parseTableBorders(tblBordersEl)
    const tblCellMarginFromPr = parseCellMarginFromPr(tblCellMarEl)
    tableBorders = mergeTableBorders(tableBorders, tblBordersFromPr)
    tableCellMargin = mergeCellMargin(tableCellMargin, tblCellMarginFromPr)

    const tblLayout = tblPr.getElementsByTagName('w:tblLayout')[0]
    if (tblLayout) {
      const layoutType = tblLayout.getAttribute('w:type')
      if (layoutType) {
        tableLayoutType = layoutType
        if (layoutType === 'autofit') {
          tableStyle = tableStyle.replace(/table-layout:\s*fixed;?/i, 'table-layout: auto;')
        }
      }
    }
  } else {
    tableStyle += ' width: 100%;'
  }
  
  const dataAttrs: string[] = []
  if (columnWidths.length > 0) {
    dataAttrs.push(`data-tbl-grid="${columnWidths.join(',')}"`)
  }
  if (totalWidth > 0) {
    dataAttrs.push(`data-tbl-grid-total="${totalWidth}"`)
  }
  if (tableLayoutType) {
    dataAttrs.push(`data-tbl-layout="${tableLayoutType}"`)
  }
  let html = `<table style="${tableStyle}" ${dataAttrs.join(' ')}>`
  
  // 如果有列宽定义，添加 colgroup
  if (columnWidths.length > 0) {
    html += '<colgroup>'
    for (const width of columnWidths) {
      // 转换为百分比或固定宽度
      if (totalWidth > 0) {
        const pct = (width / totalWidth * 100).toFixed(2)
        html += `<col style="width: ${pct}%;">`
      } else {
        html += `<col style="width: ${width / 20}pt;">`
      }
    }
    html += '</colgroup>'
  }
  
  // 解析表格行（直接子元素，避免嵌套表格问题）
  const children = tbl.childNodes
  let rowIndex = 0
  const totalRows = Array.from(children).filter(c => (c as Element).nodeName === 'w:tr').length
  
  for (let i = 0; i < children.length; i++) {
    const child = children[i] as Element
    if (child.nodeName === 'w:tr') {
      const isFirstRow = rowIndex === 0
      const isHeaderRow = isFirstRow && tblLookOptions.firstRow
      const isLastRow = rowIndex === totalRows - 1
      html += parseTableRow(
        child, 
        styles, 
        isHeaderRow,
        isFirstRow,
        columnWidths, 
        totalWidth, 
        imageMap, 
        footnoteMap, 
        imageCtx,
        headerRowStyle || undefined,
        tblLookOptions.firstColumn,
        isLastRow,
        tableBorders || undefined,
        tableCellMargin
      )
      rowIndex++
    }
  }
  
  html += '</table>'
  return html
}

// 解析表格行
function parseTableRow(
  tr: Element,
  styles: Record<string, any>,
  isHeaderRow: boolean,
  isFirstRow: boolean,
  columnWidths: number[] = [],
  totalWidth: number = 0,
  imageMap: ImageMap = {},
  footnoteMap?: FootnoteMap,
  imageCtx?: ImageParseContext,
  headerRowStyle?: TableStylePr,
  firstColumnStyle?: boolean,
  isLastRow?: boolean,
  tableBorders?: TableBorders,
  tableCellMargin?: { top?: number; right?: number; bottom?: number; left?: number }
): string {
  let rowStyle = ''
  
  // 解析行属性
  const trPr = tr.getElementsByTagName('w:trPr')[0]
  if (trPr) {
    // 行高
    const trHeight = trPr.getElementsByTagName('w:trHeight')[0]
    if (trHeight) {
      const val = trHeight.getAttribute('w:val')
      if (val) {
        rowStyle += `height: ${parseInt(val) / 20}pt;`
      }
    }
  }
  
  let html = rowStyle ? `<tr style="${rowStyle}">` : '<tr>'
  
  // 解析单元格（只处理直接子元素）
  let colIndex = 0
  const children = tr.childNodes
  const totalCols = Array.from(children).filter(c => (c as Element).nodeName === 'w:tc').length
  
  for (let i = 0; i < children.length; i++) {
    const child = children[i] as Element
    if (child.nodeName === 'w:tc') {
      const isFirstCol = colIndex === 0 && firstColumnStyle
      const cellResult = parseTableCell(
        child, 
        styles, 
        isHeaderRow, 
        columnWidths, 
        totalWidth, 
        colIndex, 
        imageMap, 
        footnoteMap, 
        imageCtx,
        isFirstCol,
        headerRowStyle,
        tableBorders,
        tableCellMargin,
        totalCols,
        !!isLastRow,
        isFirstRow
      )
      html += cellResult.html
      colIndex += cellResult.colSpan
    }
  }
  
  html += '</tr>'
  return html
}

function parseCellMarginFromPr(pr?: Element | null) {
  if (!pr) return undefined
  const pick = (tag: string) => {
    const node = pr.getElementsByTagName(tag)[0]
    if (!node) return undefined
    const w = node.getAttribute('w:w') || node.getAttribute('w:val')
    const type = node.getAttribute('w:type')
    if (!w) return undefined
    if (type && type !== 'dxa') return undefined
    const twips = parseInt(w)
    if (!Number.isFinite(twips) || twips <= 0) return undefined
    return twips / 20
  }
  return {
    top: pick('w:top'),
    right: pick('w:right'),
    bottom: pick('w:bottom'),
    left: pick('w:left'),
  }
}

// 解析表格单元格
function parseTableCell(
  tc: Element,
  styles: Record<string, any>,
  isHeader: boolean,
  columnWidths: number[] = [],
  totalWidth: number = 0,
  colIndex: number = 0,
  imageMap: ImageMap = {},
  footnoteMap?: FootnoteMap,
  imageCtx?: ImageParseContext,
  isFirstColumn?: boolean,
  headerRowStyle?: TableStylePr,
  tableBorders?: TableBorders,
  tableCellMargin?: { top?: number; right?: number; bottom?: number; left?: number },
  totalCols: number = 0,
  isLastRow: boolean = false,
  isFirstRow: boolean = false
): { html: string; colSpan: number } {
  let cellStyle = 'padding: 2pt 5pt; vertical-align: middle;'
  let colspan = ''
  let colSpan = 1
  let hasCellBackground = false
  let hasBorderStyle = false
  
  // 解析单元格属性
  const tcPr = tc.getElementsByTagName('w:tcPr')[0]
  const applyCellMargin = (margin?: { top?: number; right?: number; bottom?: number; left?: number }) => {
    if (!margin) return
    const top = margin.top
    const right = margin.right
    const bottom = margin.bottom
    const left = margin.left
    if (top == null && right == null && bottom == null && left == null) return
    cellStyle = cellStyle.replace(/padding:[^;]+;?/i, '')
    if (top != null) cellStyle += ` padding-top: ${top}pt;`
    if (right != null) cellStyle += ` padding-right: ${right}pt;`
    if (bottom != null) cellStyle += ` padding-bottom: ${bottom}pt;`
    if (left != null) cellStyle += ` padding-left: ${left}pt;`
  }
  
  if (tcPr) {
    // 合并列
    const gridSpan = tcPr.getElementsByTagName('w:gridSpan')[0]
    if (gridSpan) {
      const val = gridSpan.getAttribute('w:val')
      if (val && parseInt(val) > 1) {
        colSpan = parseInt(val)
        colspan = ` colspan="${val}"`
      }
    }
    
    // 合并行 - 检查是否是被合并的单元格
    const vMerge = tcPr.getElementsByTagName('w:vMerge')[0]
    if (vMerge) {
      const val = vMerge.getAttribute('w:val')
      // 如果没有 val 属性或 val="continue"，说明是被合并的单元格，跳过
      if (!val || val === 'continue') {
        return { html: '', colSpan }
      }
      // val="restart" 表示这是合并的起始单元格
    }
    
    // 单元格宽度 - 优先使用 tcW，否则从 columnWidths 计算
    const tcW = tcPr.getElementsByTagName('w:tcW')[0]
    if (tcW) {
      const w = tcW.getAttribute('w:w')
      const type = tcW.getAttribute('w:type')
      if (w && (!type || type === 'dxa')) {
        cellStyle += ` width: ${parseInt(w) / 20}pt;`
      } else if (w && type === 'pct') {
        cellStyle += ` width: ${parseInt(w) / 50}%;`
      }
    }
    
    // 单元格背景色（优先级：单元格定义 > 行定义）
    const shd = tcPr.getElementsByTagName('w:shd')[0]
    if (shd) {
      const fill = shd.getAttribute('w:fill')
      const themeFill = shd.getAttribute('w:themeFill')
      
      // 优先使用主题颜色
      if (themeFill) {
        const resolvedColor = resolveThemeColor(themeFill)
        if (resolvedColor && resolvedColor.toUpperCase() !== 'FFFFFF') {
          cellStyle += ` background-color: #${resolvedColor};`
          hasCellBackground = true
        }
      } else if (fill && fill !== 'auto' && fill.toUpperCase() !== 'FFFFFF') {
        cellStyle += ` background-color: #${fill};`
        hasCellBackground = true
      }
    }
    
    // 单元格边框
    const tcBorders = tcPr.getElementsByTagName('w:tcBorders')[0]
    if (tcBorders) {
      const borders: string[] = []
      const addBorder = (tag: string, cssProp: string) => {
        const b = tcBorders.getElementsByTagName(tag)[0]
        if (!b) return
        const val = (b.getAttribute('w:val') || '').toLowerCase()
        if (!val || val === 'none' || val === 'nil') return
        const sz = b.getAttribute('w:sz')
        const color = b.getAttribute('w:color')
        const widthPt = sz ? parseInt(sz) / 8 : 0.5
        const cssColor = color && color !== 'auto' ? `#${color}` : 'var(--word-rule)'
        borders.push(`${cssProp}: ${widthPt}pt solid ${cssColor};`)
      }
      addBorder('w:top', 'border-top')
      addBorder('w:right', 'border-right')
      addBorder('w:bottom', 'border-bottom')
      addBorder('w:left', 'border-left')
      if (borders.length) {
        cellStyle += borders.join('')
        hasBorderStyle = true
      }
    }
    
    // 单元格内边距
    const tcMar = tcPr.getElementsByTagName('w:tcMar')[0]
    const cellMargin = parseCellMarginFromPr(tcMar) || tableCellMargin
    applyCellMargin(cellMargin)

    // 垂直对齐
    const vAlign = tcPr.getElementsByTagName('w:vAlign')[0]
    if (vAlign) {
      const val = vAlign.getAttribute('w:val')
      if (val === 'center') {
        cellStyle += ' vertical-align: middle;'
      } else if (val === 'bottom') {
        cellStyle += ' vertical-align: bottom;'
      } else if (val === 'top') {
        cellStyle += ' vertical-align: top;'
      }
    }
  }
  
  if (!tcPr && tableCellMargin) {
    applyCellMargin(tableCellMargin)
  }

  const hasWidthStyle = /(?:^|;)\s*width:\s*/i.test(cellStyle)
  if (!hasWidthStyle && columnWidths.length && totalWidth > 0) {
    const widthTwips = columnWidths
      .slice(colIndex, colIndex + colSpan)
      .reduce((sum, w) => sum + w, 0)
    if (widthTwips > 0) {
      cellStyle += ` width: ${widthTwips / 20}pt;`
    }
  }

  if (!hasBorderStyle && tableBorders) {
    const isLastCol = totalCols > 0 && colIndex + colSpan >= totalCols
    const topBorder = isFirstRow ? tableBorders.top : (tableBorders.insideH || tableBorders.top)
    const bottomBorder = isLastRow ? tableBorders.bottom : (tableBorders.insideH || tableBorders.bottom)
    const leftBorder = isFirstColumn ? tableBorders.left : (tableBorders.insideV || tableBorders.left)
    const rightBorder = isLastCol ? tableBorders.right : (tableBorders.insideV || tableBorders.right)
    if (topBorder) cellStyle += ` border-top: ${topBorder.widthPt}pt solid ${topBorder.color};`
    if (bottomBorder) cellStyle += ` border-bottom: ${bottomBorder.widthPt}pt solid ${bottomBorder.color};`
    if (leftBorder) cellStyle += ` border-left: ${leftBorder.widthPt}pt solid ${leftBorder.color};`
    if (rightBorder) cellStyle += ` border-right: ${rightBorder.widthPt}pt solid ${rightBorder.color};`
    hasBorderStyle = !!(topBorder || bottomBorder || leftBorder || rightBorder)
  }

  if (!hasBorderStyle) {
    cellStyle += ' border: 0.5pt solid var(--word-rule);'
  }
  
  // 解析单元格内容（只处理直接子段落）
  let content = ''
  const children = tc.childNodes
  let paragraphCount = 0
  
  for (let i = 0; i < children.length; i++) {
    const child = children[i] as Element
    if (child.nodeName === 'w:p') {
      if (paragraphCount > 0) {
        content += '<br>'
      }
      content += parseParagraphContent(child, styles, true, imageMap, footnoteMap, imageCtx)
      paragraphCount++
    }
  }
  
  const tag = isHeader ? 'th' : 'td'
  
  // 首行样式（使用 th 标签并应用样式）
  if (isHeader) {
    if (headerRowStyle?.bold !== false) {
      cellStyle += ' font-weight: bold;'
    }
    if (!hasCellBackground && headerRowStyle?.backgroundColor) {
      cellStyle += ` background-color: ${headerRowStyle.backgroundColor};`
      if (headerRowStyle.textColor) {
        cellStyle += ` color: ${headerRowStyle.textColor};`
      } else if (isColorDark(headerRowStyle.backgroundColor)) {
        cellStyle += ' color: white;'
      }
    } else if (headerRowStyle?.textColor) {
      cellStyle += ` color: ${headerRowStyle.textColor};`
    }
  }
  
  // 首列样式
  if (isFirstColumn && !isHeader) {
    cellStyle += ' font-weight: bold;'
  }
  
  return { 
    html: `<${tag} style="${cellStyle}"${colspan}>${content || '&nbsp;'}</${tag}>`,
    colSpan 
  }
}

// 判断颜色是否为深色
function isColorDark(color: string): boolean {
  // 移除 # 号
  const hex = color.replace('#', '')
  if (hex.length !== 6) return false
  
  const r = parseInt(hex.slice(0, 2), 16)
  const g = parseInt(hex.slice(2, 4), 16)
  const b = parseInt(hex.slice(4, 6), 16)
  
  // 计算亮度
  const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
  return luminance < 0.5
}

// 解析段落内容（不含外层标签，用于表格单元格）
function parseParagraphContent(
  para: Element,
  styles: Record<string, any>,
  inTable: boolean = false,
  imageMap: ImageMap = {},
  footnoteMap?: FootnoteMap,
  imageCtx?: ImageParseContext
): string {
  const pPr = para.getElementsByTagName('w:pPr')[0]
  let alignment = ''
  
  if (pPr) {
    const jc = pPr.getElementsByTagName('w:jc')[0]
    if (jc) {
      const val = jc.getAttribute('w:val')
      if (val === 'center') alignment = 'text-align: center;'
      else if (val === 'right') alignment = 'text-align: right;'
    }
  }
  
  let content = ''
  const childNodes = para.childNodes
  
  for (let i = 0; i < childNodes.length; i++) {
    const child = childNodes[i] as Element
    if (child.nodeName === 'w:r') {
      content += parseRun(child, imageMap, footnoteMap, imageCtx)
    } else if (child.nodeName === 'w:hyperlink') {
      const linkRuns = child.getElementsByTagName('w:r')
      let linkContent = ''
      for (let j = 0; j < linkRuns.length; j++) {
        linkContent += parseRun(linkRuns[j], imageMap, footnoteMap, imageCtx)
      }
      content += linkContent
    } else if (child.nodeName === 'w:drawing' || child.localName === 'drawing') {
      content += parseDrawing(child, imageMap, imageCtx)
    } else if (child.nodeName === 'mc:AlternateContent' || child.localName === 'AlternateContent') {
      const choice = findElementByLocalName(child, 'Choice') || findElementByLocalName(child, 'Fallback')
      if (choice) {
        const drawing = findElementByLocalName(choice, 'drawing')
        if (drawing) {
          content += parseDrawing(drawing, imageMap, imageCtx)
        }
      }
    }
  }
  
  if (alignment && content) {
    return `<span style="${alignment}">${content}</span>`
  }
  
  return content
}

// 解析样式定义
function parseStyles(stylesXml: string, themeColors?: Record<string, string>): Record<string, any> {
  const parser = new DOMParser()
  const doc = parser.parseFromString(stylesXml, 'application/xml')
  const styles: Record<string, any> = {}

  const styleElements = doc.getElementsByTagName('w:style')
  for (let i = 0; i < styleElements.length; i++) {
    const style = styleElements[i]
    const styleId = style.getAttribute('w:styleId')
    if (styleId) {
      styles[styleId] = parseStyleElement(style, themeColors)
    }
  }

  const mergeStyle = (base: any, override: any): any => {
    if (!base) return { ...override }
    const result: any = { ...base }
    for (const [key, value] of Object.entries(override || {})) {
      if (key === 'basedOn') continue
      if (
        value &&
        typeof value === 'object' &&
        !Array.isArray(value) &&
        typeof result[key] === 'object' &&
        result[key] !== null
      ) {
        result[key] = mergeStyle(result[key], value)
      } else {
        result[key] = value
      }
    }
    return result
  }

  const resolved: Record<string, any> = {}
  const resolving = new Set<string>()
  const resolveStyle = (styleId: string): any => {
    if (resolved[styleId]) return resolved[styleId]
    const style = styles[styleId]
    if (!style) return {}
    if (resolving.has(styleId)) return style
    resolving.add(styleId)
    const baseId = style.basedOn
    const baseStyle = baseId ? resolveStyle(baseId) : {}
    const merged = mergeStyle(baseStyle, style)
    delete merged.basedOn
    resolved[styleId] = merged
    resolving.delete(styleId)
    return merged
  }

  for (const styleId of Object.keys(styles)) {
    resolveStyle(styleId)
  }

  return resolved
}

function parseStyleElement(style: Element, themeColors?: Record<string, string>): any {
  const result: any = {}
  
  const basedOn = style.getElementsByTagName('w:basedOn')[0]
  if (basedOn) {
    const baseId = basedOn.getAttribute('w:val')
    if (baseId) result.basedOn = baseId
  }
  
  const pPr = style.getElementsByTagName('w:pPr')[0]
  if (pPr) {
    const jc = pPr.getElementsByTagName('w:jc')[0]
    if (jc) {
      const val = jc.getAttribute('w:val')
      if (val === 'center') result.alignment = 'center'
      else if (val === 'right') result.alignment = 'right'
      else if (val === 'both') result.alignment = 'justify'
    }
    
    // 段落背景色
    const pShd = pPr.getElementsByTagName('w:shd')[0]
    if (pShd) {
      const fill = pShd.getAttribute('w:fill')
      if (fill && fill !== 'auto' && fill.toUpperCase() !== 'FFFFFF') {
        result.backgroundColor = `#${fill}`
      }
    }
    
    // 段落间距
    const spacing = pPr.getElementsByTagName('w:spacing')[0]
    if (spacing) {
      const before = spacing.getAttribute('w:before')
      const after = spacing.getAttribute('w:after')
      const line = spacing.getAttribute('w:line')
      if (before) result.marginTop = `${parseInt(before) / 20}pt`
      if (after) result.marginBottom = `${parseInt(after) / 20}pt`
      if (line) {
        const lineRule = spacing.getAttribute('w:lineRule')
        if (lineRule === 'auto') {
          result.lineHeight = (parseInt(line) / 240).toFixed(2)
        } else {
          result.lineHeight = `${parseInt(line) / 20}pt`
        }
      }
    }
  }

  const rPr = style.getElementsByTagName('w:rPr')[0]
  if (rPr) {
    const sz = rPr.getElementsByTagName('w:sz')[0]
    if (sz) {
      const val = sz.getAttribute('w:val')
      if (val) result.fontSize = halfPointsToPt(parseInt(val))
    }

    const rFonts = rPr.getElementsByTagName('w:rFonts')[0]
    if (rFonts) {
      const fontName = resolveFontNameFromRFonts(rFonts)
      if (fontName) {
        result.fontName = fontName  // 存储原始字体名
        result.fontFamily = getSafeFontFamily(fontName)  // CSS 字体栈
      }
    }
    
    // 解析颜色（支持主题颜色）
    const color = rPr.getElementsByTagName('w:color')[0]
    if (color) {
      const val = color.getAttribute('w:val')
      const themeColor = color.getAttribute('w:themeColor')
      
      if (themeColor && themeColors && themeColors[themeColor]) {
        // 使用主题颜色
        result.color = `#${themeColors[themeColor]}`
      } else if (val && val !== 'auto') {
        result.color = `#${val}`
      }
    }
    
    // 解析粗体
    const bold = rPr.getElementsByTagName('w:b')[0]
    if (bold) {
      const val = bold.getAttribute('w:val')
      // w:b 没有值或值不为 '0'/'false' 表示粗体
      if (val !== '0' && val !== 'false') {
        result.bold = true
      }
    }
    
    // 解析斜体
    const italic = rPr.getElementsByTagName('w:i')[0]
    if (italic) {
      const val = italic.getAttribute('w:val')
      if (val !== '0' && val !== 'false') {
        result.italic = true
      }
    }
    
    // 解析下划线
    const underline = rPr.getElementsByTagName('w:u')[0]
    if (underline) {
      const val = underline.getAttribute('w:val')
      if (val && val !== 'none') {
        result.underline = true
      }
    }
    
    // 解析背景色/高亮
    const shd = rPr.getElementsByTagName('w:shd')[0]
    if (shd) {
      const fill = shd.getAttribute('w:fill')
      if (fill && fill !== 'auto' && fill.toUpperCase() !== 'FFFFFF') {
        result.highlight = `#${fill}`
      }
    }
    
    // 解析高亮颜色
    const highlight = rPr.getElementsByTagName('w:highlight')[0]
    if (highlight) {
      const val = highlight.getAttribute('w:val')
      if (val) {
        // Word 高亮颜色名称映射
        const highlightColors: Record<string, string> = {
          'yellow': '#FFFF00',
          'green': '#00FF00',
          'cyan': '#00FFFF',
          'magenta': '#FF00FF',
          'blue': '#0000FF',
          'red': '#FF0000',
          'darkBlue': '#000080',
          'darkCyan': '#008080',
          'darkGreen': '#008000',
          'darkMagenta': '#800080',
          'darkRed': '#800000',
          'darkYellow': '#808000',
          'darkGray': '#808080',
          'lightGray': '#C0C0C0',
          'black': '#000000',
        }
        if (highlightColors[val]) {
          result.highlight = highlightColors[val]
        }
      }
    }
  }

  // 表格样式（仅当 style type=table）
  const styleType = style.getAttribute('w:type')
  if (styleType === 'table') {
    const tblPr = style.getElementsByTagName('w:tblPr')[0]
    if (tblPr) {
      const tblBorders = tblPr.getElementsByTagName('w:tblBorders')[0]
      const tblCellMar = tblPr.getElementsByTagName('w:tblCellMar')[0]
      if (tblBorders) result.tblBorders = parseTableBorders(tblBorders)
      if (tblCellMar) result.tblCellMargin = parseCellMarginFromPr(tblCellMar)
    }

    const tblStylePrEls = style.getElementsByTagName('w:tblStylePr')
    if (tblStylePrEls.length) {
      result.tblStylePr = {}
      for (let i = 0; i < tblStylePrEls.length; i++) {
        const tblStylePr = tblStylePrEls[i]
        const type = tblStylePr.getAttribute('w:type') || 'default'
        const stylePr: TableStylePr = {}
        const tcPr = tblStylePr.getElementsByTagName('w:tcPr')[0]
        if (tcPr) {
          const shd = tcPr.getElementsByTagName('w:shd')[0]
          if (shd) {
            const fill = shd.getAttribute('w:fill')
            const themeFill = shd.getAttribute('w:themeFill')
            if (themeFill) {
              const resolved = resolveThemeColor(themeFill)
              if (resolved) stylePr.backgroundColor = `#${resolved}`
            } else if (fill && fill !== 'auto') {
              stylePr.backgroundColor = `#${fill}`
            }
          }
        }
        const styleRpr = tblStylePr.getElementsByTagName('w:rPr')[0]
        if (styleRpr) {
          const color = styleRpr.getElementsByTagName('w:color')[0]
          if (color) {
            const val = color.getAttribute('w:val')
            const themeColor = color.getAttribute('w:themeColor')
            if (themeColor && themeColors && themeColors[themeColor]) {
              stylePr.textColor = `#${themeColors[themeColor]}`
            } else if (val && val !== 'auto') {
              stylePr.textColor = `#${val}`
            }
          }
          const bold = styleRpr.getElementsByTagName('w:b')[0]
          if (bold) {
            const val = bold.getAttribute('w:val')
            if (val !== '0' && val !== 'false') {
              stylePr.bold = true
            }
          }
        }
        result.tblStylePr[type] = stylePr
      }
    }
  }

  return result
}

// 解析 Track Changes（w:ins / w:del）为 inline HTML
function wrapInlineWithComments(html: string, commentIds: string[]): string {
  if (!html) return ''
  if (!commentIds || commentIds.length === 0) return html
  const uniq = Array.from(new Set(commentIds.filter(Boolean)))
  if (uniq.length === 0) return html
  return `<span class="docx-comment" data-comment-ids="${escapeHtml(uniq.join(','))}">${html}</span>`
}

function parseTrackedChangeInline(
  changeEl: Element,
  styles: Record<string, any>,
  imageMap: ImageMap = {},
  footnoteMap?: FootnoteMap,
  imageCtx?: ImageParseContext,
  inheritedCommentIds: string[] = []
): string {
  const trackType = changeEl.nodeName === 'w:ins' ? 'insert' : 'delete'
  const trackId = changeEl.getAttribute('w:id') || ''
  const trackAuthor = changeEl.getAttribute('w:author') || ''
  const trackDate = changeEl.getAttribute('w:date') || ''

  let inner = ''
  const activeCommentIds: string[] = [...(inheritedCommentIds || [])]
  const tcChildren = changeEl.childNodes
  for (let j = 0; j < tcChildren.length; j++) {
    const tc = tcChildren[j] as Element
    if (tc.nodeName === 'w:commentRangeStart') {
      const id = tc.getAttribute('w:id')
      if (id) activeCommentIds.push(id)
    } else if (tc.nodeName === 'w:commentRangeEnd') {
      const id = tc.getAttribute('w:id')
      if (id) {
        const idx = activeCommentIds.lastIndexOf(id)
        if (idx >= 0) activeCommentIds.splice(idx, 1)
      }
    } else if (tc.nodeName === 'w:commentReference') {
      // ignore visual reference marker; range is handled by start/end
    } else if (tc.nodeName === 'w:r') {
      inner += wrapInlineWithComments(parseRun(tc, imageMap, footnoteMap, imageCtx), activeCommentIds)
    } else if (tc.nodeName === 'w:hyperlink') {
      const linkRuns = tc.getElementsByTagName('w:r')
      let linkContent = ''
      for (let k = 0; k < linkRuns.length; k++) {
        linkContent += parseRun(linkRuns[k], imageMap, footnoteMap, imageCtx)
      }
      const rId = tc.getAttribute('r:id')
      if (rId) {
        inner += wrapInlineWithComments(`<a href="#${rId}">${linkContent}</a>`, activeCommentIds)
      } else {
        inner += wrapInlineWithComments(linkContent, activeCommentIds)
      }
    } else if (tc.nodeName === 'w:fldSimple' || tc.nodeName === 'w:smartTag' || tc.nodeName === 'w:sdt') {
      const innerRuns = tc.getElementsByTagName('w:r')
      let tmp = ''
      for (let k = 0; k < innerRuns.length; k++) {
        tmp += parseRun(innerRuns[k], imageMap, footnoteMap, imageCtx)
      }
      inner += wrapInlineWithComments(tmp, activeCommentIds)
    } else if (tc.nodeName === 'w:drawing' || tc.localName === 'drawing') {
      inner += wrapInlineWithComments(parseDrawing(tc, imageMap, imageCtx), activeCommentIds)
    } else if (tc.nodeName === 'mc:AlternateContent' || tc.localName === 'AlternateContent') {
      const choice = findElementByLocalName(tc, 'Choice') || findElementByLocalName(tc, 'Fallback')
      if (choice) {
        const drawing = findElementByLocalName(choice, 'drawing')
        if (drawing) {
          inner += wrapInlineWithComments(parseDrawing(drawing, imageMap, imageCtx), activeCommentIds)
        }
      }
    } else if (tc.nodeName === 'w:ins' || tc.nodeName === 'w:del') {
      inner += wrapInlineWithComments(parseTrackedChangeInline(tc, styles, imageMap, footnoteMap, imageCtx, activeCommentIds), activeCommentIds)
    }
  }

  if (!inner) return ''

  const attrs: string[] = [
    `data-track-type="${escapeHtml(trackType)}"`,
  ]
  if (trackId) attrs.push(`data-track-id="${escapeHtml(trackId)}"`)
  if (trackAuthor) attrs.push(`data-track-author="${escapeHtml(trackAuthor)}"`)
  if (trackDate) attrs.push(`data-track-date="${escapeHtml(trackDate)}"`)

  return `<span class="docx-track" ${attrs.join(' ')}>${inner}</span>`
}

// 解析段落
function parseParagraph(
  para: Element,
  styles: Record<string, any>,
  imageMap: ImageMap = {},
  footnoteMap?: FootnoteMap,
  imageCtx?: ImageParseContext,
  numberingState?: NumberingState
): string {
  const pPr = para.getElementsByTagName('w:pPr')[0]
  let paraStyle: ParagraphStyle = {}
  let tag = 'p'
  const styleProps: string[] = []
  let hasIndent = false
  let paraFontSize = ''
  let paraFontFamily = ''
  let paraFontName = ''  // 原始字体名（用于 UI 显示）
  let paraColor = ''
  let isListItem = false
  let listLevel = 0
  let listMarker = ''  // 列表项标记文本
  let pendingNumbering: { numId: string; level: number } | null = null
  let isTocParagraph = false
  let skipIndent = false
  let hasPageBreakBefore = false
  let tocTabPosPt: number | null = null
  let tocTabLeader: string | null = null
  let styleData: any | null = null
  let hasMarginTop = false
  let hasMarginBottom = false
  let hasLineHeight = false
  const dataAttrs: string[] = []

  if (pPr) {
    // 段前分页
    const pageBreakBefore = pPr.getElementsByTagName('w:pageBreakBefore')[0]
    if (pageBreakBefore) {
      hasPageBreakBefore = true
    }
    // 检查是否为列表项
    const numPr = pPr.getElementsByTagName('w:numPr')[0]
    if (numPr) {
      isListItem = true
      let useManualMarker = false
      const ilvl = numPr.getElementsByTagName('w:ilvl')[0]
      if (ilvl) {
        listLevel = parseInt(ilvl.getAttribute('w:val') || '0')
      }
      const numIdEl = numPr.getElementsByTagName('w:numId')[0]
      if (numIdEl) {
        const numId = numIdEl.getAttribute('w:val')
        const numberingInfo = numId ? getNumberingInfo(numId, listLevel) : null
        // 所有编号段落都使用手动 marker（按 numId 隔离计数器，避免 CSS list-item 共享计数）
        if (numId && numberingState) {
          useManualMarker = true
          pendingNumbering = { numId, level: listLevel }
        }
        
        // 从 numbering 映射获取缩进
        if (numberingInfo) {
          const leftPt = numberingInfo.indLeftTwips ? numberingInfo.indLeftTwips / 20 : undefined
          const hangingPt = numberingInfo.indHangingTwips ? numberingInfo.indHangingTwips / 20 : undefined
          const firstLinePt = numberingInfo.indFirstLineTwips ? numberingInfo.indFirstLineTwips / 20 : undefined
          if (leftPt != null) {
            styleProps.push(`padding-left: ${leftPt}pt`)
            skipIndent = true
          }
          if (hangingPt != null && hangingPt > 0) {
            styleProps.push(`text-indent: -${hangingPt}pt`)
            skipIndent = true
          } else if (firstLinePt != null && firstLinePt > 0) {
            styleProps.push(`text-indent: ${firstLinePt}pt`)
            skipIndent = true
          }
        }

        if (!skipIndent) {
          // 回退为层级缩进
          styleProps.push(`padding-left: ${(listLevel + 1) * 1.5}em`)
        }
      }
    }
    
    const pStyle = pPr.getElementsByTagName('w:pStyle')[0]
    if (pStyle) {
      const styleId = pStyle.getAttribute('w:val')
      if (styleId) {
        const lowerStyleId = styleId.toLowerCase()
        if (lowerStyleId.includes('toc') || styleId.includes('目录')) {
          isTocParagraph = true
        }
        // 检查是否为列表样式（无 numPr 时仅标记，不添加 CSS list-item 避免干扰）
        if (styleId.includes('ListParagraph') || styleId.includes('列表段落')) {
          if (!isListItem) {
            isListItem = true
          }
        }
        if (styleId.includes('Heading') || styleId.includes('标题')) {
          const level = styleId.match(/\d/)?.[0] || '1'
          tag = `h${level}`
          paraStyle.heading = parseInt(level)
        }
        if (styles[styleId]) {
          styleData = styles[styleId]
          Object.assign(paraStyle, styleData)
          // 从样式中继承字体信息（如果段落自身没有定义）
          if (styleData.fontFamily && !paraFontFamily) {
            paraFontFamily = styleData.fontFamily
            // 同时继承原始字体名
            if (styleData.fontName && !paraFontName) {
              paraFontName = styleData.fontName
            }
          }
          if (styleData.fontSize && !paraFontSize) {
            paraFontSize = styleData.fontSize
          }
          if (styleData.color && !paraColor) {
            paraColor = styleData.color
          }
        }
      }
    }

    // 制表位（TOC 点引导线）
    const tabs = pPr.getElementsByTagName('w:tabs')[0]
    if (tabs) {
      const tabEls = tabs.getElementsByTagName('w:tab')
      for (let i = 0; i < tabEls.length; i++) {
        const tab = tabEls[i]
        const leader = tab.getAttribute('w:leader')
        const pos = tab.getAttribute('w:pos')
        if (leader && leader !== 'none') {
          tocTabLeader = leader
        }
        if (pos) {
          const twips = parseInt(pos)
          if (Number.isFinite(twips)) {
            tocTabPosPt = twips / 20
          }
        }
      }
    }

    // 段落级别的文字样式 (rPr in pPr)
    const pRpr = pPr.getElementsByTagName('w:rPr')[0]
    if (pRpr) {
      const sz = pRpr.getElementsByTagName('w:sz')[0]
      if (sz) {
        const val = sz.getAttribute('w:val')
        if (val) {
          paraFontSize = halfPointsToPt(parseInt(val))
        }
      }
      const rFonts = pRpr.getElementsByTagName('w:rFonts')[0]
      if (rFonts) {
        const fontName = resolveFontNameFromRFonts(rFonts)
        if (fontName) {
          paraFontName = fontName  // 存储原始字体名
          paraFontFamily = getSafeFontFamily(fontName) || ''
        }
      }
      const color = pRpr.getElementsByTagName('w:color')[0]
      if (color) {
        const themeColor = color.getAttribute('w:themeColor')
        const val = color.getAttribute('w:val')
        
        // 优先使用主题颜色
        if (themeColor) {
          const resolvedColor = resolveThemeColor(themeColor)
          if (resolvedColor) {
            const colorHex = `#${resolvedColor}`
            if (!shouldIgnoreColorInDarkMode(colorHex)) {
              paraColor = colorHex
            }
          }
        } else if (val && val !== 'auto') {
          const colorHex = `#${val}`
          // 在深色模式下，忽略黑色或接近黑色的颜色，让 CSS 变量自动处理
          if (!shouldIgnoreColorInDarkMode(colorHex)) {
            paraColor = colorHex
          }
        }
      }
    }

    // 对齐方式
    const jc = pPr.getElementsByTagName('w:jc')[0]
    if (jc) {
      const val = jc.getAttribute('w:val')
      if (val === 'center') {
        paraStyle.alignment = 'center'
        styleProps.push('text-align: center')
      } else if (val === 'right') {
        paraStyle.alignment = 'right'
        styleProps.push('text-align: right')
      } else if (val === 'both' || val === 'distribute') {
        paraStyle.alignment = 'justify'
        styleProps.push('text-align: justify')
      }
    }

    // 缩进
    const ind = pPr.getElementsByTagName('w:ind')[0]
    if (ind && !skipIndent) {
      const firstLineChars = ind.getAttribute('w:firstLineChars')
      const firstLine = ind.getAttribute('w:firstLine')
      const left = ind.getAttribute('w:left') || ind.getAttribute('w:start')
      const leftChars = ind.getAttribute('w:leftChars') || ind.getAttribute('w:startChars')
      const hanging = ind.getAttribute('w:hanging')
      
      // 首行缩进
      if (firstLineChars) {
        const chars = parseInt(firstLineChars) / 100
        if (chars > 0) {
          styleProps.push(`text-indent: ${chars}em`)
          hasIndent = true
        }
      } else if (firstLine) {
        const twips = parseInt(firstLine)
        if (twips > 0) {
          const em = twips / 240
          styleProps.push(`text-indent: ${em.toFixed(2)}em`)
          hasIndent = true
        }
      }
      
      // 悬挂缩进（负缩进）
      if (hanging) {
        const twips = parseInt(hanging)
        if (twips > 0) {
          const em = twips / 240
          styleProps.push(`text-indent: -${em.toFixed(2)}em`)
          hasIndent = true
        }
      }
      
      // 左缩进
      if (leftChars) {
        const chars = parseInt(leftChars) / 100
        if (chars > 0) {
          styleProps.push(`padding-left: ${chars}em`)
        }
      } else if (left) {
        const twips = parseInt(left)
        if (twips > 0) {
          const em = twips / 240
          styleProps.push(`padding-left: ${em.toFixed(2)}em`)
        }
      }
    }

    // 段落间距和行距
    const spacing = pPr.getElementsByTagName('w:spacing')[0]
    if (spacing) {
      const before = spacing.getAttribute('w:before')
      const beforeLines = spacing.getAttribute('w:beforeLines')
      const after = spacing.getAttribute('w:after')
      const afterLines = spacing.getAttribute('w:afterLines')
      const line = spacing.getAttribute('w:line')
      const lineRule = spacing.getAttribute('w:lineRule')
      
      // 段前间距
      if (beforeLines) {
        // 以行为单位（100 = 1行）
        const lines = parseInt(beforeLines) / 100
        if (lines > 0) {
          styleProps.push(`margin-top: ${lines}em`)
          hasMarginTop = true
        }
      } else if (before) {
        const twips = parseInt(before)
        if (twips > 0) {
          styleProps.push(`margin-top: ${(twips / 20).toFixed(1)}pt`)
          hasMarginTop = true
        }
      }
      
      // 段后间距
      if (afterLines) {
        // 以行为单位（100 = 1行）
        const lines = parseInt(afterLines) / 100
        if (lines > 0) {
          styleProps.push(`margin-bottom: ${lines}em`)
          hasMarginBottom = true
        }
      } else if (after) {
        const twips = parseInt(after)
        if (twips > 0) {
          styleProps.push(`margin-bottom: ${(twips / 20).toFixed(1)}pt`)
          hasMarginBottom = true
        }
      }
      
      // 行距处理
      if (line) {
        const lineVal = parseInt(line)
        if (lineRule === 'exact') {
          // 固定值行距
          styleProps.push(`line-height: ${(lineVal / 20).toFixed(1)}pt`)
          hasLineHeight = true
        } else if (lineRule === 'atLeast') {
          // 最小值行距
          styleProps.push(`line-height: ${(lineVal / 20).toFixed(1)}pt`)
          hasLineHeight = true
        } else if (!lineRule || lineRule === 'auto') {
          // 倍数行距：240 = 单倍行距
          const multiplier = lineVal / 240
          styleProps.push(`line-height: ${multiplier.toFixed(2)}`)
          hasLineHeight = true
        }
      }
    } else {
      // 若段落未显式设置 spacing，则继承样式里的间距/行距
      if (styleData) {
        if (!hasMarginTop && styleData.marginTop) {
          styleProps.push(`margin-top: ${styleData.marginTop}`)
          hasMarginTop = true
        }
        if (!hasMarginBottom && styleData.marginBottom) {
          styleProps.push(`margin-bottom: ${styleData.marginBottom}`)
          hasMarginBottom = true
        }
        if (!hasLineHeight && styleData.lineHeight) {
          styleProps.push(`line-height: ${styleData.lineHeight}`)
          hasLineHeight = true
        }
      }

      // 仍未定义时，回退为 0
      if (!hasMarginTop) styleProps.push('margin-top: 0')
      if (!hasMarginBottom) styleProps.push('margin-bottom: 0')
    }
    
    // 检查 contextualSpacing（相同样式段落间不加间距）
    const contextualSpacing = pPr.getElementsByTagName('w:contextualSpacing')[0]
    if (contextualSpacing) {
      const val = contextualSpacing.getAttribute('w:val')
      // 如果存在且不是 false/0，则标记这个段落
      if (val !== 'false' && val !== '0') {
        styleProps.push('--contextual-spacing: 1')
      }
    }
  }

  let styleAttr = ''

  let content = ''
  const activeCommentIds: string[] = []
  const childNodes = para.childNodes
  
  for (let i = 0; i < childNodes.length; i++) {
    const child = childNodes[i] as Element
    if (child.nodeName === 'w:commentRangeStart') {
      const id = child.getAttribute('w:id')
      if (id) activeCommentIds.push(id)
    } else if (child.nodeName === 'w:commentRangeEnd') {
      const id = child.getAttribute('w:id')
      if (id) {
        const idx = activeCommentIds.lastIndexOf(id)
        if (idx >= 0) activeCommentIds.splice(idx, 1)
      }
    } else if (child.nodeName === 'w:commentReference') {
      // ignore
    } else if (child.nodeName === 'w:r') {
      content += wrapInlineWithComments(parseRun(child, imageMap, footnoteMap, imageCtx), activeCommentIds)
    } else if (child.nodeName === 'w:ins' || child.nodeName === 'w:del') {
      // 修订：插入/删除（Track Changes）
      content += wrapInlineWithComments(
        parseTrackedChangeInline(child, styles, imageMap, footnoteMap, imageCtx, activeCommentIds),
        activeCommentIds
      )
    } else if (child.nodeName === 'w:hyperlink') {
      const linkRuns = child.getElementsByTagName('w:r')
      let linkContent = ''
      for (let j = 0; j < linkRuns.length; j++) {
        linkContent += parseRun(linkRuns[j], imageMap, footnoteMap, imageCtx)
      }
      const rId = child.getAttribute('r:id')
      const wAnchor = child.getAttribute('w:anchor')
      if (rId) {
        content += wrapInlineWithComments(`<a href="#${rId}">${linkContent}</a>`, activeCommentIds)
      } else if (wAnchor) {
        content += wrapInlineWithComments(`<a href="#${wAnchor}">${linkContent}</a>`, activeCommentIds)
      } else {
        content += wrapInlineWithComments(linkContent, activeCommentIds)
      }
    } else if (child.nodeName === 'w:fldSimple' || child.nodeName === 'w:smartTag' || child.nodeName === 'w:sdt') {
      const innerRuns = child.getElementsByTagName('w:r')
      let tmp = ''
      for (let j = 0; j < innerRuns.length; j++) {
        tmp += parseRun(innerRuns[j], imageMap, footnoteMap, imageCtx)
      }
      content += wrapInlineWithComments(tmp, activeCommentIds)
    } else if (child.nodeName === 'w:drawing' || child.localName === 'drawing') {
      // 图片可能直接在段落中
      content += wrapInlineWithComments(parseDrawing(child, imageMap, imageCtx), activeCommentIds)
    } else if (child.nodeName === 'mc:AlternateContent' || child.localName === 'AlternateContent') {
      // 处理备用内容（通常包含图片）
      const choice = findElementByLocalName(child, 'Choice') || findElementByLocalName(child, 'Fallback')
      if (choice) {
        const drawing = findElementByLocalName(choice, 'drawing')
        if (drawing) {
          content += wrapInlineWithComments(parseDrawing(drawing, imageMap, imageCtx), activeCommentIds)
        }
      }
    }
  }

  const plainText = content
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ')
    .trim()
  const hasImage = /<img\b/i.test(content)
  const hasVisibleContent = plainText.length > 0 || hasImage
  const isImageOnly = hasImage && plainText.length === 0

  if (pendingNumbering && hasVisibleContent) {
    const marker = buildListMarkerText(pendingNumbering.numId, pendingNumbering.level, numberingState || {})
    if (marker) {
      listMarker = marker
    }
  }

  if (isImageOnly) {
    styleProps.push('line-height: 0')
    styleProps.push('font-size: 0')
  }

  if (!hasVisibleContent && isListItem) {
    styleProps.push('list-style-type: none')
    styleProps.push('display: block')
  }

  if (listMarker) {
    content = `<span class="docx-list-marker">${escapeHtml(listMarker)}</span>${content}`
  }

  const buildTocMarkup = (raw: string) => {
    const segments = raw.split(DOCX_TAB_TOKEN)
    const left = segments[0] || ''
    const right = segments.slice(1).join('')
    // 生成足够多的可见点字符，CSS 用 overflow:hidden 裁剪到实际宽度
    const dots = '.'.repeat(80)
    return `<span class="docx-toc-left">${left}</span><span class="docx-toc-leader">${dots}</span><span class="docx-toc-right">${right}</span>`
  }

  const isTocWithTabs = isTocParagraph && content.includes(DOCX_TAB_TOKEN)

  if (isTocWithTabs) {
    const anchorMatch = content.match(/^<a\s+[^>]*>[\s\S]*<\/a>$/)
    if (anchorMatch) {
      const attrsMatch = content.match(/^<a\s+([^>]+)>/)
      const attrs = attrsMatch?.[1] || ''
      const inner = content.replace(/^<a[^>]*>/, '').replace(/<\/a>$/, '')
      content = `<a ${attrs}>${buildTocMarkup(inner)}</a>`
    } else {
      content = buildTocMarkup(content)
    }
    styleProps.push('display: flex')
    styleProps.push('align-items: center')
    styleProps.push('gap: 0.5em')
    styleProps.push('white-space: nowrap')
  } else if (content.includes(DOCX_TAB_TOKEN)) {
    const segments = content.split(DOCX_TAB_TOKEN)
    content = segments
      .map((seg, idx) => (idx === 0 ? seg : `<span class="docx-tab"></span>${seg}`))
      .join('')
  }

  if (tocTabPosPt != null) {
    styleProps.push(`--docx-tab-pos: ${tocTabPosPt}pt`)
  }
  if (tocTabLeader) {
    dataAttrs.push(`data-docx-tab-leader="${tocTabLeader}"`)
  }

  styleAttr = styleProps.length > 0 ? ` style="${styleProps.join('; ')}"` : ''

  const classNames: string[] = []
  if (isTocWithTabs) classNames.push('docx-toc')
  if (isImageOnly) classNames.push('docx-image-only')
  const classAttr = classNames.length > 0 ? ` class="${classNames.join(' ')}"` : ''
  const dataAttr = dataAttrs.length > 0 ? ` ${dataAttrs.join(' ')}` : ''

  const wrapWithPageBreak = (html: string) =>
    hasPageBreakBefore ? `<hr class="page-break" />${html}` : html

  if (!content.trim()) {
    return wrapWithPageBreak(`<${tag}${classAttr}${dataAttr}${styleAttr}><br></${tag}>`)
  }

  // 关键：段落级字体/字号/颜色不要挂在 block(<p>/<h*>) 上，
  // 因为 Tiptap 导入 HTML 时通常不会保留 block 的 style（会导致字体信息丢失）。
  // 下沉到 inline(<span>) 后，FontFamily/TextStyle 才能稳定解析并渲染。
  const paraSpanStyle: string[] = []
  if (paraFontFamily) paraSpanStyle.push(`font-family: ${paraFontFamily}`)
  if (paraFontSize) paraSpanStyle.push(`font-size: ${paraFontSize}`)
  // 在深色模式下，忽略黑色或接近黑色的颜色，让 CSS 变量自动处理
  if (paraColor && !shouldIgnoreColorInDarkMode(paraColor)) {
    paraSpanStyle.push(`color: ${paraColor}`)
  }

  if (isImageOnly) {
    paraSpanStyle.length = 0
  }

  if (paraSpanStyle.length > 0) {
    // 添加 data-font-name 属性存储原始字体名，方便 UI 检测
    const fontNameAttr = paraFontName ? ` data-font-name="${paraFontName}"` : ''
    const innerClassAttr = isTocWithTabs ? ' class="docx-toc-inner"' : ''
    // 转义 style 属性值中的双引号，避免与 HTML 属性引号冲突
    const styleValue = paraSpanStyle.join('; ').replace(/"/g, '&quot;')
    return wrapWithPageBreak(
      `<${tag}${classAttr}${dataAttr}${styleAttr}><span data-para-font="1"${fontNameAttr}${innerClassAttr} style="${styleValue}">${content}</span></${tag}>`
    )
  }

  return wrapWithPageBreak(`<${tag}${classAttr}${dataAttr}${styleAttr}>${content}</${tag}>`)
}

// 通用的元素查找函数，处理命名空间问题
function findElementByLocalName(parent: Element, localName: string): Element | null {
  // 尝试多种方式查找元素
  // 1. 直接用带命名空间前缀的名称
  const withPrefix = parent.getElementsByTagName(`*`);
  for (let i = 0; i < withPrefix.length; i++) {
    const el = withPrefix[i]
    if (el.localName === localName || el.nodeName.endsWith(':' + localName)) {
      return el
    }
  }
  return null
}

// 查找所有匹配的元素
function findAllElementsByLocalName(parent: Element, localName: string): Element[] {
  const results: Element[] = []
  const all = parent.getElementsByTagName(`*`);
  for (let i = 0; i < all.length; i++) {
    const el = all[i]
    if (el.localName === localName || el.nodeName.endsWith(':' + localName)) {
      results.push(el)
    }
  }
  return results
}

// 获取带命名空间的属性值
function getAttributeNS(el: Element, localName: string): string | null {
  // 尝试多种命名空间前缀
  const prefixes = ['r', 'relationships', '']
  for (const prefix of prefixes) {
    const attrName = prefix ? `${prefix}:${localName}` : localName
    const value = el.getAttribute(attrName)
    if (value) return value
  }
  // 遍历所有属性查找
  for (let i = 0; i < el.attributes.length; i++) {
    const attr = el.attributes[i]
    if (attr.localName === localName || attr.name.endsWith(':' + localName)) {
      return attr.value
    }
  }
  return null
}

// 解析图片元素 (w:drawing)
function parseDrawing(drawing: Element, imageMap: ImageMap, imageCtx?: ImageParseContext): string {
  // 查找图片引用 (a:blip 或 blip)
  const blip = findElementByLocalName(drawing, 'blip')
  if (!blip) {
    return ''
  }
  
  // 获取图片的 rId (可能是 r:embed 或 r:link)
  const rId = getAttributeNS(blip, 'embed') || getAttributeNS(blip, 'link')
  if (!rId) {
    return ''
  }
  
  // 获取图片尺寸
  let width = 0
  let height = 0
  
  // 尝试从 wp:extent 获取尺寸（内联图片）
  const extent = findElementByLocalName(drawing, 'extent')
  if (extent) {
    const cx = extent.getAttribute('cx')
    const cy = extent.getAttribute('cy')
    if (cx) width = emuToPixels(parseInt(cx))
    if (cy) height = emuToPixels(parseInt(cy))
  }
  
  // 如果没有从 extent 获取到，尝试从 a:ext 获取（图片本身的尺寸）
  if (!width || !height) {
    const aExt = findElementByLocalName(drawing, 'ext')
    if (aExt) {
      const cx = aExt.getAttribute('cx')
      const cy = aExt.getAttribute('cy')
      if (cx && !width) width = emuToPixels(parseInt(cx))
      if (cy && !height) height = emuToPixels(parseInt(cy))
    }
  }
  
  // 构建样式
  const styleProps: string[] = ['max-width: 100%']
  if (width > 0) styleProps.push(`width: ${width}px`)
  if (height > 0) styleProps.push(`height: ${height}px`)
  if (height <= 0) styleProps.push('height: auto')
  
  // 检查是否是浮动图片 (wp:anchor)
  const anchor = findElementByLocalName(drawing, 'anchor')
  const inline = findElementByLocalName(drawing, 'inline')
  if (anchor) {
    // 浮动图片：后续由 WordEditor 根据锚点坐标做绝对定位
    styleProps.push('position: absolute', 'left: 0', 'top: 0', 'visibility: hidden')
  } else {
    // 内联图片：应用 Word distT/distB 间距（EMU）
    if (inline) {
      const distT = inline.getAttribute('distT')
      const distB = inline.getAttribute('distB')
      if (distT) {
        const px = emuToPixels(parseInt(distT))
        if (px) styleProps.push(`margin-top: ${px}px`)
      }
      if (distB) {
        const px = emuToPixels(parseInt(distB))
        if (px) styleProps.push(`margin-bottom: ${px}px`)
      }
    }
    // 内联图片：保留在文字流里，不要强制 block 居中
    styleProps.push('display: inline-block', 'vertical-align: baseline')
  }
  
  const styleAttr = styleProps.join('; ')

  // alt/title（如果有的话）
  const docPr = findElementByLocalName(drawing, 'docPr')
  const altRaw =
    docPr?.getAttribute('descr') ||
    docPr?.getAttribute('title') ||
    docPr?.getAttribute('name') ||
    ''
  const alt = altRaw ? escapeHtml(altRaw) : '文档图片'

  const floating = !!anchor
  const target = imageCtx?.relsMap?.[rId]

  // 收集元信息（避免重复堆积：同一个 rid 多次出现时也记录，便于定位）
  if (imageCtx?.images) {
    imageCtx.images.push({
      rId,
      target,
      widthPx: width > 0 ? width : undefined,
      heightPx: height > 0 ? height : undefined,
      alt: altRaw || undefined,
      floating,
    })
  }

  // Agent/轻量模式：不内联图片二进制
  const shouldEmbed = imageCtx?.embedImages !== false
  const imageData = shouldEmbed ? imageMap[rId] : undefined
  if (!imageData && shouldEmbed) {
    return `<span style="color: #999; font-style: italic;">[图片: ${escapeHtml(rId)}]</span>`
  }

  const src = imageData || 'about:blank'
  const dataAttrs: string[] = [`data-rid="${escapeHtml(rId)}"`]
  if (target) dataAttrs.push(`data-target="${escapeHtml(target)}"`)
  if (width > 0) dataAttrs.push(`data-w="${width}"`)
  if (height > 0) dataAttrs.push(`data-h="${height}"`)
  if (floating) dataAttrs.push(`data-floating="1"`)

  // Anchor positioning metadata (for more accurate placement)
  if (anchor) {
    try {
      const posH = findElementByLocalName(anchor, 'positionH')
      const posV = findElementByLocalName(anchor, 'positionV')
      const relH = posH?.getAttribute('relativeFrom') || ''
      const relV = posV?.getAttribute('relativeFrom') || ''
      if (relH) dataAttrs.push(`data-rel-h="${escapeHtml(relH)}"`)
      if (relV) dataAttrs.push(`data-rel-v="${escapeHtml(relV)}"`)

      const xOff = findElementByLocalName(posH || anchor, 'posOffset')?.textContent
      const yOff = findElementByLocalName(posV || anchor, 'posOffset')?.textContent
      if (xOff) dataAttrs.push(`data-x-emu="${escapeHtml(xOff.trim())}"`)
      if (yOff) dataAttrs.push(`data-y-emu="${escapeHtml(yOff.trim())}"`)

      const distL = anchor.getAttribute('distL')
      const distR = anchor.getAttribute('distR')
      const distT = anchor.getAttribute('distT')
      const distB = anchor.getAttribute('distB')
      if (distL) dataAttrs.push(`data-dist-l="${escapeHtml(distL)}"`)
      if (distR) dataAttrs.push(`data-dist-r="${escapeHtml(distR)}"`)
      if (distT) dataAttrs.push(`data-dist-t="${escapeHtml(distT)}"`)
      if (distB) dataAttrs.push(`data-dist-b="${escapeHtml(distB)}"`)

      const wrapType =
        findElementByLocalName(anchor, 'wrapNone') ? 'none' :
        findElementByLocalName(anchor, 'wrapSquare') ? 'square' :
        findElementByLocalName(anchor, 'wrapTight') ? 'tight' :
        findElementByLocalName(anchor, 'wrapTopAndBottom') ? 'topAndBottom' :
        findElementByLocalName(anchor, 'wrapThrough') ? 'through' :
        ''
      if (wrapType) dataAttrs.push(`data-wrap="${wrapType}"`)
    } catch {
      // ignore metadata parse failures
    }
  }

  return `<img src="${src}" alt="${alt}" style="${styleAttr}" ${dataAttrs.join(' ')} />`
}

// 解析 run（文本块）
function parseRun(run: Element, imageMap: ImageMap = {}, footnoteMap?: FootnoteMap, imageCtx?: ImageParseContext): string {
  const rPr = run.getElementsByTagName('w:rPr')[0]
  let style: RunStyle = {}

  if (rPr) {
    if (rPr.getElementsByTagName('w:b').length > 0) {
      style.bold = true
    }

    if (rPr.getElementsByTagName('w:i').length > 0) {
      style.italic = true
    }

    const u = rPr.getElementsByTagName('w:u')[0]
    if (u) {
      const uVal = u.getAttribute('w:val')
      if (uVal && uVal !== 'none') {
        style.underline = true
        if (uVal === 'dotted' || uVal === 'dottedHeavy') {
          style.underlineStyle = 'dotted'
        } else if (uVal === 'dash' || uVal === 'dashLong') {
          style.underlineStyle = 'dashed'
        } else if (uVal === 'double') {
          style.underlineStyle = 'double'
        } else if (uVal === 'wave' || uVal === 'wavyDouble') {
          style.underlineStyle = 'wavy'
        } else {
          style.underlineStyle = 'solid'
        }
      }
    }

    if (rPr.getElementsByTagName('w:strike').length > 0) {
      style.strike = true
    }

    const sz = rPr.getElementsByTagName('w:sz')[0]
    if (sz) {
      const val = sz.getAttribute('w:val')
      if (val) {
        style.fontSize = halfPointsToPt(parseInt(val))
      }
    }

    const rFonts = rPr.getElementsByTagName('w:rFonts')[0]
    if (rFonts) {
      const fontName = resolveFontNameFromRFonts(rFonts)
      // 调试日志：检查字体解析结果
      if (fontName) {
        style.fontName = fontName  // 存储原始字体名（用于 UI 显示）
        style.fontFamily = getSafeFontFamily(fontName)  // CSS font-family（包含回退）
      }
    }

    const color = rPr.getElementsByTagName('w:color')[0]
    if (color) {
      const themeColor = color.getAttribute('w:themeColor')
      const val = color.getAttribute('w:val')
      
      // 优先使用主题颜色
      if (themeColor) {
        const resolvedColor = resolveThemeColor(themeColor)
        if (resolvedColor) {
          const colorHex = `#${resolvedColor}`
          if (!shouldIgnoreColorInDarkMode(colorHex)) {
            style.color = colorHex
          }
        }
      } else if (val && val !== 'auto') {
        const colorHex = `#${val}`
        // 在深色模式下，忽略黑色或接近黑色的颜色，让 CSS 变量自动处理
        if (!shouldIgnoreColorInDarkMode(colorHex)) {
          style.color = colorHex
        }
      }
    }

    // 解析背景色/高亮
    const shd = rPr.getElementsByTagName('w:shd')[0]
    if (shd) {
      const fill = shd.getAttribute('w:fill')
      if (fill && fill !== 'auto' && fill.toUpperCase() !== 'FFFFFF') {
        style.highlight = `#${fill}`
      }
    }

    const highlight = rPr.getElementsByTagName('w:highlight')[0]
    if (highlight) {
      const val = highlight.getAttribute('w:val')
      if (val) {
        const highlightColors: Record<string, string> = {
          yellow: '#FFFF00',
          green: '#00FF00',
          cyan: '#00FFFF',
          magenta: '#FF00FF',
          blue: '#0000FF',
          red: '#FF0000',
          darkBlue: '#000080',
          darkCyan: '#008080',
          darkGreen: '#008000',
          darkMagenta: '#800080',
          darkRed: '#800000',
          darkYellow: '#808000',
          darkGray: '#808080',
          lightGray: '#C0C0C0',
          black: '#000000',
        }
        if (highlightColors[val]) {
          style.highlight = highlightColors[val]
        }
      }
    }
  }

  let text = ''
  let imageHtml = ''
  let hasSpecialChars = false
  const children = run.childNodes
  for (let i = 0; i < children.length; i++) {
    const child = children[i] as Element
    if (child.nodeName === 'w:t') {
      text += child.textContent || ''
    } else if (child.nodeName === 'w:delText') {
      // Track Changes 删除内容
      text += child.textContent || ''
    } else if (child.nodeName === 'w:tab') {
      text += DOCX_TAB_TOKEN
      hasSpecialChars = true
    } else if (child.nodeName === 'w:br' || child.nodeName === 'w:cr') {
      // 检查是否是分页符
      const brType = child.getAttribute('w:type')
      if (brType === 'page') {
        text += '[[PAGE_BREAK]]' // 分页符占位符
      } else if (brType === 'column') {
        text += '[[COLUMN_BREAK]]' // 分栏符占位符
      } else {
        text += '[[BR]]' // 普通换行占位符
      }
      hasSpecialChars = true
    } else if (child.nodeName === 'w:sym') {
      const char = child.getAttribute('w:char')
      if (char) {
        text += String.fromCharCode(parseInt(char, 16))
      }
    } else if (child.nodeName === 'w:ptab') {
      text += DOCX_TAB_TOKEN
      hasSpecialChars = true
    } else if (child.nodeName === 'w:drawing' || child.localName === 'drawing') {
      // 处理图片
      imageHtml += parseDrawing(child, imageMap, imageCtx)
    } else if (child.nodeName === 'w:pict' || child.localName === 'pict') {
      // 旧版图片格式 (VML)，尝试提取
      const imageData = findElementByLocalName(child, 'imagedata')
      if (imageData) {
        const rId = getAttributeNS(imageData, 'id') || getAttributeNS(imageData, 'embed')
        if (rId && imageMap[rId]) {
          imageHtml += `<img src="${imageMap[rId]}" alt="文档图片" style="max-width: 100%; height: auto; display: block; margin: 10px auto;" />`
        }
      }
    } else if (child.nodeName === 'mc:AlternateContent' || child.localName === 'AlternateContent') {
      // 处理备用内容（通常包含图片）
      const choice = findElementByLocalName(child, 'Choice') || findElementByLocalName(child, 'Fallback')
      if (choice) {
        const drawing = findElementByLocalName(choice, 'drawing')
        if (drawing) {
          imageHtml += parseDrawing(drawing, imageMap, imageCtx)
        }
      }
    } else if (child.nodeName === 'w:footnoteReference') {
      // 脚注引用
      const footnoteId = child.getAttribute('w:id')
      if (footnoteId) {
        text += `[[FOOTNOTE_REF:${footnoteId}]]`
        hasSpecialChars = true
      }
    } else if (child.nodeName === 'w:endnoteReference') {
      // 尾注引用
      const endnoteId = child.getAttribute('w:id')
      if (endnoteId) {
        text += `[[ENDNOTE_REF:${endnoteId}]]`
        hasSpecialChars = true
      }
    }
  }

  // 如果有图片，优先返回图片
  if (imageHtml) {
    return imageHtml
  }

  if (!text) return ''

  // Tab-only run：返回裸 token，不做 span 包裹，
  // 避免 buildTocMarkup 的 split 在 <span> 内部断裂 HTML
  if (text === DOCX_TAB_TOKEN) {
    return DOCX_TAB_TOKEN
  }

  // 先转义 HTML，然后处理占位符
  let html = escapeHtml(text)
  
  // 将占位符替换回 HTML 标签
  html = html.replace(/\[\[BR\]\]/g, '<br>')
  html = html.replace(/\[\[PAGE_BREAK\]\]/g, '<hr class="page-break" />')
  html = html.replace(/\[\[COLUMN_BREAK\]\]/g, '<span class="column-break"></span>')
  
  // 脚注/尾注引用标记
  html = html.replace(/\[\[FOOTNOTE_REF:(\d+)\]\]/g, '<sup class="footnote-ref"><a href="#footnote-$1">[$1]</a></sup>')
  html = html.replace(/\[\[ENDNOTE_REF:(\d+)\]\]/g, '<sup class="endnote-ref"><a href="#endnote-$1">[$1]</a></sup>')

  const styleProps: string[] = []
  if (style.fontSize) {
    styleProps.push(`font-size: ${style.fontSize}`)
  }
  if (style.fontFamily) {
    styleProps.push(`font-family: ${style.fontFamily}`)
  }
  if (style.highlight) {
    styleProps.push(`background-color: ${style.highlight}`)
  }
  if (style.color) {
    styleProps.push(`color: ${style.color}`)
  }

  if (style.underline && style.underlineStyle && style.underlineStyle !== 'solid') {
    styleProps.push('text-decoration: underline')
    styleProps.push(`text-decoration-style: ${style.underlineStyle}`)
  }

  if (styleProps.length > 0) {
    // 添加 data-font-name 属性存储原始字体名，方便 UI 检测
    const fontNameAttr = style.fontName ? ` data-font-name="${style.fontName}"` : ''
    // 转义 style 属性值中的双引号，避免与 HTML 属性引号冲突
    const styleValue = styleProps.join('; ').replace(/"/g, '&quot;')
    html = `<span${fontNameAttr} style="${styleValue}">${html}</span>`
  }

  if (style.bold) html = `<strong>${html}</strong>`
  if (style.italic) html = `<em>${html}</em>`
  if (style.underline && (!style.underlineStyle || style.underlineStyle === 'solid')) {
    html = `<u>${html}</u>`
  }
  if (style.strike) html = `<s>${html}</s>`

  return html
}

// HTML 转义
function escapeHtml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;')
}
