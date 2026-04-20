/**
 * Electron 环境下从 Fonts/ 目录加载字体到浏览器 document.fonts
 * 使用 FontFace API 注册，让 Tiptap/预览能命中 DOCX 的 font-family
 */

declare global {
  interface Window {
    electronAPI?: {
      isElectron?: boolean
      fontsList?: () => Promise<{ success: boolean; fonts: FontFileInfo[]; error?: string }>
      fontsRead?: (fileName: string) => Promise<{ success: boolean; base64?: string; size?: number; error?: string }>
    }
    __workspaceFontsLoaded?: Set<string>
    __workspaceFontsFailed?: Set<string>
    __workspaceFontsLoading?: Map<string, Promise<boolean>>
  }
}

export interface FontFileInfo {
  name: string
  ext: string
  size: number
}

/**
 * 字体别名映射：DOCX 常见字体名 → 对应的字体文件（按优先级）
 * 键为 CSS font-family 可能出现的名称（中文/英文/别名），值为 Fonts/ 下可用文件列表
 */
const FONT_ALIAS_MAP: Record<string, string[]> = {
  // 宋体系列
  '宋体': ['simsun.ttc', 'STSONG.TTF', 'NotoSerifSC-VF.ttf'],
  'SimSun': ['simsun.ttc', 'STSONG.TTF', 'NotoSerifSC-VF.ttf'],
  '新宋体': ['simsun.ttc', 'NotoSerifSC-VF.ttf'],
  'NSimSun': ['simsun.ttc', 'NotoSerifSC-VF.ttf'],
  '华文宋体': ['STSONG.TTF', 'simsun.ttc', 'NotoSerifSC-VF.ttf'],
  'STSong': ['STSONG.TTF', 'simsun.ttc', 'NotoSerifSC-VF.ttf'],
  '宋体-简': ['simsun.ttc', 'STSONG.TTF', 'NotoSerifSC-VF.ttf'],
  'Songti SC': ['simsun.ttc', 'STSONG.TTF', 'NotoSerifSC-VF.ttf'],

  // 黑体系列
  '黑体': ['STXIHEI.TTF', 'msyh.ttc', 'NotoSansSC-VF.ttf'],
  'SimHei': ['STXIHEI.TTF', 'msyh.ttc', 'NotoSansSC-VF.ttf'],
  'STXIHEI': ['STXIHEI.TTF', 'NotoSansSC-VF.ttf'],
  '华文黑体': ['STXIHEI.TTF', 'msyh.ttc', 'NotoSansSC-VF.ttf'],
  'STHeiti': ['STXIHEI.TTF', 'msyh.ttc', 'NotoSansSC-VF.ttf'],
  '华文细黑': ['STXIHEI.TTF', 'NotoSansSC-VF.ttf'],
  'STXihei': ['STXIHEI.TTF', 'NotoSansSC-VF.ttf'],

  // 微软雅黑
  '微软雅黑': ['msyh.ttc', 'NotoSansSC-VF.ttf'],
  'Microsoft YaHei': ['msyh.ttc', 'NotoSansSC-VF.ttf'],
  '微软雅黑 Light': ['msyhl.ttc', 'msyh.ttc', 'NotoSansSC-VF.ttf'],
  'Microsoft YaHei Light': ['msyhl.ttc', 'msyh.ttc', 'NotoSansSC-VF.ttf'],

  // 等线
  '等线': ['Deng.ttf', 'Dengl.ttf', 'NotoSansSC-VF.ttf'],
  'DengXian': ['Deng.ttf', 'Dengl.ttf', 'NotoSansSC-VF.ttf'],
  '等线 Light': ['Dengl.ttf', 'Deng.ttf', 'NotoSansSC-VF.ttf'],
  'DengXian Light': ['Dengl.ttf', 'Deng.ttf', 'NotoSansSC-VF.ttf'],

  // 仿宋
  '仿宋': ['STFANGSO.TTF', 'NotoSerifSC-VF.ttf'],
  'FangSong': ['STFANGSO.TTF', 'NotoSerifSC-VF.ttf'],
  '仿宋_GB2312': ['STFANGSO.TTF', 'NotoSerifSC-VF.ttf'],
  '华文仿宋': ['STFANGSO.TTF', 'NotoSerifSC-VF.ttf'],
  'STFangsong': ['STFANGSO.TTF', 'NotoSerifSC-VF.ttf'],
  'STFANGSO': ['STFANGSO.TTF', 'NotoSerifSC-VF.ttf'],

  // 楷体
  '楷体': ['STKAITI.TTF', 'NotoSerifSC-VF.ttf'],
  'KaiTi': ['STKAITI.TTF', 'NotoSerifSC-VF.ttf'],
  '楷体_GB2312': ['STKAITI.TTF', 'NotoSerifSC-VF.ttf'],
  '华文楷体': ['STKAITI.TTF', 'NotoSerifSC-VF.ttf'],
  'STKaiti': ['STKAITI.TTF', 'NotoSerifSC-VF.ttf'],
  'STKAITI': ['STKAITI.TTF', 'NotoSerifSC-VF.ttf'],

  // 华文系列其他
  '华文中宋': ['STZHONGS.TTF', 'STSONG.TTF', 'NotoSerifSC-VF.ttf'],
  'STZhongsong': ['STZHONGS.TTF', 'STSONG.TTF', 'NotoSerifSC-VF.ttf'],
  '华文新魏': ['STXINWEI.TTF', 'NotoSerifSC-VF.ttf'],
  'STXinwei': ['STXINWEI.TTF', 'NotoSerifSC-VF.ttf'],
  '华文行楷': ['STXINGKA.TTF', 'NotoSerifSC-VF.ttf'],
  'STXingkai': ['STXINGKA.TTF', 'NotoSerifSC-VF.ttf'],
  '华文琥珀': ['STHUPO.TTF', 'NotoSansSC-VF.ttf'],
  'STHupo': ['STHUPO.TTF', 'NotoSansSC-VF.ttf'],
  '华文隶书': ['STLITI.TTF', 'NotoSerifSC-VF.ttf'],
  'STLiti': ['STLITI.TTF', 'NotoSerifSC-VF.ttf'],
  '华文彩云': ['STCAIYUN.TTF', 'NotoSansSC-VF.ttf'],
  'STCaiyun': ['STCAIYUN.TTF', 'NotoSansSC-VF.ttf'],

  // 方正系列
  '方正小标宋简体': ['simsun.ttc', 'simsunb.ttf', 'NotoSerifSC-VF.ttf'],
  '方正小标宋_GBK': ['simsun.ttc', 'simsunb.ttf', 'NotoSerifSC-VF.ttf'],
  '方正仿宋简体': ['STFANGSO.TTF', 'simsun.ttc', 'NotoSerifSC-VF.ttf'],
  '方正仿宋_GBK': ['STFANGSO.TTF', 'simsun.ttc', 'NotoSerifSC-VF.ttf'],
  '方正姚体': ['FZYTK.TTF', 'NotoSerifSC-VF.ttf'],
  'FZYaoti': ['FZYTK.TTF', 'NotoSerifSC-VF.ttf'],
  '方正舒体': ['FZSTK.TTF', 'NotoSerifSC-VF.ttf'],
  'FZShuTi': ['FZSTK.TTF', 'NotoSerifSC-VF.ttf'],

  // 隶书/幼圆
  '隶书': ['SIMLI.TTF', 'NotoSerifSC-VF.ttf'],
  'LiSu': ['SIMLI.TTF', 'NotoSerifSC-VF.ttf'],
  '幼圆': ['SIMYOU.TTF', 'NotoSansSC-VF.ttf'],
  'YouYuan': ['SIMYOU.TTF', 'NotoSansSC-VF.ttf'],

  // 英文字体
  'Times New Roman': ['times.ttf', 'NotoSerifSC-VF.ttf'],
  'Arial': ['segoeui.ttf', 'NotoSansSC-VF.ttf'],
  'Calibri': ['calibrili.ttf', 'NotoSansSC-VF.ttf'],
  'Segoe UI': ['segoeui.ttf', 'NotoSansSC-VF.ttf'],
  'Cascadia Code': ['CascadiaCode.ttf', 'CascadiaMono.ttf'],
  'Cascadia Mono': ['CascadiaMono.ttf', 'CascadiaCode.ttf'],
  'Ubuntu Mono': ['UbuntuMono[wght].ttf'],

  // 日文字体
  'Yu Gothic': ['YuGothM.ttc', 'YuGothR.ttc', 'NotoSansSC-VF.ttf'],
  'Yu Gothic Medium': ['YuGothM.ttc', 'NotoSansSC-VF.ttf'],
  'Yu Gothic Bold': ['YuGothB.ttc', 'NotoSansSC-VF.ttf'],
  'Yu Gothic Light': ['YuGothL.ttc', 'NotoSansSC-VF.ttf'],

  // Century / Garamond / Book Antiqua 等西文字体
  'Century': ['CENTURY.TTF', 'times.ttf'],
  'Garamond': ['GARA.TTF', 'times.ttf'],
  'Book Antiqua': ['BKANT.TTF', 'times.ttf'],
  'Bookman Old Style': ['BOOKOS.TTF', 'times.ttf'],
  'Century Gothic': ['GOTHIC.TTF', 'segoeui.ttf'],
  'Mistral': ['MISTRAL.TTF'],
  'Papyrus': ['PAPYRUS.TTF'],
  'Pristina': ['PRISTINA.TTF'],
}

// 已加载/失败的字体缓存（全局）
function getLoadedSet(): Set<string> {
  if (!window.__workspaceFontsLoaded) {
    window.__workspaceFontsLoaded = new Set()
  }
  return window.__workspaceFontsLoaded
}

function getFailedSet(): Set<string> {
  if (!window.__workspaceFontsFailed) {
    window.__workspaceFontsFailed = new Set()
  }
  return window.__workspaceFontsFailed
}

function getLoadingMap(): Map<string, Promise<boolean>> {
  if (!window.__workspaceFontsLoading) {
    window.__workspaceFontsLoading = new Map()
  }
  return window.__workspaceFontsLoading
}

// 可用字体文件缓存
let availableFontFiles: FontFileInfo[] | null = null

/**
 * 获取 Fonts/ 目录下可用字体列表（缓存）
 */
async function getAvailableFonts(): Promise<FontFileInfo[]> {
  if (availableFontFiles !== null) return availableFontFiles

  const api = window.electronAPI
  if (!api?.fontsList) {
    availableFontFiles = []
    return []
  }

  try {
    const result = await api.fontsList()
    if (result.success && result.fonts) {
      availableFontFiles = result.fonts
      console.log(`[WorkspaceFonts] 可用字体文件: ${result.fonts.length} 个`)
    } else {
      availableFontFiles = []
    }
  } catch (e) {
    console.warn('[WorkspaceFonts] 获取字体列表失败:', e)
    availableFontFiles = []
  }

  return availableFontFiles
}

/**
 * Base64 转 ArrayBuffer
 */
function base64ToArrayBuffer(base64: string): ArrayBuffer {
  const binary = atob(base64)
  const len = binary.length
  const bytes = new Uint8Array(len)
  for (let i = 0; i < len; i++) {
    bytes[i] = binary.charCodeAt(i)
  }
  return bytes.buffer
}

/**
 * 注册单个字体文件到 document.fonts
 * @param familyName CSS font-family 名称
 * @param fileName Fonts/ 下的文件名
 * @param weight 字重（可选）
 * @param style 字体样式（可选）
 */
async function registerFontFile(
  familyName: string,
  fileName: string,
  weight: string = 'normal',
  style: string = 'normal'
): Promise<boolean> {
  const cacheKey = `${familyName}|${weight}|${style}`
  const loaded = getLoadedSet()
  const failed = getFailedSet()
  const loading = getLoadingMap()

  // 已加载或已失败，跳过
  if (loaded.has(cacheKey)) return true
  if (failed.has(cacheKey)) return false

  // 正在加载中，等待
  if (loading.has(cacheKey)) {
    return loading.get(cacheKey)!
  }

  const api = window.electronAPI
  if (!api?.fontsRead) {
    failed.add(cacheKey)
    return false
  }

  const promise = (async () => {
    try {
      const result = await api.fontsRead!(fileName)
      if (!result.success || !result.base64) {
        console.warn(`[WorkspaceFonts] 读取字体失败: ${fileName}`, result.error)
        failed.add(cacheKey)
        return false
      }

      const buffer = base64ToArrayBuffer(result.base64)
      const fontFace = new FontFace(familyName, buffer, {
        weight,
        style,
        display: 'swap',
      })

      await fontFace.load()
      document.fonts.add(fontFace)
      loaded.add(cacheKey)
      console.log(`[WorkspaceFonts] 注册成功: "${familyName}" (${fileName})`)
      return true
    } catch (e) {
      // .ttc 在某些 Chromium 版本可能不支持
      console.warn(`[WorkspaceFonts] 注册失败: "${familyName}" (${fileName})`, e)
      failed.add(cacheKey)
      return false
    } finally {
      loading.delete(cacheKey)
    }
  })()

  loading.set(cacheKey, promise)
  return promise
}

/**
 * 根据 DOCX 字体名加载对应字体（按别名映射 + 回退）
 * @param fontName DOCX 中出现的字体名
 */
export async function loadFontByName(fontName: string): Promise<boolean> {
  const trimmed = (fontName || '').trim().replace(/^["']|["']$/g, '')
  if (!trimmed) return false

  const loaded = getLoadedSet()
  const failed = getFailedSet()

  // 已经为这个 family 加载过
  const cacheKey = `${trimmed}|normal|normal`
  if (loaded.has(cacheKey)) return true
  if (failed.has(cacheKey)) return false

  // 查找别名映射
  const candidates = FONT_ALIAS_MAP[trimmed]
  if (!candidates || candidates.length === 0) {
    // 没有映射，尝试按文件名猜测（比如 "Deng" → "Deng.ttf"）
    const availableFonts = await getAvailableFonts()
    const guessFile = availableFonts.find(
      (f) => f.name.toLowerCase().startsWith(trimmed.toLowerCase())
    )
    if (guessFile) {
      return registerFontFile(trimmed, guessFile.name)
    }
    // 无法加载，标记失败
    failed.add(cacheKey)
    return false
  }

  // 按优先级尝试候选文件
  const availableFonts = await getAvailableFonts()
  // 建立小写名 -> 实际文件名的映射
  const nameToActual = new Map<string, string>()
  for (const f of availableFonts) {
    nameToActual.set(f.name.toLowerCase(), f.name)
  }

  for (const candidate of candidates) {
    const actualName = nameToActual.get(candidate.toLowerCase())
    if (actualName) {
      const success = await registerFontFile(trimmed, actualName)
      if (success) return true
    }
  }

  // 所有候选都失败
  failed.add(cacheKey)
  return false
}

/**
 * 批量加载多个字体名（去重 + 并发控制）
 */
export async function loadFontsByNames(fontNames: string[]): Promise<void> {
  const unique = [...new Set(fontNames.map((n) => n.trim()).filter(Boolean))]
  if (unique.length === 0) {
    return
  }

  // 简单并发：最多 4 个同时加载
  const concurrency = 4
  for (let i = 0; i < unique.length; i += concurrency) {
    const batch = unique.slice(i, i + concurrency)
    await Promise.all(batch.map((name) => loadFontByName(name)))
  }
}

/**
 * 从 HTML 字符串中提取所有 font-family 值
 */
export function extractFontFamiliesFromHtml(html: string): string[] {
  const families: string[] = []
  
  // 方法1：匹配 data-font-name 属性（优先，更准确）
  const dataFontRegex = /data-font-name=["']([^"']+)["']/gi
  let match: RegExpExecArray | null
  while ((match = dataFontRegex.exec(html)) !== null) {
    const fontName = match[1].trim()
    if (fontName && !fontName.startsWith('var(')) {
      families.push(fontName)
    }
  }
  
  // 方法2：匹配 font-family 样式（兼容没有 data-font-name 的情况）
  // 改进的正则：匹配 font-family: 后面直到 ; 或 " 的内容
  // 需要处理引号嵌套的情况：font-family: "微软雅黑", "Arial";
  const styleRegex = /font-family:\s*([^;]+?)(?:;|"(?:\s*>|\s*\/))/gi
  while ((match = styleRegex.exec(html)) !== null) {
    const value = match[1].trim()
    // 分割逗号分隔的字体栈
    const parts = value.split(',')
    for (const part of parts) {
      // 去掉首尾的引号（单引号、双引号、中文引号）
      const cleaned = part.trim().replace(/^["'""'']+|["'""'']+$/g, '')
      if (cleaned && !cleaned.startsWith('var(') && !families.includes(cleaned)) {
        families.push(cleaned)
      }
    }
  }
  
  // 去重
  return [...new Set(families)]
}

/**
 * 预加载常用中文字体（后台非阻塞）
 */
export async function preloadCommonChineseFonts(): Promise<void> {
  const commonFonts = [
    '宋体',
    '黑体',
    '微软雅黑',
    '等线',
    '仿宋',
    '楷体',
    'Times New Roman',
    'Arial',
  ]
  await loadFontsByNames(commonFonts)
}

/**
 * 文件名 → 应注册的别名列表（反向映射）
 * 这样后台加载时能用正确的 font-family 名称注册
 */
const FILE_TO_ALIASES: Record<string, string[]> = {
  // 宋体
  'simsun.ttc': ['宋体', 'SimSun', '新宋体', 'NSimSun'],
  'simsunb.ttf': ['宋体', 'SimSun'],
  'STSONG.TTF': ['华文宋体', 'STSong', '宋体'],
  
  // 黑体
  'STXIHEI.TTF': ['黑体', 'SimHei', '华文黑体', 'STHeiti', '华文细黑', 'STXihei', 'STXIHEI'],
  
  // 微软雅黑
  'msyh.ttc': ['微软雅黑', 'Microsoft YaHei'],
  'msyhbd.ttc': ['微软雅黑', 'Microsoft YaHei'],
  'msyhl.ttc': ['微软雅黑 Light', 'Microsoft YaHei Light'],
  
  // 等线
  'Deng.ttf': ['等线', 'DengXian'],
  'Dengl.ttf': ['等线 Light', 'DengXian Light', '等线', 'DengXian'],
  'Dengb.ttf': ['等线', 'DengXian'],
  
  // 仿宋
  'STFANGSO.TTF': ['仿宋', 'FangSong', '仿宋_GB2312', '华文仿宋', 'STFangsong', 'STFANGSO'],
  
  // 楷体
  'STKAITI.TTF': ['楷体', 'KaiTi', '楷体_GB2312', '华文楷体', 'STKaiti', 'STKAITI'],
  
  // 华文系列
  'STZHONGS.TTF': ['华文中宋', 'STZhongsong'],
  'STXINWEI.TTF': ['华文新魏', 'STXinwei'],
  'STXINGKA.TTF': ['华文行楷', 'STXingkai'],
  'STHUPO.TTF': ['华文琥珀', 'STHupo'],
  'STLITI.TTF': ['华文隶书', 'STLiti'],
  'STCAIYUN.TTF': ['华文彩云', 'STCaiyun'],
  
  // 方正/其他中文
  'FZYTK.TTF': ['方正姚体', 'FZYaoti'],
  'FZSTK.TTF': ['方正舒体', 'FZShuTi'],
  'SIMLI.TTF': ['隶书', 'LiSu'],
  'SIMYOU.TTF': ['幼圆', 'YouYuan'],
  
  // Noto（通用兜底）
  'NotoSansSC-VF.ttf': ['Noto Sans SC', 'NotoSansSC'],
  'NotoSerifSC-VF.ttf': ['Noto Serif SC', 'NotoSerifSC'],
  
  // 英文字体
  'times.ttf': ['Times New Roman'],
  'timesbd.ttf': ['Times New Roman'],
  'timesi.ttf': ['Times New Roman'],
  'segoeui.ttf': ['Segoe UI', 'Arial'],
  'segoeuib.ttf': ['Segoe UI'],
  'calibrili.ttf': ['Calibri'],
  'CascadiaCode.ttf': ['Cascadia Code'],
  'CascadiaMono.ttf': ['Cascadia Mono'],
  
  // 日文
  'YuGothM.ttc': ['Yu Gothic', 'Yu Gothic Medium'],
  'YuGothR.ttc': ['Yu Gothic', 'Yu Gothic Regular'],
  'YuGothB.ttc': ['Yu Gothic Bold'],
  'YuGothL.ttc': ['Yu Gothic Light'],
}

/**
 * 加载 Fonts/ 目录下所有字体（后台低优先级）
 * 用正确的字体别名注册，而不是文件名
 */
export async function loadAllWorkspaceFonts(): Promise<void> {
  const fonts = await getAvailableFonts()
  if (fonts.length === 0) return

  console.log(`[WorkspaceFonts] 开始后台加载全部 ${fonts.length} 个字体...`)

  // 按文件大小排序，小文件优先
  const sorted = [...fonts].sort((a, b) => a.size - b.size)

  // 低并发后台加载
  const concurrency = 2
  for (let i = 0; i < sorted.length; i += concurrency) {
    const batch = sorted.slice(i, i + concurrency)
    await Promise.all(
      batch.map(async (f) => {
        // 查找这个文件对应的别名
        const aliases = FILE_TO_ALIASES[f.name] || FILE_TO_ALIASES[f.name.toLowerCase()]
        
        if (aliases && aliases.length > 0) {
          // 用所有别名注册这个字体文件
          for (const alias of aliases) {
            try {
              await registerFontFile(alias, f.name)
            } catch {
              // 忽略单个别名的失败
            }
          }
        } else {
          // 没有别名映射，用文件名注册（兜底）
          const baseName = f.name.replace(/\.[^.]+$/, '')
          try {
            await registerFontFile(baseName, f.name)
          } catch {
            // 忽略
          }
        }
      })
    )
    // 让出主线程
    await new Promise((r) => setTimeout(r, 10))
  }

  console.log('[WorkspaceFonts] 后台加载完成')
}

/**
 * 主入口：在 Electron 环境下初始化字体加载
 * - 先预加载常用字体
 * - 然后后台加载全部
 */
export async function initWorkspaceFonts(): Promise<void> {
  if (typeof window === 'undefined') return
  if (!window.electronAPI?.isElectron) {
    console.log('[WorkspaceFonts] 非 Electron 环境，跳过')
    return
  }

  console.log('[WorkspaceFonts] 开始初始化...')

  // 1. 先加载常用字体（较快）
  await preloadCommonChineseFonts()

  // 2. 后台加载全部（不阻塞 UI）
  requestIdleCallback
    ? requestIdleCallback(() => void loadAllWorkspaceFonts())
    : setTimeout(() => void loadAllWorkspaceFonts(), 1000)
}

