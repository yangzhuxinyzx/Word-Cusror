import type { BundledFontEntry } from './fontManifest'
import { bundledFonts } from './fontManifest'

declare global {
  interface Window {
    __onlyofficeBundledFonts?: BundledFontEntry[]
  }
}

const STYLE_ID = 'onlyoffice-bundled-fonts-style'

type ManifestShape = { fonts?: BundledFontEntry[] } | BundledFontEntry[]

function normalizeManifest(data: ManifestShape | null | undefined): BundledFontEntry[] {
  if (!data) return []
  if (Array.isArray(data)) return data
  if (Array.isArray((data as any).fonts)) return (data as any).fonts
  return []
}

function buildFontFaceCss(fonts: BundledFontEntry[]): string {
  const lines: string[] = []
  for (const f of fonts) {
    if (!f?.family || !Array.isArray(f.sources) || f.sources.length === 0) continue
    const family = String(f.family).replace(/"/g, '\\"')
    const style = f.style || 'normal'
    const weight = f.weight ?? 'normal'
    const src = f.sources
      .filter((s) => s?.url && s?.format)
      .map((s) => {
        let url = String(s.url).trim()
        // 允许用户在 manifest 里只写文件名：如 "SimSun.woff2"
        if (url && !/^(?:[a-z]+:|\/|\.\/|\.\.\/)/i.test(url)) {
          url = `./fonts/onlyoffice/${url}`
        } else if (url.startsWith('fonts/')) {
          // 避免生产环境 file:// 下的绝对路径问题
          url = `./${url}`
        }
        const fmt =
          s.format === 'truetype'
            ? 'truetype'
            : s.format === 'opentype'
              ? 'opentype'
              : s.format
        return `url("${url}") format("${fmt}")`
      })
      .join(', ')
    if (!src) continue
    lines.push(
      [
        '@font-face {',
        `  font-family: "${family}";`,
        `  font-style: ${style};`,
        `  font-weight: ${weight};`,
        `  font-display: swap;`,
        `  src: ${src};`,
        '}',
      ].join('\n')
    )
  }
  return lines.join('\n\n')
}

async function tryLoadPublicManifest(): Promise<BundledFontEntry[] | null> {
  try {
    const resp = await fetch('./fonts/onlyoffice/manifest.json', { cache: 'no-store' })
    if (!resp.ok) return null
    const data = (await resp.json()) as ManifestShape
    const fonts = normalizeManifest(data)
    return fonts
  } catch {
    return null
  }
}

/**
 * 注入 @font-face 并暴露 fonts 列表给 UI（WordEditor 字体选择器）。
 *
 * - 优先读取 `public/fonts/onlyoffice/manifest.json`
 * - 回退到 `src/fonts/fontManifest.ts` 的 `bundledFonts`
 */
export async function loadOnlyOfficeBundledFonts(): Promise<BundledFontEntry[]> {
  if (typeof window === 'undefined') return []

  const existing = document.getElementById(STYLE_ID)
  if (existing) {
    return window.__onlyofficeBundledFonts || []
  }

  const publicFonts = await tryLoadPublicManifest()
  const fonts = (publicFonts && publicFonts.length ? publicFonts : bundledFonts) || []

  window.__onlyofficeBundledFonts = fonts

  if (!fonts.length) return []

  const css = buildFontFaceCss(fonts)
  if (!css.trim()) return fonts

  const styleEl = document.createElement('style')
  styleEl.id = STYLE_ID
  styleEl.textContent = css
  document.head.appendChild(styleEl)

  // 通知 UI 刷新字体列表
  window.dispatchEvent(new CustomEvent('onlyoffice-fonts-loaded', { detail: { fonts } }))

  return fonts
}


