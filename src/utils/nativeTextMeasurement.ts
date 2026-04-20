import type {
  DocxCompatSettings,
  DocxInspectionReport,
  NativeTextMeasureEntry,
  NativeTextMeasureResult,
} from '../types'

const PT_TO_PX = 96 / 72

type NativeTextStyle = {
  fontFamily: string
  fontSize: number
  fontWeight?: 'normal' | 'bold' | number
  fontStyle?: 'normal' | 'italic'
  lineHeight?: number
  letterSpacing?: number
  scaleFactor?: number
}

type CachedStyleMetric = {
  lineHeight: number
  ascent: number
  descent: number
  baseline: number
  resolvedFontFamily?: string
}

const textMetricCache = new Map<string, NativeTextMeasureResult>()
const styleMetricCache = new Map<string, CachedStyleMetric>()

let currentCompatSettings: DocxCompatSettings = {}

const normalizeFontFamily = (fontFamily: string) =>
  (fontFamily || '')
    .split(',')
    .map((segment) => segment.trim().replace(/^['"]|['"]$/g, ''))
    .filter(Boolean)
    .join(', ')

const styleSignature = (style: NativeTextStyle) => {
  const scaleFactor = style.scaleFactor || 1
  return [
    normalizeFontFamily(style.fontFamily),
    Number(style.fontSize || 10.5).toFixed(3),
    String(style.fontWeight || 'normal'),
    String(style.fontStyle || 'normal'),
    Number(style.letterSpacing || 0).toFixed(3),
    Number(style.lineHeight || 0).toFixed(3),
    scaleFactor.toFixed(3),
  ].join('|')
}

const textSignature = (text: string, style: NativeTextStyle) =>
  `${text}|${styleSignature(style)}`

const parseStyleAttr = (styleAttr: string, tagName?: string): NativeTextStyle => {
  const result: NativeTextStyle = {
    fontFamily: 'DengXian, 等线, Microsoft YaHei, SimSun, serif',
    fontSize: tagName === 'h1' ? 22 : tagName === 'h2' ? 16 : tagName === 'h3' ? 14 : 10.5,
    fontWeight: ['h1', 'h2', 'h3', 'h4', 'h5', 'h6'].includes(String(tagName || '').toLowerCase())
      ? 'bold'
      : 'normal',
    fontStyle: 'normal',
    lineHeight: 1.15,
    letterSpacing: 0,
  }

  if (!styleAttr) return result

  const fontFamily = styleAttr.match(/font-family:\s*([^;]+)/i)?.[1]?.trim()
  if (fontFamily) result.fontFamily = normalizeFontFamily(fontFamily)

  const fontSize = styleAttr.match(/font-size:\s*(\d+(?:\.\d+)?)(px|pt|em)/i)
  if (fontSize) {
    const value = parseFloat(fontSize[1])
    const unit = fontSize[2].toLowerCase()
    if (unit === 'pt') result.fontSize = value
    else if (unit === 'px') result.fontSize = value / PT_TO_PX
    else result.fontSize = value * 12
  }

  const fontWeight = styleAttr.match(/font-weight:\s*(bold|\d{3})/i)?.[1]
  if (fontWeight) {
    result.fontWeight = /^\d+$/.test(fontWeight) ? Number(fontWeight) : 'bold'
  }

  if (/font-style:\s*italic/i.test(styleAttr)) {
    result.fontStyle = 'italic'
  }

  const lineHeight = styleAttr.match(/line-height:\s*(\d+(?:\.\d+)?)(px|pt)?/i)
  if (lineHeight) {
    const value = parseFloat(lineHeight[1])
    const unit = lineHeight[2]?.toLowerCase()
    if (unit === 'pt') {
      result.lineHeight = result.fontSize > 0 ? value / result.fontSize : 1.15
    } else if (unit === 'px') {
      result.lineHeight = result.fontSize > 0 ? (value / PT_TO_PX) / result.fontSize : 1.15
    } else {
      result.lineHeight = value
    }
  }

  const letterSpacing = styleAttr.match(/letter-spacing:\s*(-?\d+(?:\.\d+)?)(px|pt|em)/i)
  if (letterSpacing) {
    const value = parseFloat(letterSpacing[1])
    const unit = letterSpacing[2].toLowerCase()
    if (unit === 'pt') result.letterSpacing = value
    else if (unit === 'px') result.letterSpacing = value / PT_TO_PX
    else result.letterSpacing = value * (result.fontSize || 10.5)
  }

  return result
}

const collectMeasurementEntries = (html: string, scale: number): NativeTextMeasureEntry[] => {
  if (typeof DOMParser === 'undefined') return []
  const parser = new DOMParser()
  const doc = parser.parseFromString(`<div>${html}</div>`, 'text/html')
  const root = doc.body.firstElementChild as HTMLElement | null
  if (!root) return []

  const entryMap = new Map<string, NativeTextMeasureEntry>()

  const addEntry = (text: string, style: NativeTextStyle) => {
    if (!text) return
    const normalizedStyle: NativeTextStyle = {
      ...style,
      scaleFactor: scale,
    }
    const key = textSignature(text, normalizedStyle)
    if (entryMap.has(key)) return
    entryMap.set(key, {
      id: key,
      text,
      fontFamily: normalizedStyle.fontFamily,
      fontSize: normalizedStyle.fontSize * scale,
      fontWeight: normalizedStyle.fontWeight,
      fontStyle: normalizedStyle.fontStyle,
      letterSpacing: (normalizedStyle.letterSpacing || 0) * scale,
    })
  }

  const elements = Array.from(root.querySelectorAll<HTMLElement>('*'))
  if (root instanceof HTMLElement) {
    elements.unshift(root)
  }

  elements.forEach((element) => {
    const style = parseStyleAttr(element.getAttribute('style') || '', element.tagName.toLowerCase())
    addEntry('Hg心桥', style)

    const text = Array.from(element.childNodes)
      .filter((node) => node.nodeType === Node.TEXT_NODE)
      .map((node) => node.textContent || '')
      .join('')
      .replace(/\s+/g, ' ')
      .trim()

    if (!text) return

    const words = text.match(/[A-Za-z0-9._-]+/g) || []
    words.forEach((word) => addEntry(word, style))
    Array.from(text).forEach((char) => {
      if (char === '\n' || char === '\r') return
      addEntry(char, style)
    })
  })

  return Array.from(entryMap.values())
}

export function configureDocxCompatSettings(report?: DocxInspectionReport | null) {
  currentCompatSettings = report?.summary?.compat || {}
}

export function getDocxCompatSettings(): DocxCompatSettings {
  return currentCompatSettings
}

export async function prewarmNativeTextMeasurementsFromHtml(
  html: string,
  scale: number,
  report?: DocxInspectionReport | null,
): Promise<void> {
  if (typeof window === 'undefined' || !window.electronAPI?.textMeasureNative || !html) return
  configureDocxCompatSettings(report)

  const entries = collectMeasurementEntries(html, scale)
  if (!entries.length) return

  const response = await window.electronAPI.textMeasureNative({
    mode: 'measure',
    entries,
  })

  if (!response.success || !response.measurements?.length) return

  response.measurements.forEach((measurement) => {
    textMetricCache.set(measurement.id, measurement)
  })

  entries.forEach((entry) => {
    const styleKey = styleSignature({
      fontFamily: entry.fontFamily,
      fontSize: entry.fontSize,
      fontWeight: entry.fontWeight,
      fontStyle: entry.fontStyle,
      letterSpacing: entry.letterSpacing,
      lineHeight: entry.lineHeight,
      scaleFactor: scale,
    })
    if (styleMetricCache.has(styleKey)) return
    const metrics = textMetricCache.get(entry.id)
    if (!metrics) return
    styleMetricCache.set(styleKey, {
      lineHeight: metrics.lineHeight,
      ascent: metrics.ascent,
      descent: metrics.descent,
      baseline: metrics.baseline,
      resolvedFontFamily: metrics.resolvedFontFamily,
    })
  })
}

export function getNativeTextMetrics(text: string, style: NativeTextStyle): NativeTextMeasureResult | null {
  return textMetricCache.get(textSignature(text, style)) || null
}

export function getNativeStyleMetrics(style: NativeTextStyle): CachedStyleMetric | null {
  return styleMetricCache.get(styleSignature(style)) || null
}
