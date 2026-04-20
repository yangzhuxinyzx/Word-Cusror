/**
 * chartParser.ts — OOXML Chart XML → Chart.js 配置转换器
 *
 * 解析 word/charts/chartN.xml，提取图表类型、系列数据、标签、颜色，
 * 转换为 chart.js 可直接使用的配置对象。
 */

import type { ChartConfiguration, ChartDataset } from 'chart.js'

// ============== 类型定义 ==============

export type ChartKind = 'bar' | 'line' | 'pie' | 'doughnut' | 'scatter' | 'radar'

export interface ChartSeries {
  name: string
  values: number[]
  color?: string // hex like '#4472C4'
  /** 饼图/环形图：每个数据点的颜色 */
  pointColors?: string[]
}

export interface ChartConfig {
  type: ChartKind
  title?: string
  categories: string[]
  series: ChartSeries[]
  stacking?: 'stacked' | 'percentStacked'
  legendPosition?: 'top' | 'bottom' | 'left' | 'right'
  /** area chart 需要 fill */
  fill?: boolean
  /** 值轴格式 (如 '0%') */
  valFormat?: string
  widthPx: number
  heightPx: number
}

// ============== 默认调色板 ==============

/** Office 默认 accent 颜色（当 theme 不可用时的 fallback） */
const DEFAULT_ACCENT_COLORS: Record<string, string> = {
  accent1: '#4472C4',
  accent2: '#ED7D31',
  accent3: '#A5A5A5',
  accent4: '#FFC000',
  accent5: '#5B9BD5',
  accent6: '#70AD47',
}

// ============== XML 辅助函数 ==============

/** 按 localName 查找第一个后代元素 */
function findByLocal(el: Element, localName: string): Element | null {
  return el.getElementsByTagNameNS('*', localName)[0] || null
}

/** 按 localName 查找所有直接子元素 */
function childrenByLocal(el: Element, localName: string): Element[] {
  const result: Element[] = []
  for (let i = 0; i < el.children.length; i++) {
    if (el.children[i].localName === localName) result.push(el.children[i])
  }
  return result
}

/** 按 localName 查找所有后代元素 */
function allByLocal(el: Element, localName: string): Element[] {
  return Array.from(el.getElementsByTagNameNS('*', localName))
}

// ============== Excel 日期转换 ==============

/** Excel 序列号 → 日期字符串 (yyyy/M/d) */
function excelSerialToDateStr(serial: number): string {
  // Excel 日期基准：serial 1 = 1900-01-01
  // Lotus 1-2-3 bug：serial 60 = 1900-02-29（不存在），serial > 60 需要减 1
  const adjusted = serial > 60 ? serial - 1 : serial
  // 用 UTC 避免时区偏移
  const ms = Date.UTC(1900, 0, 1) + (adjusted - 1) * 86400000
  const d = new Date(ms)
  return `${d.getUTCFullYear()}/${d.getUTCMonth() + 1}/${d.getUTCDate()}`
}

// ============== 颜色解析 ==============

/** 从 spPr 或 dPt 中提取颜色 */
function extractColor(
  el: Element,
  themeColors?: Record<string, string>
): string | undefined {
  const solidFill = findByLocal(el, 'solidFill')
  if (!solidFill) return undefined

  // 直接 sRGB 颜色
  const srgb = findByLocal(solidFill, 'srgbClr')
  if (srgb) {
    return '#' + (srgb.getAttribute('val') || '000000')
  }

  // scheme 颜色
  const scheme = findByLocal(solidFill, 'schemeClr')
  if (scheme) {
    const val = scheme.getAttribute('val') || ''
    const hex = themeColors?.[val] || DEFAULT_ACCENT_COLORS[val]
    if (hex) return hex.startsWith('#') ? hex : '#' + hex
  }

  return undefined
}

// ============== 缓存数据提取 ==============

/** 从 numCache/strCache 提取数据点 */
function extractCacheValues(cacheEl: Element): string[] {
  const pts = allByLocal(cacheEl, 'pt')
  const result: string[] = []
  for (const pt of pts) {
    const idx = parseInt(pt.getAttribute('idx') || '0')
    const v = findByLocal(pt, 'v')
    if (v) result[idx] = v.textContent || ''
  }
  return result
}

/** 从 numRef 或 strRef 提取缓存值 */
function extractRefValues(refEl: Element): { values: string[]; formatCode?: string } {
  const numCache = findByLocal(refEl, 'numCache')
  if (numCache) {
    const fmt = findByLocal(numCache, 'formatCode')
    return {
      values: extractCacheValues(numCache),
      formatCode: fmt?.textContent || undefined,
    }
  }
  const strCache = findByLocal(refEl, 'strCache')
  if (strCache) {
    return { values: extractCacheValues(strCache) }
  }
  return { values: [] }
}

// ============== 标题提取 ==============

/** 从 c:title 提取文本 */
function extractTitle(titleEl: Element): string {
  // rich text: <c:tx><c:rich><a:p><a:r><a:t>
  const runs = allByLocal(titleEl, 't')
  if (runs.length > 0) {
    return runs.map(r => r.textContent || '').join('')
  }
  // strRef fallback
  const strCache = findByLocal(titleEl, 'strCache')
  if (strCache) {
    const vals = extractCacheValues(strCache)
    return vals.join('')
  }
  return ''
}

// ============== 图表类型检测 ==============

interface ChartTypeInfo {
  kind: ChartKind
  fill?: boolean
  element: Element
}

/** OOXML 图表类型 → ChartKind 映射 */
const CHART_TYPE_MAP: Record<string, { kind: ChartKind; fill?: boolean }> = {
  barChart: { kind: 'bar' },
  bar3DChart: { kind: 'bar' },
  lineChart: { kind: 'line' },
  line3DChart: { kind: 'line' },
  areaChart: { kind: 'line', fill: true },
  area3DChart: { kind: 'line', fill: true },
  pieChart: { kind: 'pie' },
  pie3DChart: { kind: 'pie' },
  doughnutChart: { kind: 'doughnut' },
  scatterChart: { kind: 'scatter' },
  radarChart: { kind: 'radar' },
}

function detectChartType(plotArea: Element): ChartTypeInfo | null {
  for (const [xmlName, info] of Object.entries(CHART_TYPE_MAP)) {
    const el = findByLocal(plotArea, xmlName)
    if (el) return { ...info, element: el }
  }
  return null
}

// ============== 系列解析 ==============

function parseSeries(
  serEl: Element,
  themeColors?: Record<string, string>
): { series: ChartSeries; categories?: string[]; catFormat?: string } {
  // 系列名称
  const txEl = findByLocal(serEl, 'tx')
  let name = ''
  if (txEl) {
    const ref = findByLocal(txEl, 'strRef') || findByLocal(txEl, 'numRef')
    if (ref) {
      const { values } = extractRefValues(ref)
      name = values[0] || ''
    }
  }

  // 类别（categories）
  const catEl = findByLocal(serEl, 'cat')
  let categories: string[] | undefined
  let catFormat: string | undefined
  if (catEl) {
    const ref = findByLocal(catEl, 'numRef') || findByLocal(catEl, 'strRef')
    if (ref) {
      const { values, formatCode } = extractRefValues(ref)
      catFormat = formatCode
      // 如果是日期格式，转换 Excel 序列号
      if (formatCode && /y{2,4}/.test(formatCode)) {
        categories = values.map(v => {
          const n = parseFloat(v)
          return isNaN(n) ? v : excelSerialToDateStr(n)
        })
      } else {
        categories = values
      }
    }
    // 也可能是 multiLvlStrRef
    if (!categories) {
      const strCache = findByLocal(catEl, 'strCache')
      if (strCache) categories = extractCacheValues(strCache)
    }
  }

  // 数值
  const valEl = findByLocal(serEl, 'val')
  let values: number[] = []
  if (valEl) {
    const ref = findByLocal(valEl, 'numRef') || findByLocal(valEl, 'strRef')
    if (ref) {
      const { values: raw } = extractRefValues(ref)
      values = raw.map(v => parseFloat(v) || 0)
    }
  }

  // 系列颜色
  const spPr = findByLocal(serEl, 'spPr')
  let color: string | undefined
  if (spPr) color = extractColor(spPr, themeColors)

  // 饼图/环形图：每个数据点可能有独立颜色
  const dPts = childrenByLocal(serEl, 'dPt')
  let pointColors: string[] | undefined
  if (dPts.length > 0) {
    pointColors = []
    for (const dPt of dPts) {
      const idx = parseInt(findByLocal(dPt, 'idx')?.getAttribute('val') || '0')
      const dPtSpPr = findByLocal(dPt, 'spPr')
      if (dPtSpPr) {
        const c = extractColor(dPtSpPr, themeColors)
        if (c) pointColors[idx] = c
      }
    }
  }

  return {
    series: { name, values, color, pointColors },
    categories,
    catFormat,
  }
}

// ============== 主解析函数 ==============

/**
 * 解析 OOXML chart XML → ChartConfig
 */
export function parseChartXml(
  xml: string,
  themeColors?: Record<string, string>
): ChartConfig {
  const parser = new DOMParser()
  const doc = parser.parseFromString(xml, 'application/xml')
  const root = doc.documentElement

  // 标题
  const titleEl = findByLocal(root, 'title')
  const title = titleEl ? extractTitle(titleEl) : undefined

  // plotArea
  const plotArea = findByLocal(root, 'plotArea')
  if (!plotArea) {
    return { type: 'bar', categories: [], series: [], widthPx: 400, heightPx: 250, title }
  }

  // 检测图表类型
  const typeInfo = detectChartType(plotArea)
  if (!typeInfo) {
    return { type: 'bar', categories: [], series: [], widthPx: 400, heightPx: 250, title }
  }

  // 分组/堆叠
  const groupingEl = findByLocal(typeInfo.element, 'grouping')
  const groupingVal = groupingEl?.getAttribute('val') || ''
  let stacking: ChartConfig['stacking']
  if (groupingVal === 'stacked') stacking = 'stacked'
  else if (groupingVal === 'percentStacked') stacking = 'percentStacked'

  // barDir: 'bar' 表示水平条形图
  const barDirEl = findByLocal(typeInfo.element, 'barDir')
  const barDir = barDirEl?.getAttribute('val') || 'col'
  const kind: ChartKind = typeInfo.kind === 'bar' && barDir === 'bar' ? 'bar' : typeInfo.kind

  // 解析所有系列
  const serElements = childrenByLocal(typeInfo.element, 'ser')
  const allSeries: ChartSeries[] = []
  let categories: string[] = []
  let valFormat: string | undefined

  for (const serEl of serElements) {
    const { series, categories: cats, catFormat } = parseSeries(serEl, themeColors)
    allSeries.push(series)
    if (cats && cats.length > categories.length) categories = cats

    // 提取值格式
    if (!valFormat) {
      const valEl = findByLocal(serEl, 'val')
      if (valEl) {
        const numRef = findByLocal(valEl, 'numRef')
        if (numRef) {
          const numCache = findByLocal(numRef, 'numCache')
          if (numCache) {
            const fmt = findByLocal(numCache, 'formatCode')
            if (fmt?.textContent && fmt.textContent !== 'General') {
              valFormat = fmt.textContent
            }
          }
        }
      }
    }
  }

  // 为没有颜色的系列分配默认颜色
  const accentKeys = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6']
  for (let i = 0; i < allSeries.length; i++) {
    if (!allSeries[i].color) {
      const key = accentKeys[i % accentKeys.length]
      allSeries[i].color = themeColors?.[key]
        ? (themeColors[key].startsWith('#') ? themeColors[key] : '#' + themeColors[key])
        : DEFAULT_ACCENT_COLORS[key]
    }
  }

  // 饼图/环形图：如果没有 pointColors，用 accent 颜色填充
  if ((kind === 'pie' || kind === 'doughnut') && allSeries.length > 0) {
    const ser = allSeries[0]
    if (!ser.pointColors || ser.pointColors.length === 0) {
      ser.pointColors = ser.values.map((_, i) => {
        const key = accentKeys[i % accentKeys.length]
        return themeColors?.[key]
          ? (themeColors[key].startsWith('#') ? themeColors[key] : '#' + themeColors[key])
          : DEFAULT_ACCENT_COLORS[key]
      })
    }
  }

  // 图例位置
  const legendEl = findByLocal(root, 'legend')
  const legendPosEl = legendEl ? findByLocal(legendEl, 'legendPos') : null
  const legendPosMap: Record<string, ChartConfig['legendPosition']> = {
    t: 'top', b: 'bottom', l: 'left', r: 'right',
  }
  const legendPosition = legendPosMap[legendPosEl?.getAttribute('val') || ''] || 'bottom'

  return {
    type: kind,
    title,
    categories,
    series: allSeries,
    stacking,
    fill: typeInfo.fill,
    legendPosition,
    valFormat,
    widthPx: 400,
    heightPx: 250,
  }
}

// ============== Chart.js 配置转换 ==============

/** 将颜色加透明度 */
function withAlpha(hex: string, alpha: number): string {
  const r = parseInt(hex.slice(1, 3), 16)
  const g = parseInt(hex.slice(3, 5), 16)
  const b = parseInt(hex.slice(5, 7), 16)
  return `rgba(${r},${g},${b},${alpha})`
}

/**
 * ChartConfig → Chart.js ChartConfiguration
 */
export function chartConfigToChartJs(config: ChartConfig): ChartConfiguration {
  const { type, title, categories, series, stacking, fill, legendPosition, valFormat } = config

  // 构建 datasets
  const datasets: ChartDataset[] = series.map((s, i) => {
    const ds: any = {
      label: s.name,
      data: s.values,
    }

    if (type === 'pie' || type === 'doughnut') {
      // 饼图/环形图：每个数据点不同颜色
      ds.backgroundColor = s.pointColors || s.values.map((_, j) => {
        const keys = Object.keys(DEFAULT_ACCENT_COLORS)
        return DEFAULT_ACCENT_COLORS[keys[j % keys.length]]
      })
      ds.borderColor = '#fff'
      ds.borderWidth = 2
    } else {
      const color = s.color || DEFAULT_ACCENT_COLORS[`accent${(i % 6) + 1}`]
      ds.borderColor = color
      ds.backgroundColor = fill ? withAlpha(color, 0.6) : withAlpha(color, 0.8)
      ds.borderWidth = 2
      if (fill) ds.fill = true
      if (type === 'bar') ds.backgroundColor = withAlpha(color, 0.8)
    }

    return ds
  })

  // 构建 options
  const options: any = {
    responsive: false,
    maintainAspectRatio: false,
    animation: false,
    plugins: {
      title: title ? {
        display: true,
        text: title,
        font: { size: 14 },
      } : { display: false },
      legend: {
        display: true,
        position: legendPosition || 'bottom',
      },
    },
  }

  // 坐标轴（饼图/环形图/雷达图不需要）
  if (type !== 'pie' && type !== 'doughnut' && type !== 'radar') {
    const xAxis: any = {}
    const yAxis: any = {}

    if (stacking === 'stacked' || stacking === 'percentStacked') {
      xAxis.stacked = true
      yAxis.stacked = true
    }

    if (stacking === 'percentStacked') {
      yAxis.max = 1
      yAxis.ticks = {
        callback: (value: number) => Math.round(value * 100) + '%',
      }
    } else if (valFormat && valFormat.includes('%')) {
      yAxis.ticks = {
        callback: (value: number) => Math.round(value * 100) + '%',
      }
    }

    options.scales = { x: xAxis, y: yAxis }
  }

  return {
    type: type as any,
    data: {
      labels: categories,
      datasets,
    },
    options,
  }
}
