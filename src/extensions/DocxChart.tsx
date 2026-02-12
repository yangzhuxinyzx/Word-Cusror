/**
 * DocxChart — Tiptap Node 扩展，用 chart.js 渲染 OOXML 图表
 */
import { Node, mergeAttributes } from '@tiptap/core'
import { ReactNodeViewRenderer, NodeViewWrapper } from '@tiptap/react'
import { useRef, useEffect, memo } from 'react'
import {
  Chart,
  CategoryScale,
  LinearScale,
  BarController,
  BarElement,
  LineController,
  LineElement,
  PointElement,
  ArcElement,
  PieController,
  DoughnutController,
  RadarController,
  RadialLinearScale,
  ScatterController,
  Filler,
  Legend,
  Title,
  Tooltip,
} from 'chart.js'
import { chartConfigToChartJs } from '../utils/chartParser'
import type { ChartConfig } from '../utils/chartParser'

// 注册需要的 chart.js 组件（tree-shakeable）
Chart.register(
  CategoryScale, LinearScale,
  BarController, BarElement,
  LineController, LineElement, PointElement,
  ArcElement, PieController, DoughnutController,
  RadarController, RadialLinearScale,
  ScatterController,
  Filler, Legend, Title, Tooltip,
)

// ============== React NodeView 组件 ==============

interface DocxChartViewProps {
  node: { attrs: Record<string, any> }
}

const DocxChartView = memo(function DocxChartView({ node }: DocxChartViewProps) {
  const canvasRef = useRef<HTMLCanvasElement>(null)
  const chartRef = useRef<Chart | null>(null)

  const configStr = node.attrs.chartConfig as string | null
  const width = (node.attrs.widthPx as number) || 400
  const height = (node.attrs.heightPx as number) || 250

  useEffect(() => {
    if (!canvasRef.current || !configStr) return

    try {
      const config: ChartConfig = JSON.parse(decodeURIComponent(configStr))
      const chartJsConfig = chartConfigToChartJs(config)

      // 图表始终用白底深色文字（与 Word 一致），不跟随暗色主题
      const textColor = '#333'
      const gridColor = 'rgba(0,0,0,0.1)'
      const opts = chartJsConfig.options as any
      if (opts?.plugins?.title) opts.plugins.title.color = textColor
      if (opts?.plugins?.legend) {
        opts.plugins.legend.labels = { ...opts.plugins.legend.labels, color: textColor }
      }
      if (opts?.scales?.x) {
        opts.scales.x.ticks = { ...opts.scales.x.ticks, color: textColor }
        opts.scales.x.grid = { ...opts.scales.x.grid, color: gridColor }
      }
      if (opts?.scales?.y) {
        opts.scales.y.ticks = { ...opts.scales.y.ticks, color: textColor }
        opts.scales.y.grid = { ...opts.scales.y.grid, color: gridColor }
      }

      chartRef.current?.destroy()
      chartRef.current = new Chart(canvasRef.current, chartJsConfig as any)
    } catch (e) {
      console.warn('[DocxChart] render error:', e)
    }

    return () => {
      chartRef.current?.destroy()
      chartRef.current = null
    }
  }, [configStr])

  if (!configStr) {
    return (
      <NodeViewWrapper
        as="div"
        style={{
          width, height, maxWidth: '100%', margin: '8px auto',
          display: 'flex', alignItems: 'center', justifyContent: 'center',
          background: '#f8f9fa', border: '1px dashed #ccc', borderRadius: 4,
          color: '#666', fontSize: 13,
        }}
      >
        [图表]
      </NodeViewWrapper>
    )
  }

  return (
    <NodeViewWrapper
      as="div"
      style={{
        width, height, maxWidth: '100%', margin: '8px auto',
        background: '#fff', borderRadius: 4, padding: 4,
      }}
    >
      <canvas ref={canvasRef} width={width - 8} height={height - 8} />
    </NodeViewWrapper>
  )
})

// ============== Tiptap Node 扩展 ==============

export const DocxChart = Node.create({
  name: 'docxChart',
  group: 'block',
  atom: true,

  addAttributes() {
    return {
      chartConfig: {
        default: null,
        parseHTML: (el: HTMLElement) => el.getAttribute('data-chart-config'),
        renderHTML: (attrs: Record<string, any>) =>
          attrs.chartConfig ? { 'data-chart-config': attrs.chartConfig } : {},
      },
      widthPx: {
        default: 400,
        parseHTML: (el: HTMLElement) => {
          const w = el.style.width
          return w ? parseInt(w) || 400 : 400
        },
        renderHTML: () => ({}),
      },
      heightPx: {
        default: 250,
        parseHTML: (el: HTMLElement) => {
          const h = el.style.height
          return h ? parseInt(h) || 250 : 250
        },
        renderHTML: () => ({}),
      },
    }
  },

  parseHTML() {
    return [{ tag: 'div[data-type="docx-chart"]' }]
  },

  renderHTML({ HTMLAttributes }) {
    const w = HTMLAttributes.widthPx || 400
    const h = HTMLAttributes.heightPx || 250
    return [
      'div',
      mergeAttributes(HTMLAttributes, {
        'data-type': 'docx-chart',
        style: `width:${w}px;height:${h}px;max-width:100%;margin:8px auto;`,
      }),
    ]
  },

  addNodeView() {
    return ReactNodeViewRenderer(DocxChartView as any)
  },
})
