/**
 * useMcpBridge — 监听主进程 MCP Bridge 请求，调用 DocumentContext 方法并回复结果
 */
import { useEffect, useRef } from 'react'
import { htmlToDsl } from '../utils/htmlToDsl'
import { serializeDslForAI } from '../utils/dslSerializer'
import type { DslBlock } from '../types/docDsl'
import type { ChartConfig } from '../utils/chartParser'
import type { FileItem } from '../types'

export interface McpBridgeActions {
  getContent: () => string
  workspacePath: string | null
  insertViaDsl: (position: string, blocks: DslBlock[]) => { success: boolean; message: string }
  replaceViaDsl: (search: string, replace: string, options?: { blockIndex?: number }) => { success: boolean; count: number; message: string }
  deleteViaDsl: (target: string, options?: { blockIndex?: number }) => { success: boolean; count: number; message: string }
  insertInDocument: (position: string, content: string) => { success: boolean; message: string }
  silentSaveToFile: () => Promise<{ success: boolean; error?: string }>
  openFile: (file: FileItem) => Promise<void>
  getTiptapDocumentStructure: () => string
  currentFile: FileItem | null
}

export function useMcpBridge(actions: McpBridgeActions) {
  const actionsRef = useRef(actions)
  actionsRef.current = actions

  useEffect(() => {
    const api = window.electronAPI
    if (!api?.onMcpBridgeRequest) return

    const cleanup = api.onMcpBridgeRequest(async (payload) => {
      const { requestId, action, params } = payload
      const a = actionsRef.current
      let result: any

      try {
        switch (action) {
          case 'get_workspace_path':
            result = { success: true, data: a.workspacePath }
            break

          case 'read_document': {
            const content = a.getContent()
            if (!content) {
              result = { success: false, error: '没有打开的文档' }
              break
            }
            const format = params.format || 'dsl'
            if (format === 'structure') {
              result = { success: true, data: a.getTiptapDocumentStructure() }
            } else if (format === 'plain') {
              const dsl = htmlToDsl(content)
              const text = dsl.blocks.map(b => {
                if ('content' in b) {
                  if (typeof b.content === 'string') return b.content
                  if (Array.isArray(b.content)) return b.content.map((r: any) => typeof r === 'string' ? r : r.text || '').join('')
                }
                return ''
              }).join('\n')
              result = { success: true, data: text }
            } else {
              const dsl = htmlToDsl(content)
              result = { success: true, data: serializeDslForAI(dsl) }
            }
            break
          }

          case 'insert': {
            const insertResult = a.insertViaDsl(params.position || 'end', params.blocks || [])
            result = insertResult
            break
          }

          case 'replace': {
            const replaceResult = a.replaceViaDsl(params.search, params.replace, {
              blockIndex: params.options?.blockIndex,
            })
            result = replaceResult
            break
          }

          case 'delete': {
            const deleteResult = a.deleteViaDsl(params.target, {
              blockIndex: params.options?.blockIndex,
            })
            result = deleteResult
            break
          }

          case 'save': {
            result = await a.silentSaveToFile()
            break
          }

          case 'open_file': {
            const filePath = params.filePath as string
            if (!filePath) {
              result = { success: false, error: '缺少 filePath 参数' }
              break
            }
            const name = filePath.split(/[\\/]/).pop() || ''
            const ext = name.includes('.') ? name.split('.').pop()?.toLowerCase() : ''
            const fileItem: FileItem = {
              name,
              path: filePath,
              type: 'file',
              extension: ext ? `.${ext}` : undefined,
            }
            await a.openFile(fileItem)
            result = { success: true, message: `已打开: ${name}` }
            break
          }

          case 'insert_chart': {
            const chartConfig: ChartConfig = {
              type: params.chartType || 'bar',
              title: params.title,
              categories: params.data?.labels || [],
              series: (params.data?.datasets || []).map((ds: any) => ({
                name: ds.label || '',
                values: ds.data || [],
              })),
              widthPx: params.width || 500,
              heightPx: params.height || 300,
            }
            const encoded = encodeURIComponent(JSON.stringify(chartConfig))
            const w = chartConfig.widthPx || 500
            const h = chartConfig.heightPx || 300
            const html = `<div data-type="docx-chart" data-chart-config="${encoded}" style="width:${w}px;height:${h}px"></div>`
            result = a.insertInDocument(params.position || 'end', html)
            break
          }

          default:
            result = { success: false, error: `未知操作: ${action}` }
        }
      } catch (e: any) {
        result = { success: false, error: e.message || String(e) }
      }

      api.mcpBridgeResponse(requestId, result)
    })

    return cleanup
  }, [])
}
