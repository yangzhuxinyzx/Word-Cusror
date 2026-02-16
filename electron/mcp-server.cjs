#!/usr/bin/env node
/**
 * word-cursor-mcp — 独立 MCP Server（stdio transport）
 * 外部 agent（OpenClaw / Claude Code）通过 stdio 启动此脚本，
 * 它再通过 HTTP 调用运行中的 Word-Cursor Electron 应用的 bridge API。
 */
const http = require('http')
const path = require('path')

const BRIDGE_PORT = parseInt(process.env.WORD_CURSOR_PORT || '19527', 10)
const BRIDGE_URL = `http://127.0.0.1:${BRIDGE_PORT}`

// ─── SDK require（CJS 路径，与 main.cjs 一致）───
const sdkBase = path.join(__dirname, '..', 'node_modules', '@modelcontextprotocol', 'sdk', 'dist', 'cjs')
const { McpServer } = require(path.join(sdkBase, 'server', 'mcp.js'))
const { StdioServerTransport } = require(path.join(sdkBase, 'server', 'stdio.js'))
const { z } = require('zod')

// ─── Bridge HTTP 调用 ───
function callBridge(action, params) {
  return new Promise((resolve, reject) => {
    const body = JSON.stringify({ action, params: params || {} })
    const req = http.request(BRIDGE_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(body) },
    }, (res) => {
      let data = ''
      res.on('data', chunk => { data += chunk })
      res.on('end', () => {
        try { resolve(JSON.parse(data)) }
        catch { reject(new Error(`无效响应: ${data.slice(0, 200)}`)) }
      })
    })
    req.on('error', (e) => reject(new Error(`无法连接 Word-Cursor (${BRIDGE_URL}): ${e.message}`)))
    req.setTimeout(20000, () => { req.destroy(); reject(new Error('Bridge 请求超时')) })
    req.write(body)
    req.end()
  })
}

// ─── 工具结果格式化 ───
function ok(text) { return { content: [{ type: 'text', text }] } }
function err(text) { return { content: [{ type: 'text', text }], isError: true } }

async function main() {
  const server = new McpServer({
    name: 'word-cursor',
    version: '1.0.0',
  }, {
    capabilities: { tools: { listChanged: false } },
    instructions: 'MCP server for Word-Cursor document editor. Read, edit, and manage Word documents through a running Word-Cursor instance.',
  })

  // ─── Tool: word_read_document ───
  server.registerTool('word_read_document', {
    title: 'Read Document',
    description: '读取当前打开的文档内容。返回 DSL 格式的结构化 JSON（带 _i 索引），可用于后续编辑操作。',
    inputSchema: {
      format: z.enum(['dsl', 'structure', 'plain']).optional()
        .describe('输出格式: dsl=完整DSL JSON, structure=大纲, plain=纯文本'),
    },
    annotations: { readOnlyHint: true, destructiveHint: false },
  }, async ({ format }) => {
    const result = await callBridge('read_document', { format: format || 'dsl' })
    if (!result.success) return err(result.error || '读取失败')
    return ok(typeof result.data === 'string' ? result.data : JSON.stringify(result.data, null, 2))
  })

  // ─── Tool: word_insert ───
  server.registerTool('word_insert', {
    title: 'Insert Content',
    description: '在文档指定位置插入 DSL blocks。position 使用 "start"/"end" 或锚点文字。',
    inputSchema: {
      position: z.string().describe('插入位置: "start", "end", 或锚点文字'),
      blocks: z.string().describe('DSL blocks 的 JSON 数组字符串，如 [{"type":"paragraph","content":"Hello"}]'),
    },
    annotations: { readOnlyHint: false, destructiveHint: false },
  }, async ({ position, blocks }) => {
    let parsed
    try { parsed = JSON.parse(blocks) } catch { return err('blocks 参数不是有效 JSON') }
    const result = await callBridge('insert', { position, blocks: parsed })
    return result.success ? ok(result.message || '插入成功') : err(result.message || result.error || '插入失败')
  })

  // ─── Tool: word_replace ───
  server.registerTool('word_replace', {
    title: 'Replace Text',
    description: '查找并替换文档中的文本。可选指定 blockIndex 精确定位。',
    inputSchema: {
      search: z.string().describe('要查找的文本'),
      replace: z.string().describe('替换为的文本'),
      blockIndex: z.number().int().optional().describe('可选: 目标 block 的 _i 索引'),
    },
    annotations: { readOnlyHint: false, destructiveHint: false },
  }, async ({ search, replace, blockIndex }) => {
    const result = await callBridge('replace', { search, replace, options: { blockIndex } })
    return result.success ? ok(result.message || `替换了 ${result.count || 0} 处`) : err(result.message || result.error || '替换失败')
  })

  // ─── Tool: word_delete ───
  server.registerTool('word_delete', {
    title: 'Delete Text',
    description: '删除文档中匹配的文本或指定 block。',
    inputSchema: {
      target: z.string().describe('要删除的文本'),
      blockIndex: z.number().int().optional().describe('可选: 目标 block 的 _i 索引'),
    },
    annotations: { readOnlyHint: false, destructiveHint: true },
  }, async ({ target, blockIndex }) => {
    const result = await callBridge('delete', { target, options: { blockIndex } })
    return result.success ? ok(result.message || `删除了 ${result.count || 0} 处`) : err(result.message || result.error || '删除失败')
  })

  // ─── Tool: word_insert_chart ───
  server.registerTool('word_insert_chart', {
    title: 'Insert Chart',
    description: '在文档中插入图表（柱状图/折线图/饼图/环形图/散点图/雷达图）。',
    inputSchema: {
      chartType: z.enum(['bar', 'line', 'pie', 'doughnut', 'scatter', 'radar']).describe('图表类型'),
      title: z.string().optional().describe('图表标题'),
      labels: z.string().describe('X轴标签 JSON 数组，如 ["Q1","Q2","Q3","Q4"]'),
      datasets: z.string().describe('数据集 JSON 数组，如 [{"label":"销售额","data":[100,200,300,400]}]'),
      position: z.string().optional().describe('插入位置，默认 "end"'),
      width: z.number().optional().describe('宽度像素，默认 500'),
      height: z.number().optional().describe('高度像素，默认 300'),
    },
    annotations: { readOnlyHint: false, destructiveHint: false },
  }, async ({ chartType, title, labels, datasets, position, width, height }) => {
    let parsedLabels, parsedDatasets
    try { parsedLabels = JSON.parse(labels) } catch { return err('labels 不是有效 JSON') }
    try { parsedDatasets = JSON.parse(datasets) } catch { return err('datasets 不是有效 JSON') }
    const result = await callBridge('insert_chart', {
      chartType, title, position: position || 'end',
      width: width || 500, height: height || 300,
      data: { labels: parsedLabels, datasets: parsedDatasets },
    })
    return result.success ? ok(result.message || '图表已插入') : err(result.message || result.error || '插入图表失败')
  })

  // ─── Tool: word_save ───
  server.registerTool('word_save', {
    title: 'Save Document',
    description: '保存当前文档到磁盘。',
    inputSchema: {},
    annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: true },
  }, async () => {
    const result = await callBridge('save', {})
    return result.success ? ok('文档已保存') : err(result.error || '保存失败')
  })

  // ─── Tool: word_open ───
  server.registerTool('word_open', {
    title: 'Open Document',
    description: '在编辑器中打开一个文档文件。',
    inputSchema: {
      filePath: z.string().describe('.docx 文件的绝对路径'),
    },
    annotations: { readOnlyHint: false, destructiveHint: false },
  }, async ({ filePath }) => {
    const result = await callBridge('open_file', { filePath })
    return result.success ? ok(result.message || '已打开') : err(result.message || result.error || '打开失败')
  })

  // ─── Tool: word_list_files ───
  server.registerTool('word_list_files', {
    title: 'List Files',
    description: '列出当前工作区的文件。',
    inputSchema: {
      filter: z.string().optional().describe('文件扩展名过滤，如 ".docx"'),
    },
    annotations: { readOnlyHint: true, destructiveHint: false },
  }, async ({ filter }) => {
    const result = await callBridge('list_files', { filter })
    if (!result.success) return err(result.error || '列出文件失败')
    return ok(JSON.stringify(result.data, null, 2))
  })

  // ─── 启动 stdio transport ───
  const transport = new StdioServerTransport()
  await server.connect(transport)
}

main().catch(e => {
  process.stderr.write(`[word-cursor-mcp] 启动失败: ${e.message}\n`)
  process.exit(1)
})
