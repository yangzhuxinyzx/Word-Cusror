import { generateDocxAgentContextFromFilePath } from '../../../../../utils/docxAgentContext'
import { parseDocxToHtmlForAgent } from '../../../../../utils/docxParser'
import { htmlToDsl } from '../../../../../utils/htmlToDsl'
import { serializeDslForAI } from '../../../../../utils/dslSerializer'

export async function summarizeDocxFile(params: {
  fileName: string
  filePath: string
  maxChars: number
  format?: string
  readFile: (filePath: string) => Promise<{
    success: boolean
    data?: string
    error?: string
  }>
}): Promise<string> {
  if (params.format === 'dsl') {
    try {
      const result = await params.readFile(params.filePath)
      if (result.success && result.data) {
        const parsed = await parseDocxToHtmlForAgent(result.data)
        const dsl = htmlToDsl(parsed.html, { stripDiffMarkers: true })
        return `【Word 文档 DSL】${params.fileName}\n${serializeDslForAI(dsl, {
          maxLength: params.maxChars - 100,
        })}`
      }
      return `⚠️ 无法读取 .docx：${result.error || '未知错误'}`
    } catch (error) {
      return `⚠️ DSL 解析失败：${(error as Error).message}`
    }
  }

  return generateDocxAgentContextFromFilePath(params.fileName, params.filePath, {
    maxLength: params.maxChars,
    maxParagraphs: 30,
    maxParagraphLength: 120,
  })
}
