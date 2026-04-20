import type { FileItem } from '../../../types'
import { summarizeDocxFile } from '../../tools/packs/workspace/summarizers/DocxSummarizer'
import { summarizePdfFilePlaceholder } from '../../tools/packs/workspace/summarizers/PdfSummarizer'
import { summarizePptxBase64 } from '../../tools/packs/workspace/summarizers/PptxSummarizer'
import { summarizeTextContent } from '../../tools/packs/workspace/summarizers/TextSummarizer'
import { summarizeWorkbookPreview } from '../../tools/packs/workspace/summarizers/XlsxSummarizer'

export interface WorkspaceSummaryOptions {
  maxChars?: number
  maxSlides?: number
  format?: string
}

export interface WorkspaceSummaryAdapterOptions {
  isElectron: boolean
  workspaceSummaryMaxChars: number
  workspacePptxMaxSlides: number
  truncateWithNote: (text: string, maxLen: number, note: string) => string
  readFile?: (filePath: string) => Promise<{
    success: boolean
    data?: string
    type?: 'text' | 'docx' | 'doc-html' | 'pptx'
    error?: string
  }>
  getFileInfo?: (filePath: string) => Promise<{
    success: boolean
    data?: { modified?: string | Date; size?: number }
  }>
  excelListSheets?: (
    filePath: string,
  ) => Promise<{
    success: boolean
    sheets?: Array<{ name: string; rowCount: number; columnCount: number }>
    error?: string
  }>
  excelReadCells?: (
    filePath: string,
    sheetName: string,
    range: string,
  ) => Promise<{
    success: boolean
    cells?: Array<{
      r: number
      c: number
      text?: string
      value?: unknown
    }>
  }>
}

function htmlToPlainText(html: string): string {
  if (!html) return ''
  let text = html.replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
  text = text.replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
  text = text.replace(/&nbsp;/g, ' ')
  text = text.replace(/&amp;/g, '&')
  text = text.replace(/&lt;/g, '<')
  text = text.replace(/&gt;/g, '>')
  text = text.replace(/&quot;/g, '"')
  text = text.replace(/&#39;/g, "'")
  text = text.replace(/<[^>]+>/g, ' ')
  text = text.replace(/\s+/g, ' ').trim()
  return text
}

export class WorkspaceSummaryAdapter {
  private cache = new Map<string, { key: string; summary: string }>()

  constructor(private readonly options: WorkspaceSummaryAdapterOptions) {}

  async summarizeWorkspaceFile(
    file: FileItem,
    summaryOptions?: WorkspaceSummaryOptions,
  ): Promise<string> {
    if (!this.options.isElectron || !this.options.readFile) {
      return '⚠️ 当前环境不支持读取工作夹文件'
    }

    const maxChars =
      summaryOptions?.maxChars || this.options.workspaceSummaryMaxChars
    const maxSlides =
      summaryOptions?.maxSlides || this.options.workspacePptxMaxSlides
    const format = summaryOptions?.format || 'summary'

    let cacheKey = `${file.path}:${maxChars}:${maxSlides}:${format}`
    if (this.options.getFileInfo) {
      try {
        const infoResult = await this.options.getFileInfo(file.path)
        const info = infoResult?.data
        if (infoResult?.success && info) {
          cacheKey = `${file.path}:${String(info.modified || '')}:${info.size || ''}:${maxChars}:${maxSlides}:${format}`
        }
      } catch {
        // ignore
      }
    }

    const cached = this.cache.get(file.path)
    if (cached?.key === cacheKey) {
      return cached.summary
    }

    const ext = (file.extension || file.name.split('.').pop() || '').toLowerCase()
    let summary = ''

    if (ext === 'docx') {
      summary = await summarizeDocxFile({
        fileName: file.name,
        filePath: file.path,
        maxChars,
        format,
        readFile: this.options.readFile,
      })
    } else if (ext === 'doc') {
      const result = await this.options.readFile(file.path)
      if (result.success && result.data) {
        summary = summarizeTextContent(
          htmlToPlainText(result.data),
          maxChars,
          this.options.truncateWithNote,
          file.name,
        )
      } else {
        summary = `⚠️ 无法读取 .doc：${result.error || '未知错误'}`
      }
    } else if (ext === 'xlsx' || ext === 'xls') {
      if (!this.options.excelListSheets || !this.options.excelReadCells) {
        summary = '⚠️ Excel 工具不可用'
      } else {
        summary = await summarizeWorkbookPreview({
          filePath: file.path,
          maxChars,
          truncateWithNote: this.options.truncateWithNote,
          excelListSheets: this.options.excelListSheets,
          excelReadCells: this.options.excelReadCells,
        })
      }
    } else if (ext === 'pptx' || ext === 'ppt') {
      const result = await this.options.readFile(file.path)
      if (result.success && result.data && result.type === 'pptx') {
        summary = await summarizePptxBase64({
          base64: result.data,
          maxSlides,
          maxChars,
          truncateWithNote: this.options.truncateWithNote,
        })
      } else {
        summary = `⚠️ 无法读取 PPT：${result.error || '未知错误'}`
      }
    } else if (ext === 'pdf') {
      summary = summarizePdfFilePlaceholder(file.name)
    } else {
      const result = await this.options.readFile(file.path)
      if (result.success && result.data) {
        summary = summarizeTextContent(
          result.data,
          maxChars,
          this.options.truncateWithNote,
          file.name,
        )
      } else {
        summary = `⚠️ 无法读取文件：${result.error || '未知错误'}`
      }
    }

    this.cache.set(file.path, { key: cacheKey, summary })
    return summary
  }

  async buildWorkspaceAutoSummaries(
    flatFiles: FileItem[],
    currentFilePath?: string | null,
    maxFiles = 3,
  ): Promise<string> {
    if (!flatFiles.length) return ''
    const candidates = flatFiles.filter(
      (file) => file.type === 'file' && file.path !== currentFilePath,
    )
    if (!candidates.length) return ''

    const priority = (file: FileItem) => {
      const ext = (file.extension || file.name.split('.').pop() || '').toLowerCase()
      if (ext === 'docx') return 1
      if (ext === 'xlsx' || ext === 'xls') return 2
      if (ext === 'pptx' || ext === 'ppt') return 3
      if (ext === 'md' || ext === 'txt') return 4
      return 9
    }

    const selected = candidates
      .sort((left, right) => priority(left) - priority(right))
      .slice(0, maxFiles)

    if (!selected.length) return ''

    const summaries = await Promise.all(
      selected.map(async (file) => {
        const summary = await this.summarizeWorkspaceFile(file, {
          maxChars: this.options.workspaceSummaryMaxChars,
        })
        return `【${file.name}】\n${summary}`
      }),
    )

    return `【自动摘要】\n${summaries.join('\n\n')}`
  }
}
