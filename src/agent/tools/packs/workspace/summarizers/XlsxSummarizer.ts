export async function summarizeWorkbookPreview(params: {
  filePath: string
  maxChars: number
  truncateWithNote: (text: string, maxLen: number, note: string) => string
  excelListSheets: (
    filePath: string,
  ) => Promise<{
    success: boolean
    sheets?: Array<{ name: string; rowCount: number; columnCount: number }>
    error?: string
  }>
  excelReadCells: (
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
}): Promise<string> {
  const list = await params.excelListSheets(params.filePath)
  if (!list.success || !list.sheets?.length) {
    return `⚠️ 无法读取工作表：${list.error || '未知错误'}`
  }

  const sheetNames = list.sheets
    .map((sheet) => `${sheet.name}(${sheet.rowCount}x${sheet.columnCount})`)
    .join(', ')

  const lines: string[] = [`【工作表】${sheetNames}`]
  const firstSheet = list.sheets[0]?.name
  if (firstSheet) {
    const preview = await params.excelReadCells(params.filePath, firstSheet, 'A1:E8')
    if (preview.success && preview.cells?.length) {
      const maxRows = 6
      const maxCols = 5
      const cellMap = new Map<string, string>()
      for (const cell of preview.cells) {
        if (cell.r < maxRows && cell.c < maxCols) {
          cellMap.set(`${cell.r}-${cell.c}`, String(cell.text || cell.value || ''))
        }
      }
      const rows: string[] = []
      for (let r = 0; r < maxRows; r += 1) {
        const cols: string[] = []
        for (let c = 0; c < maxCols; c += 1) {
          cols.push(cellMap.get(`${r}-${c}`) || '')
        }
        if (cols.some(Boolean)) {
          rows.push(cols.join('\t'))
        }
      }
      if (rows.length) {
        lines.push(`【${firstSheet} 预览】`)
        lines.push(rows.join('\n'))
      }
    }
  }

  return params.truncateWithNote(lines.join('\n'), params.maxChars, 'Excel 摘要')
}
