import type { FileItem } from '../../../types'
import type {
  ElectronAPI,
  ExcelCellUpdate,
  ExcelCreateOptions,
  ExcelSearchOptions,
} from '../../../types/electron'

export interface ExcelWorkbookContextResult {
  ok: true
  excelFilePath: string
  currentFileName: string | null
}

export interface ExcelWorkbookContextError {
  ok: false
  message: string
}

export type ExcelWorkbookContext =
  | ExcelWorkbookContextResult
  | ExcelWorkbookContextError

export interface ExcelDomainAdapterOptions {
  currentFileName: string | null
  currentFilePath: string | null
  workspacePath: string | null
  refreshExcelData: () => Promise<boolean>
  refreshFiles: () => Promise<void>
  openFile: (file: FileItem) => Promise<void>
}

export class ExcelDomainAdapter {
  constructor(private readonly options: ExcelDomainAdapterOptions) {}

  private getApi(): ElectronAPI {
    if (!window.electronAPI) {
      throw new Error('Excel APIs are unavailable in the current environment.')
    }
    return window.electronAPI
  }

  ensureWorkbook(): ExcelWorkbookContext {
    const isExcelFile =
      this.options.currentFileName?.toLowerCase().endsWith('.xlsx') ||
      this.options.currentFileName?.toLowerCase().endsWith('.xls')
    if (!isExcelFile || !this.options.currentFilePath) {
      return {
        ok: false,
        message: 'Please open an Excel workbook first.',
      }
    }

    return {
      ok: true,
      excelFilePath: this.options.currentFilePath,
      currentFileName: this.options.currentFileName,
    }
  }

  getWorkspacePath(): string | null {
    return this.options.workspacePath
  }

  async refreshWorkbookPreview(): Promise<boolean> {
    return this.options.refreshExcelData()
  }

  async refreshWorkspaceFiles(): Promise<void> {
    await this.options.refreshFiles()
  }

  async openWorkbook(file: FileItem): Promise<void> {
    await this.options.openFile(file)
  }

  async readCells(filePath: string, sheetName: string, range: string) {
    return this.getApi().excelReadCells(filePath, sheetName, range)
  }

  async search(
    filePath: string,
    sheetName: string,
    searchText: string,
    options?: ExcelSearchOptions,
  ) {
    return this.getApi().excelSearch(filePath, sheetName, searchText, options)
  }

  async writeCells(
    filePath: string,
    sheetName: string,
    cellUpdates: ExcelCellUpdate[],
  ) {
    return this.getApi().excelWriteCells(filePath, sheetName, cellUpdates)
  }

  async insertRows(
    filePath: string,
    sheetName: string,
    startRow: number,
    count?: number,
    data?: unknown[][],
  ) {
    return this.getApi().excelInsertRows(filePath, sheetName, startRow, count, data)
  }

  async insertColumns(
    filePath: string,
    sheetName: string,
    startCol: number,
    count?: number,
  ) {
    return this.getApi().excelInsertColumns(filePath, sheetName, startCol, count)
  }

  async deleteRows(
    filePath: string,
    sheetName: string,
    startRow: number,
    count?: number,
  ) {
    return this.getApi().excelDeleteRows(filePath, sheetName, startRow, count)
  }

  async deleteColumns(
    filePath: string,
    sheetName: string,
    startCol: number,
    count?: number,
  ) {
    return this.getApi().excelDeleteColumns(filePath, sheetName, startCol, count)
  }

  async addSheet(filePath: string, sheetName: string) {
    return this.getApi().excelAddSheet(filePath, sheetName)
  }

  async deleteSheet(filePath: string, sheetName: string) {
    return this.getApi().excelDeleteSheet(filePath, sheetName)
  }

  async mergeCells(filePath: string, sheetName: string, range: string) {
    return this.getApi().excelMergeCells(filePath, sheetName, range)
  }

  async unmergeCells(filePath: string, sheetName: string, range: string) {
    return this.getApi().excelUnmergeCells(filePath, sheetName, range)
  }

  async createWorkbook(filePath: string, options?: ExcelCreateOptions) {
    return this.getApi().excelCreate(filePath, options)
  }

  async setFormula(filePath: string, sheetName: string, formulas: unknown) {
    return this.getApi().excelSetFormula(filePath, sheetName, formulas)
  }

  async sort(filePath: string, sheetName: string, options: unknown) {
    return this.getApi().excelSort(filePath, sheetName, options)
  }

  async autoFill(filePath: string, sheetName: string, options: unknown) {
    return this.getApi().excelAutoFill(filePath, sheetName, options)
  }

  async setDimensions(filePath: string, sheetName: string, options: unknown) {
    return this.getApi().excelSetDimensions(filePath, sheetName, options)
  }

  async conditionalFormat(
    filePath: string,
    sheetName: string,
    options: unknown,
  ) {
    return this.getApi().excelConditionalFormat(filePath, sheetName, options)
  }

  async calculate(filePath: string, sheetName: string, addresses: unknown) {
    return this.getApi().excelCalculate(filePath, sheetName, addresses)
  }

  async setFilter(filePath: string, sheetName: string, options: unknown) {
    return this.getApi().excelSetFilter(filePath, sheetName, options as {
      range?: string
      remove?: boolean
    })
  }

  async setValidation(filePath: string, sheetName: string, options: unknown) {
    return this.getApi().excelSetValidation(filePath, sheetName, options as {
      range: string
      type?: 'list' | 'whole' | 'decimal' | 'date' | 'textLength'
      values?: string[]
      min?: number
      max?: number
      allowBlank?: boolean
      showError?: boolean
      errorTitle?: string
      errorMessage?: string
      remove?: boolean
    })
  }

  async setHyperlink(filePath: string, sheetName: string, options: unknown) {
    return this.getApi().excelSetHyperlink(filePath, sheetName, options as {
      cell: string
      url?: string
      text?: string
      tooltip?: string
      remove?: boolean
    })
  }

  async findReplace(filePath: string, sheetName: string, options: unknown) {
    return this.getApi().excelFindReplace(filePath, sheetName, options as {
      find: string
      replace?: string
      matchCase?: boolean
      matchWholeCell?: boolean
      allSheets?: boolean
    })
  }

  async insertChart(filePath: string, sheetName: string, options: unknown) {
    return this.getApi().excelInsertChart(filePath, sheetName, options as {
      type?: 'column' | 'bar' | 'line' | 'pie' | 'area' | 'scatter'
      dataRange: string
      title?: string
      position?: string
      width?: number
      height?: number
    })
  }
}
