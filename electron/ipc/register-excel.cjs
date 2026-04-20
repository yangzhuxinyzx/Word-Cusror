function registerExcelIpc(ipcMain, excelService) {
  ipcMain.handle('check-libreoffice', async () => excelService.checkLibreOffice())

  ipcMain.handle('excel-convert-xls-to-xlsx', async (_event, xlsPath) =>
    excelService.excelConvertXlsToXlsx(xlsPath),
  )

  ipcMain.handle('excel-open', async (_event, filePath) =>
    excelService.excelOpen(filePath),
  )

  ipcMain.handle('excel-read-cells', async (_event, filePath, sheetName, rangeOrCell) =>
    excelService.excelReadCells(filePath, sheetName, rangeOrCell),
  )

  ipcMain.handle('excel-search', async (_event, filePath, sheetName, searchText, options = {}) =>
    excelService.excelSearch(filePath, sheetName, searchText, options),
  )

  ipcMain.handle('excel-write-cells', async (_event, filePath, sheetName, cellUpdates) =>
    excelService.excelWriteCells(filePath, sheetName, cellUpdates),
  )

  ipcMain.handle('excel-insert-rows', async (_event, filePath, sheetName, startRow, count = 1, data = null) =>
    excelService.excelInsertRows(filePath, sheetName, startRow, count, data),
  )

  ipcMain.handle('excel-insert-columns', async (_event, filePath, sheetName, startCol, count = 1) =>
    excelService.excelInsertColumns(filePath, sheetName, startCol, count),
  )

  ipcMain.handle('excel-add-sheet', async (_event, filePath, sheetName) =>
    excelService.excelAddSheet(filePath, sheetName),
  )

  ipcMain.handle('excel-delete-rows', async (_event, filePath, sheetName, startRow, count = 1) =>
    excelService.excelDeleteRows(filePath, sheetName, startRow, count),
  )

  ipcMain.handle('excel-delete-columns', async (_event, filePath, sheetName, startCol, count = 1) =>
    excelService.excelDeleteColumns(filePath, sheetName, startCol, count),
  )

  ipcMain.handle('excel-delete-sheet', async (_event, filePath, sheetName) =>
    excelService.excelDeleteSheet(filePath, sheetName),
  )

  ipcMain.handle('excel-list-sheets', async (_event, filePath) =>
    excelService.excelListSheets(filePath),
  )

  ipcMain.handle('excel-merge-cells', async (_event, filePath, sheetName, range) =>
    excelService.excelMergeCells(filePath, sheetName, range),
  )

  ipcMain.handle('excel-unmerge-cells', async (_event, filePath, sheetName, range) =>
    excelService.excelUnmergeCells(filePath, sheetName, range),
  )

  ipcMain.handle('excel-set-formula', async (_event, filePath, sheetName, formulas) =>
    excelService.excelSetFormula(filePath, sheetName, formulas),
  )

  ipcMain.handle('excel-sort', async (_event, filePath, sheetName, options) =>
    excelService.excelSort(filePath, sheetName, options),
  )

  ipcMain.handle('excel-conditional-format', async (_event, filePath, sheetName, options) =>
    excelService.excelConditionalFormat(filePath, sheetName, options),
  )

  ipcMain.handle('excel-auto-fill', async (_event, filePath, sheetName, options) =>
    excelService.excelAutoFill(filePath, sheetName, options),
  )

  ipcMain.handle('excel-set-dimensions', async (_event, filePath, sheetName, options) =>
    excelService.excelSetDimensions(filePath, sheetName, options),
  )

  ipcMain.handle('excel-add-chart', async (_event, filePath, sheetName, options) =>
    excelService.excelAddChart(filePath, sheetName, options),
  )

  ipcMain.handle('excel-calculate', async (_event, filePath, sheetName, addresses) =>
    excelService.excelCalculate(filePath, sheetName, addresses),
  )

  ipcMain.handle('excel-create', async (_event, filePath, options = {}) =>
    excelService.excelCreate(filePath, options),
  )

  ipcMain.handle('excel-close', async (_event, filePath) =>
    excelService.excelClose(filePath),
  )

  ipcMain.handle('excel-reload', async (_event, filePath) =>
    excelService.excelReload(filePath),
  )

  ipcMain.handle('excel-set-filter', async (_event, filePath, sheetName, options) =>
    excelService.excelSetFilter(filePath, sheetName, options),
  )

  ipcMain.handle('excel-set-validation', async (_event, filePath, sheetName, options) =>
    excelService.excelSetValidation(filePath, sheetName, options),
  )

  ipcMain.handle('excel-set-hyperlink', async (_event, filePath, sheetName, options) =>
    excelService.excelSetHyperlink(filePath, sheetName, options),
  )

  ipcMain.handle('excel-find-replace', async (_event, filePath, sheetName, options) =>
    excelService.excelFindReplace(filePath, sheetName, options),
  )

  ipcMain.handle('excel-insert-chart', async (_event, filePath, sheetName, options) =>
    excelService.excelInsertChart(filePath, sheetName, options),
  )
}

module.exports = {
  registerExcelIpc,
}
