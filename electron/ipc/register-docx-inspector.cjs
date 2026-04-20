function registerDocxInspectorIpc(ipcMain, docxInspectorService) {
  ipcMain.handle('docx-inspect', async (_event, filePath) =>
    docxInspectorService.inspect(filePath),
  )
}

module.exports = {
  registerDocxInspectorIpc,
}
