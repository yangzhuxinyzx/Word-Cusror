function registerFontsIpc(ipcMain, fontsService) {
  ipcMain.handle('fonts-list', async () => fontsService.listFonts())
  ipcMain.handle('fonts-read', async (_event, fileName) =>
    fontsService.readFont(fileName),
  )
}

module.exports = {
  registerFontsIpc,
}
