function registerNativeTextIpc(ipcMain, nativeTextService) {
  ipcMain.handle('text-measure-native', async (_event, payload = {}) =>
    nativeTextService.measure(payload),
  )
}

module.exports = {
  registerNativeTextIpc,
}
