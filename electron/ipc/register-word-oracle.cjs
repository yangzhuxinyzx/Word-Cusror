function registerWordOracleIpc(ipcMain, wordOracleService) {
  ipcMain.handle('word-oracle-export', async (_event, payload = {}) =>
    wordOracleService.export(payload),
  )

  ipcMain.handle('word-oracle-diff', async (_event, payload = {}) =>
    wordOracleService.diff(payload),
  )
}

module.exports = {
  registerWordOracleIpc,
}
