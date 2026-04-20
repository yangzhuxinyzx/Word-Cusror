function registerMemoryIpc(ipcMain, memoryService) {
  ipcMain.handle('memory-search', async (_event, options) =>
    memoryService.search(options),
  )

  ipcMain.handle('memory-append', async (_event, payload) =>
    memoryService.append(payload),
  )

  ipcMain.handle('memory-append-session', async (_event, payload) =>
    memoryService.appendSession(payload),
  )

  ipcMain.handle('memory-status', async () => memoryService.status())

  ipcMain.handle('memory-status-detail', async () =>
    memoryService.statusDetail(),
  )

  ipcMain.handle('memory-clear', async (_event, payload) =>
    memoryService.clear(payload),
  )

  ipcMain.handle('memory-rebuild-index', async () =>
    memoryService.rebuildIndex(),
  )
}

module.exports = {
  registerMemoryIpc,
}
