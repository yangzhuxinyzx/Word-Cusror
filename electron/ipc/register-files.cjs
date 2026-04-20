function registerFileIpc(ipcMain, filesService) {
  ipcMain.handle('get-file-url', async (_event, filePath) =>
    filesService.getFileUrl(filePath),
  )

  ipcMain.handle('get-local-file-url', async (_event, filePath) =>
    filesService.getLocalFileUrl(filePath),
  )

  ipcMain.handle('select-folder', async () => filesService.selectFolder())

  ipcMain.handle('read-folder', async (_event, folderPath) =>
    filesService.readFolder(folderPath),
  )

  ipcMain.handle('read-file', async (_event, filePath) =>
    filesService.readFile(filePath),
  )

  ipcMain.handle('write-file', async (_event, filePath, content) =>
    filesService.writeFile(filePath, content),
  )

  ipcMain.handle('append-file', async (_event, filePath, content) =>
    filesService.appendFile(filePath, content),
  )

  ipcMain.handle('write-binary-file', async (_event, filePath, base64Data) =>
    filesService.writeBinaryFile(filePath, base64Data),
  )

  ipcMain.handle('save-file-dialog', async (_event, defaultName) =>
    filesService.saveFileDialog(defaultName),
  )

  ipcMain.handle('create-file', async (_event, folderPath, fileName, content = '') =>
    filesService.createFile(folderPath, fileName, content),
  )

  ipcMain.handle('delete-file', async (_event, filePath) =>
    filesService.deleteFile(filePath),
  )

  ipcMain.handle('rename-file', async (_event, oldPath, newPath) =>
    filesService.renameFile(oldPath, newPath),
  )

  ipcMain.handle('show-in-folder', async (_event, filePath) =>
    filesService.showInFolder(filePath),
  )

  ipcMain.handle('get-file-info', async (_event, filePath) =>
    filesService.getFileInfo(filePath),
  )
}

module.exports = {
  registerFileIpc,
}
