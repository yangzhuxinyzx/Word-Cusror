function registerPptIpc(ipcMain, pptService) {
  ipcMain.handle('pptx-render-preview', async (_event, filePath) =>
    pptService.renderPreview(filePath),
  )

  ipcMain.handle('openrouter-gemini-ppt-prompts', async (_event, options = {}) =>
    pptService.generatePrompts(options),
  )

  ipcMain.handle('ppt-generate-deck', async (_event, options = {}) =>
    pptService.generateDeck(options),
  )

  ipcMain.handle('ppt-edit-slides', async (_event, options = {}) =>
    pptService.editSlides(options),
  )

  ipcMain.handle('ppt-detect-text-layer', async (_event, options = {}) =>
    pptService.detectTextLayer(options),
  )

  ipcMain.handle('ppt-apply-text-edits', async (_event, options = {}) =>
    pptService.applyTextEdits(options),
  )

  ipcMain.handle('ppt-text-edit-health', async (_event, options = {}) =>
    pptService.textEditHealth(options),
  )
}

module.exports = {
  registerPptIpc,
}
