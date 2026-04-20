function registerAIProxyIpc(ipcMain, aiProxyService) {
  ipcMain.handle('ai-chat-completions', async (event, payload = {}) =>
    aiProxyService.chatCompletions(event, payload),
  )

  ipcMain.handle('ai-cancel', async (_event, requestId) =>
    aiProxyService.cancel(requestId),
  )
}

module.exports = {
  registerAIProxyIpc,
}
