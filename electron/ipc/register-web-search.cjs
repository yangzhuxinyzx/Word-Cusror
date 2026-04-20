function registerWebSearchIpc(ipcMain, webSearchService) {
  ipcMain.handle('web-search', async (_event, options = {}) => {
    const query = (options.query || '').trim()
    if (!query) {
      return { success: false, message: 'Missing query parameter' }
    }

    try {
      return await webSearchService.performWebSearch(query, {
        locale: options.locale,
        region: options.region,
        num: options.num,
        braveApiKey: options.braveApiKey,
      })
    } catch (error) {
      console.error('Brave web search failed:', error)
      return { success: false, error: error.message || String(error) }
    }
  })
}

module.exports = {
  registerWebSearchIpc,
}
