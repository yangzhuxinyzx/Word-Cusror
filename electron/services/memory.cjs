function createMemoryService(options) {
  const { getMemoryManager } = options

  return {
    async search(searchOptions = {}) {
      const mgr = getMemoryManager()
      const query = searchOptions.query || ''
      const topK = searchOptions.topK || 5
      const textWeight =
        typeof searchOptions.textWeight === 'number'
          ? searchOptions.textWeight
          : 0.6
      const vectorWeight =
        typeof searchOptions.vectorWeight === 'number'
          ? searchOptions.vectorWeight
          : 0.4
      const workspaceKey = searchOptions.workspaceKey || ''
      const sources = searchOptions.sources
      return mgr.search({ query, topK, textWeight, vectorWeight, workspaceKey, sources })
    },

    async append(payload = {}) {
      const mgr = getMemoryManager()
      const text = payload.text || ''
      if (!text.trim()) {
        return { success: false, error: 'Text is empty' }
      }
      return mgr.appendDaily({
        text,
        source: payload.source || 'chat',
        tags: Array.isArray(payload.tags) ? payload.tags : [],
      })
    },

    async appendSession(payload = {}) {
      const mgr = getMemoryManager()
      const sessionId = payload.sessionId || ''
      const text = payload.text || ''
      if (!sessionId || !text.trim()) {
        return { success: false, error: 'sessionId and text are required' }
      }
      return mgr.appendSession({
        sessionId,
        text,
        meta: payload.meta || {},
      })
    },

    async status() {
      const mgr = getMemoryManager()
      return mgr.getStatus()
    },

    async statusDetail() {
      const mgr = getMemoryManager()
      return mgr.getStatusDetail()
    },

    async clear(payload = {}) {
      const mgr = getMemoryManager()
      const scope = payload.scope || 'all'
      return mgr.clear(scope)
    },

    async rebuildIndex() {
      const mgr = getMemoryManager()
      mgr.rebuildAll()
      return mgr.getStatusDetail()
    },
  }
}

module.exports = {
  createMemoryService,
}
