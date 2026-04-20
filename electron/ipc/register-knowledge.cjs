function registerKnowledgeIpc(ipcMain, knowledgeService) {
  ipcMain.handle('knowledge-configure', async (_event, payload = {}) =>
    knowledgeService.configure(payload),
  )

  ipcMain.handle('knowledge-set-active-workspace', async (_event, payload = {}) =>
    knowledgeService.setActiveWorkspace(payload),
  )

  ipcMain.handle('knowledge-status', async () => knowledgeService.status())

  ipcMain.handle('knowledge-retrieve', async (_event, payload = {}) =>
    knowledgeService.retrieve(payload),
  )

  ipcMain.handle('knowledge-rebuild', async (_event, payload = {}) =>
    knowledgeService.rebuild(payload),
  )

  ipcMain.handle('knowledge-list-pending-profile', async () =>
    knowledgeService.listPendingProfile(),
  )

  ipcMain.handle('knowledge-resolve-pending-profile', async (_event, payload = {}) =>
    knowledgeService.resolvePendingProfile(payload),
  )

  ipcMain.handle('knowledge-list-profile-facts', async () =>
    knowledgeService.listProfileFacts(),
  )

  ipcMain.handle('knowledge-queue-profile-candidates', async (_event, payload = {}) =>
    knowledgeService.queueProfileCandidates(payload),
  )
}

module.exports = {
  registerKnowledgeIpc,
}
