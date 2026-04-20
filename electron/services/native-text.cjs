const path = require('path')
const { runSwiftJson } = require('./macos-utils.cjs')

function createNativeTextService(options = {}) {
  const { app } = options
  const scriptPath = path.join(__dirname, '..', 'native', 'macos_text_bridge.swift')

  return {
    async measure(payload = {}) {
      if (process.platform !== 'darwin') {
        return { success: false, error: 'text-measure-native 仅支持 macOS' }
      }

      try {
        const result = runSwiftJson({
          scriptPath,
          payload,
          timeoutMs: 30000,
        })
        return result
      } catch (error) {
        return { success: false, error: error?.message || String(error) }
      }
    },
  }
}

module.exports = {
  createNativeTextService,
}
