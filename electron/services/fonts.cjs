function createFontsService(options) {
  const { fs, path } = options

  const supportedExts = ['.ttf', '.otf', '.ttc', '.woff', '.woff2']

  const getFontsDir = () => path.join(__dirname, '..', '..', 'Fonts')

  return {
    async listFonts() {
      try {
        const fontsDir = getFontsDir()
        if (!fs.existsSync(fontsDir)) {
          return { success: true, fonts: [] }
        }

        const entries = fs.readdirSync(fontsDir, { withFileTypes: true })
        const fontFiles = []

        for (const entry of entries) {
          if (!entry.isFile()) continue
          const ext = path.extname(entry.name).toLowerCase()
          if (!supportedExts.includes(ext)) continue
          fontFiles.push({
            name: entry.name,
            ext,
            size: fs.statSync(path.join(fontsDir, entry.name)).size,
          })
        }

        return { success: true, fonts: fontFiles }
      } catch (error) {
        return { success: false, error: error.message, fonts: [] }
      }
    },

    async readFont(fileName) {
      let safeName = ''
      try {
        const fontsDir = getFontsDir()
        safeName = path.basename(String(fileName || ''))
        if (!safeName) {
          return { success: false, error: 'Invalid font file name' }
        }

        const filePath = path.join(fontsDir, safeName)
        if (!filePath.startsWith(fontsDir)) {
          return { success: false, error: 'Invalid font path' }
        }
        if (!fs.existsSync(filePath)) {
          return { success: false, error: `Font file not found: ${safeName}` }
        }

        const buffer = fs.readFileSync(filePath)
        return {
          success: true,
          base64: buffer.toString('base64'),
          size: buffer.length,
        }
      } catch (error) {
        return { success: false, error: error.message || `Failed to read font: ${safeName}` }
      }
    },
  }
}

module.exports = {
  createFontsService,
}
