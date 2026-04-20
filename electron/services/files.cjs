function escapeHtml(text) {
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;')
}

function createFilesService(options) {
  const {
    fs,
    path,
    dialog,
    shell,
    WordExtractor,
    getMainWindow,
    getFileServerPort,
  } = options

  const supportedExts = [
    '.docx',
    '.doc',
    '.txt',
    '.md',
    '.json',
    '.xml',
    '.xlsx',
    '.xls',
    '.pptx',
    '.ppt',
  ]

  const ignoredDirNames = new Set([
    '.git',
    'node_modules',
    'dist',
    'release',
    'outputs',
    'logs',
    'docx_temp',
    '_rels',
    'docProps',
    'customXml',
  ])

  const ensureDir = (filePath) => {
    const dir = path.dirname(filePath)
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true })
    }
  }

  const buildFileUrl = (filePath, host) => {
    const encodedPath = encodeURIComponent(filePath.replace(/\\/g, '/'))
    return `http://${host}:${getFileServerPort()}/file/${encodedPath}`
  }

  const readFolderRecursive = async (basePath, currentPath, depth = 0) => {
    if (depth > 5) return []

    const items = []
    const entries = fs.readdirSync(currentPath, { withFileTypes: true })

    for (const entry of entries) {
      if (entry.name.startsWith('.') || ignoredDirNames.has(entry.name)) continue

      const fullPath = path.join(currentPath, entry.name)
      const relativePath = path.relative(basePath, fullPath)

      if (entry.isDirectory()) {
        const children = await readFolderRecursive(basePath, fullPath, depth + 1)
        if (!children.length) continue
        items.push({
          name: entry.name,
          path: fullPath,
          relativePath,
          type: 'folder',
          children,
        })
      } else {
        const ext = path.extname(entry.name).toLowerCase()
        if (supportedExts.includes(ext)) {
          items.push({
            name: entry.name,
            path: fullPath,
            relativePath,
            type: 'file',
            extension: ext,
          })
        }
      }
    }

    items.sort((a, b) => {
      if (a.type !== b.type) return a.type === 'folder' ? -1 : 1
      return a.name.localeCompare(b.name)
    })

    return items
  }

  return {
    async getFileUrl(filePath) {
      return buildFileUrl(filePath, 'host.docker.internal')
    },

    async getLocalFileUrl(filePath) {
      return buildFileUrl(filePath, 'localhost')
    },

    async selectFolder() {
      const result = await dialog.showOpenDialog(getMainWindow?.(), {
        properties: ['openDirectory'],
        title: '选择工作文件夹',
      })

      if (result.canceled) return null
      return result.filePaths[0]
    },

    async readFolder(folderPath) {
      try {
        const items = await readFolderRecursive(folderPath, folderPath)
        return { success: true, data: items }
      } catch (error) {
        return { success: false, error: error.message }
      }
    },

    async readFile(filePath) {
      try {
        const ext = path.extname(filePath).toLowerCase()
        const fileName = path.basename(filePath)

        if (fileName.startsWith('~$')) {
          return {
            success: true,
            data: '<p style="text-align: center; color: #888; padding: 40px;">这是一个 Word 临时文件，无法打开。</p>',
            type: 'html',
          }
        }

        if (ext === '.docx') {
          const buffer = fs.readFileSync(filePath)
          return { success: true, data: buffer.toString('base64'), type: 'docx' }
        }

        if (ext === '.pptx') {
          const buffer = fs.readFileSync(filePath)
          return { success: true, data: buffer.toString('base64'), type: 'pptx' }
        }

        if (ext === '.doc') {
          try {
            const extractor = new WordExtractor()
            const extracted = await extractor.extract(filePath)
            const body = extracted.getBody() || ''

            let html = ''
            const paragraphs = body.split(/\n\n+/)
            for (const para of paragraphs) {
              const trimmed = para.trim()
              if (!trimmed) continue
              const lines = trimmed.split(/\n/)
              const formattedPara = lines.map((line) => escapeHtml(line)).join('<br>')
              html += `<p>${formattedPara}</p>`
            }

            if (!html) {
              html = '<p></p>'
            }

            return { success: true, data: html, type: 'doc-html' }
          } catch (extractorError) {
            return {
              success: true,
              data: `<div style="padding: 40px; text-align: center; color: #888;">
            <p style="font-size: 18px; margin-bottom: 10px;">⚠️ 无法解析此 .doc 文件</p>
            <p style="font-size: 14px;">此文件可能已损坏或使用了不支持的格式。</p>
            <p style="font-size: 12px; margin-top: 15px; color: #666;">
              建议：使用 Microsoft Word 打开此文件，然后另存为 .docx 格式。
            </p>
          </div>`,
              type: 'doc-html',
            }
          }
        }

        const content = fs.readFileSync(filePath, 'utf-8')
        return { success: true, data: content, type: 'text' }
      } catch (error) {
        return { success: false, error: error.message }
      }
    },

    async writeFile(filePath, content) {
      try {
        ensureDir(filePath)
        fs.writeFileSync(filePath, content, 'utf-8')
        return { success: true }
      } catch (error) {
        return { success: false, error: error.message }
      }
    },

    async appendFile(filePath, content) {
      try {
        ensureDir(filePath)
        fs.appendFileSync(filePath, content, 'utf-8')
        return { success: true }
      } catch (error) {
        return { success: false, error: error.message }
      }
    },

    async writeBinaryFile(filePath, base64Data) {
      try {
        ensureDir(filePath)
        const buffer = Buffer.from(base64Data, 'base64')
        fs.writeFileSync(filePath, buffer)
        return { success: true }
      } catch (error) {
        return { success: false, error: error.message }
      }
    },

    async saveFileDialog(defaultName) {
      const result = await dialog.showSaveDialog(getMainWindow?.(), {
        defaultPath: defaultName,
        filters: [
          { name: 'Word 文档', extensions: ['docx'] },
          { name: 'Markdown', extensions: ['md'] },
          { name: '文本文件', extensions: ['txt'] },
          { name: '所有文件', extensions: ['*'] },
        ],
      })

      if (result.canceled) return null
      return result.filePath
    },

    async createFile(folderPath, fileName, content = '') {
      try {
        const filePath = path.join(folderPath, fileName)
        ensureDir(filePath)
        fs.writeFileSync(filePath, content, 'utf-8')
        return { success: true, path: filePath }
      } catch (error) {
        return { success: false, error: error.message }
      }
    },

    async deleteFile(filePath) {
      try {
        fs.unlinkSync(filePath)
        return { success: true }
      } catch (error) {
        return { success: false, error: error.message }
      }
    },

    async renameFile(oldPath, newPath) {
      try {
        ensureDir(newPath)
        fs.renameSync(oldPath, newPath)
        return { success: true }
      } catch (error) {
        return { success: false, error: error.message }
      }
    },

    async showInFolder(filePath) {
      shell.showItemInFolder(filePath)
      return { success: true }
    },

    async getFileInfo(filePath) {
      try {
        const stats = fs.statSync(filePath)
        return {
          success: true,
          data: {
            size: stats.size,
            created: stats.birthtime,
            modified: stats.mtime,
            isFile: stats.isFile(),
            isDirectory: stats.isDirectory(),
          },
        }
      } catch (error) {
        return { success: false, error: error.message }
      }
    },
  }
}

module.exports = {
  createFilesService,
}
