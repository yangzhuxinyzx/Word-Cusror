const path = require('path')
const { spawnSync } = require('child_process')
const sharp = require('sharp')
const { ensureDir, runSwiftJson, stableHash } = require('./macos-utils.cjs')

function createWordOracleService(options = {}) {
  const { fs, path: pathLib, app, docxInspectorService } = options
  const pdfBridgePath = path.join(__dirname, '..', 'native', 'macos_pdf_bridge.swift')
  const textBridgePath = path.join(__dirname, '..', 'native', 'macos_text_bridge.swift')

  const toArray = (value) => {
    if (Array.isArray(value)) return value
    if (value == null) return []
    return [value]
  }

  const findWordApp = () => {
    if (process.platform !== 'darwin') {
      return { installed: false, reason: 'Word Oracle 仅支持 macOS' }
    }

    const result = spawnSync('/usr/bin/osascript', [
      '-e',
      'try',
      '-e',
      'POSIX path of (path to application id "com.microsoft.Word")',
      '-e',
      'on error',
      '-e',
      'return ""',
      '-e',
      'end try',
    ], {
      encoding: 'utf8',
      timeout: 8000,
    })

    const appPath = (result.stdout || '').trim()
    if (!appPath) {
      return { installed: false, reason: '未安装 Microsoft Word for Mac' }
    }

    return { installed: true, appPath }
  }

  const checkFonts = (referencedFonts = []) => {
    const fontCandidates = []
    referencedFonts.forEach((item) => {
      const candidates = [item.name, ...(item.alternates || [])].filter(Boolean)
      candidates.forEach((candidate) => {
        if (!fontCandidates.includes(candidate)) {
          fontCandidates.push(candidate)
        }
      })
    })

    const response = runSwiftJson({
      scriptPath: textBridgePath,
      payload: {
        mode: 'font-check',
        fonts: fontCandidates,
      },
      timeoutMs: 30000,
    })

    const availabilityMap = new Map((response.fonts || []).map((item) => [item.name, item]))
    const missingFonts = []

    referencedFonts.forEach((item) => {
      const candidates = [item.name, ...(item.alternates || [])].filter(Boolean)
      const resolved = candidates
        .map((candidate) => availabilityMap.get(candidate))
        .find((candidate) => candidate?.available)

      if (!resolved) {
        missingFonts.push({
          name: item.name,
          alternates: item.alternates || [],
        })
      }
    })

    return missingFonts
  }

  const runAppleScript = ({ filePath, pdfPath, refreshFields }) => {
    const script = `
on run argv
  set inputPath to item 1 of argv
  set outputPath to item 2 of argv
  set shouldRefresh to item 3 of argv
  set docAlias to POSIX file inputPath
  tell application "Microsoft Word"
    activate
    set activeDoc to open docAlias
  end tell
  delay 1
  if shouldRefresh is "true" then
    try
      tell application "System Events"
        tell process "Microsoft Word"
          keystroke "a" using {command down}
          delay 0.2
          keystroke "u" using {command down, shift down}
        end tell
      end tell
      delay 1
    end try
  end if
  tell application "Microsoft Word"
    save as activeDoc file name outputPath file format format PDF
    close activeDoc saving no
  end tell
  return outputPath
end run
`.trim()

    const result = spawnSync('/usr/bin/osascript', ['-l', 'AppleScript', '-e', script, filePath, pdfPath, refreshFields ? 'true' : 'false'], {
      encoding: 'utf8',
      timeout: 120000,
      maxBuffer: 16 * 1024 * 1024,
    })

    if (result.error) throw result.error
    if (result.status !== 0) {
      const stderr = (result.stderr || '').trim()
      throw new Error(stderr || 'Word AppleScript 导出失败')
    }

    if (!fs.existsSync(pdfPath)) {
      throw new Error('Word 未生成 PDF 输出')
    }
  }

  const renderPdfPages = ({ pdfPath, outputDir, dpi }) => {
    const result = runSwiftJson({
      scriptPath: pdfBridgePath,
      payload: {
        pdfPath,
        outputDir,
        dpi,
      },
      timeoutMs: 120000,
    })

    if (!result.success) {
      throw new Error(result.error || 'PDF 渲染失败')
    }

    return result.pages || []
  }

  const buildArtifactDir = (filePath) => {
    const stats = fs.statSync(filePath)
    const root = path.join(app.getPath('temp'), 'word-cursor', 'word-oracle')
    ensureDir(fs, root)
    const stamp = stableHash(`${filePath}:${stats.size}:${stats.mtimeMs}:${Date.now()}`)
    const dir = path.join(root, stamp)
    ensureDir(fs, dir)
    return { dir, exportId: stamp }
  }

  const dataUrlToBuffer = (dataUrl) => {
    const match = String(dataUrl || '').match(/^data:image\/png;base64,(.+)$/)
    if (!match) {
      throw new Error('当前页面截图不是有效的 PNG data URL')
    }
    return Buffer.from(match[1], 'base64')
  }

  const comparePair = async ({ oraclePath, currentInput, diffPath, thresholdRatio }) => {
    const oracleMeta = await sharp(oraclePath).metadata()
    const currentMeta = currentInput.path
      ? await sharp(currentInput.path).metadata()
      : await sharp(dataUrlToBuffer(currentInput.dataUrl)).metadata()

    const width = Math.max(oracleMeta.width || 0, currentMeta.width || 0)
    const height = Math.max(oracleMeta.height || 0, currentMeta.height || 0)
    if (!width || !height) {
      throw new Error('无法比较空图像')
    }

    const prepare = async (inputPathOrBuffer, metadata) => {
      const source = typeof inputPathOrBuffer === 'string'
        ? sharp(inputPathOrBuffer)
        : sharp(inputPathOrBuffer)
      const buffer = await source.ensureAlpha().png().toBuffer()
      return sharp({
        create: {
          width,
          height,
          channels: 4,
          background: { r: 255, g: 255, b: 255, alpha: 1 },
        },
      })
        .composite([{ input: buffer, left: 0, top: 0 }])
        .raw()
        .toBuffer({ resolveWithObject: true })
    }

    const oraclePrepared = await prepare(oraclePath, oracleMeta)
    const currentPrepared = await prepare(currentInput.path ? currentInput.path : dataUrlToBuffer(currentInput.dataUrl), currentMeta)

    const diff = Buffer.alloc(width * height * 4)
    let mismatchPixels = 0
    let minX = width
    let minY = height
    let maxX = -1
    let maxY = -1

    for (let index = 0; index < diff.length; index += 4) {
      const dr = Math.abs(oraclePrepared.data[index] - currentPrepared.data[index])
      const dg = Math.abs(oraclePrepared.data[index + 1] - currentPrepared.data[index + 1])
      const db = Math.abs(oraclePrepared.data[index + 2] - currentPrepared.data[index + 2])
      const da = Math.abs(oraclePrepared.data[index + 3] - currentPrepared.data[index + 3])
      const mismatch = Math.max(dr, dg, db, da) > 24

      if (mismatch) {
        mismatchPixels += 1
        const pixelIndex = index / 4
        const x = pixelIndex % width
        const y = Math.floor(pixelIndex / width)
        minX = Math.min(minX, x)
        minY = Math.min(minY, y)
        maxX = Math.max(maxX, x)
        maxY = Math.max(maxY, y)
        diff[index] = 255
        diff[index + 1] = 59
        diff[index + 2] = 48
        diff[index + 3] = 255
      } else {
        diff[index] = Math.round(oraclePrepared.data[index] * 0.7 + 255 * 0.3)
        diff[index + 1] = Math.round(oraclePrepared.data[index + 1] * 0.7 + 255 * 0.3)
        diff[index + 2] = Math.round(oraclePrepared.data[index + 2] * 0.7 + 255 * 0.3)
        diff[index + 3] = 255
      }
    }

    await sharp(diff, {
      raw: {
        width,
        height,
        channels: 4,
      },
    }).png().toFile(diffPath)

    const mismatchRatio = Number((mismatchPixels / (width * height)).toFixed(6))
    return {
      mismatchPixels,
      mismatchRatio,
      thresholdExceeded: mismatchRatio > thresholdRatio,
      oracleSize: {
        width: oracleMeta.width || width,
        height: oracleMeta.height || height,
      },
      currentSize: {
        width: currentMeta.width || width,
        height: currentMeta.height || height,
      },
      bbox: maxX >= 0
        ? {
            x: minX,
            y: minY,
            width: maxX - minX + 1,
            height: maxY - minY + 1,
          }
        : null,
    }
  }

  return {
    async export(options = {}) {
      try {
        const filePath = options.filePath
        if (!filePath || !fs.existsSync(filePath)) {
          return { success: false, error: '源文件不存在' }
        }
        if (pathLib.extname(filePath).toLowerCase() !== '.docx') {
          return { success: false, error: 'word-oracle-export 仅支持 .docx 文件' }
        }

        const appInfo = findWordApp()
        if (!appInfo.installed) {
          return { success: false, unavailableReason: appInfo.reason, error: appInfo.reason }
        }

        const inspection = docxInspectorService
          ? await docxInspectorService.inspect(filePath)
          : { success: false }
        const referencedFonts = inspection.success
          ? inspection.report.summary.referencedFonts
          : []

        const missingFonts = checkFonts(referencedFonts)
        if (missingFonts.length > 0) {
          return {
            success: false,
            unavailableReason: 'missing-fonts',
            error: `缺少关键字体：${missingFonts.map((item) => item.name).join('、')}`,
          }
        }

        const { dir, exportId } = buildArtifactDir(filePath)
        const pdfPath = path.join(dir, `${pathLib.basename(filePath, '.docx')}.oracle.pdf`)
        const imageDir = path.join(dir, 'pages')
        ensureDir(fs, imageDir)

        runAppleScript({
          filePath,
          pdfPath,
          refreshFields: options.refreshFields !== false,
        })

        const pages = renderPdfPages({
          pdfPath,
          outputDir: imageDir,
          dpi: options.dpi || 144,
        })

        return {
          success: true,
          artifact: {
            exportId,
            sourcePath: filePath,
            pdfPath,
            imageDir,
            pageCount: pages.length,
            pages,
            exportedAt: new Date().toISOString(),
            wordAppPath: appInfo.appPath,
            inspectorExtractedDir: inspection.success ? inspection.report.extractedDir : undefined,
            missingFonts,
          },
        }
      } catch (error) {
        return { success: false, error: error?.message || String(error) }
      }
    },

    async diff(options = {}) {
      try {
        const oraclePages = toArray(options.oraclePages)
        const currentPages = toArray(options.currentPages)
        const thresholdRatio = Number(options.thresholdRatio || 0.0025)
        const diffDir = path.join(app.getPath('temp'), 'word-cursor', 'word-oracle-diff', stableHash(`${options.sourcePath}:${options.artifactId}:${Date.now()}`))
        ensureDir(fs, diffDir)

        const pageCount = Math.max(oraclePages.length, currentPages.length)
        const pages = []

        for (let index = 0; index < pageCount; index += 1) {
          const oraclePage = oraclePages[index]
          const currentPage = currentPages[index]
          if (!oraclePage || !currentPage) {
            continue
          }

          const diffPath = path.join(diffDir, `page-${String(index + 1).padStart(3, '0')}.diff.png`)
          const diffResult = await comparePair({
            oraclePath: oraclePage.path,
            currentInput: currentPage,
            diffPath,
            thresholdRatio,
          })

          pages.push({
            pageIndex: oraclePage.pageIndex || index + 1,
            oracleImagePath: oraclePage.path,
            diffImagePath: diffPath,
            thresholdRatio,
            ...diffResult,
          })
        }

        const currentPageIndicesOverThreshold = pages
          .filter((page) => page.thresholdExceeded)
          .map((page) => page.pageIndex)

        const report = {
          artifactId: options.artifactId,
          sourcePath: options.sourcePath,
          createdAt: new Date().toISOString(),
          expectedPageCount: oraclePages.length,
          actualPageCount: currentPages.length,
          pageCountMatches: oraclePages.length === currentPages.length,
          thresholdRatio,
          mismatchCount: currentPageIndicesOverThreshold.length,
          pages,
          currentPageIndicesOverThreshold,
          status: currentPageIndicesOverThreshold.length === 0 && oraclePages.length === currentPages.length
            ? 'aligned'
            : 'misaligned',
        }

        return { success: true, report }
      } catch (error) {
        return { success: false, error: error?.message || String(error) }
      }
    },
  }
}

module.exports = {
  createWordOracleService,
}
