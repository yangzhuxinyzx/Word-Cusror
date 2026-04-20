import { spawn } from 'node:child_process'
import fs from 'node:fs/promises'
import path from 'node:path'
import process from 'node:process'
import { fileURLToPath } from 'node:url'

import { chromium } from 'playwright'
import sharp from 'sharp'

const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)
const repoRoot = path.resolve(__dirname, '..', '..')

function parseArgs(argv) {
  const args = {}
  for (let i = 0; i < argv.length; i += 1) {
    const token = argv[i]
    if (!token.startsWith('--')) continue
    const key = token.slice(2)
    const next = argv[i + 1]
    if (!next || next.startsWith('--')) {
      args[key] = 'true'
      continue
    }
    args[key] = next
    i += 1
  }
  return args
}

function resolveFixtureArg(value) {
  const raw = (value || 'render-lab-sample-cn.docx').trim()
  if (!raw) return 'render-lab-sample-cn.docx'
  return raw
}

function buildRenderLabUrl(baseUrl, fixture) {
  const url = new URL(baseUrl)
  url.searchParams.set('render-lab', '1')
  url.searchParams.set('fixture', fixture)
  url.searchParams.set('render-mode', 'tiptap')
  return url.toString()
}

async function waitForHttpReady(url, timeoutMs = 30000) {
  const deadline = Date.now() + timeoutMs
  let lastError = null

  while (Date.now() < deadline) {
    try {
      const response = await fetch(url)
      if (response.ok) return
      lastError = new Error(`HTTP ${response.status}`)
    } catch (error) {
      lastError = error
    }
    await new Promise((resolve) => setTimeout(resolve, 500))
  }

  throw new Error(`等待开发服务器超时：${url} ${lastError ? `(${String(lastError)})` : ''}`)
}

async function ensureDir(dirPath) {
  await fs.mkdir(dirPath, { recursive: true })
}

function nowStamp() {
  return new Date().toISOString().replace(/[:.]/g, '-')
}

async function startViteServer(port) {
  const viteBin = path.join(repoRoot, 'node_modules', 'vite', 'bin', 'vite.js')
  const child = spawn(process.execPath, [viteBin, '--host', 'localhost', '--port', String(port)], {
    cwd: repoRoot,
    env: {
      ...process.env,
      BROWSER: 'none',
    },
    stdio: 'inherit',
  })

  const baseUrl = `http://localhost:${port}`
  await waitForHttpReady(baseUrl, 30000)
  return { child, baseUrl }
}

async function renderOnWhiteCanvas(filePath, width, height) {
  const input = await sharp(filePath).ensureAlpha().png().toBuffer()
  const metadata = await sharp(input).metadata()
  const canvas = sharp({
    create: {
      width,
      height,
      channels: 4,
      background: { r: 255, g: 255, b: 255, alpha: 1 },
    },
  })

  const left = 0
  const top = 0
  const composed = await canvas
    .composite([{ input, left, top }])
    .raw()
    .toBuffer({ resolveWithObject: true })

  return {
    data: composed.data,
    width: metadata.width || width,
    height: metadata.height || height,
  }
}

async function compareImages(actualPath, referencePath, diffPath) {
  const actualMeta = await sharp(actualPath).metadata()
  const referenceMeta = await sharp(referencePath).metadata()
  const width = Math.max(actualMeta.width || 0, referenceMeta.width || 0)
  const height = Math.max(actualMeta.height || 0, referenceMeta.height || 0)

  if (!width || !height) {
    throw new Error('无法读取截图尺寸，不能生成 diff')
  }

  const actual = await renderOnWhiteCanvas(actualPath, width, height)
  const reference = await renderOnWhiteCanvas(referencePath, width, height)
  const diff = Buffer.alloc(width * height * 4)
  const threshold = 24
  let mismatchPixels = 0

  for (let i = 0; i < diff.length; i += 4) {
    const dr = Math.abs(actual.data[i] - reference.data[i])
    const dg = Math.abs(actual.data[i + 1] - reference.data[i + 1])
    const db = Math.abs(actual.data[i + 2] - reference.data[i + 2])
    const da = Math.abs(actual.data[i + 3] - reference.data[i + 3])
    const mismatch = Math.max(dr, dg, db, da) > threshold

    if (mismatch) {
      mismatchPixels += 1
      diff[i] = 255
      diff[i + 1] = 59
      diff[i + 2] = 48
      diff[i + 3] = 255
      continue
    }

    diff[i] = Math.round(actual.data[i] * 0.65 + 255 * 0.35)
    diff[i + 1] = Math.round(actual.data[i + 1] * 0.65 + 255 * 0.35)
    diff[i + 2] = Math.round(actual.data[i + 2] * 0.65 + 255 * 0.35)
    diff[i + 3] = 255
  }

  await sharp(diff, {
    raw: {
      width,
      height,
      channels: 4,
    },
  }).png().toFile(diffPath)

  return {
    width,
    height,
    mismatchPixels,
    mismatchRatio: Number((mismatchPixels / (width * height)).toFixed(6)),
    actualSize: {
      width: actualMeta.width || 0,
      height: actualMeta.height || 0,
    },
    referenceSize: {
      width: referenceMeta.width || 0,
      height: referenceMeta.height || 0,
    },
  }
}

async function screenshotWithRetry(page, selector, outputPath, attempts = 4) {
  let lastError = null

  for (let attempt = 1; attempt <= attempts; attempt += 1) {
    try {
      const canvasDataUrl = await page.evaluate((nextSelector) => {
        const element = document.querySelector(nextSelector)
        if (!(element instanceof HTMLCanvasElement)) return null
        return element.toDataURL('image/png')
      }, selector)
      if (canvasDataUrl) {
        const base64 = canvasDataUrl.replace(/^data:image\/png;base64,/, '')
        await fs.writeFile(outputPath, Buffer.from(base64, 'base64'))
        return
      }

      const clip = await page.evaluate((nextSelector) => {
        const element = document.querySelector(nextSelector)
        if (!element) return null
        element.scrollIntoView({ block: 'start', inline: 'nearest' })
        const rect = element.getBoundingClientRect()
        return {
          x: Math.max(0, Math.floor(rect.left + window.scrollX)),
          y: Math.max(0, Math.floor(rect.top + window.scrollY)),
          width: Math.max(1, Math.ceil(rect.width)),
          height: Math.max(1, Math.ceil(rect.height)),
        }
      }, selector)
      if (!clip) {
        throw new Error(`未找到截图目标: ${selector}`)
      }
      await page.waitForTimeout(300)
      await page.screenshot({ path: outputPath, clip })
      return
    } catch (error) {
      lastError = error
      if (attempt === attempts) break
      await page.waitForTimeout(400)
    }
  }

  throw lastError || new Error(`截图失败: ${selector}`)
}

async function resolveCaptureTarget(page, pageNumber) {
  return page.evaluate((nextPageNumber) => {
    const toPx = (rawValue, fallback = 0) => {
      const value = String(rawValue || '').trim()
      if (!value) return fallback
      const number = Number.parseFloat(value)
      if (!Number.isFinite(number)) return fallback
      if (value.endsWith('mm')) return number * 96 / 25.4
      if (value.endsWith('cm')) return number * 96 / 2.54
      if (value.endsWith('in')) return number * 96
      if (value.endsWith('pt')) return number * 96 / 72
      return number
    }

    const editorDebug = window.__wordCursorWordEditorDebug || null
    if (editorDebug?.printRenderMode === 'canvas') {
      return {
        kind: 'selector',
        selector:
          nextPageNumber === 'all'
            ? '[data-testid="word-render-canvas-preview"]'
            : `[data-testid="word-render-page-${nextPageNumber}"] canvas`,
      }
    }

    const pageEl = document.querySelector('[data-testid="word-render-tiptap-page"]')
    if (!(pageEl instanceof HTMLElement)) return null

    const rect = pageEl.getBoundingClientRect()
    const computed = window.getComputedStyle(pageEl)
    const pageHeight = Math.max(1, Math.ceil(toPx(computed.getPropertyValue('--page-height'), rect.height)))
    const x = Math.max(0, Math.floor(rect.left + window.scrollX))
    const width = Math.max(1, Math.ceil(rect.width))

    if (nextPageNumber === 'all') {
      return {
        kind: 'clip',
        clip: {
          x,
          y: Math.max(0, Math.floor(rect.top + window.scrollY)),
          width,
          height: Math.max(1, Math.ceil(rect.height)),
        },
      }
    }

    const pageIndex = Math.max(1, Number(nextPageNumber))
    let y = Math.max(0, Math.floor(rect.top + window.scrollY))

    if (pageIndex > 1) {
      const breaks = Array.from(document.querySelectorAll('.word-editor-content .pm-page-break'))
      const breakEl = breaks[pageIndex - 2]
      if (!(breakEl instanceof HTMLElement)) return null
      const breakRect = breakEl.getBoundingClientRect()
      const breakComputed = window.getComputedStyle(breakEl)
      const fill = toPx(breakComputed.getPropertyValue('--fill'), 0)
      const gap = toPx(breakComputed.getPropertyValue('--gap'), 0)
      y = Math.max(0, Math.floor(breakRect.top + window.scrollY + fill + gap))
    }

    return {
      kind: 'clip',
      clip: {
        x,
        y,
        width,
        height: pageHeight,
      },
    }
  }, pageNumber)
}

async function main() {
  const args = parseArgs(process.argv.slice(2))
  const port = Number(args.port || 3000)
  const baseUrl = args.url || `http://localhost:${port}`
  const fixture = resolveFixtureArg(args.fixture)
  const pageNumber = args.page === 'all' ? 'all' : String(Math.max(1, Number(args.page || 1)))
  const runDir = path.join(repoRoot, 'outputs', 'render-lab', nowStamp())
  const screenshotName = pageNumber === 'all' ? 'render-lab-full.png' : `render-lab-page-${pageNumber}.png`
  const screenshotPath = path.resolve(args.output || path.join(runDir, screenshotName))
  const metadataPath = screenshotPath.replace(/\.png$/i, '.json')
  const diffPath = screenshotPath.replace(/\.png$/i, '.diff.png')
  const referencePath = args.reference ? path.resolve(args.reference) : null

  let serverProcess = null

  try {
    await ensureDir(path.dirname(screenshotPath))

    if (args['no-server'] !== 'true' && !args.url) {
      try {
        await waitForHttpReady(baseUrl, 1500)
      } catch {
        const server = await startViteServer(port)
        serverProcess = server.child
      }
    } else {
      await waitForHttpReady(baseUrl, 30000)
    }

    const browser = await chromium.launch({
      headless: true,
    })

    try {
      const page = await browser.newPage({
        viewport: { width: 1440, height: 1800 },
        deviceScaleFactor: 1,
      })

      const targetUrl = buildRenderLabUrl(baseUrl, fixture)
      await page.goto(targetUrl, { waitUntil: 'domcontentloaded' })
      await page.waitForSelector('[data-testid="render-lab-root"]', { timeout: 30000 })
      const snapshot = await page.evaluate(async () => {
        if (!window.__wordCursorRenderLab?.waitUntilReady) {
          throw new Error('window.__wordCursorRenderLab.waitUntilReady 不可用')
        }
        await window.__wordCursorRenderLab.waitUntilReady(30000)
        return {
          renderLab: window.__wordCursorRenderLab?.getSnapshot?.() || null,
          layout: window.__wordCursorLayoutDebug || null,
          editor: (window).__wordCursorWordEditorDebug || null,
        }
      })

      await page.waitForFunction(() => {
        const text = document.body?.innerText || ''
        return !text.includes('正在渲染页面...') && !text.includes('正在加载文档...')
      }, { timeout: 30000 })

      let captureTarget = await resolveCaptureTarget(page, pageNumber)
      if (!captureTarget) {
        const requestedPage = pageNumber === 'all' ? null : Number(pageNumber)
        if (snapshot.editor?.printRenderMode === 'tiptap' && requestedPage && requestedPage > 1) {
          await page.waitForFunction((nextPage) => {
            return document.querySelectorAll('.word-editor-content .pm-page-break').length >= nextPage - 1
          }, requestedPage, { timeout: 30000 })
        }

        captureTarget = await resolveCaptureTarget(page, pageNumber)
        if (!captureTarget) {
          throw new Error(`无法确定截图目标（page=${pageNumber}）`)
        }
      }

      await page.waitForTimeout(800)
      if (captureTarget.kind === 'selector') {
        await page.waitForFunction(() => {
          const preview = document.querySelector('[data-testid="word-render-canvas-preview"]')
          return preview?.getAttribute('data-render-status') === 'ready'
        }, { timeout: 30000 })
        await page.waitForSelector(captureTarget.selector, { timeout: 30000 })
        await screenshotWithRetry(page, captureTarget.selector, screenshotPath)
      } else {
        let activeTarget = captureTarget
        const viewport = page.viewportSize()
        const requiredHeight = Math.ceil(activeTarget.clip.y + activeTarget.clip.height + 120)
        if (viewport && requiredHeight > viewport.height) {
          await page.setViewportSize({
            width: viewport.width,
            height: Math.min(requiredHeight, 12000),
          })
          await page.waitForTimeout(300)
          const resizedTarget = await resolveCaptureTarget(page, pageNumber)
          if (!resizedTarget || resizedTarget.kind !== 'clip') {
            throw new Error(`视口调整后无法确定截图目标（page=${pageNumber}）`)
          }
          activeTarget = resizedTarget
        }
        await page.screenshot({ path: screenshotPath, clip: activeTarget.clip })
      }

      const metadata = {
        createdAt: new Date().toISOString(),
        targetUrl,
        fixture,
        captureTarget,
        output: path.relative(repoRoot, screenshotPath),
        snapshot,
      }

      if (referencePath) {
        const diff = await compareImages(screenshotPath, referencePath, diffPath)
        metadata.reference = path.relative(repoRoot, referencePath)
        metadata.diff = {
          output: path.relative(repoRoot, diffPath),
          ...diff,
        }
      }

      await fs.writeFile(metadataPath, JSON.stringify(metadata, null, 2), 'utf8')
      console.log(`[render-lab] screenshot saved: ${path.relative(repoRoot, screenshotPath)}`)
      console.log(`[render-lab] metadata saved: ${path.relative(repoRoot, metadataPath)}`)
      if (referencePath) {
        console.log(`[render-lab] diff saved: ${path.relative(repoRoot, diffPath)}`)
      }
    } finally {
      await browser.close()
    }
  } finally {
    if (serverProcess && !serverProcess.killed) {
      serverProcess.kill('SIGTERM')
    }
  }
}

main().catch((error) => {
  console.error('[render-lab] capture failed:', error)
  process.exitCode = 1
})
