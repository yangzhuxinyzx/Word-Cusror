const crypto = require('crypto')
const http = require('http')
const https = require('https')
const { BrowserWindow } = require('electron')
const sharp = require('sharp')
const PptxGenJS = require('pptxgenjs')
const { createPptTextSidecarService } = require('./ppt-text-sidecar.cjs')

function pLimit(concurrency) {
  const queue = []
  let activeCount = 0

  const next = () => {
    activeCount -= 1
    if (queue.length > 0) {
      const run = queue.shift()
      run()
    }
  }

  return (fn) =>
    new Promise((resolve, reject) => {
      const run = () => {
        activeCount += 1
        Promise.resolve()
          .then(fn)
          .then(resolve, reject)
          .finally(next)
      }

      if (activeCount < concurrency) {
        run()
      } else {
        queue.push(run)
      }
    })
}

function createPptService(options = {}) {
  const {
    fs,
    path,
    app,
    findLibreOffice,
    getLibreOfficeDownloadUrl,
  } = options
  const textSidecar = createPptTextSidecarService({ fs, path, app })
  let adobeTokenCache = null
  let deterministicRenderWindow = null
  let deterministicRenderWindowReady = false

  function emitPptTextEditProgress(detail = {}) {
    try {
      const windows = BrowserWindow.getAllWindows()
      for (const win of windows) {
        if (!win.isDestroyed()) {
          win.webContents.send('ppt-text-edit-progress', {
            active: true,
            ...detail,
          })
        }
      }
    } catch {}
  }

  function clearPptTextEditProgress(detail = {}) {
    try {
      const windows = BrowserWindow.getAllWindows()
      for (const win of windows) {
        if (!win.isDestroyed()) {
          win.webContents.send('ppt-text-edit-progress', {
            active: false,
            ...detail,
          })
        }
      }
    } catch {}
  }

  function destroyDeterministicRenderWindow() {
    if (deterministicRenderWindow && !deterministicRenderWindow.isDestroyed()) {
      deterministicRenderWindow.destroy()
    }
    deterministicRenderWindow = null
    deterministicRenderWindowReady = false
  }

  app.on('before-quit', () => {
    destroyDeterministicRenderWindow()
  })

  function hashForFileCache(filePath) {
    const st = fs.statSync(filePath)
    const key = `${filePath}|${st.size}|${st.mtimeMs}`
    return crypto.createHash('sha1').update(key).digest('hex')
  }

  function getPptxPreviewCacheDir(filePath) {
    const hash = hashForFileCache(filePath)
    const tempDir = app.getPath('temp')
    return path.join(tempDir, 'word-cursor-ppt-preview', hash)
  }

  function listPngFilesSorted(dir) {
    const files = fs.readdirSync(dir).filter((f) => f.toLowerCase().endsWith('.png'))
    const withMeta = files.map((name) => {
      const m = name.match(/(\d+)(?=\.png$)/)
      const idx = m ? parseInt(m[1], 10) : 0
      return { name, idx }
    })
    withMeta.sort((a, b) => (a.idx - b.idx) || a.name.localeCompare(b.name))
    return withMeta.map((x) => path.join(dir, x.name))
  }

  async function renderPptxToPngsWithLibreOffice(pptxPath, outDir) {
    const libreOfficePath = findLibreOffice()
    if (!libreOfficePath) {
      return { success: false, error: 'LibreOffice 未安装', downloadUrl: getLibreOfficeDownloadUrl() }
    }

    if (!fs.existsSync(outDir)) {
      fs.mkdirSync(outDir, { recursive: true })
    }

    const { execFile } = require('child_process')
    // LibreOffice 将每页导出为 PNG（文件名规则依版本不同，导出后我们扫描目录排序）
    return new Promise((resolve) => {
      execFile(
        libreOfficePath,
        ['--headless', '--nologo', '--nolockcheck', '--norestore', '--convert-to', 'png', '--outdir', outDir, pptxPath],
        { timeout: 180000 },
        (error, stdout, stderr) => {
          if (error) {
            console.error('[PPTX] LibreOffice 转换失败:', error)
            resolve({ success: false, error: 'LibreOffice 转换失败', details: stderr || stdout })
            return
          }
          const pngs = listPngFilesSorted(outDir)
          if (!pngs.length) {
            resolve({ success: false, error: 'LibreOffice 转换未生成 PNG' })
            return
          }
          resolve({ success: true, images: pngs })
        }
      )
    })
  }

  function getDashScopeEndpoint(region = 'cn') {
    return region === 'intl'
      ? 'https://dashscope-intl.aliyuncs.com/api/v1/services/aigc/multimodal-generation/generation'
      : 'https://dashscope.aliyuncs.com/api/v1/services/aigc/multimodal-generation/generation'
  }

  // 负面词基线：用于"去水印/去UI/去乱码/去廉价/去AI味/防字体畸变"
  const NEGATIVE_PROMPT_BASELINE =
    // 防止字体/文字畸变（最重要）
    'deformed text, broken text, malformed letters, illegible text, unreadable text, distorted characters, corrupted text, warped text, melted text, stretched text, squished text, overlapping text, cropped text, cut off text, truncated text, incomplete text, missing letters, extra letters, wrong stroke order, bad stroke, messy strokes, ' +
    // 防止中文乱码/错字
    'garbled Chinese, wrong Chinese characters, simplified-traditional mix, mojibake, wrong characters, misspelling, random letters, gibberish, extra text, unwanted text, english text mixed, ' +
    // 防止排版问题
    'ugly typography, amateur typography, bad kerning, bad tracking, uneven spacing, inconsistent font size, font size mismatch, bad line height, crowded text, text too small, ' +
    // 去水印/去UI/去品牌
    'watermark, logo, brand name, badge, QR code, UI elements, screenshot, buttons, interface, HUD, sci-fi interface, holographic UI, futuristic dashboard, ' +
    // 去廉价科技风
    'neon cyberpunk, neon cyan, bright cyan, fluorescent cyan, neon teal, cheap turquoise, neon glow, laser lines, circuit board, generic isometric city, isometric cityscape, circuit-board city, cheap sci-fi, ' +
    // 去AI味/低质量
    'lowres, low resolution, blurry, jpeg artifacts, compression artifacts, noisy, grainy, pixelated, worst quality, low quality, normal quality, bad quality, amateur, unprofessional, amateur layout, noisy background, oversaturated, cheap plastic, toy-like, glossy, harsh specular, overbloom, stock 3d icons, generic template, ai artifacts, uncanny, artificial looking, cgi looking, ' +
    // 去结构问题
    'bad composition, cluttered, messy layout, unbalanced, asymmetric in bad way, empty space, too much whitespace, boring layout, generic layout'

  function mergeNegativePrompt(userNegativePrompt) {
    const set = new Set()
    const add = (s) => {
      String(s || '')
        .split(',')
        .map((t) => t.trim())
        .filter(Boolean)
        .forEach((t) => set.add(t))
    }
    add(userNegativePrompt)
    add(NEGATIVE_PROMPT_BASELINE)
    return Array.from(set).join(', ')
  }

  function requestJson(urlStr, { method = 'GET', headers = {}, body } = {}) {
    return new Promise((resolve, reject) => {
      const urlObj = new URL(urlStr)
      const isHttps = urlObj.protocol === 'https:'
      const lib = isHttps ? https : http

      const req = lib.request(
        {
          protocol: urlObj.protocol,
          hostname: urlObj.hostname,
          port: urlObj.port || (isHttps ? 443 : 80),
          path: urlObj.pathname + urlObj.search,
          method,
          headers,
        },
        (res) => {
          const chunks = []
          res.on('data', (c) => chunks.push(c))
          res.on('end', () => {
            const text = Buffer.concat(chunks).toString('utf-8')
            resolve({ statusCode: res.statusCode || 0, headers: res.headers, text })
          })
        }
      )
      req.on('error', reject)
      if (body) req.write(body)
      req.end()
    })
  }

  async function downloadToBuffer(urlStr, redirectLeft = 5) {
    const urlObj = new URL(urlStr)
    const lib = urlObj.protocol === 'https:' ? https : http

    return new Promise((resolve, reject) => {
      const req = lib.request(
        {
          protocol: urlObj.protocol,
          hostname: urlObj.hostname,
          port: urlObj.port || (urlObj.protocol === 'https:' ? 443 : 80),
          path: urlObj.pathname + urlObj.search,
          method: 'GET',
          headers: {
            'User-Agent': 'word-cursor/1.0',
          },
        },
        (res) => {
          const status = res.statusCode || 0
          const location = res.headers.location
          if ([301, 302, 303, 307, 308].includes(status) && location && redirectLeft > 0) {
            res.resume()
            const nextUrl = new URL(location, urlStr).toString()
            downloadToBuffer(nextUrl, redirectLeft - 1).then(resolve).catch(reject)
            return
          }
          if (status < 200 || status >= 300) {
            const chunks = []
            res.on('data', (c) => chunks.push(c))
            res.on('end', () => reject(new Error(`下载失败: HTTP ${status} ${Buffer.concat(chunks).toString('utf-8').slice(0, 200)}`)))
            return
          }
          const chunks = []
          res.on('data', (c) => chunks.push(c))
          res.on('end', () => resolve(Buffer.concat(chunks)))
        }
      )
      req.on('error', reject)
      req.end()
    })
  }

  function extractDashScopeImageUrl(json) {
    // sync multimodal-generation format
    const maybe1 = json?.output?.choices?.[0]?.message?.content?.find?.((c) => c?.image)?.image
    if (maybe1) return maybe1
    // async / ImageSynthesis format
    const maybe2 = json?.output?.results?.[0]?.url
    if (maybe2) return maybe2
    // task query format
    const maybe3 = json?.output?.results?.[0]?.url || json?.output?.results?.[0]?.image
    if (maybe3) return maybe3
    return null
  }

  function extractDashScopeTaskId(json) {
    return (
      json?.output?.task_id ||
      json?.output?.taskId ||
      json?.output?.taskID ||
      json?.task_id ||
      json?.taskId ||
      null
    )
  }

  function getDashScopeTaskEndpoint(region = 'cn', taskId) {
    const origin =
      region === 'intl' ? 'https://dashscope-intl.aliyuncs.com' : 'https://dashscope.aliyuncs.com'
    return `${origin}/api/v1/tasks/${encodeURIComponent(String(taskId))}`
  }

  async function dashscopeWaitForImageUrlByTaskId({ taskId, region, apiKey, timeoutMs = 120000 }) {
    const started = Date.now()
    let delay = 800
    let lastText = ''
    while (Date.now() - started < timeoutMs) {
      const endpoint = getDashScopeTaskEndpoint(region, taskId)
      const { statusCode, text } = await requestJson(endpoint, {
        method: 'GET',
        headers: {
          'Content-Type': 'application/json',
          Authorization: `Bearer ${apiKey}`,
        },
      })
      lastText = text
      if (statusCode >= 200 && statusCode < 300) {
        try {
          const json = JSON.parse(text)
          const url = extractDashScopeImageUrl(json)
          if (url) return { url, raw: json }
          const status =
            json?.output?.task_status || json?.output?.taskStatus || json?.output?.status || json?.status
          if (String(status).toUpperCase().includes('FAILED')) {
            const msg = json?.message || json?.output?.message || 'DashScope 任务失败'
            throw new Error(`DashScope 任务失败: ${msg}`)
          }
        } catch {
          // ignore JSON parse errors; will retry
        }
      }
      await new Promise((r) => setTimeout(r, delay))
      delay = Math.min(Math.floor(delay * 1.35), 5000)
    }
    throw new Error(`DashScope 异步任务超时，taskId=${taskId}，last=${String(lastText).slice(0, 200)}`)
  }

  async function dashscopeGenerateImageUrl({
    prompt,
    negativePrompt = '',
    size = '2048*1152',
    promptExtend = false,
    watermark = false,
    model = 'z-image-turbo',
    region = 'cn',
    apiKey: apiKeyOverride,
  }) {
    const apiKey =
      apiKeyOverride ||
      process.env.DASHSCOPE_API_KEY ||
      process.env.BAILIAN_API_KEY ||
      process.env.DASHSCOPE_KEY ||
      process.env.API_KEY
    if (!apiKey) {
      throw new Error('缺少 DashScope API Key：请在“AI 设置”里填写 apiKey（与 LLM 相同也可），或在 .env 中配置 DASHSCOPE_API_KEY')
    }
    if (!prompt || !String(prompt).trim()) {
      throw new Error('缺少 prompt')
    }

    const endpoint = getDashScopeEndpoint(region)
    const payload = {
      model,
      input: {
        messages: [
          {
            role: 'user',
            content: [{ text: String(prompt) }],
          },
        ],
      },
      parameters: {
        negative_prompt: String(negativePrompt || ''),
        prompt_extend: !!promptExtend,
        watermark: !!watermark,
        size,
      },
    }

    const { statusCode, text } = await requestJson(endpoint, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${apiKey}`,
      },
      body: JSON.stringify(payload),
    })

    let json
    try {
      json = JSON.parse(text)
    } catch {
      throw new Error(`DashScope 返回非 JSON: HTTP ${statusCode} ${text.slice(0, 200)}`)
    }

    if (statusCode < 200 || statusCode >= 300) {
      const msg = json?.message || json?.error?.message || text
      throw new Error(`DashScope 调用失败: HTTP ${statusCode} ${msg}`)
    }

    const url = extractDashScopeImageUrl(json)
    if (url) return { url, raw: json }

    // 兼容异步任务：如果返回 task_id，则轮询直到拿到图片 URL
    const taskId = extractDashScopeTaskId(json)
    if (taskId) {
      return await dashscopeWaitForImageUrlByTaskId({ taskId, region, apiKey })
    }

    throw new Error(`DashScope 返回中未找到 image url/task_id: ${text.slice(0, 500)}`)
  }

  /**
   * DashScope 图像编辑 API（qwen-image-edit-plus）
   * 用于局部编辑 PPT 页面（换背景、改文字等）
   */
  async function dashscopeImageEdit({
    imageBase64,         // 当前页图片 base64（不含 data:... 前缀）
    prompt,              // 编辑指令
    negativePrompt = '',
    n = 1,
    watermark = false,
    model = 'qwen-image-edit-plus',
    region = 'cn',
    apiKey: apiKeyOverride,
  }) {
    const apiKey =
      apiKeyOverride ||
      process.env.DASHSCOPE_API_KEY ||
      process.env.BAILIAN_API_KEY ||
      process.env.DASHSCOPE_KEY ||
      process.env.API_KEY
    if (!apiKey) {
      throw new Error('缺少 DashScope API Key')
    }
    if (!prompt || !String(prompt).trim()) {
      throw new Error('缺少编辑 prompt')
    }
    if (!imageBase64 || !String(imageBase64).trim()) {
      throw new Error('缺少待编辑的图片 base64')
    }

    const endpoint = getDashScopeEndpoint(region)
    
    // qwen-image-edit-plus 使用 MultiModalConversation 格式
    // 图片可以是 URL 或 data URI
    const imageDataUri = imageBase64.startsWith('data:')
      ? imageBase64
      : `data:image/png;base64,${imageBase64}`

    const payload = {
      model,
      input: {
        messages: [
          {
            role: 'user',
            content: [
              { image: imageDataUri },
              { text: String(prompt) },
            ],
          },
        ],
      },
      parameters: {
        negative_prompt: String(negativePrompt || ''),
        n: Math.max(1, Math.min(4, n)),
        watermark: !!watermark,
      },
    }

    const { statusCode, text } = await requestJson(endpoint, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${apiKey}`,
      },
      body: JSON.stringify(payload),
    })

    let json
    try {
      json = JSON.parse(text)
    } catch {
      throw new Error(`DashScope ImageEdit 返回非 JSON: HTTP ${statusCode} ${text.slice(0, 200)}`)
    }

    if (statusCode < 200 || statusCode >= 300) {
      const msg = json?.message || json?.error?.message || text
      throw new Error(`DashScope ImageEdit 调用失败: HTTP ${statusCode} ${msg}`)
    }

    const url = extractDashScopeImageUrl(json)
    if (url) return { url, raw: json }

    // 兼容异步任务
    const taskId = extractDashScopeTaskId(json)
    if (taskId) {
      return await dashscopeWaitForImageUrlByTaskId({ taskId, region, apiKey })
    }

    throw new Error(`DashScope ImageEdit 返回中未找到 image url/task_id: ${text.slice(0, 500)}`)
  }

  /**
   * 保存 PPT 生成的元数据到 _assets 目录
   * 用于后续编辑时恢复上下文
   */
  function saveDeckMetadata(assetsDir, metadata) {
    try {
      if (!fs.existsSync(assetsDir)) {
        fs.mkdirSync(assetsDir, { recursive: true })
      }

      if (metadata.deckContext) {
        fs.writeFileSync(
          path.join(assetsDir, 'deck_context.json'),
          JSON.stringify(metadata.deckContext, null, 2)
        )
      }

      if (metadata.slidesPrompts) {
        fs.writeFileSync(
          path.join(assetsDir, 'slides_prompts.json'),
          JSON.stringify(metadata.slidesPrompts, null, 2)
        )
      }

      if (metadata.outline) {
        fs.writeFileSync(
          path.join(assetsDir, 'outline.json'),
          JSON.stringify(metadata.outline, null, 2)
        )
      }

      console.log('[PPTX] 元数据已保存到:', assetsDir)
    } catch (e) {
      console.warn('[PPTX] 保存元数据失败:', e?.message || e)
    }
  }

  /**
   * 从 _assets 目录加载 PPT 元数据
   */
  function loadDeckMetadata(assetsDir) {
    const result = {
      deckContext: null,
      slidesPrompts: null,
      outline: null,
    }

    try {
      const contextPath = path.join(assetsDir, 'deck_context.json')
      if (fs.existsSync(contextPath)) {
        result.deckContext = JSON.parse(fs.readFileSync(contextPath, 'utf-8'))
      }
    } catch {}

    try {
      const promptsPath = path.join(assetsDir, 'slides_prompts.json')
      if (fs.existsSync(promptsPath)) {
        result.slidesPrompts = JSON.parse(fs.readFileSync(promptsPath, 'utf-8'))
      }
    } catch {}

    try {
      const outlinePath = path.join(assetsDir, 'outline.json')
      if (fs.existsSync(outlinePath)) {
        result.outline = JSON.parse(fs.readFileSync(outlinePath, 'utf-8'))
      }
    } catch {}

    return result
  }

  function getDeckAssetsDir(pptxPath) {
    const baseName = path.basename(pptxPath, '.pptx')
    return path.join(path.dirname(pptxPath), `${baseName}_assets`)
  }

  function getTextLayersDir(assetsDir) {
    return path.join(assetsDir, 'text_layers')
  }

  function getTextEditLogsDir(assetsDir) {
    return path.join(assetsDir, 'text_edit_logs')
  }

  function getTextLayerCachePath(assetsDir, pageNumber) {
    return path.join(getTextLayersDir(assetsDir), `slide_${String(pageNumber).padStart(2, '0')}.v2.json`)
  }

  function getTextLayerSourceImagePath(assetsDir, pageNumber) {
    return path.join(getTextLayersDir(assetsDir), `slide_${String(pageNumber).padStart(2, '0')}_source.png`)
  }

  function readTextLayerCache(cachePath) {
    if (!fs.existsSync(cachePath)) return null
    try {
      return JSON.parse(fs.readFileSync(cachePath, 'utf8'))
    } catch {
      return null
    }
  }

  function writeTextLayerCache(cachePath, payload) {
    const dir = path.dirname(cachePath)
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true })
    }
    fs.writeFileSync(cachePath, JSON.stringify(payload, null, 2))
  }

  function hashBuffer(buffer) {
    return crypto.createHash('sha1').update(buffer).digest('hex')
  }

  function clamp(value, min, max) {
    return Math.max(min, Math.min(max, value))
  }

  function slugify(value) {
    return String(value || '')
      .trim()
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, '-')
      .replace(/^-+|-+$/g, '') || 'candidate'
  }

  function normalizeCompareText(value) {
    return String(value || '')
      .replace(/\s+/g, '')
      .replace(/[“”"]/g, '')
      .trim()
  }

  function sequenceRatio(a, b) {
    const left = normalizeCompareText(a)
    const right = normalizeCompareText(b)
    if (!left && !right) return 1
    if (!left || !right) return 0
    const rows = Array.from({ length: left.length + 1 }, () => new Uint16Array(right.length + 1))
    for (let i = 1; i <= left.length; i += 1) {
      for (let j = 1; j <= right.length; j += 1) {
        rows[i][j] = left[i - 1] === right[j - 1]
          ? rows[i - 1][j - 1] + 1
          : Math.max(rows[i - 1][j], rows[i][j - 1])
      }
    }
    const lcs = rows[left.length][right.length]
    return (2 * lcs) / (left.length + right.length)
  }

  function safeFileUrl(filePath) {
    return `file://${String(filePath).replace(/#/g, '%23').replace(/\?/g, '%3F')}`
  }

  function getExternalAdapters() {
    const adobeAvailable = !!(process.env.ADOBE_FIREFLY_CLIENT_ID && process.env.ADOBE_FIREFLY_CLIENT_SECRET)
    const fluxAvailable = !!(process.env.BFL_API_KEY || process.env.FLUX_BFL_API_KEY)
    return [
      {
        name: 'adobe_firefly',
        available: adobeAvailable,
        reason: adobeAvailable ? undefined : 'missing_credentials',
      },
      {
        name: 'flux_refine',
        available: fluxAvailable,
        reason: fluxAvailable ? undefined : 'missing_credentials',
      },
    ]
  }

  async function getAdobeFireflyAccessToken() {
    const clientId = String(process.env.ADOBE_FIREFLY_CLIENT_ID || '').trim()
    const clientSecret = String(process.env.ADOBE_FIREFLY_CLIENT_SECRET || '').trim()
    if (!clientId || !clientSecret) {
      throw new Error('missing_adobe_credentials')
    }
    const now = Date.now()
    if (adobeTokenCache?.accessToken && adobeTokenCache?.expiresAt && adobeTokenCache.expiresAt > now + 60_000) {
      return adobeTokenCache.accessToken
    }
    const body = new URLSearchParams({
      grant_type: 'client_credentials',
      client_id: clientId,
      client_secret: clientSecret,
      scope: 'openid,AdobeID,session,additional_info,read_organizations,firefly_api,ff_apis',
    })
    const res = await fetch('https://ims-na1.adobelogin.com/ims/token/v3', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
      },
      body,
    })
    if (!res.ok) {
      throw new Error(`adobe_token_failed:${res.status}:${(await res.text()).slice(0, 300)}`)
    }
    const data = await res.json()
    const accessToken = String(data?.access_token || '')
    const expiresIn = Number(data?.expires_in || 0)
    if (!accessToken) {
      throw new Error('adobe_token_missing')
    }
    adobeTokenCache = {
      accessToken,
      expiresAt: now + Math.max(60, expiresIn - 120) * 1000,
    }
    return accessToken
  }

  async function adobeUploadImageBuffer(buffer, mimeType, accessToken) {
    const clientId = String(process.env.ADOBE_FIREFLY_CLIENT_ID || '').trim()
    const res = await fetch('https://firefly-api.adobe.io/v2/storage/image', {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'X-API-Key': clientId,
        'Content-Type': mimeType,
        'Content-Length': String(buffer.length),
      },
      body: buffer,
    })
    if (!res.ok) {
      throw new Error(`adobe_upload_failed:${res.status}:${(await res.text()).slice(0, 300)}`)
    }
    const data = await res.json()
    const uploadId = data?.images?.[0]?.id
    if (!uploadId) {
      throw new Error('adobe_upload_missing_id')
    }
    return uploadId
  }

  async function adobeFillCrop({ sourceBuffer, maskBuffer, prompt }) {
    const accessToken = await getAdobeFireflyAccessToken()
    const clientId = String(process.env.ADOBE_FIREFLY_CLIENT_ID || '').trim()
    const [sourceId, maskId] = await Promise.all([
      adobeUploadImageBuffer(sourceBuffer, 'image/png', accessToken),
      adobeUploadImageBuffer(maskBuffer, 'image/png', accessToken),
    ])
    const res = await fetch('https://firefly-api.adobe.io/v3/images/fill', {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'X-API-Key': clientId,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        image: {
          source: { uploadId: sourceId },
          mask: { uploadId: maskId },
        },
        prompt,
      }),
    })
    if (!res.ok) {
      throw new Error(`adobe_fill_failed:${res.status}:${(await res.text()).slice(0, 400)}`)
    }
    const data = await res.json()
    const imageUrl = data?.outputs?.[0]?.image?.url
    if (!imageUrl) {
      throw new Error('adobe_fill_missing_output')
    }
    return downloadToBuffer(imageUrl)
  }

  async function readImageRaw(buffer) {
    return sharp(buffer).ensureAlpha().raw().toBuffer({ resolveWithObject: true })
  }

  function averageRgbDifference(left, right, alphaMask = null) {
    const length = Math.min(left.length, right.length)
    if (!length) return 0
    let diff = 0
    let count = 0
    for (let idx = 0, pixel = 0; idx < length; idx += 4, pixel += 1) {
      const alpha = alphaMask ? alphaMask[pixel] / 255 : 1
      if (alpha <= 0.001) continue
      diff += (
        Math.abs(left[idx] - right[idx]) +
        Math.abs(left[idx + 1] - right[idx + 1]) +
        Math.abs(left[idx + 2] - right[idx + 2])
      ) / 3 * alpha
      count += alpha
    }
    return count > 0 ? diff / count : 0
  }

  function getTextLayerCacheVersion() {
    return 'v2'
  }

  function buildFontFaceCss(fontCandidate) {
    if (!fontCandidate?.fontPath) return ''
    const family = `WC_${slugify(fontCandidate.candidateId || fontCandidate.family)}`
    return {
      family,
      css: `
        @font-face {
          font-family: '${family}';
          src: url('${safeFileUrl(fontCandidate.fontPath)}');
          font-display: block;
        }
      `,
    }
  }

  /**
   * 从 PPTX 或 _assets 目录读取指定页的图片
   * @returns {Promise<Buffer|null>}
   */
  async function getSlideImageFromPptx(pptxPath, pageIndex, assetsDir) {
    // 优先从 _assets 读取最新的 processed PNG
    if (assetsDir) {
      const seq = String(pageIndex + 1).padStart(2, '0')
      // 查找最新 attempt 的 1920x1080 PNG（兼容旧的 1920x1200）
      const files = fs.existsSync(assetsDir) ? fs.readdirSync(assetsDir) : []
      const pngFiles = files
        .filter((f) => (f.startsWith(`slide_${seq}_1920x1080_`) || f.startsWith(`slide_${seq}_1920x1200_`)) && f.endsWith('.png'))
        .sort()
        .reverse()
      if (pngFiles.length > 0) {
        const pngPath = path.join(assetsDir, pngFiles[0])
        return fs.readFileSync(pngPath)
      }
    }

    // 从 PPTX 解压读取
    try {
      const JSZip = require('jszip')
      const pptxBuffer = fs.readFileSync(pptxPath)
      const zip = await JSZip.loadAsync(pptxBuffer)

      // 找到对应页的图片
      const slideNum = pageIndex + 1
      const relPath = `ppt/slides/_rels/slide${slideNum}.xml.rels`
      const relFile = zip.file(relPath)
      if (!relFile) return null

      const relXml = await relFile.async('string')
      // 找第一个 image 关系
      const match = relXml.match(/Relationship[^>]*Type="[^"]*image[^"]*"[^>]*Target="([^"]+)"/)
      if (!match) return null

      let imagePath = match[1]
      // 解析相对路径
      if (imagePath.startsWith('..')) {
        imagePath = 'ppt/' + imagePath.replace(/^\.\.\//g, '')
      } else if (!imagePath.startsWith('ppt/')) {
        imagePath = 'ppt/slides/' + imagePath
      }

      const imgFile = zip.file(imagePath)
      if (!imgFile) return null

      return Buffer.from(await imgFile.async('arraybuffer'))
    } catch (e) {
      console.warn('[PPTX] 从 PPTX 读取图片失败:', e?.message || e)
      return null
    }
  }

  /**
   * 替换 PPTX 中指定页的图片并覆盖写回
   * @param {string} pptxPath - PPTX 文件路径
   * @param {Array<{pageIndex: number, imageBuffer: Buffer}>} replacements - 替换列表
   * @param {boolean} backup - 是否备份原文件
   */
  async function replaceSlideImagesInPptx(pptxPath, replacements, backup = true) {
    const JSZip = require('jszip')
    const pptxBuffer = fs.readFileSync(pptxPath)
    const zip = await JSZip.loadAsync(pptxBuffer)

    // 备份原文件
    if (backup) {
      const dir = path.dirname(pptxPath)
      const baseName = path.basename(pptxPath, '.pptx')
      const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19)
      const backupPath = path.join(dir, `${baseName}_backup_${timestamp}.pptx`)
      fs.copyFileSync(pptxPath, backupPath)
      console.log('[PPTX] 已备份到:', backupPath)
    }

    for (const { pageIndex, imageBuffer } of replacements) {
      const slideNum = pageIndex + 1
      const relPath = `ppt/slides/_rels/slide${slideNum}.xml.rels`
      const relFile = zip.file(relPath)
      if (!relFile) {
        console.warn(`[PPTX] 未找到 slide${slideNum} 的 rels`)
        continue
      }

      const relXml = await relFile.async('string')
      const match = relXml.match(/Relationship[^>]*Type="[^"]*image[^"]*"[^>]*Target="([^"]+)"/)
      if (!match) {
        console.warn(`[PPTX] slide${slideNum} 未找到图片关系`)
        continue
      }

      let imagePath = match[1]
      if (imagePath.startsWith('..')) {
        imagePath = 'ppt/' + imagePath.replace(/^\.\.\//g, '')
      } else if (!imagePath.startsWith('ppt/')) {
        imagePath = 'ppt/slides/' + imagePath
      }

      // 替换图片
      zip.file(imagePath, imageBuffer)
      console.log(`[PPTX] 已替换 slide${slideNum} 图片: ${imagePath}`)
    }

    // 写回 PPTX
    const newBuffer = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
    fs.writeFileSync(pptxPath, newBuffer)
    console.log('[PPTX] 已覆盖写回:', pptxPath)

    return { success: true, path: pptxPath }
  }

  async function cropImageBuffer(buffer, bounds, padding = 28) {
    const meta = await sharp(buffer).metadata()
    const width = meta.width || 0
    const height = meta.height || 0
    if (!width || !height) {
      throw new Error('无法读取图片尺寸')
    }
    const left = Math.max(0, Math.floor(bounds.left - padding))
    const top = Math.max(0, Math.floor(bounds.top - padding))
    const right = Math.min(width, Math.ceil(bounds.left + bounds.width + padding))
    const bottom = Math.min(height, Math.ceil(bounds.top + bounds.height + padding))
    const extract = {
      left,
      top,
      width: Math.max(1, right - left),
      height: Math.max(1, bottom - top),
    }
    const cropBuffer = await sharp(buffer)
      .extract(extract)
      .png()
      .toBuffer()
    return {
      cropBuffer,
      extract,
      imageWidth: width,
      imageHeight: height,
    }
  }

  async function buildCleanupMaskCropBuffer({ box, cropExtract, strategy }) {
    const style = box?.styleEstimate || box?.styleHint || {}
    const bounds = style.textBounds || box.bounds
    const fontSize = Math.max(8, Number(style.fontSize || bounds.height || 18))
    const strokeWidth = Math.max(0, Number(style.strokeWidth || 0))
    const shadowBlur = Math.max(0, Number(style.shadowBlur || 0))
    const shadowOffsetX = Math.abs(Number(style.shadowOffsetX || 0))
    const shadowOffsetY = Math.abs(Number(style.shadowOffsetY || 0))

    const rects = []
    if (strategy === 'adobe_firefly_fill' && Array.isArray(box?.charBoxes) && box.charBoxes.length > 0) {
      const padX = Math.max(4, Math.round(fontSize * 0.12 + strokeWidth * 2 + shadowOffsetX * 0.5 + shadowBlur * 0.4))
      const padY = Math.max(4, Math.round(fontSize * 0.16 + strokeWidth * 2 + shadowOffsetY + shadowBlur * 0.5))
      for (const charBox of box.charBoxes) {
        const cb = charBox.bounds
        rects.push({
          x: clamp(Math.floor(cb.left - cropExtract.left - padX), 0, cropExtract.width),
          y: clamp(Math.floor(cb.top - cropExtract.top - padY), 0, cropExtract.height),
          width: clamp(Math.ceil(cb.width + padX * 2), 1, cropExtract.width),
          height: clamp(Math.ceil(cb.height + padY * 2), 1, cropExtract.height),
          rx: Math.max(1, Math.round(fontSize * 0.04)),
        })
      }
    }
    if (!rects.length) {
      const padX = Math.max(8, Math.round(fontSize * 0.22 + strokeWidth * 3 + shadowOffsetX + shadowBlur))
      const padY = Math.max(6, Math.round(fontSize * 0.24 + strokeWidth * 3 + shadowOffsetY + shadowBlur))
      rects.push({
        x: clamp(Math.floor(bounds.left - cropExtract.left - padX), 0, cropExtract.width),
        y: clamp(Math.floor(bounds.top - cropExtract.top - padY), 0, cropExtract.height),
        width: clamp(Math.ceil(bounds.width + padX * 2), 1, cropExtract.width),
        height: clamp(Math.ceil(bounds.height + padY * 2), 1, cropExtract.height),
        rx: Math.max(2, Math.round(fontSize * 0.06)),
      })
    }

    const svg = `
      <svg xmlns="http://www.w3.org/2000/svg" width="${cropExtract.width}" height="${cropExtract.height}" viewBox="0 0 ${cropExtract.width} ${cropExtract.height}">
        <rect x="0" y="0" width="${cropExtract.width}" height="${cropExtract.height}" fill="black" />
        ${rects.map((rect) => (
          `<rect x="${rect.x}" y="${rect.y}" width="${rect.width}" height="${rect.height}" rx="${rect.rx}" ry="${rect.rx}" fill="white" />`
        )).join('\n')}
      </svg>
    `
    return sharp(Buffer.from(svg)).png().toBuffer()
  }

  async function compositeCropBack(baseBuffer, cropBuffer, extract) {
    return sharp(baseBuffer)
      .composite([
        {
          input: cropBuffer,
          left: extract.left,
          top: extract.top,
        },
      ])
      .png()
      .toBuffer()
  }

  async function ensurePageSourceImage(pptxPath, pageNumber, assetsDir) {
    const pageIndex = Math.max(0, pageNumber - 1)
    const imageBuffer = await getSlideImageFromPptx(pptxPath, pageIndex, assetsDir)
    if (!imageBuffer) {
      throw new Error(`无法读取第 ${pageNumber} 页图片`)
    }

    const textLayersDir = getTextLayersDir(assetsDir)
    if (!fs.existsSync(textLayersDir)) {
      fs.mkdirSync(textLayersDir, { recursive: true })
    }

    const sourcePath = getTextLayerSourceImagePath(assetsDir, pageNumber)
    fs.writeFileSync(sourcePath, imageBuffer)

    return {
      imageBuffer,
      imageHash: hashBuffer(imageBuffer),
      sourcePath,
    }
  }

  function enrichDetectedTextBox(box) {
    const styleHint = {
      shadowColor: '#000000',
      shadowOpacity: 0,
      shadowOffsetX: 0,
      shadowOffsetY: 0,
      shadowBlur: 0,
      strokeColor: '#000000',
      strokeWidth: 0,
      letterSpacing: 0,
      lineHeight: 1,
      opacity: 1,
      blendMode: 'normal',
      ...(box.styleHint || {}),
    }
    const styleEstimate = {
      ...styleHint,
      textDirection: box.textDirection || box.styleEstimate?.textDirection || 'ltr',
      skewX: box.styleEstimate?.skewX || 0,
      skewY: box.styleEstimate?.skewY || 0,
      ...(box.styleEstimate || {}),
    }
    const fontCandidates = Array.isArray(box.fontCandidates) && box.fontCandidates.length
      ? box.fontCandidates
      : [
          {
            candidateId: `system:${styleHint.familyHint || 'default'}`,
            family: styleHint.familyHint || 'PingFang SC',
            confidence: 0.6,
            source: 'system',
          },
        ]
    return {
      ...box,
      rotation: Number.isFinite(Number(box.rotation)) ? Number(box.rotation) : Number(styleHint.rotation || 0),
      skew: Number.isFinite(Number(box.skew)) ? Number(box.skew) : 0,
      textDirection: box.textDirection || 'ltr',
      backgroundComplexity: box.backgroundComplexity || 'medium',
      styleComplexity: box.styleComplexity || 'plain',
      charBoxes: Array.isArray(box.charBoxes) ? box.charBoxes : [],
      styleHint,
      styleEstimate,
      fontCandidates,
    }
  }

  function enrichDetectedTextBoxes(boxes = []) {
    return boxes.map(enrichDetectedTextBox)
  }

  function getTempTextEditDir() {
    const dir = path.join(app.getPath('temp'), 'word-cursor-ppt-text-v2')
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true })
    }
    return dir
  }

  function makeTempFilePath(prefix, ext = '.png') {
    const stamp = `${Date.now()}-${Math.random().toString(16).slice(2, 8)}`
    return path.join(getTempTextEditDir(), `${prefix}-${stamp}${ext}`)
  }

  function colorToRgba(hex, alpha = 1) {
    const safe = String(hex || '#000000').trim()
    const value = /^#[0-9a-f]{6}$/i.test(safe) ? safe : '#000000'
    const r = parseInt(value.slice(1, 3), 16)
    const g = parseInt(value.slice(3, 5), 16)
    const b = parseInt(value.slice(5, 7), 16)
    return `rgba(${r}, ${g}, ${b}, ${clamp(alpha, 0, 1)})`
  }

  async function cleanupTextBoxesOnBuffer({ imageBuffer, boxes, boxIds, forceStrategy }) {
    const inputPath = makeTempFilePath('cleanup-input')
    const outputPath = makeTempFilePath('cleanup-output')
    fs.writeFileSync(inputPath, imageBuffer)
    const result = await textSidecar.cleanupTextBoxes({
      image_path: inputPath,
      output_path: outputPath,
      boxes,
      box_ids: boxIds,
      force_strategy: forceStrategy,
    })
    try { fs.unlinkSync(inputPath) } catch {}
    if (!result?.success || !result.outputPath || !fs.existsSync(result.outputPath)) {
      return {
        success: false,
        error: result?.error || 'cleanup failed',
        logs: result?.logs || [],
      }
    }
    const buffer = fs.readFileSync(result.outputPath)
    try { fs.unlinkSync(outputPath) } catch {}
    return {
      success: true,
      buffer,
      logs: result.logs || [],
    }
  }

  async function recognizeTextFromBuffer(buffer) {
    const inputPath = makeTempFilePath('ocr-input')
    fs.writeFileSync(inputPath, buffer)
    const result = await textSidecar.recognizeText({ image_path: inputPath })
    try { fs.unlinkSync(inputPath) } catch {}
    return result
  }

  async function recognizeTextsFromBuffers(buffers) {
    const valid = Array.isArray(buffers) ? buffers.filter((buffer) => Buffer.isBuffer(buffer) && buffer.length > 0) : []
    if (!valid.length) return []
    const inputPaths = valid.map((buffer, index) => {
      const filePath = makeTempFilePath(`ocr-batch-${index}`)
      fs.writeFileSync(filePath, buffer)
      return filePath
    })
    try {
      const result = await textSidecar.recognizeTextsBatch({ image_paths: inputPaths })
      const items = Array.isArray(result?.items) ? result.items : []
      return items
    } finally {
      for (const filePath of inputPaths) {
        try { fs.unlinkSync(filePath) } catch {}
      }
    }
  }

  async function getDeterministicRenderWindow(width, height) {
    const targetWidth = Math.max(64, Math.ceil(width))
    const targetHeight = Math.max(64, Math.ceil(height))
    if (!deterministicRenderWindow || deterministicRenderWindow.isDestroyed()) {
      deterministicRenderWindow = new BrowserWindow({
        show: false,
        width: targetWidth,
        height: targetHeight,
        transparent: true,
        frame: false,
        resizable: false,
        movable: false,
        fullscreenable: false,
        paintWhenInitiallyHidden: true,
        backgroundColor: '#00000000',
        webPreferences: {
          sandbox: false,
          contextIsolation: true,
        },
      })
      deterministicRenderWindowReady = false
    } else {
      deterministicRenderWindow.setBounds({
        x: 0,
        y: 0,
        width: targetWidth,
        height: targetHeight,
      })
    }

    if (!deterministicRenderWindowReady) {
      await deterministicRenderWindow.loadURL('data:text/html,<html><body style="margin:0;background:transparent"></body></html>')
      deterministicRenderWindowReady = true
    }
    return deterministicRenderWindow
  }

  async function renderDeterministicTextLayer({
    width,
    height,
    relativeBounds,
    text,
    styleEstimate,
    fontCandidate,
  }) {
    const fontFace = buildFontFaceCss(fontCandidate)
    const fontFamily = fontFace?.family || fontCandidate?.family || styleEstimate.fontFamily || 'PingFang SC'
    const shadowOpacity = clamp(Number(styleEstimate.shadowOpacity || 0), 0, 1)
    const textShadow = shadowOpacity > 0
      ? `${Number(styleEstimate.shadowOffsetX || 0)}px ${Number(styleEstimate.shadowOffsetY || 0)}px ${Math.max(0, Number(styleEstimate.shadowBlur || 0))}px ${colorToRgba(styleEstimate.shadowColor || '#000000', shadowOpacity)}`
      : 'none'
    const html = `<!doctype html>
      <html>
        <head>
          <meta charset="utf-8" />
          <style>
            html, body {
              margin: 0;
              width: ${Math.ceil(width)}px;
              height: ${Math.ceil(height)}px;
              overflow: hidden;
              background: transparent;
            }
            ${fontFace ? fontFace.css : ''}
            #stage {
              position: relative;
              width: ${Math.ceil(width)}px;
              height: ${Math.ceil(height)}px;
              background: transparent;
            }
            #box {
              position: absolute;
              left: ${relativeBounds.left}px;
              top: ${relativeBounds.top}px;
              width: ${relativeBounds.width}px;
              height: ${relativeBounds.height}px;
              display: flex;
              align-items: center;
              justify-content: ${styleEstimate.align === 'right' ? 'flex-end' : styleEstimate.align === 'center' ? 'center' : 'flex-start'};
              transform: rotate(${Number(styleEstimate.rotation || 0)}deg) skewX(${Number(styleEstimate.skewX || 0)}deg);
              transform-origin: center center;
            }
            #text {
              box-sizing: border-box;
              width: 100%;
              color: ${styleEstimate.textColor || '#000000'};
              font-family: '${fontFamily}', '${fontCandidate?.family || ''}', ${styleEstimate.familyHint === 'serif' ? 'serif' : 'sans-serif'};
              font-size: ${Math.max(8, Number(styleEstimate.fontSize || 18))}px;
              line-height: ${Math.max(0.8, Number(styleEstimate.lineHeight || 1))};
              letter-spacing: ${Number(styleEstimate.letterSpacing || 0)}px;
              text-align: ${styleEstimate.align || 'left'};
              white-space: pre-wrap;
              word-break: break-word;
              overflow-wrap: anywhere;
              text-shadow: ${textShadow};
              -webkit-text-stroke: ${Math.max(0, Number(styleEstimate.strokeWidth || 0))}px ${styleEstimate.strokeColor || 'transparent'};
              opacity: ${clamp(Number(styleEstimate.opacity || 1), 0, 1)};
              mix-blend-mode: ${styleEstimate.blendMode || 'normal'};
            }
          </style>
        </head>
        <body>
          <div id="stage">
            <div id="box"><div id="text">${String(text || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')}</div></div>
          </div>
        </body>
      </html>`

    const win = await getDeterministicRenderWindow(width, height)

    await win.loadURL(`data:text/html;charset=UTF-8,${encodeURIComponent(html)}`)
    const metrics = await win.webContents.executeJavaScript(`
        new Promise(async (resolve) => {
          try {
            await (document.fonts?.ready || Promise.resolve())
          } catch {}
          const box = document.getElementById('box')
          const text = document.getElementById('text')
          let fontSize = parseFloat(getComputedStyle(text).fontSize || '16')
          let guard = 0
          while (guard < 80 && (text.scrollWidth > box.clientWidth + 1 || text.scrollHeight > box.clientHeight + 1) && fontSize > 8) {
            fontSize -= 1
            text.style.fontSize = fontSize + 'px'
            guard += 1
          }
          resolve({
            finalFontSize: fontSize,
            scrollWidth: text.scrollWidth,
            scrollHeight: text.scrollHeight,
            boxWidth: box.clientWidth,
            boxHeight: box.clientHeight,
            overflowX: Math.max(0, text.scrollWidth - box.clientWidth),
            overflowY: Math.max(0, text.scrollHeight - box.clientHeight),
          })
        })
      `)
    await new Promise((resolve) => setTimeout(resolve, 20))
    const image = await win.webContents.capturePage()
    return {
      success: true,
      buffer: image.toPNG(),
      metrics: {
        ...metrics,
        fontFamily,
      },
    }
  }

  async function scoreEditCandidate({
    originalCropBuffer,
    cleanCropBuffer,
    compositeCropBuffer,
    renderLayerBuffer,
    box,
    targetText,
    cleanupResidualText,
    cleanupStrategy,
    fontCandidate,
    renderMetrics,
    ocrTextOverride,
  }) {
    const ocrText = ocrTextOverride != null
      ? String(ocrTextOverride || '')
      : (await recognizeTextFromBuffer(compositeCropBuffer))?.text || ''
    const ocrExactness = clamp(sequenceRatio(ocrText, targetText), 0, 1)
    const cleanupResidualSimilarity = clamp(sequenceRatio(cleanupResidualText || '', box?.text || ''), 0, 1)
    const fontConfidence = clamp(Number(fontCandidate?.confidence || 0.4), 0, 1)
    const expectedFontSize = Math.max(8, Number(box?.styleEstimate?.fontSize || box?.styleHint?.fontSize || 16))
    const fontSizeRatio = clamp(1 - Math.abs((Number(renderMetrics?.finalFontSize || expectedFontSize) - expectedFontSize) / expectedFontSize), 0, 1)
    const overflowPenalty = clamp(
      ((Number(renderMetrics?.overflowX || 0) / Math.max(1, Number(renderMetrics?.boxWidth || 1))) +
      (Number(renderMetrics?.overflowY || 0) / Math.max(1, Number(renderMetrics?.boxHeight || 1)))) / 2,
      0,
      1,
    )

    const originalRaw = await readImageRaw(originalCropBuffer)
    const cleanRaw = await readImageRaw(cleanCropBuffer)
    const compositeRaw = await readImageRaw(compositeCropBuffer)
    const renderRaw = await readImageRaw(renderLayerBuffer)
    const alphaMask = new Uint8Array(renderRaw.info.width * renderRaw.info.height)
    const inverseMask = new Uint8Array(alphaMask.length)
    for (let idx = 0, pixel = 0; idx < renderRaw.data.length; idx += 4, pixel += 1) {
      const alpha = renderRaw.data[idx + 3]
      alphaMask[pixel] = alpha
      inverseMask[pixel] = 255 - alpha
    }

    const cleanupImpact = averageRgbDifference(originalRaw.data, cleanRaw.data, inverseMask)
    const outsideDiff = averageRgbDifference(cleanRaw.data, compositeRaw.data, inverseMask)
    const edgeDiff = averageRgbDifference(cleanRaw.data, compositeRaw.data, alphaMask)
    const backgroundPreservation = clamp(1 - outsideDiff / 28, 0, 1)
    const cleanupImpactPenalty = clamp(cleanupImpact / 40, 0, 1)
    const edgeArtifactScore = clamp(1 - edgeDiff / 84, 0, 1)
    const fontStyleSimilarity = clamp(0.6 * fontConfidence + 0.4 * fontSizeRatio, 0, 1)
    const cleanupStrategyPenalty =
      cleanupStrategy === 'analytic_fill'
        ? box?.backgroundComplexity === 'complex'
          ? 0.1
          : box?.backgroundComplexity === 'medium'
            ? 0.04
            : 0
        : 0
    const total = clamp(
      0.38 * ocrExactness +
      0.24 * fontStyleSimilarity +
      0.18 * backgroundPreservation +
      0.14 * edgeArtifactScore -
      0.20 * cleanupResidualSimilarity -
      0.18 * cleanupImpactPenalty -
      cleanupStrategyPenalty -
      0.16 * overflowPenalty,
      0,
      1,
    )

    return {
      total,
      ocrExactness,
      fontStyleSimilarity,
      backgroundPreservation,
      edgeArtifactScore,
      overflowPenalty,
      cleanupResidualSimilarity,
      cleanupImpactPenalty,
      cleanupStrategyPenalty,
      detectedText: ocrText,
    }
  }

  function chooseBlendStrategy(box) {
    return 'deterministic'
  }

  async function detectTextLayerInternal({ pptxPath, pageNumber, useCache = true, cacheOnly = false }) {
    emitPptTextEditProgress({
      stage: 'detecting',
      progress: 0.05,
      message: `正在准备第 ${pageNumber} 页文字识别...`,
      pageNumber,
    })
    const assetsDir = getDeckAssetsDir(pptxPath)
    const cachePath = getTextLayerCachePath(assetsDir, pageNumber)
    const source = await ensurePageSourceImage(pptxPath, pageNumber, assetsDir)

    if (useCache) {
      const cached = readTextLayerCache(cachePath)
      if (cached && cached.imageHash === source.imageHash && Array.isArray(cached.boxes)) {
        return {
          success: true,
          cached: true,
          cacheVersion: cached.cacheVersion || getTextLayerCacheVersion(),
          canvasWidth: cached.canvasWidth,
          canvasHeight: cached.canvasHeight,
          boxes: enrichDetectedTextBoxes(cached.boxes),
          sourceImagePath: source.sourcePath,
          cachePath,
          assetsDir,
        }
      }

      if (cacheOnly) {
        return {
          success: false,
          cached: false,
          error: '当前页没有可用的文字层缓存',
          cachePath,
          assetsDir,
        }
      }
    }

    emitPptTextEditProgress({
      stage: 'detecting',
      progress: 0.35,
      message: `正在识别第 ${pageNumber} 页文字...`,
      pageNumber,
    })
    const result = await textSidecar.detectTextBoxes({
      image_path: source.sourcePath,
    })

    if (!result?.success) {
      return {
        success: false,
        error: result?.error || 'OCR 识别失败',
        cachePath,
        assetsDir,
      }
    }

    const nextBoxes = enrichDetectedTextBoxes(result.boxes || [])
    emitPptTextEditProgress({
      stage: 'detecting',
      progress: 0.9,
      message: `正在整理文字框与样式候选...`,
      pageNumber,
    })
    writeTextLayerCache(cachePath, {
      cacheVersion: getTextLayerCacheVersion(),
      imageHash: source.imageHash,
      canvasWidth: result.canvasWidth,
      canvasHeight: result.canvasHeight,
      boxes: nextBoxes,
      updatedAt: new Date().toISOString(),
      engine: result.engine || 'paddleocr',
    })

    return {
      success: true,
      cached: false,
      cacheVersion: getTextLayerCacheVersion(),
      canvasWidth: result.canvasWidth,
      canvasHeight: result.canvasHeight,
      boxes: nextBoxes,
      sourceImagePath: source.sourcePath,
      cachePath,
      assetsDir,
    }
  }

  async function postprocessTo1920x1200(buffer, mode = 'letterbox') {
    // 更新为 16:9 分辨率以匹配 z-image-turbo 输出 (2048*1152)
    const targetW = 1920
    const targetH = 1080 // 16:9 比例
    if (mode === 'cover') {
      return await sharp(buffer).resize(targetW, targetH, { fit: 'cover', position: 'attention' }).png().toBuffer()
    }
    // default: letterbox (no crop)
    return await sharp(buffer)
      .resize(targetW, targetH, { fit: 'contain', background: { r: 0, g: 0, b: 0, alpha: 1 } })
      .png()
      .toBuffer()
  }

  function makePptx16x10FromImagesBase64(imageBase64List, outputPath) {
    const pptx = new PptxGenJS()
    // 更新为 16:9 布局以匹配 z-image-turbo 输出
    const w = 13.333 // 10 英寸 * 1.333
    const h = 7.5    // 16:9 比例 (10 英寸宽 * 9/16 = 5.625 英寸，但 PPT 标准是 13.333 x 7.5)
    pptx.defineLayout({ name: 'LAYOUT_16X9', width: w, height: h })
    pptx.layout = 'LAYOUT_16X9'
    pptx.author = '智启文档'

    for (const img of imageBase64List) {
      const slide = pptx.addSlide()
      // 显式设置背景，避免部分前端预览器解析时出现 background undefined
      slide.background = { color: '000000' }
      slide.addImage({ data: img, x: 0, y: 0, w, h })
    }
    return pptx.writeFile({ fileName: outputPath, compression: true })
  }

  // OpenRouter Gemini: 调用（支持 messages 或 system+user）
  async function callOpenRouterGemini({ apiKey, model, systemPrompt, userPrompt, messages }) {
    const baseUrl = 'https://openrouter.ai/api/v1/chat/completions'
    // 使用 Gemini 3 Pro Preview（最新最强）
    const selectedModel = model || 'google/gemini-3-pro-preview'
    const finalMessages = Array.isArray(messages) && messages.length > 0
      ? messages
      : [
          { role: 'system', content: systemPrompt },
          { role: 'user', content: userPrompt }
        ]
    const body = {
      model: selectedModel,
      messages: finalMessages,
      temperature: 0.7,
      max_tokens: 16000, // Gemini 3 Pro 支持更长输出，PPT 提示词需要足够空间
    }
    console.log('[OpenRouter] Calling Gemini 3 Pro:', selectedModel)
    const res = await fetch(baseUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`,
        'HTTP-Referer': 'https://word-cursor.app',
        'X-Title': 'ZhiQi Docs PPT Generator',
      },
      body: JSON.stringify(body),
    })
    if (!res.ok) {
      const text = await res.text()
      throw new Error(`OpenRouter API error: ${res.status} - ${text}`)
    }
    const data = await res.json()
    console.log('[OpenRouter] Gemini 3 Pro response, finish_reason:', data.choices?.[0]?.finish_reason, 'tokens:', data.usage?.total_tokens)
    return data.choices?.[0]?.message?.content || ''
  }

  // 通用 OpenAI 兼容 API 调用（用于主模型回退）
  async function callOpenAICompatible({ apiKey, baseUrl, model, systemPrompt, userPrompt, messages }) {
    // 清理 baseUrl，确保正确格式
    let endpoint = String(baseUrl || 'https://api.openai.com/v1').trim()
    if (endpoint.endsWith('/')) {
      endpoint = endpoint.slice(0, -1)
    }
    if (!endpoint.endsWith('/chat/completions')) {
      endpoint = `${endpoint}/chat/completions`
    }
    
    const finalMessages = Array.isArray(messages) && messages.length > 0
      ? messages
      : [
          { role: 'system', content: systemPrompt },
          { role: 'user', content: userPrompt }
        ]
    
    const body = {
      model: model || 'gpt-4',
      messages: finalMessages,
      temperature: 0.7,
      max_tokens: 16000,
    }
    
    console.log('[OpenAI Compatible] Calling:', model, 'at', endpoint)
    const res = await fetch(endpoint, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`,
      },
      body: JSON.stringify(body),
    })
    
    if (!res.ok) {
      const text = await res.text()
      throw new Error(`API error: ${res.status} - ${text}`)
    }
    
    const data = await res.json()
    console.log('[OpenAI Compatible] Response, finish_reason:', data.choices?.[0]?.finish_reason, 'tokens:', data.usage?.total_tokens)
    return data.choices?.[0]?.message?.content || ''
  }

  // 通用主模型调用：用于 PPT 提示词生成，优先走当前设置中的 OpenAI 兼容接口
  async function callConfiguredGeminiText({ apiKey, baseUrl, model, systemPrompt, userPrompt, messages }) {
    let endpoint = String(baseUrl || 'https://api.linapi.net/v1').trim()
    if (endpoint.endsWith('/')) endpoint = endpoint.slice(0, -1)
    if (!endpoint.endsWith('/chat/completions')) {
      endpoint = `${endpoint}/chat/completions`
    }
    const selectedModel = model || 'gemini-3.1-pro-preview'
    const finalMessages = Array.isArray(messages) && messages.length > 0
      ? messages
      : [
          { role: 'system', content: systemPrompt },
          { role: 'user', content: userPrompt }
        ]
    const body = {
      model: selectedModel,
      messages: finalMessages,
      temperature: 0.7,
      max_tokens: 16000,
    }
    
    // 添加重试逻辑（网络不稳定时最多重试 3 次）
    const maxRetries = 3
    let lastError = null
    
    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      try {
        console.log(`[PPT Prompts] Calling main model: ${selectedModel} @ ${endpoint} (attempt ${attempt}/${maxRetries})`)
        const res = await fetch(endpoint, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${apiKey}`,
          },
          body: JSON.stringify(body),
        })
        if (!res.ok) {
          const text = await res.text()
          throw new Error(`Main model error: ${res.status} - ${text}`)
        }
        const data = await res.json()
        console.log('[PPT Prompts] Main model response, finish_reason:', data.choices?.[0]?.finish_reason, 'tokens:', data.usage?.total_tokens)
        return data.choices?.[0]?.message?.content || ''
      } catch (err) {
        lastError = err
        const isNetworkError = err?.cause?.code === 'ECONNRESET' || 
                               err?.cause?.code === 'UND_ERR_SOCKET' ||
                               err?.message?.includes('fetch failed')
        if (isNetworkError && attempt < maxRetries) {
          console.warn(`[PPT Prompts] 网络错误，${attempt}s 后重试... (${err?.cause?.code || err.message})`)
          await new Promise(r => setTimeout(r, attempt * 1000)) // 递增等待时间
          continue
        }
        throw err
      }
    }
    throw lastError
  }

  // LinAPI Gemini 生图: 调用 gemini-3-pro-image-preview-2K 生成图片（带重试）
  async function linapiGenerateImage({ apiKey, prompt, aspectRatio = '16:9' }) {
    const endpoint = 'https://api.linapi.net/v1beta/models/gemini-3-pro-image-preview-2K:generateContent'
    
    const body = {
      contents: [{
        parts: [{ text: prompt }]
      }],
      generationConfig: {
        imageConfig: {
          aspectRatio: aspectRatio,
          imageSize: '1K'
        }
      }
    }
    
    // 添加重试逻辑（网络不稳定时最多重试 3 次）
    const maxRetries = 3
    let lastError = null
    
    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      try {
        console.log(`\n${'-'.repeat(40)}`)
        console.log(`[LinAPI Image] Generating image (attempt ${attempt}/${maxRetries})`)
        console.log(`[LinAPI Image] FULL PROMPT:\n${prompt}`)
        console.log(`${'-'.repeat(40)}\n`)
        
        const res = await fetch(endpoint, {
          method: 'POST',
          headers: {
            'x-goog-api-key': apiKey,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify(body),
        })
        
        if (!res.ok) {
          const text = await res.text()
          throw new Error(`LinAPI Image error: ${res.status} - ${text}`)
        }
        
        const data = await res.json()
        
        // 提取生成的图片 Base64 数据
        const candidate = data.candidates?.[0]
        const parts = candidate?.content?.parts || []
        
        for (const part of parts) {
          if (part.inlineData?.data) {
            const mimeType = part.inlineData.mimeType || 'image/png'
            const base64Data = part.inlineData.data
            console.log('[LinAPI Image] Got image, mimeType:', mimeType, 'size:', base64Data.length)
            return {
              url: `data:${mimeType};base64,${base64Data}`,
              base64: base64Data,
              mimeType
            }
          }
        }
        
        throw new Error('LinAPI Image: 未在响应中找到图片数据')
      } catch (err) {
        lastError = err
        const isNetworkError = err?.cause?.code === 'ECONNRESET' || 
                               err?.cause?.code === 'UND_ERR_SOCKET' ||
                               err?.message?.includes('fetch failed') ||
                               err?.message?.includes('socket')
        if (isNetworkError && attempt < maxRetries) {
          console.warn(`[LinAPI Image] 网络错误，${attempt * 2}s 后重试... (${err?.cause?.code || err.message})`)
          await new Promise(r => setTimeout(r, attempt * 2000)) // 生图需要更长等待时间
          continue
        }
        throw err
      }
    }
    throw lastError
  }

  function enhancePromptForGeminiImage({ prompt, negativePrompt }) {
    const safePrompt = String(prompt || '').trim()
    const safeNeg = String(negativePrompt || '').trim()
    // Gemini 生图对"负面词"没有单独字段，这里把 negativePrompt 作为 Avoid 列表融入同一条文本指令里
    // 目标：高端杂志级设计、丰富细节、精致纹理、专业排版
    
    const designDirectives = `## IMAGE GENERATION REQUIREMENTS (STRICT)

  ### 1. FORMAT & PURPOSE
  You are generating a PRESENTATION SLIDE image (16:9 aspect ratio).
  This is FLAT GRAPHIC DESIGN for business/editorial use — NOT a 3D render, NOT a game scene, NOT dark cyberpunk art.

  ### 2. COLOR PALETTE (MANDATORY)
  - PRIMARY BACKGROUND: Off-white (#F8F7F4), warm gray (#E8E6E1), soft ivory (#FDFBF7), or pale cream (#F5F3EE)
  - Background must have SUBTLE TEXTURE: fine paper grain, linen weave, very light noise, or faint watercolor wash — NEVER flat solid color
  - TEXT COLOR: Rich charcoal (#2D2D2D) or warm dark gray (#3A3A3A) — NOT pure black
  - ACCENT COLOR: ONE sophisticated accent only (muted blue #4A7C9B, terracotta #C4785A, sage green #7D9B76, or warm gold #B8976A) — used sparingly (≤5% of area)
  - FORBIDDEN: Neon colors, saturated blues/purples, glowing effects, gradients that look like cheap stock photos

  ### 3. LAYOUT & COMPOSITION (MANDATORY)
  - Use strict GRID SYSTEM: Bento grid, modular grid, or classic editorial columns
  - Strong ALIGNMENT: all elements snap to baseline grid
  - Generous WHITE SPACE: minimum 15% margins, breathing room between elements
  - Clear VISUAL HIERARCHY: primary → secondary → tertiary information levels
  - Professional TYPOGRAPHY: elegant sans-serif for Chinese text, proper kerning, comfortable line-height (1.4-1.6)

  ### 4. MICRO-DETAILS (CRITICAL — This creates richness)
  Add these LOW-OPACITY decorative elements throughout the design:
  - Ultra-thin grid lines (0.5px, 5-10% opacity)
  - Corner registration marks (like print crop marks)
  - Small page numbers or serial codes (No.01, VOL.25)
  - Tiny geometric accents: dots, crosses, small squares, subtle lines
  - Abstract data visualization elements (thin connecting lines, small nodes)
  - Faint geometric patterns in background (hexagons, circles, triangles at 3-5% opacity)
  - Subtle shadow layers for depth
  - Delicate dividing lines between content sections
  - Small iconographic elements relevant to the topic (minimal line-art style)
  - Grain/noise texture overlay (very subtle, 2-5% opacity)

  ### 5. MATERIALITY & TEXTURE
  - Frosted glass cards (glassmorphism) for content containers — with realistic blur and soft edges
  - Soft drop shadows (not harsh, offset 0-4px, blur 8-16px, 8-15% opacity)
  - Paper-like texture on background
  - Subtle embossing or debossing effects on key elements
  - Matte finish aesthetic — no glossy/plastic look

  ### 6. TYPOGRAPHY REQUIREMENTS
  - ALL Chinese characters must be PERFECTLY LEGIBLE, correctly formed, elegantly spaced
  - Use modern Chinese sans-serif aesthetic (like PingFang, Source Han Sans style)
  - Proper text hierarchy through size, weight, and spacing — NOT color variation
  - Headlines: bold/semibold, generous tracking
  - Body text: regular weight, comfortable line-height
  - NO random English text, NO gibberish, NO Lorem Ipsum — only the exact content provided

  ### 7. STRICT AVOIDANCE LIST
  NEVER include these elements:
  - 3D rendered spheres, cubes, or geometric shapes that look like stock 3D assets
  - Dark/black backgrounds
  - Neon glows, lens flares, or light leaks
  - Cyberpunk/sci-fi hologram aesthetics
  - Cheap-looking gradients (especially blue-purple)
  - Generic stock photo elements (handshake, lightbulb, puzzle pieces)
  - Watermarks, logos, or brand marks
  - Cluttered layouts with no breathing room
  - Toy-like or plastic textures
  - Overly complex 3D scenes or realistic photo compositions
  - Random English abbreviations or placeholder text

  ---

  ## CONTENT TO VISUALIZE:`
    
    const avoidList = [
      '3D spheres', '3D balls', '3D cubes', '3D geometric primitives',
      'dark background', 'black background', 'neon glow', 'cyberpunk',
      'hologram', 'sci-fi UI', 'circuit board', 'matrix code',
      'cheap gradient', 'stock photo', 'plastic texture', 'toy-like',
      'blurry text', 'deformed text', 'broken text', 'garbled Chinese',
      'wrong characters', 'illegible text', 'ugly typography',
      'watermark', 'logo', 'brand mark', 'lowres', 'amateur design',
      safeNeg
    ].filter(Boolean).join(', ')

    return [
      designDirectives,
      '',
      safePrompt,
      '',
      `## MUST AVOID: ${avoidList}`
    ].join('\n')
  }

  function isDashscopeInappropriateContentError(err) {
    const msg = String(err?.message || err || '').toLowerCase()
    return msg.includes('inappropriate content') || msg.includes('inappropriate-content')
  }

  function extractHttpStatusFromErrorMessage(err) {
    const msg = String(err?.message || err || '')
    const m = msg.match(/\bHTTP\s+(\d{3})\b/i)
    return m ? Number(m[1]) : null
  }

  function parseJsonFromModelText(text) {
    if (!text) return null
    try {
      const jsonMatch = String(text).match(/```json\s*([\s\S]*?)\s*```/i)
      if (jsonMatch?.[1]) return JSON.parse(jsonMatch[1])
      return JSON.parse(String(text))
    } catch {
      return null
    }
  }

  // IPC: 调用 Gemini 生成文生图提示词（统一使用主模型 API）

  return {
    async renderPreview(filePath) {
        try {
          if (!filePath || typeof filePath !== 'string') {
            return { success: false, error: '缺少 filePath' }
          }
          if (!fs.existsSync(filePath)) {
            return { success: false, error: '文件不存在' }
          }
          if (path.extname(filePath).toLowerCase() !== '.pptx') {
            return { success: false, error: '仅支持 .pptx' }
          }

          const cacheDir = getPptxPreviewCacheDir(filePath)
          if (fs.existsSync(cacheDir)) {
            const cached = listPngFilesSorted(cacheDir)
            if (cached.length > 0) {
              return { success: true, images: cached, cacheDir, cached: true }
            }
          }

          const result = await renderPptxToPngsWithLibreOffice(filePath, cacheDir)
          if (!result.success) {
            return result
          }
          return { success: true, images: result.images, cacheDir, cached: false }
        } catch (error) {
          console.error('[PPTX] render preview failed:', error)
          return { success: false, error: error.message || String(error) }
        }
    },

    async textEditHealth(options = {}) {
        const result = await textSidecar.health(options)
        if (!result?.success) return result
        return {
          ...result,
          deterministicRendererAvailable: true,
          cleanupEngines: Array.isArray(result.cleanupEngines) ? result.cleanupEngines : ['analytic_fill', 'local_inpaint'],
          externalAdapters: getExternalAdapters(),
        }
    },

    async detectTextLayer(options = {}) {
        try {
          const { pptxPath, pageNumber, useCache = true, cacheOnly = false } = options
          if (!pptxPath || typeof pptxPath !== 'string') {
            return { success: false, error: '缺少 pptxPath' }
          }
          if (!fs.existsSync(pptxPath)) {
            return { success: false, error: `PPTX 文件不存在: ${pptxPath}` }
          }
          if (!Number.isFinite(Number(pageNumber)) || Number(pageNumber) <= 0) {
            return { success: false, error: '缺少有效页码' }
          }

          const result = await detectTextLayerInternal({
            pptxPath,
            pageNumber: Number(pageNumber),
            useCache,
            cacheOnly,
          })
          clearPptTextEditProgress({
            stage: 'idle',
            progress: 1,
            message: '文字识别完成',
            pageNumber: Number(pageNumber),
          })
          return result
        } catch (error) {
          console.error('[PPT Text] detect failed:', error)
          clearPptTextEditProgress({
            stage: 'error',
            message: error.message || String(error),
            pageNumber: Number(options.pageNumber || 0),
          })
          return { success: false, error: error.message || String(error) }
        }
    },

    async applyTextEdits(options = {}) {
        try {
          const { pptxPath, pageNumber, edits = [] } = options
          if (!pptxPath || typeof pptxPath !== 'string') {
            return { success: false, error: '缺少 pptxPath' }
          }
          if (!fs.existsSync(pptxPath)) {
            return { success: false, error: `PPTX 文件不存在: ${pptxPath}` }
          }
          if (!Number.isFinite(Number(pageNumber)) || Number(pageNumber) <= 0) {
            return { success: false, error: '缺少有效页码' }
          }
          if (!Array.isArray(edits) || edits.length === 0) {
            return { success: false, error: '缺少 edits' }
          }

          emitPptTextEditProgress({
            stage: 'applying',
            progress: 0.02,
            message: `正在加载第 ${pageNumber} 页改字上下文...`,
            pageNumber: Number(pageNumber),
            editsTotal: edits.length,
          })

          const detection = await detectTextLayerInternal({
            pptxPath,
            pageNumber: Number(pageNumber),
            useCache: true,
            cacheOnly: false,
          })
          if (!detection.success) {
            clearPptTextEditProgress({
              stage: 'error',
              message: detection.error || '文字识别失败',
              pageNumber: Number(pageNumber),
            })
            return detection
          }

          const assetsDir = detection.assetsDir || getDeckAssetsDir(pptxPath)
          const source = await ensurePageSourceImage(pptxPath, Number(pageNumber), assetsDir)
          const logsDir = getTextEditLogsDir(assetsDir)
          if (!fs.existsSync(logsDir)) {
            fs.mkdirSync(logsDir, { recursive: true })
          }

          const seq = String(pageNumber).padStart(2, '0')
          const timestamp = new Date().toISOString().replace(/[:.]/g, '-')
          const outputPath = path.join(assetsDir, `slide_${seq}_1920x1200_textedit_${timestamp}.png`)
          const boxes = Array.isArray(detection.boxes) ? detection.boxes : []
          const boxMap = new Map(boxes.map((box) => [box.boxId, box]))
          let workingBuffer = source.imageBuffer
          const logs = []
          const perBoxCandidates = {}
          const perBoxApplied = {}
          const totalCandidateEstimate = edits.reduce((sum, edit) => {
            const box = boxMap.get(edit.boxId)
            if (!box) return sum
            const cleanupCount = box.backgroundComplexity === 'simple' ? 2 : 2
            const fontCount = Math.min(3, Array.isArray(box.fontCandidates) ? box.fontCandidates.length : 3)
            return sum + cleanupCount * fontCount
          }, 0) || Math.max(1, edits.length * 3)
          let completedCandidateUnits = 0

          for (const edit of edits) {
            const box = boxMap.get(edit.boxId)
            if (!box) {
              logs.push({ boxId: edit.boxId, success: false, error: 'box not found' })
              continue
            }

            emitPptTextEditProgress({
              stage: 'cleanup',
              progress: 0.08 + (completedCandidateUnits / totalCandidateEstimate) * 0.6,
              message: `正在清理文字框 ${box.readingOrder} 的原文字...`,
              pageNumber: Number(pageNumber),
              currentBoxId: box.boxId,
              editsTotal: edits.length,
              completedCandidates: completedCandidateUnits,
              totalCandidates: totalCandidateEstimate,
            })

            const styleOverride = edit.styleOverride || {}
            const styleEstimate = {
              ...(box.styleEstimate || box.styleHint || {}),
              ...styleOverride,
            }
            const rankedFontCandidates = (() => {
              const base = Array.isArray(box.fontCandidates) && box.fontCandidates.length
                ? [...box.fontCandidates]
                : [{
                    candidateId: `fallback:${slugify(styleEstimate.familyHint || 'default')}`,
                    family: styleEstimate.familyHint || 'PingFang SC',
                    confidence: 0.45,
                    source: 'system',
                  }]
              if (styleOverride.fontFamily && !base.some((item) => item.family === styleOverride.fontFamily)) {
                base.unshift({
                  candidateId: `manual:${slugify(styleOverride.fontFamily)}`,
                  family: styleOverride.fontFamily,
                  confidence: 0.99,
                  source: 'system',
                })
              }
              base.sort((left, right) => {
                if (styleOverride.fontCandidateId) {
                  if (left.candidateId === styleOverride.fontCandidateId) return -1
                  if (right.candidateId === styleOverride.fontCandidateId) return 1
                }
                return Number(right.confidence || 0) - Number(left.confidence || 0)
              })
              return base.slice(0, 3)
            })()

            const cleanupStrategies = styleOverride.cleanupStrategy
              ? [styleOverride.cleanupStrategy]
              : box.backgroundComplexity === 'simple'
                ? ['analytic_fill', 'local_inpaint']
                : box.backgroundComplexity === 'medium'
                  ? ['local_inpaint', 'analytic_fill']
                  : (() => {
                      const strategies = ['local_inpaint', 'analytic_fill']
                      if (getExternalAdapters().find((item) => item.name === 'adobe_firefly')?.available) {
                        strategies.unshift('adobe_firefly_fill')
                      }
                      return strategies
                    })()

            const candidateResults = []
            const cleanupRecognitionItems = []
            for (const cleanupStrategy of cleanupStrategies) {
              const blendStrategy = styleOverride.blendStrategy || chooseBlendStrategy(box)
              let cleanPageBuffer = workingBuffer
              let cleanupResidualText = ''
              let crop = await cropImageBuffer(workingBuffer, edit.bounds || box.bounds, 32)
              const originalCropBuffer = crop.cropBuffer

              if (cleanupStrategy === 'adobe_firefly_fill') {
                try {
                  const maskBuffer = await buildCleanupMaskCropBuffer({
                    box,
                    cropExtract: crop.extract,
                    strategy: cleanupStrategy,
                  })
                  const adobeCleanCropBuffer = await adobeFillCrop({
                    sourceBuffer: crop.cropBuffer,
                    maskBuffer,
                    prompt: 'Remove the masked text and continue the original background naturally. Preserve the poster layout, texture, lighting, and soft gradients. Do not add any new text, shapes, frames, or decorations.',
                  })
                  crop.cropBuffer = adobeCleanCropBuffer
                } catch (error) {
                  logs.push({
                    boxId: edit.boxId,
                    success: false,
                    error: `adobe_fill_failed:${error.message || String(error)}`,
                  })
                  continue
                }
              } else {
                const cleanupResult = await cleanupTextBoxesOnBuffer({
                  imageBuffer: workingBuffer,
                  boxes,
                  boxIds: [edit.boxId],
                  forceStrategy: cleanupStrategy,
                })
                if (!cleanupResult.success || !cleanupResult.buffer) {
                  continue
                }
                cleanPageBuffer = cleanupResult.buffer
                crop = await cropImageBuffer(cleanPageBuffer, edit.bounds || box.bounds, 32)
              }

              const cleanupRecognitionKey = `${edit.boxId}:${cleanupStrategy}`
              cleanupRecognitionItems.push({
                key: cleanupRecognitionKey,
                buffer: crop.cropBuffer,
              })

              const relativeBounds = {
                left: Math.max(0, (edit.bounds || box.bounds).left - crop.extract.left),
                top: Math.max(0, (edit.bounds || box.bounds).top - crop.extract.top),
                width: Math.max(1, (edit.bounds || box.bounds).width),
                height: Math.max(1, (edit.bounds || box.bounds).height),
              }

              for (const fontCandidate of rankedFontCandidates) {
                emitPptTextEditProgress({
                  stage: 'rendering',
                  progress: 0.1 + (completedCandidateUnits / totalCandidateEstimate) * 0.6,
                  message: `正在生成候选：${fontCandidate.family} · ${cleanupStrategy}...`,
                  pageNumber: Number(pageNumber),
                  currentBoxId: box.boxId,
                  completedCandidates: completedCandidateUnits,
                  totalCandidates: totalCandidateEstimate,
                })
                const render = await renderDeterministicTextLayer({
                  width: crop.extract.width,
                  height: crop.extract.height,
                  relativeBounds,
                  text: edit.toText,
                  styleEstimate: {
                    ...styleEstimate,
                    fontFamily: fontCandidate.family,
                  },
                  fontCandidate,
                })
                if (!render.success || !render.buffer) continue
                const normalizedRenderBuffer = await sharp(render.buffer)
                  .resize(crop.extract.width, crop.extract.height, {
                    fit: 'fill',
                  })
                  .png()
                  .toBuffer()

                const compositeCropBuffer = await sharp(crop.cropBuffer)
                  .composite([{ input: normalizedRenderBuffer, left: 0, top: 0 }])
                  .png()
                  .toBuffer()

                candidateResults.push({
                  candidateId: `${edit.boxId}:${cleanupStrategy}:${fontCandidate.candidateId}`,
                  boxId: edit.boxId,
                  label: `Deterministic · ${cleanupStrategy} · ${fontCandidate.family}`,
                  previewDataUrl: `data:image/png;base64,${compositeCropBuffer.toString('base64')}`,
                  fontCandidateId: fontCandidate.candidateId,
                  cleanupStrategy,
                  blendStrategy,
                  score: {
                    total: 0,
                    ocrExactness: 0,
                    fontStyleSimilarity: 0,
                    backgroundPreservation: 0,
                    edgeArtifactScore: 0,
                    overflowPenalty: 0,
                  },
                  applied: false,
                  metrics: {
                    fontFamily: fontCandidate.family,
                    fontConfidence: fontCandidate.confidence,
                    finalFontSize: render.metrics?.finalFontSize,
                    overflowX: render.metrics?.overflowX,
                    overflowY: render.metrics?.overflowY,
                  },
                  _compositeCropBuffer: compositeCropBuffer,
                  _cleanPageBuffer: cleanPageBuffer,
                  _extract: crop.extract,
                  _cleanupRecognitionKey: cleanupRecognitionKey,
                  _originalCropBuffer: originalCropBuffer,
                  _cleanCropBuffer: crop.cropBuffer,
                  _renderLayerBuffer: normalizedRenderBuffer,
                  _renderMetrics: render.metrics,
                  _fontCandidate: fontCandidate,
                })
                completedCandidateUnits += 1
              }
            }

            emitPptTextEditProgress({
              stage: 'scoring',
              progress: 0.72,
              message: `正在批量评分与选优...`,
              pageNumber: Number(pageNumber),
              currentBoxId: box.boxId,
              completedCandidates: completedCandidateUnits,
              totalCandidates: totalCandidateEstimate,
            })
            const cleanupRecognitionResults = await recognizeTextsFromBuffers(
              cleanupRecognitionItems.map((item) => item.buffer),
            )
            const cleanupRecognitionMap = new Map(
              cleanupRecognitionItems.map((item, index) => [
                item.key,
                cleanupRecognitionResults[index]?.text || '',
              ]),
            )

            const candidateRecognitionResults = await recognizeTextsFromBuffers(
              candidateResults.map((candidate) => candidate._compositeCropBuffer),
            )

            for (const [index, candidate] of candidateResults.entries()) {
              const cleanupResidualTextValue = cleanupRecognitionMap.get(candidate._cleanupRecognitionKey) || ''
              const score = await scoreEditCandidate({
                originalCropBuffer: candidate._originalCropBuffer,
                cleanCropBuffer: candidate._cleanCropBuffer,
                compositeCropBuffer: candidate._compositeCropBuffer,
                renderLayerBuffer: candidate._renderLayerBuffer,
                box,
                targetText: edit.toText,
                cleanupResidualText: cleanupResidualTextValue,
                cleanupStrategy: candidate.cleanupStrategy,
                fontCandidate: candidate._fontCandidate,
                renderMetrics: candidate._renderMetrics,
                ocrTextOverride: candidateRecognitionResults[index]?.text || '',
              })
              candidate.score = {
                total: score.total,
                ocrExactness: score.ocrExactness,
                fontStyleSimilarity: score.fontStyleSimilarity,
                backgroundPreservation: score.backgroundPreservation,
                edgeArtifactScore: score.edgeArtifactScore,
                overflowPenalty: score.overflowPenalty,
              }
              candidate.metrics = {
                ...candidate.metrics,
                detectedText: score.detectedText,
                cleanupResidualText: cleanupResidualTextValue,
                cleanupResidualSimilarity: score.cleanupResidualSimilarity,
                cleanupImpactPenalty: score.cleanupImpactPenalty,
              }
            }

            candidateResults.sort((left, right) => right.score.total - left.score.total)
            if (!candidateResults.length) {
              logs.push({
                boxId: edit.boxId,
                success: false,
                error: 'no deterministic candidates',
              })
              continue
            }

            const bestCandidate = candidateResults[0]
            bestCandidate.applied = true
            workingBuffer = await compositeCropBack(
              bestCandidate._cleanPageBuffer,
              bestCandidate._compositeCropBuffer,
              bestCandidate._extract,
            )

            perBoxApplied[edit.boxId] = bestCandidate
            perBoxCandidates[edit.boxId] = candidateResults.map((candidate) => ({
              candidateId: candidate.candidateId,
              boxId: candidate.boxId,
              label: candidate.label,
              previewDataUrl: candidate.previewDataUrl,
              fontCandidateId: candidate.fontCandidateId,
              cleanupStrategy: candidate.cleanupStrategy,
              blendStrategy: candidate.blendStrategy,
              score: candidate.score,
              applied: candidate.applied,
              metrics: candidate.metrics,
            }))

            logs.push({
              boxId: edit.boxId,
              success: true,
              cleanupStrategy: bestCandidate.cleanupStrategy,
              blendStrategy: bestCandidate.blendStrategy,
              appliedCandidateId: bestCandidate.candidateId,
              score: bestCandidate.score.total,
            })
          }

          const successfulLogs = logs.filter((item) => item.success)
          if (successfulLogs.length === 0) {
            clearPptTextEditProgress({
              stage: 'error',
              message: '没有成功生成任何可用候选',
              pageNumber: Number(pageNumber),
            })
            return {
              success: false,
              error: '没有成功生成任何可用候选',
              fallbackSuggested: true,
              logs,
            }
          }

          const outputBuffer = workingBuffer
          emitPptTextEditProgress({
            stage: 'writing',
            progress: 0.9,
            message: `正在写回 PPT 与刷新预览...`,
            pageNumber: Number(pageNumber),
          })
          fs.writeFileSync(outputPath, outputBuffer)
          const outputResultPath = outputPath
          const engine = 'deterministic_pipeline'

          await replaceSlideImagesInPptx(
            pptxPath,
            [{ pageIndex: Number(pageNumber) - 1, imageBuffer: outputBuffer }],
            true,
          )

          const cachePath = getTextLayerCachePath(assetsDir, Number(pageNumber))
          const nextBoxes = Array.isArray(detection.boxes)
            ? enrichDetectedTextBoxes(detection.boxes.map((box) => {
                const edit = edits.find((item) => item.boxId === box.boxId)
                if (!edit) return box
                return {
                  ...box,
                  text: edit.toText,
                  bounds: edit.bounds || box.bounds,
                  styleEstimate: {
                    ...(box.styleEstimate || box.styleHint || {}),
                    ...(edit.styleOverride || {}),
                  },
                }
              }))
            : []
          writeTextLayerCache(cachePath, {
            cacheVersion: getTextLayerCacheVersion(),
            imageHash: hashBuffer(outputBuffer),
            canvasWidth: detection.canvasWidth,
            canvasHeight: detection.canvasHeight,
            boxes: nextBoxes,
            updatedAt: new Date().toISOString(),
            engine,
          })

          fs.writeFileSync(
            path.join(logsDir, `slide_${seq}_text_edit_${timestamp}.json`),
            JSON.stringify(
              {
                pageNumber: Number(pageNumber),
                edits,
                logs,
                perBoxCandidates,
                outputPath: outputResultPath,
                appliedAt: new Date().toISOString(),
                engine,
              },
              null,
              2,
            ),
          )

          return {
            success: true,
            path: pptxPath,
            imageDataUrl: `data:image/png;base64,${outputBuffer.toString('base64')}`,
            editedBoxes: edits.map((item) => item.boxId),
            appliedCandidateId: edits.length === 1 ? perBoxApplied[edits[0].boxId]?.candidateId : 'multi',
            candidateCount: Object.values(perBoxCandidates).reduce((sum, list) => sum + list.length, 0),
            fontMatchConfidence: edits.length
              ? edits.reduce((sum, edit) => {
                  const applied = perBoxApplied[edit.boxId]
                  const value = Number(applied?.metrics?.fontConfidence || 0)
                  return sum + value
                }, 0) / edits.length
              : 0,
            cleanupStrategy: edits.length === 1 ? perBoxApplied[edits[0].boxId]?.cleanupStrategy : 'none',
            blendStrategy: edits.length === 1 ? perBoxApplied[edits[0].boxId]?.blendStrategy : 'none',
            candidates: edits.length === 1 ? (perBoxCandidates[edits[0].boxId] || []) : [],
            perBoxCandidates,
            logs,
          }
        } catch (error) {
          console.error('[PPT Text] apply failed:', error)
          clearPptTextEditProgress({
            stage: 'error',
            message: error.message || String(error),
            pageNumber: Number(options.pageNumber || 0),
          })
          return {
            success: false,
            error: error.message || String(error),
            fallbackSuggested: true,
          }
        } finally {
          clearPptTextEditProgress({
            stage: 'idle',
            progress: 1,
            message: '改字流程结束',
            pageNumber: Number(options.pageNumber || 0),
          })
        }
    },

    async generatePrompts(options = {}) {
        try {
          const { outline, theme, style, mainApiKey } = options
          
          if (!mainApiKey) {
            return { success: false, error: '缺少主模型 API Key，请在设置中配置' }
          }
          if (!outline) {
            return { success: false, error: '缺少 PPT 大纲' }
          }

          const systemPrompt = `你是一位专精于 **高端品牌视觉设计** 的顶级艺术总监。你的任务是编写 **极其详细、视觉元素丰富** 的 AI 绘画提示词，用于生成一张 **直接包含完整 PPT 内容** 的幻灯片成片。

      ⚠️ **核心目标：视觉丰富度（VISUAL RICHNESS）** ⚠️
      - 每张图片必须像 **Dribbble/Behance 上获奖的品牌提案** 那样精致、层次丰富、细节饱满。
      - 绝不能是"文字+简单背景"的单调设计，必须有 **多层次视觉元素堆叠**。

      ## 🎨 视觉丰富度铁律（每页必须全部满足）

      ### 1) 多层次构图（Layered Composition）- 必须 5+ 层
      每页至少包含以下层次（从后到前）：
      - **Layer 1 背景层**：渐变/纹理/图案（绝不能是纯色）
      - **Layer 2 氛围层**：大面积模糊光斑、柔和渐变云、抽象几何形状
      - **Layer 3 装饰层**：网格线、几何图形、抽象元素、图标阵列
      - **Layer 4 主视觉层**：与主题相关的核心插图/图形/3D元素
      - **Layer 5 内容层**：毛玻璃卡片承载的文字内容

      ### 2) 主题创意元素（Thematic Visual Elements）- 必须 2-3 个
      根据 PPT 主题，必须加入 **与内容直接相关的创意视觉元素**：
      - 科技主题：电路线条、数据流粒子、代码片段装饰、芯片纹理、光纤线
      - 商业主题：图表元素、上升箭头、齿轮连接、网络节点、增长曲线
      - 教育主题：书本元素、灯泡图标、知识树、公式装饰、学术符号
      - 创意主题：画笔笔触、色彩飞溅、艺术纹理、创意工具图标
      - 自然主题：植物剪影、水波纹理、有机曲线、自然光影
      - **必须在 prompt 中明确描述这些元素的位置、大小、颜色和透明度**

      ### 3) 微元素密度（Micro-Detail Density）- 必须 8+ 种
      每页必须包含大量低透明度装饰元素（10-30% opacity），从以下清单中选择至少 8 种：
      - □ 极细网格线（ultra-thin grid lines, 0.5px）
      - □ 角标/裁切标记（corner marks, registration marks）
      - □ 页码序列号（No.01, VOL.25, SLIDE 03）
      - □ 微型图标阵列（tiny icons array, 16px）
      - □ 抽象条形码/二维码装饰（abstract barcode pattern）
      - □ 点阵图案（dot matrix pattern）
      - □ 细分隔线/引导线（thin dividers, guide lines）
      - □ 浮动几何小块（floating geometric shapes）
      - □ 数据可视化元素（mini charts, data points, progress bars）
      - □ 渐变光晕/光斑（gradient orbs, soft glows）
      - □ 纹理叠加（noise texture, paper grain, fabric weave）
      - □ 连接线/流程线（connecting lines, flow paths）
      - □ 时间轴元素（timeline markers, date stamps）
      - □ 标签/徽章装饰（label badges, status indicators）
      - □ 波形/脉冲线（waveforms, pulse lines）

      ### 4) 色彩层次（Color Depth）
      - 主背景：Off-white (#F5F3EE) 到 Warm Gray (#E8E4DF) 的微妙渐变
      - 必须有 2-3 个不同透明度的装饰色层
      - 一个鲜明但克制的强调色（面积≤8%）
      - 阴影必须是暖灰色调，不能是纯黑

      ### 5) 材质与质感（Materiality）
      - 毛玻璃卡片：blur 20-40px, 白色 60-80% 透明度, 1px 白色边框
      - 多层柔和阴影：近影 + 远影 创造立体感
      - 背景必须有可见纹理：纸纹/布纹/噪点（5-15% opacity）

      ### 6) 文字规范
      - 中文必须清晰可读（crisp Chinese text, elegant sans-serif）
      - 只包含大纲提供的文字，禁止随机内容
      - 文字有呼吸空间，行距 1.5+

      ## ✅ Prompt 格式要求

      每条 prompt 必须 **600-900 字符**，结构如下：
      1. 整体场景描述（overall scene）
      2. 背景层详细描述（background layer details）
      3. 装饰元素详细描述（decorative elements with positions）
      4. 主题视觉元素描述（thematic visual elements）
      5. 内容卡片描述（content card with glassmorphism）
      6. 完整的中文文字内容（exact Chinese text）
      7. 色彩和光影描述（colors, lighting, shadows）
      8. 风格关键词（style keywords）

      ## ✅ 输出格式（严格 JSON）
      你必须只输出 JSON（可用 \`\`\`json 代码块包裹），结构：
      {
        "designConcept": "整体视觉策略：选用的风格 + 核心视觉元素 + 统一的配色方案",
        "colorPalette": "具体色值：背景色 + 装饰色 + 强调色",
        "slides": [
          {
            "pageNumber": 1,
            "pageType": "cover/content/summary",
            "visualConcept": "本页视觉创意：使用哪些主题元素 + 如何体现内容",
            "prompt": "极其详细的英文提示词（600-900字符）",
            "negativePrompt": "负面词"
          }
        ]
      }

      ## 负面词（必须包含）
      negativePrompt 必须包含：deformed text, broken text, malformed letters, illegible text, garbled Chinese, wrong Chinese characters, ugly typography, dark background, pure black, neon glow, cyberpunk, hologram, messy layout, cluttered, watermark, logo, brand mark, lowres, blurry, cheap, plastic, amateur, empty, minimal, simple, plain, boring, flat design without depth

      ---

      下面是你要处理的具体大纲与风格偏好。`

          const userPrompt = `请为以下 PPT 大纲设计视觉方案并生成文生图提示词（每页一张成片）：

      ## PPT 主题/用途
      ${theme || '（未指定，请根据大纲内容判断）'}

      ## 用户期望的风格倾向
      ${style || '你可在风格库 A-F 中自动选择一个最匹配大纲的高级风格；也可以在同体系内做少量变奏（保持统一审美）'}

      ## PPT 大纲内容
      ${outline}

      ---
      要求：
      1) 先选择最匹配该大纲的主风格 preset（A-F），在 designConcept 里说明原因；必要时可“同体系变奏”，但不要乱混风格  
      2) **反廉价/反AI味（强制）**：避免“塑料感/玩具感/廉价霓虹/模板化等距城市/素材库风 3D 图标”。整体要像品牌 KV / 杂志海报  
      3) **配色（强制）**：给出“美术生审美”的配色——低饱和主色 + 中性色 + 1 个点睛色；避免过饱和、刺眼荧光、廉价蓝紫霓虹  
      4) **主题创意元素（强制）**：每页除背景+文字外，至少加入 1-2 个与主题直接相关的创意视觉元素/隐喻（例如：城市路网拓扑线、地图纹理、时间轴丝带、印章纹理、建筑剖面线稿、数据流粒子等），而不是随机几何装饰  
      5) 每页 prompt 必须包含该页所有中文文案（标题/副标题/要点/页脚），并强调中文清晰可读；**禁止出现大纲之外的任何文字**（尤其随机英文缩写/乱码）  
      6) 图片中只能有设计元素和用户内容：**禁止任何品牌/Logo/软件界面/水印/角标**  
      7) 每条 prompt ≤900 chars，negativePrompt ≤400 chars，并在 negativePrompt 中**必须**加入：deformed text, broken text, malformed letters, illegible text, garbled Chinese, wrong Chinese characters, ugly typography, cheap plastic, toy-like, lowres, blurry, amateur, neon cyberpunk, circuit board, watermark, logo  
      8) **最终输出必须是严格的 JSON 格式**（参考 system prompt 中的格式），不要输出任何解释性文字，只输出 JSON 代码块。`

          // 统一使用当前主模型配置调用 Gemini 类文本模型生成 PPT 提示词
          let response = ''
          
          if (!mainApiKey) {
            return { success: false, error: '缺少主模型 API Key，请在设置中配置' }
          }
          
          console.log('[PPT Prompts] 使用当前主模型配置生成 PPT 提示词')
          response = await callConfiguredGeminiText({
            apiKey: mainApiKey,
            baseUrl: options.mainBaseUrl || 'https://api.linapi.net/v1',
            model: options.mainModel || 'gemini-3.1-pro-preview',
            systemPrompt,
            userPrompt,
          })

          // 解析JSON响应
          let parsed = null
          try {
            // 尝试提取JSON块
            const jsonMatch = response.match(/```json\s*([\s\S]*?)\s*```/)
            if (jsonMatch) {
              parsed = JSON.parse(jsonMatch[1])
            } else {
              // 尝试直接解析
              parsed = JSON.parse(response)
            }
          } catch (parseError) {
            console.error('Gemini response parse error:', parseError)
            return { success: false, error: 'Gemini 返回的内容无法解析为JSON', raw: response }
          }

          const normalizedSlides = Array.isArray(parsed?.slides)
            ? parsed.slides.map((s, idx) => ({
                pageNumber: Number(s?.pageNumber) || idx + 1,
                pageType: String(s?.pageType || 'content'),
                visualConcept: typeof s?.visualConcept === 'string' ? s.visualConcept : '',
                prompt: String(s?.prompt || ''),
                negativePrompt: mergeNegativePrompt(s?.negativePrompt),
              }))
            : parsed?.slides

          return {
            success: true,
            slides: normalizedSlides,
            designConcept: parsed.designConcept || '',
            colorPalette: parsed.colorPalette || '',
            raw: response,
          }
        } catch (error) {
          console.error('openrouter-gemini-ppt-prompts error:', error)
          return { success: false, error: error.message || String(error) }
        }
    },

    async generateDeck(options = {}) {
        try {
          const {
            outputPath,
            slides = [],
            mainApiKey = '', // 主模型 API Key（用于 Gemini 生图）
            dashscope = {},
            postprocess = { mode: 'letterbox' },
            repair = {},
            outline = null, // 原始大纲（用于保存元数据）
          } = options
          
          // 用于收集每页最终使用的 prompt（含修复后的）
          const finalSlidesPrompts = []

          if (!outputPath || typeof outputPath !== 'string') {
            return { success: false, error: '缺少 outputPath' }
          }
          if (!outputPath.toLowerCase().endsWith('.pptx')) {
            return { success: false, error: 'outputPath 必须以 .pptx 结尾' }
          }
          if (!Array.isArray(slides) || slides.length === 0) {
            return { success: false, error: 'slides 不能为空' }
          }

          const limit = pLimit(2) // DashScope: RPS=2 且并发=2（两张两张生成）
          const geminiRepairLimit = pLimit(1) // Gemini 维修：串行，保证上下文一致
          // 用户选择的图像生成模型
          const imageModel = dashscope.model || 'z-image-turbo'
          // 根据模型选择默认分辨率
          const defaultSize = imageModel === 'z-image-turbo' ? '2048*1152' : '1664*928'
          const size = dashscope.size || defaultSize
          const promptExtend = !!dashscope.promptExtend
          const watermark = dashscope.watermark === true
          const negativePromptDefault = dashscope.negativePromptDefault || ''
          const region = dashscope.region || 'cn'
          const apiKey = dashscope.apiKey || ''
          const saveImages = dashscope.saveImages !== false // 默认保存，便于排查"是否真的生成了图片"
          
          console.log(`[PPT Generate] 使用模型: ${imageModel}, 分辨率: ${size}`)

          const repairEnabled =
            repair?.enabled !== false &&
            !!repair?.openRouterApiKey &&
            typeof repair?.openRouterApiKey === 'string' &&
            repair.openRouterApiKey.trim().length > 10
          const repairMaxAttempts = Math.max(0, Math.min(5, Number(repair?.maxAttempts ?? 2)))
          const repairModel = repair?.model || 'google/gemini-3-pro-preview'
          const deckContext = repair?.deckContext || {}

          // 维修会话上下文：用于“只修失败页”的连续对话（串行执行，避免并发污染上下文）
          const geminiRepairMessages = repairEnabled
            ? [
                {
                  role: 'system',
                  content:
                    'You are a world-class presentation designer and prompt engineer. ' +
                    'We are generating poster-style PPT slide images (text is part of the image). ' +
                    'Some slides may fail DashScope safety moderation (inappropriate content). ' +
                    'Your job: REWRITE ONLY the failed slide prompt to pass moderation while keeping the same deck style.\n' +
                    '\n' +
                    'Rules:\n' +
                    '- Keep the overall style consistent with the deck design concept and color palette.\n' +
                    '- Keep Chinese text crisp & legible. Prefer keeping the exact Chinese copy; if any phrase is likely to trigger moderation, paraphrase into neutral, compliant wording while preserving meaning.\n' +
                    '- Avoid any violence/politics/sensitive content. Avoid brand names, logos, UI, watermarks.\n' +
                    '- Output JSON ONLY: {"prompt":"...","negativePrompt":"...","textEdits":[{"from":"...","to":"..."}]}\n' +
                    '- prompt <= 800 chars, negativePrompt <= 300 chars.',
                },
                {
                  role: 'user',
                  content:
                    'Deck context (keep consistent):\n' +
                    `- designConcept: ${String(deckContext.designConcept || '').slice(0, 800)}\n` +
                    `- colorPalette: ${String(deckContext.colorPalette || '').slice(0, 200)}\n` +
                    'Remember: do NOT regenerate the whole deck. We will request single-slide repairs as needed.',
                },
              ]
            : []

          async function repairSlidePromptWithGemini({ idx, attempt, prompt, negativePrompt, errorMessage }) {
            return await geminiRepairLimit(async () => {
              const slideNo = idx + 1
              const userMsg =
                `Slide repair request:\n` +
                `- slideNumber: ${slideNo}\n` +
                `- dashscopeError: ${String(errorMessage).slice(0, 800)}\n` +
                `- previousPrompt: ${String(prompt).slice(0, 4000)}\n` +
                `- previousNegativePrompt: ${String(negativePrompt || '').slice(0, 1200)}\n` +
                '\n' +
                'Rewrite a safer prompt that preserves layout and typography, keeps deck style consistent, and avoids moderation triggers. Output JSON only.'

              geminiRepairMessages.push({ role: 'user', content: userMsg })
              
              let responseText = ''
              try {
                responseText = await callOpenRouterGemini({
                  apiKey: repair.openRouterApiKey,
                  model: repairModel,
                  messages: geminiRepairMessages,
                })
              } catch (geminiErr) {
                console.warn(`[PPT Repair] Gemini 调用异常 (slide=${slideNo}, attempt=${attempt}):`, geminiErr?.message || geminiErr)
                throw new Error(`Gemini 修复调用失败（slide=${slideNo}）: ${geminiErr?.message || '网络错误'}`)
              }
              
              // 检查空响应（Gemini 有时会返回空内容）
              if (!responseText || responseText.trim().length < 10) {
                console.warn(`[PPT Repair] Gemini 返回空响应 (slide=${slideNo}, attempt=${attempt})`)
                throw new Error(`Gemini 修复返回空响应（slide=${slideNo}），请重试`)
              }
              
              geminiRepairMessages.push({ role: 'assistant', content: responseText })

              const parsed = parseJsonFromModelText(responseText)
              const newPrompt = parsed?.prompt
              const newNegative = parsed?.negativePrompt
              if (!newPrompt || typeof newPrompt !== 'string' || newPrompt.trim().length < 20) {
                console.warn(`[PPT Repair] Gemini 返回无效 JSON (slide=${slideNo}):`, responseText?.slice(0, 500))
                throw new Error(`Gemini 修复提示词失败：无法解析 prompt（slide=${slideNo}, attempt=${attempt}）`)
              }
              return {
                prompt: String(newPrompt).trim(),
                negativePrompt: typeof newNegative === 'string' ? String(newNegative).trim() : String(negativePrompt || '').trim(),
                textEdits: Array.isArray(parsed?.textEdits) ? parsed.textEdits : [],
                raw: responseText,
              }
            })
          }

          // 把每页下载到的原始图片 & 后处理后的 1920x1080 PNG 保存到本地，便于排查
          const outDir = path.dirname(outputPath)
          const baseName = path.basename(outputPath, path.extname(outputPath))
          const assetsDir = path.join(outDir, `${baseName}_assets`)
          if (saveImages && !fs.existsSync(assetsDir)) {
            fs.mkdirSync(assetsDir, { recursive: true })
          }
          const results = await Promise.all(
            slides.map((s, idx) =>
              limit(async () => {
                let prompt = s.prompt || s.finalPrompt || s.finalPromptCNorEN || ''
                let negativePrompt = s.negativePrompt ?? negativePromptDefault

                const seq = String(idx + 1).padStart(2, '0')
                const promptPathBase = saveImages ? path.join(assetsDir, `slide_${seq}_prompt`) : null

                // 单页重试：遇到审核失败（inappropriate content）→ 把该页失败信息交给 Gemini 改写提示词 → 仅重试该页
                let attempt = 0
                while (true) {
                  try {
                    // 保存当前尝试的 prompt（便于对比）
                    if (saveImages && promptPathBase) {
                      try {
                        fs.writeFileSync(`${promptPathBase}_attempt${attempt}.txt`, String(prompt))
                        fs.writeFileSync(`${promptPathBase}_neg_attempt${attempt}.txt`, String(negativePrompt || ''))
                      } catch {}
                    }

                    let raw
                    let imageSource = '' // 记录图片来源（URL 或 'gemini-base64'）
                    if (imageModel === 'gemini-image') {
                      // 使用 LinAPI Gemini 生图（需要主模型 API Key）
                      // 直接使用 Gemini 3 Pro 生成的原始提示词，不做额外增强
                      if (!mainApiKey) {
                        throw new Error('使用 Gemini 生图需要配置主模型 API Key（LinAPI）')
                      }
                      console.log(`\n${'='.repeat(60)}`)
                      console.log(`[PPT Generate] ✅ 使用 gemini-3-pro-image-preview-2K 生图`)
                      console.log(`[PPT Generate] Slide ${idx + 1}/${slides.length}`)
                      console.log(`[PPT Generate] 提示词长度: ${prompt.length} chars`)
                      console.log(`${'='.repeat(60)}\n`)
                      
                      const geminiResult = await linapiGenerateImage({
                        apiKey: mainApiKey,
                        prompt: prompt, // 直接使用原始提示词
                        aspectRatio: '16:9',
                      })
                      raw = Buffer.from(geminiResult.base64, 'base64')
                      imageSource = 'gemini-base64'
                      console.log(`[PPT Generate] Slide ${idx + 1} 生图完成，图片大小: ${raw.length} bytes`)
                    } else {
                      // 使用 DashScope 生图
                      const { url } = await dashscopeGenerateImageUrl({
                        prompt,
                        negativePrompt,
                        size,
                        promptExtend,
                        watermark,
                        model: imageModel,
                        region,
                        apiKey,
                      })
                      raw = await downloadToBuffer(url)
                      imageSource = url // 记录 URL 来源
                    }
                    if (!raw || raw.length === 0) {
                      throw new Error(`图片下载失败或为空（idx=${idx}）`)
                    }
                    const processed = await postprocessTo1920x1200(raw, postprocess?.mode || 'letterbox')
                    if (!processed || processed.length === 0) {
                      throw new Error(`图片后处理失败或为空（idx=${idx}）`)
                    }

                    // 保存图片到本地（用于排查是否真实生成/下载/后处理成功）
                    if (saveImages) {
                      // 根据图片来源决定文件扩展名
                      let ext = '.jpg' // 默认为 jpg
                      if (imageSource && imageSource !== 'gemini-base64') {
                        try {
                          const u = new URL(imageSource)
                          ext = path.extname(u.pathname).toLowerCase() || '.jpg'
                        } catch {}
                      }
                      if (!ext || ext.length > 5) ext = '.jpg'
                      const rawPath = path.join(assetsDir, `slide_${seq}_raw_attempt${attempt}${ext}`)
                      const pngPath = path.join(assetsDir, `slide_${seq}_1920x1080_attempt${attempt}.png`)
                      const sourcePath = path.join(assetsDir, `slide_${seq}_source_attempt${attempt}.txt`)
                      try {
                        fs.writeFileSync(rawPath, raw)
                        fs.writeFileSync(pngPath, processed)
                        fs.writeFileSync(sourcePath, imageSource === 'gemini-base64' ? 'gemini-3-pro-image-preview-2K (base64)' : String(imageSource))
                      } catch (e) {
                        console.warn('[PPTX] 保存图片失败:', e?.message || e)
                      }
                    }

                    const base64 = processed.toString('base64')
                    const dataUri = `image/png;base64,${base64}`

                    return { idx, dataUri, finalPrompt: prompt, finalNegativePrompt: negativePrompt, attempts: attempt + 1 }
                  } catch (e) {
                    const errorMessage = e?.message || String(e)
                    const status = extractHttpStatusFromErrorMessage(e)
                    const isInappropriate = status === 400 && isDashscopeInappropriateContentError(e)

                    if (saveImages) {
                      try {
                        fs.writeFileSync(path.join(assetsDir, `slide_${seq}_error_attempt${attempt}.txt`), String(errorMessage))
                      } catch {}
                    }

                    if (!repairEnabled || !isInappropriate || attempt >= repairMaxAttempts) {
                      throw e
                    }

                    // 触发 Gemini 单页修复
                    const repairRes = await repairSlidePromptWithGemini({
                      idx,
                      attempt,
                      prompt,
                      negativePrompt,
                      errorMessage,
                    })

                    if (saveImages) {
                      try {
                        fs.writeFileSync(path.join(assetsDir, `slide_${seq}_repair_response_attempt${attempt}.txt`), String(repairRes.raw || ''))
                        if (Array.isArray(repairRes.textEdits) && repairRes.textEdits.length) {
                          fs.writeFileSync(
                            path.join(assetsDir, `slide_${seq}_repair_text_edits_attempt${attempt}.json`),
                            JSON.stringify(repairRes.textEdits, null, 2)
                          )
                        }
                      } catch {}
                    }

                    prompt = repairRes.prompt
                    negativePrompt = repairRes.negativePrompt || negativePrompt
                    attempt += 1
                    continue
                  }
                }
              })
            )
          )

          results.sort((a, b) => a.idx - b.idx)
          const images = results.map((r) => r.dataUri)
          
          // 收集每页最终的 prompt 信息
          const slidesPromptsData = results.map((r, i) => ({
            pageNumber: i + 1,
            prompt: r.finalPrompt || slides[r.idx]?.prompt || '',
            negativePrompt: r.finalNegativePrompt || slides[r.idx]?.negativePrompt || '',
            attempts: r.attempts || 1,
            originalChineseContent: slides[r.idx]?.originalChineseContent || '',
          }))

          // Ensure directory exists
          const dir = path.dirname(outputPath)
          if (!fs.existsSync(dir)) {
            fs.mkdirSync(dir, { recursive: true })
          }

          await makePptx16x10FromImagesBase64(images, outputPath)
          
          // 保存元数据到 _assets 目录（用于后续编辑）
          if (saveImages) {
            saveDeckMetadata(assetsDir, {
              deckContext: deckContext,
              slidesPrompts: slidesPromptsData,
              outline: outline,
            })
          }

          return { success: true, path: outputPath, slideCount: slides.length }
        } catch (error) {
          console.error('ppt-generate-deck failed:', error)
          return { success: false, error: error.message || String(error) }
        }
    },

    async editSlides(options = {}) {
        try {
          const {
            pptxPath,           // PPTX 文件路径
            pageNumbers = [],   // 要编辑的页码数组（1-based）
            feedback = '',      // 用户反馈
            mode = 'regenerate', // 'regenerate' = 整页重做，'partial_edit' = 局部编辑
            openRouterApiKey,   // Gemini API Key
            dashscopeApiKey,    // DashScope API Key
            mainApiKey,         // 主模型 API Key（用于 LinAPI Gemini 生图）
            pptImageModel = 'z-image-turbo', // 生图模型选择
            deckContext: providedDeckContext, // 可选，优先使用提供的
            regionScreenshot,   // 新增：用户框选区域的截图 base64
            regionRect,         // 新增：框选区域坐标 {x, y, w, h}
          } = options

          if (!pptxPath || typeof pptxPath !== 'string') {
            return { success: false, error: '缺少 pptxPath' }
          }
          if (!fs.existsSync(pptxPath)) {
            return { success: false, error: `PPTX 文件不存在: ${pptxPath}` }
          }
          if (!Array.isArray(pageNumbers) || pageNumbers.length === 0) {
            return { success: false, error: '缺少要编辑的页码' }
          }
          if (!feedback || typeof feedback !== 'string' || !feedback.trim()) {
            return { success: false, error: '缺少用户反馈' }
          }
          if (!openRouterApiKey) {
            return { success: false, error: '缺少 OpenRouter API Key' }
          }
          if (pptImageModel === 'gemini-image') {
            // Gemini 生图走 LinAPI，需要主模型 key（或复用 dashscopeApiKey 兜底，但推荐传 mainApiKey）
            if (!mainApiKey && !dashscopeApiKey) {
              return { success: false, error: '缺少 API Key：Gemini 生图需要 mainApiKey（或至少提供 dashscopeApiKey 兜底）' }
            }
          } else {
            // DashScope 生图/编辑仍需要 DashScope Key
            if (!dashscopeApiKey) {
              return { success: false, error: '缺少 DashScope API Key' }
            }
          }

          // 读取 _assets 元数据
          const baseName = path.basename(pptxPath, '.pptx')
          const assetsDir = path.join(path.dirname(pptxPath), `${baseName}_assets`)
          const metadata = loadDeckMetadata(assetsDir)
          
          const deckContext = providedDeckContext || metadata.deckContext || {}
          const slidesPrompts = metadata.slidesPrompts || []
          const outline = metadata.outline || {}
          
          // 获取大纲中的 slides 数组
          const outlineSlides = outline.slides || outline.pages || outline.content || []

          console.log(`[PPT Edit] 模式: ${mode}, 页码: ${pageNumbers.join(', ')}, 反馈: ${feedback.slice(0, 100)}...`)
          console.log(`[PPT Edit] 大纲页数: ${outlineSlides.length}, slidesPrompts: ${slidesPrompts.length}`)

          const editLimit = pLimit(1) // 串行编辑，保证 Gemini 上下文一致
          const replacements = []
          const editLogs = []

          for (const pageNum of pageNumbers) {
            await editLimit(async () => {
              const pageIndex = pageNum - 1
              if (pageIndex < 0) {
                editLogs.push({ pageNum, success: false, error: '页码无效' })
                return
              }

              // 获取该页的原始图片
              const originalImage = await getSlideImageFromPptx(pptxPath, pageIndex, assetsDir)
              if (!originalImage) {
                editLogs.push({ pageNum, success: false, error: '无法读取该页图片' })
                return
              }

              const originalImageBase64 = originalImage.toString('base64')
              const slidePromptInfo = slidesPrompts[pageIndex] || {}
              
              // 获取该页的大纲内容（非常重要：确保生成的图片包含正确的文字）
              const outlineSlide = outlineSlides[pageIndex] || {}
              const slideHeadline = outlineSlide.headline || outlineSlide.title || outlineSlide.heading || outlineSlide.pageTitle || outlineSlide.page_title || ''
              const slideSubheadline = outlineSlide.subheadline || outlineSlide.subtitle || outlineSlide.sub_title || ''
              const slideBullets = outlineSlide.bullets || outlineSlide.points || outlineSlide.content_points || outlineSlide.keyPoints || []
              const slideFooter = outlineSlide.footerNote || outlineSlide.footer || ''
              const slidePageType = outlineSlide.pageType || outlineSlide.page_type || 'content'
              const slideLayoutIntent = outlineSlide.layoutIntent || outlineSlide.layout_intent || ''
              
              // 构建该页的完整中文内容（用于 Gemini）
              let slideChineseContent = ''
              if (slideHeadline) slideChineseContent += `标题: "${slideHeadline}"\n`
              if (slideSubheadline) slideChineseContent += `副标题: "${slideSubheadline}"\n`
              if (Array.isArray(slideBullets) && slideBullets.length > 0) {
                slideChineseContent += `要点:\n${slideBullets.map((b, i) => `  ${i + 1}. "${b}"`).join('\n')}\n`
              }
              if (slideFooter) slideChineseContent += `页脚: "${slideFooter}"\n`
              
              // 如果大纲内容为空，尝试使用 slidesPrompts 中保存的内容
              if (!slideChineseContent.trim() && slidePromptInfo.originalChineseContent) {
                slideChineseContent = slidePromptInfo.originalChineseContent
              }
              
              console.log(`[PPT Edit] 第 ${pageNum} 页大纲内容:\n${slideChineseContent.slice(0, 500)}`)

              let newImageBuffer = null

              if (mode === 'regenerate') {
                // ========== 整页重做 ==========
                // 1. 让 Gemini 根据反馈重写 prompt（必须包含该页的中文内容 + 高级设计感）
                const geminiSystemPrompt = 
                  'You are a world-class presentation designer creating PREMIUM, AWARD-WINNING slide visuals. ' +
                  'The user is not satisfied with their current slide and wants it redesigned with MUCH BETTER aesthetics. ' +
                  '\n\n' +
                  '## YOUR DESIGN PHILOSOPHY\n' +
                  '- Think like top-tier keynote design teams - clean, sophisticated, highly curated\n' +
                  '- Use LAYERED DEPTH: paper/cards/scroll-tabs/panels with soft shadows and clear hierarchy (not sci-fi HUD)\n' +
                  '- Apply PREMIUM MATERIALS: matte ceramic, fine paper (xuan paper), silk, lacquer, jade/bronze accents, restrained gold foil lines\n' +
                  '- Create VISUAL HIERARCHY: clear focal point, breathing space, balanced composition\n' +
                  '- Add REFINED DETAILS: filmic tone mapping, subtle grain, micro-texture, restrained highlights\n' +
                  '\n' +
                  '## CRITICAL RULES\n' +
                  '1. ALL Chinese text from slide content MUST appear in the image (title, bullets, footer)\n' +
                  '2. Chinese text: crisp, high-contrast, elegant typography (not plain black on white!)\n' +
                  '3. Layout: asymmetric balance, golden ratio, generous margins\n' +
                  '4. For AGENDA/TOC pages: use creative layouts like numbered cards, timeline, floating panels\n' +
                  '5. Quote each Chinese text explicitly in your prompt\n' +
                  '6. Add 1-2 THEME-RELATED creative motifs (not random decoration)\n' +
                  '7. If user feedback asks for "古风/国风/东方/典雅/宋韵/新中式": MUST switch to a premium neo-Chinese heritage aesthetic (xuan paper, ink wash, seal stamp, antique map lines) and AVOID any tech/HUD/neon look\n' +
                  '\n' +
                  '## AESTHETIC TECHNIQUES TO USE\n' +
                  '- Neo-Chinese heritage (when requested): xuan paper texture, ink wash gradients, subtle cloud patterns, antique gold foil dividers, cinnabar seal accents\n' +
                  '- Depth layering: foreground text on refined panels over textured background (paper/ink wash)\n' +
                  '- Premium color harmony: low-chroma main hue + neutrals + one accent (art-school palette, filmic grading)\n' +
                  '- Elegant motifs: thin contour lines, map topology, architectural line art, seals, minimal ornaments tied to the topic\n' +
                  '- Lighting: soft global illumination, gentle vignetting, subtle shadowing (avoid neon glow rings)\n' +
                  '\n' +
                  '## HIGH-END DESIGN VOCAB (USE SELECTIVELY, NOT KEYWORD STUFFING)\n' +
                  '- Layout/Grid: International Typographic Style (Swiss), typographic grid, baseline grid, modular grid, strong alignment, consistent gutters, generous margins\n' +
                  '- Microtypography: microtypography, optical alignment, kerning, tracking, leading, typographic scale, clean line breaks\n' +
                  '- Premium finishes/material cues: soft-touch matte lamination, paper grain, spot UV varnish, hot foil stamping, emboss/deboss, debossed foil linework, letterpress impression, duotone, spot color, subtle halftone\n' +
                  '- Cinematic lighting: three-point lighting, key light, fill light ratio, rim light, kicker light, bounce light, softbox diffusion, gentle falloff, volumetric light rays, ambient occlusion\n' +
                  '- Filmic color: filmic tone mapping, split toning (warm highlights + cool shadows), matte blacks, highlight roll-off, subtle halation, fine film grain, restrained bloom\n' +
                  '- Use 6-12 of these terms per prompt at most; keep the prompt actionable.\n' +
                  '\n' +
                  '## AVOID CHEAP/AI LOOK (MANDATORY)\n' +
                  '- Avoid: cheap plastic, toy-like, glossy, harsh specular, over-bloom, over-saturated neon\n' +
                  '- Avoid: HUD, sci-fi interface, holographic UI, futuristic dashboards, glowing rings/dials\n' +
                  '- Avoid: generic isometric city / stock 3D icon templates / cliché circuit-board city\n' +
                  '- Prefer: matte, textured, editorial poster vibe, restrained highlights, elegant palettes\n' +
                  '\n' +
                  'Output JSON ONLY: {"prompt":"...","negativePrompt":"..."}'

                // 构建框选区域描述（如果有）
                const regionHint = regionRect 
                  ? `\n\n## User Selected Region\nThe user specifically highlighted a region at: x=${regionRect.x}, y=${regionRect.y}, width=${regionRect.w}, height=${regionRect.h}.\nPlease pay special attention to improving this area in your redesign.`
                  : ''
                
                const geminiUserPrompt = 
                  `## Deck Style Context\n` +
                  `- Design Concept: ${String(deckContext.designConcept || 'Premium neo-Chinese heritage editorial with refined textures').slice(0, 800)}\n` +
                  `- Color Palette: ${String(deckContext.colorPalette || 'Ink black, warm parchment, cinnabar accent, antique gold').slice(0, 200)}\n\n` +
                  `## Page ${pageNum} Information\n` +
                  `- Page Type: ${slidePageType}\n` +
                  `- Layout Intent: ${slideLayoutIntent || 'balanced asymmetric layout with visual hierarchy'}\n\n` +
                  `## SLIDE CONTENT (Chinese text that MUST appear):\n` +
                  `${slideChineseContent || '(No content provided)'}\n\n` +
                  `## User Feedback (the problem to solve):\n${feedback}${regionHint}\n\n` +
                  `## Original Prompt (what went wrong - AVOID these issues):\n${String(slidePromptInfo.prompt || '').slice(0, 1000)}\n\n` +
                  '## YOUR TASK\n' +
                  'Create a COMPLETELY NEW, VISUALLY STUNNING prompt that:\n' +
                  '1. Addresses the user feedback (more design, more polish, more visual interest)\n' +
                  '2. Uses premium materials + theme-related motifs (neo-Chinese heritage if requested)\n' +
                  '3. Includes ALL the Chinese text content with elegant typography\n' +
                  '4. Creates a slide that looks like it belongs in a Fortune 500 keynote' +
                  (regionRect ? '\n5. Especially focus on improving the user-highlighted region' : '')

                // 带重试的 Gemini 调用（网络不稳定时自动重试）
                let geminiResponse = null
                let geminiRetries = 0
                const maxGeminiRetries = 3
                while (geminiRetries < maxGeminiRetries) {
                  try {
                    geminiResponse = await callOpenRouterGemini({
                      apiKey: openRouterApiKey,
                      model: 'google/gemini-3-pro-preview',
                      systemPrompt: geminiSystemPrompt,
                      userPrompt: geminiUserPrompt,
                    })
                    break // 成功则跳出
                  } catch (geminiErr) {
                    geminiRetries++
                    const errMsg = geminiErr?.message || String(geminiErr)
                    console.warn(`[PPT Edit] Gemini 调用失败 (尝试 ${geminiRetries}/${maxGeminiRetries}): ${errMsg.slice(0, 200)}`)
                    if (geminiRetries >= maxGeminiRetries) {
                      throw new Error(`Gemini 调用失败（已重试 ${maxGeminiRetries} 次）: ${errMsg}`)
                    }
                    // 等待后重试
                    await new Promise(r => setTimeout(r, 1000 * geminiRetries))
                  }
                }

                const parsed = parseJsonFromModelText(geminiResponse)
                if (!parsed?.prompt) {
                  editLogs.push({ pageNum, success: false, error: 'Gemini 返回的 prompt 无效' })
                  return
                }

                // 2. 生成新图（根据模型选择不同接口）
                let raw
                const maxRetries = 2
                let retries = 0
                
                while (retries < maxRetries) {
                  try {
                    if (pptImageModel === 'gemini-image') {
                      // 使用 LinAPI Gemini 生图
                      const enhancedPrompt = enhancePromptForGeminiImage({
                        prompt: parsed.prompt,
                        negativePrompt: mergeNegativePrompt(parsed.negativePrompt),
                      })
                      const geminiResult = await linapiGenerateImage({
                        apiKey: mainApiKey || dashscopeApiKey,
                        prompt: enhancedPrompt,
                        aspectRatio: '16:9',
                      })
                      raw = Buffer.from(geminiResult.base64, 'base64')
                    } else {
                      // 使用 DashScope 生图
                      const result = await dashscopeGenerateImageUrl({
                        prompt: parsed.prompt,
                        negativePrompt: mergeNegativePrompt(parsed.negativePrompt),
                        size: '2048*1152',
                        promptExtend: false,
                        watermark: false,
                        model: pptImageModel || 'z-image-turbo',
                        region: 'cn',
                        apiKey: dashscopeApiKey,
                      })
                      raw = await downloadToBuffer(result.url)
                    }
                    break
                  } catch (imgErr) {
                    retries++
                    const errMsg = imgErr?.message || String(imgErr)
                    console.warn(`[PPT Edit] 生图失败 (尝试 ${retries}/${maxRetries}): ${errMsg.slice(0, 200)}`)
                    if (retries >= maxRetries) {
                      throw new Error(`生图失败（已重试 ${maxRetries} 次）: ${errMsg}`)
                    }
                    await new Promise(r => setTimeout(r, 1500 * retries))
                  }
                }
                newImageBuffer = await postprocessTo1920x1200(raw, 'letterbox')

                // 保存编辑记录
                if (fs.existsSync(assetsDir)) {
                  const seq = String(pageNum).padStart(2, '0')
                  const timestamp = Date.now()
                  try {
                    fs.writeFileSync(path.join(assetsDir, `slide_${seq}_edit_${timestamp}_prompt.txt`), parsed.prompt)
                    fs.writeFileSync(path.join(assetsDir, `slide_${seq}_edit_${timestamp}_after.png`), newImageBuffer)
                  } catch {}
                }

                // 更新 slidesPrompts
                if (slidesPrompts[pageIndex]) {
                  slidesPrompts[pageIndex].prompt = parsed.prompt
                  slidesPrompts[pageIndex].negativePrompt = parsed.negativePrompt || ''
                }

              } else if (mode === 'partial_edit') {
                // ========== 局部编辑 ==========
                // 1. 让 Gemini 生成编辑指令（给 qwen-image-edit-plus）
                const geminiSystemPrompt = 
                  'You are an expert at image editing prompts for AI image editors. ' +
                  'The user wants to make SPECIFIC partial edits to a PPT slide image. ' +
                  '\n\n' +
                  '## EDITING GUIDELINES\n' +
                  '- Be PRECISE about what to change and what to keep\n' +
                  '- If changing background: describe the NEW background style (gradient, abstract, etc.)\n' +
                  '- If changing colors: specify exact color transitions\n' +
                  '- If changing text style: describe the NEW typography style\n' +
                  '- PRESERVE all Chinese text unless user explicitly wants to change it\n' +
                  '\n' +
                  '## QUALITY REQUIREMENTS\n' +
                  '- Maintain premium design aesthetic\n' +
                  '- Chinese text must remain crisp and readable\n' +
                  '- Changes should enhance, not diminish the design\n' +
                  '\n' +
                  '## HIGH-END DESIGN VOCAB (USE SELECTIVELY)\n' +
                  '- Microtypography: typographic grid, baseline grid, microtypography, optical alignment, kerning, tracking, leading\n' +
                  '- Premium finishes/material cues: soft-touch matte, paper grain, spot UV varnish, hot foil stamping, emboss/deboss, letterpress impression, duotone/spot color\n' +
                  '- Cinematic lighting: key light, fill ratio, rim/kicker light, softbox diffusion, bounce light, gentle falloff, volumetric rays, subtle vignetting\n' +
                  '- Filmic color: filmic tone mapping, split toning, matte blacks, highlight roll-off, fine grain, restrained bloom\n' +
                  '- Do NOT keyword-stuff. Use only what helps the requested edit.\n' +
                  '\n' +
                  'Output JSON ONLY: {"editPrompt":"...","negativePrompt":"..."}'

                const geminiUserPrompt = 
                  `## Current Slide Design\n` +
                  `- Design Concept: ${String(deckContext.designConcept || '').slice(0, 600)}\n` +
                  `- Color Palette: ${String(deckContext.colorPalette || '').slice(0, 200)}\n\n` +
                  `## Chinese Text Content (PRESERVE unless asked to change):\n` +
                  `${slideChineseContent || String(slidePromptInfo.originalChineseContent || '').slice(0, 800)}\n\n` +
                  `## User Edit Request:\n${feedback}\n\n` +
                  'Create an edit prompt that makes ONLY the requested changes while keeping everything else intact.'

                // 带重试的 Gemini 调用
                let geminiResponse = null
                let geminiRetries = 0
                const maxGeminiRetries = 3
                while (geminiRetries < maxGeminiRetries) {
                  try {
                    geminiResponse = await callOpenRouterGemini({
                      apiKey: openRouterApiKey,
                      model: 'google/gemini-3-pro-preview',
                      systemPrompt: geminiSystemPrompt,
                      userPrompt: geminiUserPrompt,
                    })
                    break
                  } catch (geminiErr) {
                    geminiRetries++
                    console.warn(`[PPT Edit] Gemini 调用失败 (尝试 ${geminiRetries}/${maxGeminiRetries})`)
                    if (geminiRetries >= maxGeminiRetries) {
                      throw new Error(`Gemini 调用失败（已重试 ${maxGeminiRetries} 次）`)
                    }
                    await new Promise(r => setTimeout(r, 1000 * geminiRetries))
                  }
                }

                const parsed = parseJsonFromModelText(geminiResponse)
                if (!parsed?.editPrompt) {
                  editLogs.push({ pageNum, success: false, error: 'Gemini 返回的 editPrompt 无效' })
                  return
                }

                // 2. 带重试的 DashScope 图像编辑
                let dashscopeUrl = null
                let dashscopeRetries = 0
                const maxDashscopeRetries = 2
                while (dashscopeRetries < maxDashscopeRetries) {
                  try {
                    const result = await dashscopeImageEdit({
                      imageBase64: originalImageBase64,
                      prompt: parsed.editPrompt,
                      negativePrompt: mergeNegativePrompt(parsed.negativePrompt),
                      n: 1,
                      watermark: false,
                      model: 'qwen-image-edit-plus',
                      region: 'cn',
                      apiKey: dashscopeApiKey,
                    })
                    dashscopeUrl = result.url
                    break
                  } catch (dsErr) {
                    dashscopeRetries++
                    console.warn(`[PPT Edit] DashScope 编辑失败 (尝试 ${dashscopeRetries}/${maxDashscopeRetries})`)
                    if (dashscopeRetries >= maxDashscopeRetries) {
                      throw new Error(`DashScope 图像编辑失败（已重试 ${maxDashscopeRetries} 次）`)
                    }
                    await new Promise(r => setTimeout(r, 1500 * dashscopeRetries))
                  }
                }
                
                const { url } = { url: dashscopeUrl }

                const raw = await downloadToBuffer(url)
                newImageBuffer = await postprocessTo1920x1200(raw, 'letterbox')

                // 保存编辑记录
                if (fs.existsSync(assetsDir)) {
                  const seq = String(pageNum).padStart(2, '0')
                  const timestamp = Date.now()
                  try {
                    fs.writeFileSync(path.join(assetsDir, `slide_${seq}_partialedit_${timestamp}_prompt.txt`), parsed.editPrompt)
                    fs.writeFileSync(path.join(assetsDir, `slide_${seq}_partialedit_${timestamp}_before.png`), originalImage)
                    fs.writeFileSync(path.join(assetsDir, `slide_${seq}_partialedit_${timestamp}_after.png`), newImageBuffer)
                  } catch {}
                }
              }

              if (newImageBuffer && newImageBuffer.length > 0) {
                replacements.push({ pageIndex, imageBuffer: newImageBuffer })
                editLogs.push({ pageNum, success: true })
              } else {
                editLogs.push({ pageNum, success: false, error: '生成的图片为空' })
              }
            })
          }

          if (replacements.length === 0) {
            return { success: false, error: '没有成功编辑任何页面', logs: editLogs }
          }

          // 替换 PPTX 中的图片并覆盖写回
          await replaceSlideImagesInPptx(pptxPath, replacements, true)

          // 更新 slides_prompts.json
          if (fs.existsSync(assetsDir) && slidesPrompts.length > 0) {
            try {
              fs.writeFileSync(
                path.join(assetsDir, 'slides_prompts.json'),
                JSON.stringify(slidesPrompts, null, 2)
              )
            } catch {}
          }

          return {
            success: true,
            path: pptxPath,
            editedPages: replacements.map((r) => r.pageIndex + 1),
            logs: editLogs,
          }
        } catch (error) {
          console.error('ppt-edit-slides failed:', error)
          return { success: false, error: error.message || String(error) }
        }
    },
  }
}

module.exports = {
  createPptService,
}
