function createWebSearchService(options) {
  const { app } = options

  const DEFAULT_RESULT_FILTER = ['web', 'query', 'faq', 'news', 'videos', 'discussions']
  const braveServerModulePromise = import('@brave/brave-search-mcp-server/dist/server.js')

  const COUNTRY_CODES = new Set([
    'ALL', 'AR', 'AU', 'AT', 'BE', 'BR', 'CA', 'CL', 'DK', 'FI', 'FR', 'DE',
    'HK', 'IN', 'ID', 'IT', 'JP', 'KR', 'MY', 'MX', 'NL', 'NZ', 'NO', 'CN',
    'PL', 'PT', 'PH', 'RU', 'SA', 'ZA', 'ES', 'SE', 'CH', 'TW', 'TR', 'GB', 'US',
  ])

  const UI_LANG_OPTIONS = new Set([
    'es-AR', 'en-AU', 'de-AT', 'nl-BE', 'fr-BE', 'pt-BR', 'en-CA', 'fr-CA',
    'es-CL', 'da-DK', 'fi-FI', 'fr-FR', 'de-DE', 'el-GR', 'zh-HK', 'en-IN',
    'en-ID', 'it-IT', 'ja-JP', 'ko-KR', 'en-MY', 'es-MX', 'nl-NL', 'en-NZ',
    'no-NO', 'zh-CN', 'pl-PL', 'en-PH', 'ru-RU', 'en-ZA', 'es-ES', 'sv-SE',
    'fr-CH', 'de-CH', 'zh-TW', 'tr-TR', 'en-GB', 'en-US', 'es-US',
  ])

  const SEARCH_LANG_OPTIONS = new Set([
    'ar', 'eu', 'bn', 'bg', 'ca', 'zh-hans', 'zh-hant', 'hr', 'cs', 'da', 'nl',
    'en', 'en-gb', 'et', 'fi', 'fr', 'gl', 'de', 'el', 'gu', 'he', 'hi', 'hu',
    'is', 'it', 'jp', 'kn', 'ko', 'lv', 'lt', 'ms', 'ml', 'mr', 'nb', 'pl',
    'pt-br', 'pt-pt', 'pa', 'ro', 'ru', 'sr', 'sk', 'sl', 'es', 'sv', 'ta',
    'te', 'th', 'tr', 'uk', 'vi',
  ])

  let braveMcpConnection = null
  let braveMcpInitPromise = null
  let braveMcpApiKey = null

  async function ensureBraveMcpClient(apiKeyOverride) {
    const apiKey = apiKeyOverride || process.env.BRAVE_API_KEY

    if (braveMcpConnection && braveMcpApiKey === apiKey) {
      return braveMcpConnection
    }

    if (braveMcpConnection && braveMcpApiKey !== apiKey) {
      try {
        await braveMcpConnection.client?.close?.()
        await braveMcpConnection.server?.close?.()
      } catch {}
      braveMcpConnection = null
      braveMcpInitPromise = null
    }

    if (braveMcpInitPromise) return braveMcpInitPromise

    braveMcpInitPromise = (async () => {
      if (!apiKey) {
        throw new Error('Missing Brave Search API key')
      }

      const serverModule = await braveServerModulePromise
      const createServer = serverModule?.default || serverModule
      const server = createServer({ config: { braveApiKey: apiKey } })

      const sdkBase = require('path').join(
        __dirname,
        '..',
        '..',
        'node_modules',
        '@modelcontextprotocol',
        'sdk',
        'dist',
        'cjs',
      )
      const { Client: McpClient } = require(require('path').join(sdkBase, 'client', 'index.js'))
      const { InMemoryTransport } = require(require('path').join(sdkBase, 'inMemory.js'))
      const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair()
      await server.connect(serverTransport)

      const client = new McpClient({
        name: 'word-cursor',
        version: app?.getVersion?.() || 'dev',
      })
      await client.connect(clientTransport)
      await client.listTools({})

      braveMcpConnection = { client, server }
      braveMcpApiKey = apiKey
      return braveMcpConnection
    })().catch((error) => {
      braveMcpInitPromise = null
      throw error
    })

    return braveMcpInitPromise
  }

  function normalizeUiLang(locale) {
    if (!locale) return null
    const normalized = locale.replace('_', '-')
    const [lang, region] = normalized.split('-')
    if (!lang) return null
    const candidate = region
      ? `${lang.toLowerCase()}-${region.toUpperCase()}`
      : `${lang.toLowerCase()}`
    if (UI_LANG_OPTIONS.has(candidate)) {
      return candidate
    }
    if (region) {
      const fallback = `${lang.toLowerCase()}-${region.toUpperCase()}`
      if (UI_LANG_OPTIONS.has(fallback)) return fallback
    }
    return UI_LANG_OPTIONS.has('en-US') ? 'en-US' : null
  }

  function normalizeSearchLang(locale) {
    if (!locale) return null
    const normalized = locale.toLowerCase()
    if (SEARCH_LANG_OPTIONS.has(normalized)) return normalized
    if (normalized.startsWith('zh')) {
      return normalized.includes('tw') || normalized.includes('hk') ? 'zh-hant' : 'zh-hans'
    }
    const base = normalized.split(/[-_]/)[0]
    if (SEARCH_LANG_OPTIONS.has(base)) return base
    return 'en'
  }

  function normalizeCountry(region) {
    if (!region) return 'US'
    const upper = region.toUpperCase()
    return COUNTRY_CODES.has(upper) ? upper : 'US'
  }

  function buildBraveWebArguments(query, searchOptions = {}) {
    const count = Math.max(1, Math.min(parseInt(searchOptions.num ?? 5, 10) || 5, 20))
    const args = {
      query,
      count,
      safesearch: 'moderate',
      spellcheck: true,
      text_decorations: true,
      summary: false,
      extra_snippets: true,
      result_filter: DEFAULT_RESULT_FILTER,
    }

    const locale = typeof searchOptions.locale === 'string' ? searchOptions.locale.trim() : ''
    const region = typeof searchOptions.region === 'string' ? searchOptions.region.trim() : ''

    const uiLang = normalizeUiLang(locale) || 'en-US'
    if (uiLang) {
      args.ui_lang = uiLang
    }

    const searchLang = normalizeSearchLang(locale) || 'en'
    if (searchLang) {
      args.search_lang = searchLang
    }

    const country = normalizeCountry(region || (uiLang?.split('-')[1] || 'US'))
    if (country) {
      args.country = country
    }

    return args
  }

  function isFaqResult(data) {
    return data && typeof data.question === 'string' && typeof data.answer === 'string'
  }

  function isNewsResult(data) {
    return data && typeof data.source === 'string' && Object.prototype.hasOwnProperty.call(data, 'breaking')
  }

  function isVideoResult(data) {
    return data && (
      Object.prototype.hasOwnProperty.call(data, 'thumbnail_url') ||
      Object.prototype.hasOwnProperty.call(data, 'duration')
    )
  }

  function isDiscussionResult(data) {
    return data && data.data && typeof data.data.forum_name === 'string'
  }

  function isWebResult(data) {
    if (!data || typeof data !== 'object') return false
    if (typeof data.title !== 'string' || typeof data.url !== 'string') return false
    if (isFaqResult(data) || isNewsResult(data) || isVideoResult(data) || isDiscussionResult(data)) {
      return false
    }
    return true
  }

  function transformBraveContent(content = [], maxWebCount = 5) {
    const sections = {
      web: [],
      faq: [],
      news: [],
      videos: [],
      discussions: [],
    }
    let summarizerKey = null

    for (const block of content || []) {
      if (!block || block.type !== 'text' || !block.text) continue
      const textBlock = block.text.trim()
      if (!textBlock) continue

      if (textBlock.startsWith('Summarizer key:')) {
        summarizerKey = textBlock.split(':').slice(1).join(':').trim()
        continue
      }

      let data
      try {
        data = JSON.parse(textBlock)
      } catch {
        continue
      }

      if (isFaqResult(data)) {
        sections.faq.push({
          question: data.question,
          answer: data.answer,
          title: data.title,
          link: data.url,
        })
        continue
      }

      if (isNewsResult(data)) {
        sections.news.push({
          title: data.title,
          link: data.url,
          source: data.source,
          description: data.description,
          breaking: Boolean(data.breaking),
          isLive: Boolean(data.is_live),
          age: data.age,
        })
        continue
      }

      if (isVideoResult(data)) {
        sections.videos.push({
          title: data.title,
          link: data.url,
          description: data.description,
          duration: data.duration,
          thumbnail: data.thumbnail_url,
          viewCount: data.view_count,
          creator: data.creator,
          publisher: data.publisher,
        })
        continue
      }

      if (isDiscussionResult(data)) {
        sections.discussions.push({
          link: data.url,
          forumName: data.data?.forum_name,
          question: data.data?.question,
          topComment: data.data?.top_comment,
        })
        continue
      }

      if (isWebResult(data)) {
        sections.web.push({
          title: data.title || 'Untitled result',
          link: data.url || '',
          snippet: data.description || '',
          extraSnippets: Array.isArray(data.extra_snippets) ? data.extra_snippets : undefined,
        })
      }
    }

    sections.web = sections.web.slice(0, maxWebCount)

    return {
      sections,
      summarizerKey,
    }
  }

  return {
    async performWebSearch(query, searchOptions = {}) {
      const { client } = await ensureBraveMcpClient(searchOptions.braveApiKey)
      const args = buildBraveWebArguments(query, searchOptions)

      const result = await client.callTool({
        name: 'brave_web_search',
        arguments: args,
      })

      if (result.isError) {
        const errorMessage = Array.isArray(result.content)
          ? result.content.map((item) => item?.text || '').join('\n')
          : 'Brave search failed'
        throw new Error(errorMessage || 'Brave search failed')
      }

      const parsedContent = transformBraveContent(result.content, args.count || 5)
      if (parsedContent.sections.web.length === 0) {
        return { success: false, message: 'Brave returned no usable web results.' }
      }

      return {
        success: true,
        results: parsedContent.sections.web,
        sections: parsedContent.sections,
        summarizerKey: parsedContent.summarizerKey,
        raw: result.content,
      }
    },

    async close() {
      if (!braveMcpConnection) return
      try {
        await braveMcpConnection.client?.close?.()
      } catch {}
      try {
        await braveMcpConnection.server?.close?.()
      } catch {}
      braveMcpConnection = null
      braveMcpInitPromise = null
      braveMcpApiKey = null
    },
  }
}

module.exports = {
  createWebSearchService,
}
