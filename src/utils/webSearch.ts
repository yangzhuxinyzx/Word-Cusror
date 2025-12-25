export interface WebSearchOptions {
  locale?: string
  region?: string
  num?: number
  braveApiKey?: string
}

export interface WebSearchResultItem {
  title: string
  link: string
  snippet: string
  extraSnippets?: string[]
}

export interface WebSearchFaqItem {
  question: string
  answer: string
  title?: string
  link?: string
}

export interface WebSearchNewsItem {
  title: string
  link: string
  source?: string
  description?: string
  breaking?: boolean
  isLive?: boolean
  age?: string
}

export interface WebSearchVideoItem {
  title: string
  link: string
  description?: string
  duration?: string
  thumbnail?: string
  viewCount?: number | string
  creator?: string
  publisher?: string
}

export interface WebSearchDiscussionItem {
  link: string
  forumName?: string
  question?: string
  topComment?: string
}

export interface WebSearchSections {
  web: WebSearchResultItem[]
  faq: WebSearchFaqItem[]
  news: WebSearchNewsItem[]
  videos: WebSearchVideoItem[]
  discussions: WebSearchDiscussionItem[]
}

export interface WebSearchResponse {
  success: boolean
  results?: WebSearchResultItem[]
  sections?: WebSearchSections
  summarizerKey?: string
  message?: string
  raw?: unknown
}

/**
 * 调用桌面端（Electron）暴露的 webSearch API。
 * 如果在纯浏览器环境运行且未配置后备接口，会返回失败信息。
 */
export async function runWebSearch(
  query: string,
  options?: WebSearchOptions
): Promise<WebSearchResponse> {
  const payload = {
    query,
    locale: options?.locale,
    region: options?.region,
    num: options?.num,
    braveApiKey: options?.braveApiKey,
  }

  if (typeof window !== 'undefined' && window.electronAPI?.webSearch) {
    return window.electronAPI.webSearch(payload)
  }

  const fallbackUrl = import.meta.env.VITE_WEB_SEARCH_ENDPOINT
  const fallbackKey = import.meta.env.VITE_WEB_SEARCH_KEY

  if (fallbackUrl && fallbackKey) {
    try {
      const response = await fetch(fallbackUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          Authorization: `Bearer ${fallbackKey}`,
        },
        body: JSON.stringify(payload),
      })
      if (!response.ok) {
        return { success: false, message: `后备搜索接口失败：${response.status}` }
      }
      const data = (await response.json()) as WebSearchResponse
      return data
    } catch (error) {
      return { success: false, message: (error as Error).message }
    }
  }

  return {
    success: false,
    message:
      'Web 搜索功能仅在桌面应用中可用。请在 Electron 环境下运行，或配置 VITE_WEB_SEARCH_ENDPOINT/VITE_WEB_SEARCH_KEY。',
  }
}

