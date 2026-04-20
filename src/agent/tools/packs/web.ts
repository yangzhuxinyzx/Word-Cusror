import { BUILTIN_KEYS } from '../../../config/models'
import type { WebSearchResponse } from '../../../utils/webSearch'
import { runWebSearch } from '../../../utils/webSearch'
import { defineAgentTool } from '../contracts'
import type { ExecutableAgentTool } from '../executor'

export interface WebToolPackDeps {
  registerToolActivity: (tool: string, label: string) => string
  completeToolActivity: (
    activityId: string,
    status: 'success' | 'error' | 'skipped',
    detail?: string,
  ) => void
  updateAgentAction: (action: string) => void
  truncateLabel: (text: string, limit?: number) => string
}

function formatSearchResults(response: WebSearchResponse, query: string) {
  const sections = response.sections
  const webResults = sections?.web ?? response.results ?? []
  const lines: string[] = []

  if (webResults.length > 0) {
    lines.push('【Brave Web】')
    lines.push(
      webResults
        .map((item, index) => {
          const snippet = item.snippet ? item.snippet.replace(/\s+/g, ' ').trim() : ''
          return `${index + 1}. ${item.title}\n${item.link}\n${snippet}`
        })
        .join('\n\n'),
    )
  }

  if (sections?.faq?.length) {
    const faqBlock = sections.faq
      .slice(0, 3)
      .map((faq, idx) => `Q${idx + 1}: ${faq.question}\nA: ${faq.answer}`)
      .join('\n\n')
    lines.push('【FAQ】')
    lines.push(faqBlock)
  }

  if (sections?.news?.length) {
    const newsBlock = sections.news
      .slice(0, 3)
      .map((news) => `${news.title}${news.source ? ` - ${news.source}` : ''}\n${news.link}`)
      .join('\n\n')
    lines.push('【新闻】')
    lines.push(newsBlock)
  }

  if (sections?.videos?.length) {
    const videoBlock = sections.videos
      .slice(0, 2)
      .map((video) => `${video.title}${video.duration ? ` (${video.duration})` : ''}\n${video.link}`)
      .join('\n\n')
    lines.push('【视频】')
    lines.push(videoBlock)
  }

  if (sections?.discussions?.length) {
    const discussionBlock = sections.discussions
      .slice(0, 2)
      .map((discussion) => `${discussion.forumName ?? '讨论'}：${discussion.question ?? ''}\n${discussion.link}`)
      .join('\n\n')
    lines.push('【讨论】')
    lines.push(discussionBlock)
  }

  if (response.summarizerKey) {
    lines.push(`Summarizer key: ${response.summarizerKey}`)
  }

  return `【Brave 搜索】${query}\n\n${lines.join('\n\n')}`
}

export function createWebToolPack(
  deps: WebToolPackDeps,
): ExecutableAgentTool[] {
  return [
    defineAgentTool({
      id: 'web_search',
      displayName: 'Web Search',
      description: 'Search the web for external information',
      domain: 'web',
      mutation: 'external',
      concurrency: 'parallel_safe',
      tags: ['phase1', 'web', 'search'],
      inputKeys: ['query', 'q', 'keyword', 'hl', 'locale', 'gl', 'region', 'num'],
      async handler(args) {
        const query = (args.query || args.q || args.keyword || '').trim()
        if (!query) {
          return { tool: 'web_search', success: false, message: '缺少 query 参数' }
        }

        const locale = args.hl || args.locale || 'zh-CN'
        const region = args.gl || args.region || 'cn'
        const num = args.num ? parseInt(args.num, 10) || 5 : 5

        const activityId = deps.registerToolActivity(
          'web_search',
          `搜索：${deps.truncateLabel(query, 28)}`,
        )
        deps.updateAgentAction(`正在检索外部信息：${deps.truncateLabel(query, 28)}`)

        const searchResponse = await runWebSearch(query, {
          locale,
          region,
          num,
          braveApiKey: BUILTIN_KEYS.braveApiKey,
        })

        const webResults = searchResponse.results ?? []
        if (!searchResponse.success || webResults.length === 0) {
          deps.completeToolActivity(
            activityId,
            'error',
            searchResponse.message || '0 条结果',
          )
          return {
            tool: 'web_search',
            success: false,
            message: searchResponse.message || '未获取到搜索结果，请稍后重试',
          }
        }

        const extraTotal =
          (searchResponse.sections?.faq?.length ?? 0) +
          (searchResponse.sections?.news?.length ?? 0) +
          (searchResponse.sections?.videos?.length ?? 0) +
          (searchResponse.sections?.discussions?.length ?? 0)
        const summaryLabel = `${webResults.length}${extraTotal ? `+${extraTotal}` : ''} 条`

        deps.completeToolActivity(activityId, 'success', summaryLabel)
        const formatted = formatSearchResults(searchResponse, query)

        return {
          tool: 'web_search',
          success: true,
          message: formatted,
          data: {
            query,
            locale,
            region,
            results: webResults,
            sections: searchResponse.sections,
            summarizerKey: searchResponse.summarizerKey,
          },
        }
      },
    }),
  ]
}

