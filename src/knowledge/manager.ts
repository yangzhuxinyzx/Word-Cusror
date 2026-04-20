import type {
  KnowledgePendingProfileItem,
  KnowledgeSearchResult,
  KnowledgeStatusResponse,
} from '../types'

export async function knowledgeConfigure(options: {
  knowledgeEnabled?: boolean
  workspaceKnowledgeEnabled?: boolean
  globalKnowledgePath?: string
  profileMemoryEnabled?: boolean
  embeddingBaseUrl?: string
  embeddingApiKey?: string
  embeddingModel?: string
  knowledgeTopK?: number
}): Promise<KnowledgeStatusResponse> {
  if (typeof window === 'undefined' || !window.electronAPI?.knowledgeConfigure) {
    return { success: false, error: 'knowledgeConfigure 不可用' }
  }
  return window.electronAPI.knowledgeConfigure(options)
}

export async function knowledgeSetActiveWorkspace(workspacePath?: string | null) {
  if (typeof window === 'undefined' || !window.electronAPI?.knowledgeSetActiveWorkspace) {
    return { success: false, error: 'knowledgeSetActiveWorkspace 不可用' }
  }
  return window.electronAPI.knowledgeSetActiveWorkspace({ workspacePath })
}

export async function knowledgeStatus(): Promise<KnowledgeStatusResponse> {
  if (typeof window === 'undefined' || !window.electronAPI?.knowledgeStatus) {
    return { success: false, error: 'knowledgeStatus 不可用' }
  }
  return window.electronAPI.knowledgeStatus()
}

export async function knowledgeRetrieve(options: {
  query: string
  topK?: number
}): Promise<{ success: boolean; results: KnowledgeSearchResult[]; error?: string }> {
  if (typeof window === 'undefined' || !window.electronAPI?.knowledgeRetrieve) {
    return { success: false, results: [], error: 'knowledgeRetrieve 不可用' }
  }
  return window.electronAPI.knowledgeRetrieve(options)
}

export async function knowledgeRebuild(scope: 'all' | 'workspace' | 'global' = 'all') {
  if (typeof window === 'undefined' || !window.electronAPI?.knowledgeRebuild) {
    return { success: false, error: 'knowledgeRebuild 不可用' }
  }
  return window.electronAPI.knowledgeRebuild({ scope })
}

export async function knowledgeListPendingProfile(): Promise<{ success: boolean; items: KnowledgePendingProfileItem[]; error?: string }> {
  if (typeof window === 'undefined' || !window.electronAPI?.knowledgeListPendingProfile) {
    return { success: false, items: [], error: 'knowledgeListPendingProfile 不可用' }
  }
  return window.electronAPI.knowledgeListPendingProfile()
}

export async function knowledgeResolvePendingProfile(options: {
  id?: string
  ids?: string[]
  action: 'accept' | 'reject'
}) {
  if (typeof window === 'undefined' || !window.electronAPI?.knowledgeResolvePendingProfile) {
    return { success: false, error: 'knowledgeResolvePendingProfile 不可用' }
  }
  return window.electronAPI.knowledgeResolvePendingProfile(options)
}

export async function knowledgeListProfileFacts(): Promise<{ success: boolean; items: KnowledgePendingProfileItem[]; error?: string }> {
  if (typeof window === 'undefined' || !window.electronAPI?.knowledgeListProfileFacts) {
    return { success: false, items: [], error: 'knowledgeListProfileFacts 不可用' }
  }
  return window.electronAPI.knowledgeListProfileFacts()
}

export async function knowledgeQueueProfileCandidates(options: {
  items: Array<{
    id?: string
    category: string
    statement: string
    evidenceText: string
    sourceScope?: string
    sourcePath?: string
    metadata?: Record<string, unknown>
  }>
}) {
  if (typeof window === 'undefined' || !window.electronAPI?.knowledgeQueueProfileCandidates) {
    return { success: false, error: 'knowledgeQueueProfileCandidates 不可用' }
  }
  return window.electronAPI.knowledgeQueueProfileCandidates(options)
}

export function formatKnowledgeResults(results: KnowledgeSearchResult[], maxChars = 2400) {
  if (!results.length) return ''
  const lines = results.map((item, index) => {
    const title = item.title ? ` ${item.title}` : ''
    const path = item.relativePath || item.sourcePath || ''
    return `#${index + 1} [${item.sourceScope}]${title}\n${path}\n${item.snippet}`
  })
  const output = lines.join('\n\n')
  if (output.length <= maxChars) return output
  return `${output.slice(0, maxChars)}\n\n... (知识库结果已截断)`
}
