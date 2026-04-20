import {
  memoryClear,
  memoryRebuildIndex,
  memorySearch,
  memoryStatus,
  memoryStatusDetail,
} from '../../../memory/manager'
import { formatMemoryResults } from '../../../memory/hybrid'
import { defineAgentTool } from '../contracts'
import type { ExecutableAgentTool } from '../executor'

export interface MemoryToolPackDeps {
  registerToolActivity: (tool: string, label: string) => string
  completeToolActivity: (
    activityId: string,
    status: 'success' | 'error' | 'skipped',
    detail?: string,
  ) => void
  updateAgentAction: (action: string) => void
  truncateLabel: (text: string, limit?: number) => string
}

function ok(tool: string, message: string, data?: Record<string, unknown>) {
  return { tool, success: true, message, data }
}

function fail(tool: string, message: string) {
  return { tool, success: false, message }
}

export function createMemoryToolPack(
  deps: MemoryToolPackDeps,
): ExecutableAgentTool[] {
  return [
    defineAgentTool({
      id: 'memory_search',
      displayName: 'Memory Search',
      description: 'Search local memory for relevant past context',
      domain: 'memory',
      mutation: 'read',
      concurrency: 'parallel_safe',
      legacyAliases: ['memory.search'],
      async handler(args) {
        const query = (args.query || args.q || '').trim()
        if (!query) return fail('memory_search', 'Missing query.')
        const activityId = deps.registerToolActivity(
          'memory_search',
          `Memory: ${deps.truncateLabel(query, 24)}`,
        )
        deps.updateAgentAction(`Searching memory for ${deps.truncateLabel(query, 24)}...`)
        try {
          const response = await memorySearch({
            query,
            topK: args.topK ? parseInt(args.topK, 10) : undefined,
            workspaceKey: args.workspaceKey,
          })
          if (!response.success) {
            deps.completeToolActivity(activityId, 'error', response.error)
            return fail('memory_search', response.error || 'Memory search failed')
          }
          const formatted = formatMemoryResults(response.results, 2000)
          deps.completeToolActivity(activityId, 'success', `${response.results.length} hits`)
          return ok(
            'memory_search',
            formatted || 'No relevant memory found.',
            { results: response.results, count: response.results.length },
          )
        } catch (error) {
          deps.completeToolActivity(activityId, 'error', 'Memory search failed')
          return fail('memory_search', `Memory search failed: ${error}`)
        }
      },
    }),
    defineAgentTool({
      id: 'memory_status',
      displayName: 'Memory Status',
      description: 'Inspect local memory index status',
      domain: 'memory',
      mutation: 'read',
      concurrency: 'parallel_safe',
      legacyAliases: ['memory.status'],
      async handler() {
        const activityId = deps.registerToolActivity('memory_status', 'Memory status')
        try {
          const response = await memoryStatus()
          if (!response.success) {
            deps.completeToolActivity(activityId, 'error', response.error)
            return fail('memory_status', response.error || 'Memory status failed')
          }
          deps.completeToolActivity(activityId, 'success')
          return ok('memory_status', `Memory dir: ${response.memoryDir || '(unknown)'}`, response as Record<string, unknown>)
        } catch (error) {
          deps.completeToolActivity(activityId, 'error', 'Memory status failed')
          return fail('memory_status', `Memory status failed: ${error}`)
        }
      },
    }),
    defineAgentTool({
      id: 'memory_status_detail',
      displayName: 'Memory Status Detail',
      description: 'Inspect detailed local memory index status',
      domain: 'memory',
      mutation: 'read',
      concurrency: 'parallel_safe',
      legacyAliases: ['memory.status_detail'],
      async handler() {
        const activityId = deps.registerToolActivity('memory_status_detail', 'Memory detail')
        try {
          const response = await memoryStatusDetail()
          if (!response.success) {
            deps.completeToolActivity(activityId, 'error', response.error)
            return fail('memory_status_detail', response.error || 'Detailed memory status failed')
          }
          deps.completeToolActivity(activityId, 'success')
          return ok('memory_status_detail', 'Loaded detailed memory status.', response as Record<string, unknown>)
        } catch (error) {
          deps.completeToolActivity(activityId, 'error', 'Detailed memory status failed')
          return fail('memory_status_detail', `Detailed memory status failed: ${error}`)
        }
      },
    }),
    defineAgentTool({
      id: 'memory_rebuild',
      displayName: 'Memory Rebuild',
      description: 'Rebuild the local memory index',
      domain: 'memory',
      mutation: 'transform',
      concurrency: 'serial',
      legacyAliases: ['memory.rebuild'],
      async handler() {
        const activityId = deps.registerToolActivity('memory_rebuild', 'Memory rebuild')
        try {
          const response = await memoryRebuildIndex()
          if (!response.success) {
            deps.completeToolActivity(activityId, 'error', response.error)
            return fail('memory_rebuild', response.error || 'Memory rebuild failed')
          }
          deps.completeToolActivity(activityId, 'success')
          return ok('memory_rebuild', 'Memory index rebuilt.', response as Record<string, unknown>)
        } catch (error) {
          deps.completeToolActivity(activityId, 'error', 'Memory rebuild failed')
          return fail('memory_rebuild', `Memory rebuild failed: ${error}`)
        }
      },
    }),
    defineAgentTool({
      id: 'memory_clear',
      displayName: 'Memory Clear',
      description: 'Clear local memory data by scope',
      domain: 'memory',
      mutation: 'delete',
      concurrency: 'serial',
      legacyAliases: ['memory.clear'],
      async handler(args) {
        const scope = (args.scope || 'all') as 'all' | 'daily' | 'long' | 'sessions'
        const activityId = deps.registerToolActivity('memory_clear', `Memory clear: ${scope}`)
        try {
          const response = await memoryClear(scope)
          if (!response.success) {
            deps.completeToolActivity(activityId, 'error', response.error)
            return fail('memory_clear', response.error || 'Memory clear failed')
          }
          deps.completeToolActivity(activityId, 'success')
          return ok('memory_clear', `Cleared memory scope: ${scope}`)
        } catch (error) {
          deps.completeToolActivity(activityId, 'error', 'Memory clear failed')
          return fail('memory_clear', `Memory clear failed: ${error}`)
        }
      },
    }),
  ]
}
