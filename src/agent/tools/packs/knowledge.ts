import {
  knowledgeRetrieve,
  knowledgeStatus,
  knowledgeRebuild,
  formatKnowledgeResults,
} from '../../../knowledge/manager'
import { defineAgentTool } from '../contracts'
import type { ExecutableAgentTool } from '../executor'

export interface KnowledgeToolPackDeps {
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

export function createKnowledgeToolPack(
  deps: KnowledgeToolPackDeps,
): ExecutableAgentTool[] {
  return [
    defineAgentTool({
      id: 'knowledge_search',
      displayName: 'Knowledge Search',
      description: 'Search the local workspace/global knowledge base and user profile memory',
      domain: 'knowledge',
      mutation: 'read',
      concurrency: 'parallel_safe',
      legacyAliases: ['knowledge.search'],
      async handler(args) {
        const query = (args.query || args.q || '').trim()
        if (!query) return fail('knowledge_search', 'Missing query.')
        const activityId = deps.registerToolActivity(
          'knowledge_search',
          `Knowledge: ${deps.truncateLabel(query, 24)}`,
        )
        deps.updateAgentAction(`Searching knowledge for ${deps.truncateLabel(query, 24)}...`)
        try {
          const response = await knowledgeRetrieve({
            query,
            topK: args.topK ? parseInt(args.topK, 10) : undefined,
          })
          if (!response.success) {
            deps.completeToolActivity(activityId, 'error', response.error)
            return fail('knowledge_search', response.error || 'Knowledge search failed')
          }
          deps.completeToolActivity(activityId, 'success', `${response.results.length} hits`)
          return ok(
            'knowledge_search',
            formatKnowledgeResults(response.results, 2200) || 'No relevant knowledge found.',
            { results: response.results, count: response.results.length },
          )
        } catch (error) {
          deps.completeToolActivity(activityId, 'error', 'Knowledge search failed')
          return fail('knowledge_search', `Knowledge search failed: ${error}`)
        }
      },
    }),
    defineAgentTool({
      id: 'knowledge_status',
      displayName: 'Knowledge Status',
      description: 'Inspect local knowledge base and profile-memory status',
      domain: 'knowledge',
      mutation: 'read',
      concurrency: 'parallel_safe',
      legacyAliases: ['knowledge.status'],
      async handler() {
        const activityId = deps.registerToolActivity('knowledge_status', 'Knowledge status')
        try {
          const response = await knowledgeStatus()
          if (!response.success) {
            deps.completeToolActivity(activityId, 'error', response.error)
            return fail('knowledge_status', response.error || 'Knowledge status failed')
          }
          deps.completeToolActivity(activityId, 'success')
          return ok('knowledge_status', 'Knowledge status loaded.', response as Record<string, unknown>)
        } catch (error) {
          deps.completeToolActivity(activityId, 'error', 'Knowledge status failed')
          return fail('knowledge_status', `Knowledge status failed: ${error}`)
        }
      },
    }),
    defineAgentTool({
      id: 'knowledge_rebuild',
      displayName: 'Knowledge Rebuild',
      description: 'Rebuild local knowledge indexes',
      domain: 'knowledge',
      mutation: 'transform',
      concurrency: 'serial',
      legacyAliases: ['knowledge.rebuild'],
      async handler(args) {
        const scope = (args.scope || 'all') as 'all' | 'workspace' | 'global'
        const activityId = deps.registerToolActivity('knowledge_rebuild', `Knowledge rebuild: ${scope}`)
        try {
          const response = await knowledgeRebuild(scope)
          if (!response.success) {
            deps.completeToolActivity(activityId, 'error', response.error)
            return fail('knowledge_rebuild', response.error || 'Knowledge rebuild failed')
          }
          deps.completeToolActivity(activityId, 'success')
          return ok('knowledge_rebuild', `Knowledge rebuild completed: ${scope}`, response as Record<string, unknown>)
        } catch (error) {
          deps.completeToolActivity(activityId, 'error', 'Knowledge rebuild failed')
          return fail('knowledge_rebuild', `Knowledge rebuild failed: ${error}`)
        }
      },
    }),
  ]
}
