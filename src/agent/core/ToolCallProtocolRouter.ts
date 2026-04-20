import type { AISettings } from '../../types'
import { AnthropicToolUseAdapter } from '../adapters/providers/AnthropicToolUseAdapter'
import { LegacyTextToolAdapter } from '../adapters/providers/LegacyTextToolAdapter'
import { OpenAICompatibleToolCallAdapter } from '../adapters/providers/OpenAICompatibleToolCallAdapter'
import type {
  ProviderConversationMessage,
  ProviderKind,
  ProviderToolResultBinding,
  ProviderToolCallAdapter,
} from '../adapters/providers/types'
import type { ToolCallIR } from '../tools/ir'

export interface ToolCallProtocolRouterOptions {
  settings: Pick<AISettings, 'baseUrl' | 'model'>
  response: unknown
  turnId?: string
  metadata?: Record<string, unknown>
  adapters?: ProviderToolCallAdapter[]
}

const DEFAULT_PROVIDER_ADAPTERS: ProviderToolCallAdapter[] = [
  OpenAICompatibleToolCallAdapter,
  AnthropicToolUseAdapter,
  LegacyTextToolAdapter,
]

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === 'object' && !Array.isArray(value)
}

function detectProviderKind(
  settings: Pick<AISettings, 'baseUrl' | 'model'>,
): ProviderKind {
  const baseUrl = String(settings.baseUrl || '').toLowerCase()
  const model = String(settings.model || '').toLowerCase()

  if (
    baseUrl.includes('anthropic') ||
    baseUrl.includes('/v1/messages') ||
    baseUrl.includes('/claude') ||
    model.startsWith('claude-')
  ) {
    return 'anthropic_messages'
  }

  return 'openai_compatible'
}

function orderAdapters(
  settings: Pick<AISettings, 'baseUrl' | 'model'>,
  adapters?: ProviderToolCallAdapter[],
): ProviderToolCallAdapter[] {
  const provider = detectProviderKind(settings)
  const available = adapters && adapters.length > 0
    ? adapters
    : DEFAULT_PROVIDER_ADAPTERS

  return [...available].sort((left, right) => {
    if (left.provider === right.provider) return 0
    if (left.provider === provider) return -1
    if (right.provider === provider) return 1
    if (left.provider === 'legacy_text') return 1
    if (right.provider === 'legacy_text') return -1
    return 0
  })
}

function buildSignature(call: ToolCallIR): string {
  const normalizedInput = Object.entries(call.input || {})
    .sort(([left], [right]) => left.localeCompare(right))
    .map(([key, value]) => {
      if (typeof value === 'string') {
        return `${key}:${value}`
      }
      return `${key}:${JSON.stringify(value)}`
    })
    .join('|')

  return `${call.source}:${call.toolName}:${normalizedInput}`
}

function normalizeForCallbacks(call: ToolCallIR): Record<string, string> {
  const result: Record<string, string> = {}

  for (const [key, value] of Object.entries(call.input || {})) {
    if (value === undefined || value === null) continue
    if (typeof value === 'string') {
      result[key] = value
      continue
    }
    if (typeof value === 'number' || typeof value === 'boolean') {
      result[key] = String(value)
      continue
    }
    if (Array.isArray(value) || isRecord(value)) {
      try {
        result[key] = JSON.stringify(value)
      } catch {
        // ignore unserializable payloads
      }
    }
  }

  return result
}

export function parseToolCallsToIRFromResponse(
  options: ToolCallProtocolRouterOptions,
): ToolCallIR[] {
  const orderedAdapters = orderAdapters(options.settings, options.adapters)
  const seen = new Set<string>()
  const parsed: ToolCallIR[] = []

  for (const adapter of orderedAdapters) {
    if (!adapter.toToolCalls) continue
    const toolCalls = adapter.toToolCalls(options.response)
    if (!toolCalls || toolCalls.length === 0) continue

    for (const toolCall of toolCalls) {
      const signature = buildSignature(toolCall)
      if (seen.has(signature)) continue
      seen.add(signature)
      parsed.push({
        ...toolCall,
        turnId: toolCall.turnId || options.turnId,
        metadata: {
          ...(toolCall.metadata || {}),
          ...(options.metadata || {}),
          parsedByProvider: adapter.provider,
        },
      })
    }
  }

  return parsed
}

export function parseToolCallsForCallbacksFromResponse(
  options: ToolCallProtocolRouterOptions,
): Array<{
  tool: string
  args: Record<string, string>
  source: ToolCallIR['source']
}> {
  return parseToolCallsToIRFromResponse(options).map((call) => ({
    tool: call.toolName,
    args: normalizeForCallbacks(call),
    source: call.source,
  }))
}

export function hasToolCallInResponse(
  options: Omit<ToolCallProtocolRouterOptions, 'turnId' | 'metadata'>,
): boolean {
  return parseToolCallsToIRFromResponse({
    ...options,
    turnId: undefined,
    metadata: undefined,
  }).length > 0
}

export function getPreferredProviderToolAdapter(
  settings: Pick<AISettings, 'baseUrl' | 'model'>,
  adapters?: ProviderToolCallAdapter[],
): ProviderToolCallAdapter {
  return orderAdapters(settings, adapters)[0] || LegacyTextToolAdapter
}

export function buildNativeAssistantConversationMessageFromResponse(options: {
  settings: Pick<AISettings, 'baseUrl' | 'model'>
  response: unknown
  adapters?: ProviderToolCallAdapter[]
}): ProviderConversationMessage | null {
  const adapter = getPreferredProviderToolAdapter(options.settings, options.adapters)
  return adapter.toAssistantConversationMessage?.(options.response) || null
}

export function buildNativeToolResultConversationMessages(options: {
  settings: Pick<AISettings, 'baseUrl' | 'model'>
  bindings: ProviderToolResultBinding[]
  adapters?: ProviderToolCallAdapter[]
}): ProviderConversationMessage[] {
  const adapter = getPreferredProviderToolAdapter(options.settings, options.adapters)
  return adapter.fromToolResults?.(options.bindings) || []
}
