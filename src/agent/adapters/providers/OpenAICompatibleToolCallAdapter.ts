import { createToolCallIR } from '../../tools/ir'
import { decodeOpenAIToolName, encodeOpenAIToolName } from './openAIToolNameCodec'
import type { ProviderToolCallAdapter } from './types'

function extractOpenAIMessageText(message: unknown): string {
  if (!message || typeof message !== 'object') return ''
  const content = (message as { content?: unknown }).content
  if (typeof content === 'string') return content
  if (Array.isArray(content)) {
    return content
      .map((item) => {
        if (!item || typeof item !== 'object') return ''
        if (typeof (item as { text?: unknown }).text === 'string') {
          return (item as { text: string }).text
        }
        return ''
      })
      .join('')
  }
  return ''
}

function normalizeOpenAIToolInput(
  value: unknown,
): Record<string, unknown> {
  if (!value) return {}
  if (typeof value === 'string') {
    try {
      const parsed = JSON.parse(value)
      return parsed && typeof parsed === 'object' && !Array.isArray(parsed)
        ? (parsed as Record<string, unknown>)
        : {}
    } catch {
      return {}
    }
  }
  if (typeof value === 'object' && !Array.isArray(value)) {
    return value as Record<string, unknown>
  }
  return {}
}

function extractOpenAIToolCalls(response: unknown): Array<{
  id?: string
  name: string
  arguments?: unknown
}> {
  if (!response || typeof response !== 'object') return []

  if (
    'tool_calls' in (response as Record<string, unknown>) &&
    Array.isArray((response as { tool_calls?: unknown[] }).tool_calls)
  ) {
    return ((response as { tool_calls?: unknown[] }).tool_calls || [])
      .map((item) => {
        const call = item as {
          id?: string
          function?: { name?: string; arguments?: unknown }
          name?: string
          arguments?: unknown
        }
        if (call.function?.name) {
          return {
            id: call.id,
            name: call.function.name,
            arguments: call.function.arguments,
          }
        }
        if (call.name) {
          return {
            id: call.id,
            name: call.name,
            arguments: call.arguments,
          }
        }
        return null
      })
      .filter((item): item is { id?: string; name: string; arguments?: unknown } => Boolean(item?.name))
  }

  const choices = (response as { choices?: unknown[] }).choices
  if (!Array.isArray(choices) || choices.length === 0) return []

  const firstChoice = choices[0] as {
    message?: {
      tool_calls?: Array<{
        id?: string
        function?: { name?: string; arguments?: unknown }
      }>
    }
  }

  const toolCalls = firstChoice.message?.tool_calls
  if (!Array.isArray(toolCalls)) return []

  return toolCalls
    .map((toolCall) => {
      const name = toolCall.function?.name
      if (!name) return null
      return {
        id: toolCall.id,
        name,
        arguments: toolCall.function?.arguments,
      }
    })
    .filter((item): item is { id?: string; name: string; arguments?: unknown } => Boolean(item?.name))
}

export const OpenAICompatibleToolCallAdapter: ProviderToolCallAdapter = {
  provider: 'openai_compatible',
  capabilities: {
    provider: 'openai_compatible',
    supportsNativeToolUse: true,
    supportsReasoning: false,
    supportsMultimodal: true,
    supportsPromptCache: false,
    supportsDeferredTools: false,
    supportsStructuredToolSchema: true,
  },
  toToolCalls(response: unknown) {
    return extractOpenAIToolCalls(response).map((call) =>
      createToolCallIR({
        toolName: decodeOpenAIToolName(call.name),
        input: normalizeOpenAIToolInput(call.arguments),
        source: 'native',
        rawInput:
          typeof call.arguments === 'string'
            ? call.arguments
            : JSON.stringify(call.arguments || {}),
        metadata: call.id
          ? {
              nativeToolCallId: call.id,
            }
          : undefined,
      }),
    )
  },
  toAssistantConversationMessage(response: unknown) {
    const choices = (response as { choices?: unknown[] })?.choices
    const message = Array.isArray(choices)
      ? (choices[0] as { message?: Record<string, unknown> })?.message
      : undefined
    if (!message || !Array.isArray((message as { tool_calls?: unknown[] }).tool_calls)) {
      return null
    }

    return {
      role: 'assistant',
      content: extractOpenAIMessageText(message),
      nativePayload: {
        tool_calls: (message as { tool_calls: unknown[] }).tool_calls,
      },
    }
  },
  fromToolResults(bindings) {
    return bindings
      .map(({ call, result }) => {
        const nativeToolCallId = typeof call.metadata?.nativeToolCallId === 'string'
          ? call.metadata.nativeToolCallId
          : null
        if (!nativeToolCallId) return null

        const payload =
          result.data && Object.keys(result.data).length > 0
            ? `${result.message}\n${JSON.stringify(result.data, null, 2)}`
            : result.message

        return {
          role: 'tool',
          content: payload,
          nativePayload: {
            tool_call_id: nativeToolCallId,
            name: encodeOpenAIToolName(call.toolName),
          },
        }
      })
      .filter(Boolean) as Array<{
      role: string
      content: string
      nativePayload?: Record<string, unknown>
    }>
  },
}
