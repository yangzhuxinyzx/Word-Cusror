import { createToolCallIR } from '../../tools/ir'
import type { ProviderToolCallAdapter } from './types'

function extractAnthropicToolBlocks(response: unknown): Array<{
  id?: string
  name: string
  input?: Record<string, unknown>
}> {
  if (!response || typeof response !== 'object') return []

  const directContent = (response as { content?: unknown[] }).content
  const blocks = Array.isArray(directContent)
    ? directContent
    : Array.isArray((response as { message?: { content?: unknown[] } }).message?.content)
      ? ((response as { message?: { content?: unknown[] } }).message?.content || [])
      : []

  return blocks
    .map((item) => {
      const block = item as {
        type?: string
        id?: string
        name?: string
        input?: Record<string, unknown>
      }
      if (block.type !== 'tool_use' || !block.name) return null
      return {
        id: block.id,
        name: block.name,
        input:
          block.input && typeof block.input === 'object'
            ? block.input
            : {},
      }
    })
    .filter((item): item is { id?: string; name: string; input?: Record<string, unknown> } => Boolean(item?.name))
}

export const AnthropicToolUseAdapter: ProviderToolCallAdapter = {
  provider: 'anthropic_messages',
  capabilities: {
    provider: 'anthropic_messages',
    supportsNativeToolUse: true,
    supportsReasoning: true,
    supportsMultimodal: true,
    supportsPromptCache: true,
    supportsDeferredTools: true,
    supportsStructuredToolSchema: true,
  },
  toToolCalls(response: unknown) {
    return extractAnthropicToolBlocks(response).map((block) =>
      createToolCallIR({
        toolName: block.name,
        input: block.input || {},
        source: 'native',
        rawInput: JSON.stringify(block.input || {}),
        metadata: block.id
          ? {
              nativeToolCallId: block.id,
            }
          : undefined,
      }),
    )
  },
  toAssistantConversationMessage(response: unknown) {
    const content = Array.isArray((response as { content?: unknown[] })?.content)
      ? (response as { content: unknown[] }).content
      : null
    if (!content || !content.some((item) => (item as { type?: string })?.type === 'tool_use')) {
      return null
    }

    return {
      role: 'assistant',
      content: '',
      nativePayload: {
        content,
      },
    }
  },
  fromToolResults(bindings) {
    const blocks = bindings
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
          type: 'tool_result',
          tool_use_id: nativeToolCallId,
          content: payload,
        }
      })
      .filter(Boolean)

    if (blocks.length === 0) return []

    return [
      {
        role: 'user',
        content: '',
        nativePayload: {
          content: blocks,
        },
      },
    ]
  },
}
