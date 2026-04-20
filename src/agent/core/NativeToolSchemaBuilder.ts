import type { AISettings } from '../../types'
import { encodeOpenAIToolName, usesEncodedOpenAIToolName } from '../adapters/providers/openAIToolNameCodec'
import type { ProviderKind } from '../adapters/providers/types'
import { getPreferredProviderToolAdapter } from './ToolCallProtocolRouter'
import type { AgentToolDefinition } from '../tools/contracts'

export type NativeToolingConfig = {
  enabled: boolean
  provider: ProviderKind
  tools: unknown[]
  systemPrompt: string
}

function supportsOpenAINativeTooling(baseUrl: string): boolean {
  const lowered = String(baseUrl || '').toLowerCase()
  if (!lowered) return false

  try {
    const { hostname } = new URL(baseUrl)
    const host = hostname.toLowerCase()
    return (
      host === 'api.openai.com' ||
      host.endsWith('.openai.com') ||
      host.endsWith('.openai.azure.com') ||
      host === 'openrouter.ai' ||
      host.endsWith('.openrouter.ai')
    )
  } catch {
    return (
      lowered.includes('api.openai.com') ||
      lowered.includes('.openai.azure.com') ||
      lowered.includes('openrouter.ai')
    )
  }
}

function describeInputKey(tool: AgentToolDefinition, key: string): string {
  return `${tool.displayName} parameter: ${key}`
}

function buildSchemaProperties(tool: AgentToolDefinition) {
  return Object.fromEntries(
    (tool.inputKeys || []).map((key) => [
      key,
      {
        type: 'string',
        description: describeInputKey(tool, key),
      },
    ]),
  )
}

function buildToolInputSchema(tool: AgentToolDefinition) {
  if (tool.inputSchema) {
    return {
      additionalProperties: true,
      ...tool.inputSchema,
    }
  }

  return {
    type: 'object',
    properties: buildSchemaProperties(tool),
    additionalProperties: true,
  }
}

function buildOpenAISchemas(tools: AgentToolDefinition[]): unknown[] {
  return tools.map((tool) => {
    const wireName = encodeOpenAIToolName(tool.id)
    const description = tool.prompt || tool.description

    return {
      type: 'function',
      function: {
        name: wireName,
        description:
          wireName === tool.id
            ? description
            : `Canonical tool id: ${tool.id}\n${description}`,
        parameters: buildToolInputSchema(tool),
      },
    }
  })
}

function buildAnthropicSchemas(tools: AgentToolDefinition[]): unknown[] {
  return tools.map((tool) => ({
    name: tool.id,
    description: tool.prompt || tool.description,
    input_schema: buildToolInputSchema(tool),
  }))
}

export function buildNativeToolingConfig(params: {
  settings: Pick<AISettings, 'baseUrl' | 'model'>
  tools: AgentToolDefinition[]
}): NativeToolingConfig {
  const adapter = getPreferredProviderToolAdapter(params.settings)

  if (
    !adapter.capabilities.supportsNativeToolUse ||
    !adapter.capabilities.supportsStructuredToolSchema ||
    (adapter.provider === 'openai_compatible' &&
      !supportsOpenAINativeTooling(params.settings.baseUrl)) ||
    params.tools.length === 0
  ) {
    return {
      enabled: false,
      provider: adapter.provider,
      tools: [],
      systemPrompt: '',
    }
  }

  const tools =
    adapter.provider === 'anthropic_messages'
      ? buildAnthropicSchemas(params.tools)
      : buildOpenAISchemas(params.tools)
  const hasEncodedToolNames =
    adapter.provider === 'openai_compatible' &&
    params.tools.some((tool) => usesEncodedOpenAIToolName(tool.id))

  return {
    enabled: true,
    provider: adapter.provider,
    tools,
    systemPrompt: [
      'Tool protocol override:',
      'Use the provider-native tool calling interface for any tool invocation.',
      'Do not emit [TOOL_CALL], [/TOOL_CALL], [TOOL_RESULT], XML tool tags, or <tool_use> text blocks.',
      'When a tool is needed, call the native tool directly with structured arguments.',
      ...(hasEncodedToolNames
        ? ['If the provider schema exposes transport-encoded tool names, use the exact schema names for tool calls.']
        : []),
      'When no tool is needed, reply with normal assistant text.',
    ].join('\n'),
  }
}
