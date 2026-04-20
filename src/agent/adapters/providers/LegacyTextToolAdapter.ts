import { parseLegacyToolCallsToIR } from '../../compat/legacyProtocol'
import type { ProviderToolCallAdapter } from './types'

export const LegacyTextToolAdapter: ProviderToolCallAdapter = {
  provider: 'legacy_text',
  capabilities: {
    provider: 'legacy_text',
    supportsNativeToolUse: false,
    supportsReasoning: false,
    supportsMultimodal: false,
    supportsPromptCache: false,
    supportsDeferredTools: false,
    supportsStructuredToolSchema: false,
  },
  toToolCalls(response: unknown) {
    if (typeof response !== 'string') return []
    return parseLegacyToolCallsToIR(response)
  },
}

