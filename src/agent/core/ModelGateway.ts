import type { AISettings } from '../../types'
import { cleanLegacyModelOutput, extractLegacyStreamText, extractLegacyThinking } from '../compat/legacyModelOutput'
import {
  emitLegacyToolCallPreviewFromRaw,
  emitLegacyToolCallStartFromRaw,
  parseLegacyAssistantOutput,
  trimTrailingOpenLegacyToolBlock,
} from '../compat/legacyProtocol'
import type { MessageContent } from './runtimeTypes'

export interface CallModelOptions {
  returnRaw?: boolean
  onToolCallStart?: (tool: string) => void
  onToolCallPreview?: (tool: string, args: Record<string, string>) => void
  maxToolCallPreviews?: number
  onResponseFinal?: (payload: { raw: string; cleaned: string }) => void
  nativeTools?: unknown[]
}

export interface ModelGatewayResponse {
  rawResponse: unknown
  rawText: string
  cleanedText: string
  displayText: string
}

export interface ModelGatewayParams {
  settings: AISettings
  allMessages: Array<{
    role: string
    content: MessageContent
    nativePayload?: Record<string, unknown>
  }>
  signal: AbortSignal
  streamPrefixRef: { current: string }
  setStreamingContent: (value: string) => void
  setStreamingReasoning: (value: string) => void
  setStreamingSummary: (value: string) => void
  setEditPhase: (value: 'idle' | 'editing' | 'done') => void
  options?: CallModelOptions
}

async function streamFallbackContent(
  content: string,
  signal: AbortSignal,
  setStreamingContent: (value: string) => void,
): Promise<void> {
  if (!content) return
  const total = content.length
  const chunkSize =
    total > 4000 ? 120 : total > 1500 ? 60 : total > 400 ? 30 : 12
  const delay = 24
  for (let i = 0; i < total; i += chunkSize) {
    if (signal.aborted) break
    setStreamingContent(content.slice(0, i + chunkSize))
    await new Promise((resolve) => setTimeout(resolve, delay))
  }
}

export async function callLegacyModelGateway({
  settings,
  allMessages,
  signal,
  streamPrefixRef,
  setStreamingContent,
  setStreamingReasoning,
  setStreamingSummary,
  setEditPhase,
  options,
}: ModelGatewayParams): Promise<ModelGatewayResponse> {
  const headers: Record<string, string> = {
    'Content-Type': 'application/json',
  }
  if (settings.apiKey) {
    headers.Authorization = `Bearer ${settings.apiKey}`
  }

  const toolPreviewState = {
    emitted: new Set<string>(),
    startedSignatures: new Set<string>(),
  }

  if (
    window.electronAPI?.isElectron &&
    window.electronAPI.aiChatCompletions &&
    window.electronAPI.onAIStreamDelta
  ) {
    const requestId = `${Date.now()}-${Math.random().toString(16).slice(2)}`
    let fullContent = ''
    let fullReasoning = ''
    let contentDeltaCount = 0

    setStreamingReasoning('')

    const unsubscribe = window.electronAPI.onAIStreamDelta((payload) => {
      if (!payload || payload.requestId !== requestId) return

      const delta = payload.delta || ''
      if (delta) {
        contentDeltaCount += 1
        fullContent += delta
        emitLegacyToolCallStartFromRaw(
          fullContent,
          toolPreviewState,
          options?.onToolCallStart,
          options?.maxToolCallPreviews,
        )
        emitLegacyToolCallPreviewFromRaw(
          fullContent,
          cleanLegacyModelOutput(fullContent),
          toolPreviewState,
          options?.onToolCallPreview,
          options?.maxToolCallPreviews,
        )
        const { thinking, cleaned } = extractLegacyThinking(fullContent)
        if (thinking && thinking !== fullReasoning) {
          fullReasoning = thinking
          setStreamingReasoning(fullReasoning)
        }
        const parsed = parseLegacyAssistantOutput(cleaned)
        if (parsed.summary) setStreamingSummary(parsed.summary)
        if (parsed.phase === 'editing') setEditPhase('editing')
        if (parsed.phase === 'done') setEditPhase('done')
        const safeDisplay = trimTrailingOpenLegacyToolBlock(parsed.displayText)
        setStreamingContent(streamPrefixRef.current + safeDisplay)
      }

      const reasoningDelta =
        (payload as { reasoningDelta?: string }).reasoningDelta || ''
      if (reasoningDelta) {
        fullReasoning += reasoningDelta
        setStreamingReasoning(fullReasoning)
      }
    })

    const abortHandler = () => {
      try {
        window.electronAPI?.aiCancel?.(requestId)
      } catch {
        // ignore
      }
    }

    signal.addEventListener('abort', abortHandler, { once: true })

    try {
      const effectiveTemperature = settings.model
        .toLowerCase()
        .includes('kimi-k2')
        ? 1
        : settings.temperature

      const result = await window.electronAPI.aiChatCompletions({
        requestId,
        baseUrl: settings.baseUrl,
        apiKey: settings.apiKey,
        model: settings.model,
        messages: allMessages,
        temperature: effectiveTemperature,
        maxTokens: settings.maxTokens,
        tools: options?.nativeTools,
      })

      if (!result?.success) {
        throw new Error(result?.error || '请求失败')
      }

      const content = result.content || fullContent
      emitLegacyToolCallStartFromRaw(
        content,
        toolPreviewState,
        options?.onToolCallStart,
        options?.maxToolCallPreviews,
      )
      emitLegacyToolCallPreviewFromRaw(
        content,
        cleanLegacyModelOutput(content),
        toolPreviewState,
        options?.onToolCallPreview,
        options?.maxToolCallPreviews,
      )
      const cleaned = cleanLegacyModelOutput(content)
      options?.onResponseFinal?.({ raw: content, cleaned })
      const parsed = parseLegacyAssistantOutput(cleaned)
      const displayText = parsed.displayText || parsed.summary || cleaned
      if (parsed.summary) setStreamingSummary(parsed.summary)
      if (parsed.phase === 'editing') setEditPhase('editing')
      if (parsed.phase === 'done') setEditPhase('done')
      if (contentDeltaCount === 0 && displayText) {
        await streamFallbackContent(displayText, signal, setStreamingContent)
      }
      return {
        rawResponse: result.rawResponse ?? content,
        rawText: content,
        cleanedText: cleaned,
        displayText: options?.returnRaw ? cleaned : displayText,
      }
    } finally {
      try {
        unsubscribe?.()
      } catch {
        // ignore
      }
    }
  }

  const effectiveTemperature = settings.model.toLowerCase().includes('kimi-k2')
    ? 1
    : settings.temperature

  const response = await fetch(`${settings.baseUrl}/chat/completions`, {
    method: 'POST',
    headers,
    signal,
    body: JSON.stringify({
      model: settings.model,
      messages: allMessages,
      temperature: effectiveTemperature,
      max_tokens: settings.maxTokens,
      stream: true,
      tools: options?.nativeTools,
    }),
  })

  if (!response.ok) {
    const errorText = await response.text()
    throw new Error(errorText || '请求失败')
  }

  const reader = response.body?.getReader()
  if (!reader) throw new Error('无法读取响应')

  const decoder = new TextDecoder()
  let fullContent = ''
  let fullReasoning = ''
  let buffer = ''

  setStreamingReasoning('')

  const readWithTimeout = async (timeoutMs: number) => {
    const timeoutPromise = new Promise<{ done: true; value: undefined }>(
      (_, reject) => {
        setTimeout(() => reject(new Error('读取超时')), timeoutMs)
      },
    )
    return Promise.race([reader.read(), timeoutPromise])
  }

  const READ_TIMEOUT = 60000

  while (true) {
    let result
    try {
      result = await readWithTimeout(READ_TIMEOUT)
    } catch {
      console.warn('[API] 流响应读取超时，返回已有内容')
      break
    }

    const { done, value } = result
    if (done) break

    buffer += decoder.decode(value, { stream: true })
    const lines = buffer.split('\n')
    buffer = lines.pop() || ''

    for (const line of lines) {
      if (!line.startsWith('data: ')) continue
      const data = line.slice(6).trim()
      if (data === '[DONE]') continue

      try {
        const json = JSON.parse(data)
        const choice = json.choices?.[0] || {}
        const delta = choice?.delta || {}
        const message = choice?.message || {}

        const reasoningDelta =
          delta?.reasoning_content ||
          delta?.thinking_content ||
          delta?.reasoning ||
          delta?.thinking ||
          delta?.thought ||
          message?.reasoning_content ||
          json?.reasoning_content ||
          ''
        if (reasoningDelta) {
          fullReasoning += reasoningDelta
          setStreamingReasoning(fullReasoning)
        }

        let contentDelta = delta?.content || delta?.text || ''
        if (!contentDelta) {
          const messageContent = extractLegacyStreamText(
            message?.content ?? message?.text,
          )
          if (messageContent) {
            contentDelta = messageContent
          } else if (choice?.text) {
            contentDelta = choice.text
          }
        }

        if (contentDelta && fullContent && contentDelta.startsWith(fullContent)) {
          contentDelta = contentDelta.slice(fullContent.length)
        }

        if (contentDelta) {
          fullContent += contentDelta
          emitLegacyToolCallStartFromRaw(
            fullContent,
            toolPreviewState,
            options?.onToolCallStart,
            options?.maxToolCallPreviews,
          )
          emitLegacyToolCallPreviewFromRaw(
            fullContent,
            cleanLegacyModelOutput(fullContent),
            toolPreviewState,
            options?.onToolCallPreview,
            options?.maxToolCallPreviews,
          )
          const { thinking, cleaned } = extractLegacyThinking(fullContent)
          if (thinking && thinking !== fullReasoning) {
            fullReasoning = thinking
            setStreamingReasoning(fullReasoning)
          }
          const parsed = parseLegacyAssistantOutput(cleaned)
          if (parsed.summary) setStreamingSummary(parsed.summary)
          if (parsed.phase === 'editing') setEditPhase('editing')
          if (parsed.phase === 'done') setEditPhase('done')
          const safeDisplay = trimTrailingOpenLegacyToolBlock(parsed.displayText)
          setStreamingContent(streamPrefixRef.current + safeDisplay)
        }
      } catch {
        // ignore stream parse errors
      }
    }
  }

  emitLegacyToolCallStartFromRaw(
    fullContent,
    toolPreviewState,
    options?.onToolCallStart,
    options?.maxToolCallPreviews,
  )
  emitLegacyToolCallPreviewFromRaw(
    fullContent,
    cleanLegacyModelOutput(fullContent),
    toolPreviewState,
    options?.onToolCallPreview,
    options?.maxToolCallPreviews,
  )

  const cleaned = cleanLegacyModelOutput(fullContent)
  options?.onResponseFinal?.({ raw: fullContent, cleaned })
  const parsed = parseLegacyAssistantOutput(cleaned)
  if (parsed.summary) setStreamingSummary(parsed.summary)
  if (parsed.phase === 'editing') setEditPhase('editing')
  if (parsed.phase === 'done') setEditPhase('done')
  const displayText = parsed.displayText || parsed.summary || cleaned
  return {
    rawResponse: fullContent,
    rawText: fullContent,
    cleanedText: cleaned,
    displayText: options?.returnRaw ? cleaned : displayText,
  }
}
