function createAIProxyService(options = {}) {
  const { fs, getDebugLogPath } = options
  const abortControllers = new Map()

  function emitDelta(event, payload) {
    try {
      if (event?.sender && !event.sender.isDestroyed()) {
        event.sender.send('ai-stream-delta', payload)
      }
    } catch {}
  }

  function isAnthropicFormat(baseUrl = '', model = '') {
    const urlLower = String(baseUrl).toLowerCase()
    const modelLower = String(model).toLowerCase()
    if (urlLower.includes('/claude') || urlLower.includes('anthropic')) {
      return true
    }
    return modelLower.startsWith('claude-') && !urlLower.includes('/v1/chat')
  }

  function toAnthropicContentPart(part) {
    if (typeof part === 'string') {
      return { type: 'text', text: part }
    }

    if (!part || typeof part !== 'object') {
      return null
    }

    if (part.type === 'text' && typeof part.text === 'string') {
      return part
    }

    if (part.type === 'image_url' && part.image_url?.url) {
      const url = part.image_url.url
      if (url.startsWith('data:')) {
        const match = url.match(/^data:(image\/[^;]+);base64,(.+)$/)
        if (match) {
          return {
            type: 'image',
            source: {
              type: 'base64',
              media_type: match[1],
              data: match[2],
            },
          }
        }
      }

      return { type: 'text', text: `[image: ${url}]` }
    }

    return null
  }

  function buildAnthropicMessages(messages) {
    let systemPrompt = ''
    const anthropicMessages = []

    for (const message of messages) {
      if (message.role === 'system') {
        systemPrompt +=
          (systemPrompt ? '\n\n' : '') +
          (typeof message.content === 'string'
            ? message.content
            : JSON.stringify(message.content))
        continue
      }

      let content = message.content
      if (
        message.nativePayload &&
        typeof message.nativePayload === 'object' &&
        Array.isArray(message.nativePayload.content)
      ) {
        content = message.nativePayload.content
      } else if (Array.isArray(content)) {
        content = content.map(toAnthropicContentPart).filter(Boolean)
      }

      anthropicMessages.push({
        role: message.role === 'assistant' ? 'assistant' : 'user',
        content,
      })
    }

    if (anthropicMessages.length > 0 && anthropicMessages[0].role === 'assistant') {
      anthropicMessages.unshift({ role: 'user', content: 'Please continue.' })
    }

    return { systemPrompt, anthropicMessages }
  }

  function buildOpenAIRequestMessages(messages) {
    return messages.map((message) => {
      const base = {
        role: message.role,
        content: message.content,
      }
      if (message.nativePayload && typeof message.nativePayload === 'object') {
        return {
          ...base,
          ...message.nativePayload,
        }
      }
      return base
    })
  }

  function extractStreamText(value) {
    if (!value) return ''
    if (typeof value === 'string') return value

    if (Array.isArray(value)) {
      return value
        .map((part) => {
          if (!part) return ''
          if (typeof part === 'string') return part
          if (typeof part.text === 'string') return part.text
          if (typeof part.content === 'string') return part.content
          return ''
        })
        .join('')
    }

    if (typeof value === 'object') {
      if (typeof value.text === 'string') return value.text
      if (typeof value.content === 'string') return value.content
    }

    return ''
  }

  function extractThinkingFromContent(fullContent, currentReasoning) {
    const thinkMatch = fullContent.match(/<think>([\s\S]*?)<\/think>/g)
    if (!thinkMatch) return currentReasoning
    const extracted = thinkMatch
      .map((item) => item.replace(/<\/?think>/g, ''))
      .join('\n')
    return extracted || currentReasoning
  }

  function writeDebugEntry(entry) {
    if (!fs || typeof getDebugLogPath !== 'function') return
    const debugLogPath = getDebugLogPath()
    if (!debugLogPath) return

    try {
      fs.appendFileSync(debugLogPath, `${JSON.stringify(entry)}\n`)
    } catch {}
  }

  async function streamAnthropicRequest({
    event,
    requestId,
    controller,
    apiKey,
    baseUrl,
    model,
    messages,
    temperature,
    maxTokens,
  }) {
    const { systemPrompt, anthropicMessages } = buildAnthropicMessages(messages)
    const url = `${baseUrl.replace(/\/$/, '')}/v1/messages`
    const headers = {
      'Content-Type': 'application/json',
      'anthropic-version': '2023-06-01',
    }

    if (apiKey) {
      headers['x-api-key'] = apiKey
    }

    const body = {
      model,
      messages: anthropicMessages,
      max_tokens: maxTokens || 4096,
      stream: true,
    }

    if (systemPrompt) body.system = systemPrompt
    if (temperature !== undefined && temperature !== null) {
      body.temperature = temperature
    }

    const response = await fetch(url, {
      method: 'POST',
      headers,
      signal: controller.signal,
      body: JSON.stringify(body),
    })

    if (!response.ok) {
      const errorText = await response.text().catch(() => '')
      return {
        success: false,
        error: errorText || `Request failed with HTTP ${response.status}`,
      }
    }

    const reader = response.body?.getReader?.()
    if (!reader) {
      const text = await response.text().catch(() => '')
      return { success: false, error: text || 'Streaming response is not readable' }
    }

    const decoder = new TextDecoder()
    let fullContent = ''
    let fullReasoning = ''
    let buffer = ''

    while (true) {
      const { done, value } = await reader.read()
      if (done) break

      buffer += decoder.decode(value, { stream: true })
      const lines = buffer.split('\n')
      buffer = lines.pop() || ''

      for (const line of lines) {
        if (!line.startsWith('data: ')) continue
        const data = line.slice(6).trim()
        if (!data || data === '[DONE]') continue

        try {
          const json = JSON.parse(data)
          let contentDelta = ''
          let reasoningDelta = ''

          if (json.type === 'content_block_delta') {
            if (json.delta?.type === 'text_delta') {
              contentDelta = json.delta.text || ''
            } else if (json.delta?.type === 'thinking_delta') {
              reasoningDelta = json.delta.thinking || ''
            }
          }

          if (!contentDelta && !reasoningDelta) continue

          if (contentDelta) fullContent += contentDelta
          if (reasoningDelta) fullReasoning += reasoningDelta

          emitDelta(event, {
            requestId,
            delta: contentDelta,
            reasoningDelta,
            fullContent,
            fullReasoning,
          })
        } catch {}
      }
    }

    return { success: true, content: fullContent }
  }

  async function streamOpenAICompatibleRequest({
    event,
    requestId,
    controller,
    apiKey,
    baseUrl,
    model,
    messages,
    temperature,
    maxTokens,
  }) {
    const url = `${baseUrl.replace(/\/$/, '')}/chat/completions`
    const headers = {
      'Content-Type': 'application/json',
    }

    if (apiKey) {
      headers.Authorization = `Bearer ${apiKey}`
    }

    const response = await fetch(url, {
      method: 'POST',
      headers,
      signal: controller.signal,
      body: JSON.stringify({
        model,
        messages: buildOpenAIRequestMessages(messages),
        temperature,
        max_tokens: maxTokens,
        stream: true,
      }),
    })

    if (!response.ok) {
      const errorText = await response.text().catch(() => '')
      return {
        success: false,
        error: errorText || `Request failed with HTTP ${response.status}`,
      }
    }

    const reader = response.body?.getReader?.()
    if (!reader) {
      const text = await response.text().catch(() => '')
      return { success: false, error: text || 'Streaming response is not readable' }
    }

    const decoder = new TextDecoder()
    let fullContent = ''
    let fullReasoning = ''
    let buffer = ''
    let loggedDeltaShape = false

    while (true) {
      const { done, value } = await reader.read()
      if (done) break

      buffer += decoder.decode(value, { stream: true })
      const lines = buffer.split('\n')
      buffer = lines.pop() || ''

      for (const line of lines) {
        if (!line.startsWith('data: ')) continue
        const data = line.slice(6).trim()
        if (!data || data === '[DONE]') continue

        try {
          const json = JSON.parse(data)
          const choice = json.choices?.[0] || {}
          const delta = choice.delta || {}
          const message = choice.message || {}

          let contentDelta = delta.content || delta.text || ''
          const reasoningDelta =
            delta.reasoning_content ||
            delta.thinking_content ||
            delta.reasoning ||
            delta.thinking ||
            delta.thought ||
            message.reasoning_content ||
            json.reasoning_content ||
            ''

          if (!contentDelta) {
            const messageContent = extractStreamText(message.content ?? message.text)
            if (messageContent) {
              contentDelta = messageContent
            } else if (choice.text) {
              contentDelta = choice.text
            }
          }

          if (contentDelta && fullContent && contentDelta.startsWith(fullContent)) {
            contentDelta = contentDelta.slice(fullContent.length)
          }

          if (!loggedDeltaShape && delta && (contentDelta || reasoningDelta)) {
            loggedDeltaShape = true
            writeDebugEntry({
              location: 'ai-proxy-stream',
              message: 'delta structure',
              data: {
                deltaKeys: Object.keys(delta || {}),
                hasReasoning: Boolean(reasoningDelta),
                sample: JSON.stringify(delta).slice(0, 300),
              },
              timestamp: Date.now(),
            })
          }

          if (!contentDelta && !reasoningDelta) continue

          if (contentDelta) fullContent += contentDelta
          if (reasoningDelta) fullReasoning += reasoningDelta
          fullReasoning = extractThinkingFromContent(fullContent, fullReasoning)

          emitDelta(event, {
            requestId,
            delta: contentDelta,
            reasoningDelta,
            fullContent,
            fullReasoning,
          })
        } catch {}
      }
    }

    return { success: true, content: fullContent }
  }

  function extractTextFromAnthropicResponse(responseBody) {
    const content = Array.isArray(responseBody?.content) ? responseBody.content : []
    return content
      .filter((item) => item?.type === 'text' && typeof item.text === 'string')
      .map((item) => item.text)
      .join('')
  }

  function extractTextFromOpenAIMessage(message) {
    if (!message) return ''
    if (typeof message.content === 'string') return message.content
    if (Array.isArray(message.content)) {
      return message.content
        .map((item) => {
          if (!item) return ''
          if (typeof item === 'string') return item
          if (typeof item.text === 'string') return item.text
          return ''
        })
        .join('')
    }
    return ''
  }

  async function executeAnthropicNativeToolRequest({
    controller,
    apiKey,
    baseUrl,
    model,
    messages,
    temperature,
    maxTokens,
    tools,
  }) {
    const { systemPrompt, anthropicMessages } = buildAnthropicMessages(messages)
    const url = `${baseUrl.replace(/\/$/, '')}/v1/messages`
    const headers = {
      'Content-Type': 'application/json',
      'anthropic-version': '2023-06-01',
    }

    if (apiKey) {
      headers['x-api-key'] = apiKey
    }

    const body = {
      model,
      messages: anthropicMessages,
      max_tokens: maxTokens || 4096,
      tools: Array.isArray(tools) ? tools : [],
    }

    if (systemPrompt) body.system = systemPrompt
    if (temperature !== undefined && temperature !== null) {
      body.temperature = temperature
    }

    const response = await fetch(url, {
      method: 'POST',
      headers,
      signal: controller.signal,
      body: JSON.stringify(body),
    })

    if (!response.ok) {
      const errorText = await response.text().catch(() => '')
      return {
        success: false,
        error: errorText || `Request failed with HTTP ${response.status}`,
      }
    }

    const data = await response.json()
    return {
      success: true,
      content: extractTextFromAnthropicResponse(data),
      rawResponse: data,
    }
  }

  async function executeOpenAINativeToolRequest({
    controller,
    apiKey,
    baseUrl,
    model,
    messages,
    temperature,
    maxTokens,
    tools,
  }) {
    const url = `${baseUrl.replace(/\/$/, '')}/chat/completions`
    const headers = {
      'Content-Type': 'application/json',
    }

    if (apiKey) {
      headers.Authorization = `Bearer ${apiKey}`
    }

    const response = await fetch(url, {
      method: 'POST',
      headers,
      signal: controller.signal,
      body: JSON.stringify({
        model,
        messages: buildOpenAIRequestMessages(messages),
        temperature,
        max_tokens: maxTokens,
        tools: Array.isArray(tools) ? tools : [],
        tool_choice: 'auto',
        stream: false,
      }),
    })

    if (!response.ok) {
      const errorText = await response.text().catch(() => '')
      return {
        success: false,
        error: errorText || `Request failed with HTTP ${response.status}`,
      }
    }

    const data = await response.json()
    const message = data?.choices?.[0]?.message
    return {
      success: true,
      content: extractTextFromOpenAIMessage(message),
      rawResponse: data,
    }
  }

  return {
    async chatCompletions(event, payload = {}) {
      const {
        requestId,
        baseUrl,
        apiKey,
        model,
        messages,
        temperature,
        maxTokens,
        tools,
      } = payload

      if (!requestId) return { success: false, error: 'Missing requestId' }
      if (!baseUrl) return { success: false, error: 'Missing baseUrl' }
      if (!model) return { success: false, error: 'Missing model' }
      if (!Array.isArray(messages) || messages.length === 0) {
        return { success: false, error: 'Missing messages' }
      }

      const controller = new AbortController()
      abortControllers.set(requestId, controller)

      try {
        if (Array.isArray(tools) && tools.length > 0) {
          if (isAnthropicFormat(baseUrl, model)) {
            return await executeAnthropicNativeToolRequest({
              controller,
              apiKey,
              baseUrl,
              model,
              messages,
              temperature,
              maxTokens,
              tools,
            })
          }

          return await executeOpenAINativeToolRequest({
            controller,
            apiKey,
            baseUrl,
            model,
            messages,
            temperature,
            maxTokens,
            tools,
          })
        }

        if (isAnthropicFormat(baseUrl, model)) {
          return await streamAnthropicRequest({
            event,
            requestId,
            controller,
            apiKey,
            baseUrl,
            model,
            messages,
            temperature,
            maxTokens,
          })
        }

        return await streamOpenAICompatibleRequest({
          event,
          requestId,
          controller,
          apiKey,
          baseUrl,
          model,
          messages,
          temperature,
          maxTokens,
        })
      } catch (error) {
        if (error?.name === 'AbortError') {
          return { success: false, error: 'Request aborted' }
        }
        return { success: false, error: error?.message || String(error) }
      } finally {
        abortControllers.delete(requestId)
      }
    },

    async cancel(requestId) {
      try {
        const controller = abortControllers.get(requestId)
        if (controller) {
          controller.abort()
        }
        abortControllers.delete(requestId)
        return { success: true }
      } catch (error) {
        return { success: false, error: error?.message || String(error) }
      }
    },

    async close() {
      for (const controller of abortControllers.values()) {
        try {
          controller.abort()
        } catch {}
      }
      abortControllers.clear()
    },
  }
}

module.exports = {
  createAIProxyService,
}
