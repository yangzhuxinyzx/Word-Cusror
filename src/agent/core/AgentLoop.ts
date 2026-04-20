import { LEGACY_TOOL_NAMES } from '../compat/legacyTools'
import type {
  CallModelOptions,
  ModelGatewayResponse,
} from './ModelGateway'
import type { MessageContent } from './runtimeTypes'
import type {
  AgentCallbacks,
  AgentDebugEvent,
  ToolResult,
} from './runtimeTypes'
import type { ToolCallIR } from '../tools/ir'

export type AgentConversationMessage = {
  role: string
  content: MessageContent
  nativePayload?: Record<string, unknown>
}

export function createAgentTurnId(): string {
  return `turn-${Date.now()}-${Math.random().toString(16).slice(2)}`
}

export function sanitizeConversationMessagesForApi(
  messages: AgentConversationMessage[],
): void {
  for (let i = messages.length - 1; i >= 0; i -= 1) {
    const content = messages[i].content
    if (Array.isArray(content)) {
      if (content.length === 0) {
        messages[i].content = '[空]'
      }
    } else if (!content || !content.trim()) {
      messages[i].content = '[空]'
    }
  }
}

export function splitToolCallsForIteration<T>(
  parsedToolCalls: T[],
  maxToolCallsPerIteration: number,
): {
  currentBatch: T[]
  deferred: T[]
} {
  return {
    currentBatch: parsedToolCalls.slice(0, maxToolCallsPerIteration),
    deferred: parsedToolCalls.slice(maxToolCallsPerIteration),
  }
}

export function hasIncompleteLegacyToolCall(content: string): boolean {
  const openCount = (content.match(/\[TOOL_CALL\]/g) || []).length
  const closeCount = (content.match(/\[\/TOOL_CALL\]/g) || []).length
  if (openCount > 0 && openCount > closeCount) {
    return true
  }

  const lowered = content.toLowerCase()
  if (lowered.includes('<tool_use>') && !lowered.includes('</tool_use>')) {
    return true
  }

  for (const tool of LEGACY_TOOL_NAMES) {
    if (lowered.includes(`<${tool}>`) && !lowered.includes(`</${tool}>`)) {
      return true
    }
  }

  return false
}

export interface LegacyParsedToolCall {
  tool: string
  args: Record<string, string>
  source?: ToolCallIR['source']
}

export interface LegacyParsedAssistantOutput {
  displayText: string
  summary: string
  phase: 'idle' | 'editing' | 'done'
}

export interface LegacyToolCallLogger {
  setIteration: (iteration: number) => void
  logToolCallsParsed: (
    calls: Array<{ tool: string; args: Record<string, string> }>,
  ) => void
  logToolCallSkipped: (
    tool: string,
    reason: string,
    args?: Record<string, string>,
  ) => void
  logToolExecStart: (tool: string, args?: Record<string, string>) => void
  logToolExecResult: (
    tool: string,
    result: ToolResult,
    durationMs?: number,
  ) => void
  logToolResultsSent: (results: string[], totalReplaceCount: number) => void
  logTurnComplete: (payload: {
    totalIterations: number
    totalToolCalls: number
    totalSkipped: number
    stopReason: string
    finalResponseLength?: number
  }) => void
}

export interface RunLegacyAgentLoopOptions {
  conversationMessages: AgentConversationMessage[]
  signal: AbortSignal
  turnId: string
  callbacks?: AgentCallbacks
  callModel: (
    messages: AgentConversationMessage[],
    signal: AbortSignal,
    options?: CallModelOptions,
  ) => Promise<ModelGatewayResponse>
  emitDebugEvent: (event: AgentDebugEvent) => Promise<void>
  streamPrefixRef: { current: string }
  setStreamingContent: (value: string) => void
  getStreamingReasoning: () => string
  parseAssistantOutput: (
    content: string,
  ) => LegacyParsedAssistantOutput
  cleanModelOutput: (content: string) => string
  hasToolCall: (content: string) => boolean
  parseToolCallsToIR: (
    content: unknown,
    context?: {
      turnId?: string
      metadata?: Record<string, unknown>
    },
  ) => ToolCallIR[]
  buildNativeAssistantMessage?: (
    response: unknown,
  ) => AgentConversationMessage | null
  buildNativeToolResultMessages?: (
    bindings: Array<{ call: ToolCallIR; result: ToolResult }>,
  ) => AgentConversationMessage[]
  buildToolFailureHint: (latestDoc: string, search: string) => string
  toolLogger: LegacyToolCallLogger
  yieldAfterToolCall?: () => Promise<void>
  limits?: {
    maxIterations?: number
    maxConsecutiveReplace?: number
    maxToolCallsPerIteration?: number
  }
}

export interface LegacyAgentLoopExecutionResult {
  iteration: number
  toolResults: ToolResult[]
  accumulatedReasoningForTurn: string
  finalContentForMemory: string
}

export async function runLegacyAgentLoop({
  conversationMessages,
  signal,
  turnId,
  callbacks,
  callModel,
  emitDebugEvent,
  streamPrefixRef,
  setStreamingContent,
  getStreamingReasoning,
  parseAssistantOutput,
  cleanModelOutput,
  hasToolCall,
  parseToolCallsToIR,
  buildNativeAssistantMessage,
  buildNativeToolResultMessages,
  buildToolFailureHint,
  toolLogger,
  yieldAfterToolCall,
  limits,
}: RunLegacyAgentLoopOptions): Promise<LegacyAgentLoopExecutionResult> {
  const maxIterations = limits?.maxIterations ?? 20
  const maxConsecutiveReplace = limits?.maxConsecutiveReplace ?? 10
  const maxToolCallsPerIteration = limits?.maxToolCallsPerIteration ?? 3

  let iteration = 0
  let accumulatedReasoningForTurn = ''
  let accumulatedContent = ''
  let latestSummary = ''
  let lastResponse = ''
  let finalContentForMemory = ''
  let shouldForceStop = false

  const allToolResults: ToolResult[] = []
  const normalizeToolText = (value: string) =>
    (value || '').replace(/\s+/g, ' ').trim()
  const buildEditKey = (searchText: string, replaceText: string) =>
    `${normalizeToolText(searchText)}=>${normalizeToolText(replaceText)}`
  const inferWordEditOperation = (call: { tool: string; args: Record<string, string> }) => {
    if (call.tool === 'review') return 'review'
    if (call.tool === 'replace') return 'replace'
    if (call.tool === 'insert') return 'insert'
    if (call.tool === 'delete') return 'delete'
    if (call.tool !== 'word.edit') return ''

    const explicitOperation = (call.args.operation || call.args.mode || call.args.action || '')
      .trim()
      .toLowerCase()
    if (
      explicitOperation === 'replace' ||
      explicitOperation === 'review' ||
      explicitOperation === 'insert' ||
      explicitOperation === 'delete'
    ) {
      return explicitOperation
    }
    if (call.args.reason || call.args.type) return 'review'
    if (call.args.position || call.args.content || call.args.dsl) return 'insert'
    if (call.args.target && !call.args.search && !call.args.replace) return 'delete'
    return 'replace'
  }
  const isReplaceLikeCall = (call: { tool: string; args: Record<string, string> }) =>
    inferWordEditOperation(call) === 'replace'
  const isReviewLikeCall = (call: { tool: string; args: Record<string, string> }) =>
    inferWordEditOperation(call) === 'review'
  const modifiedSearchTexts = new Set<string>()
  const modifiedReplaceTexts = new Set<string>()
  const successfulEditPairs = new Set<string>()
  let totalReplaceCount = 0
  let consecutiveReplaceCount = 0

  while (iteration < maxIterations && !shouldForceStop) {
    iteration += 1
    toolLogger.setIteration(iteration)

    streamPrefixRef.current = ''
    setStreamingContent('')

    sanitizeConversationMessagesForApi(conversationMessages)

    let rawResponse = ''
    const response = await callModel(conversationMessages, signal, {
      returnRaw: true,
      onToolCallStart: callbacks?.onToolCallStart,
      onToolCallPreview: callbacks?.onToolCallPreview,
      maxToolCallPreviews: maxToolCallsPerIteration,
      onResponseFinal: ({ raw }) => {
        rawResponse = raw
      },
    })

    const iterationReasoning = getStreamingReasoning().trim()
    if (iterationReasoning) {
      accumulatedReasoningForTurn = accumulatedReasoningForTurn
        ? `${accumulatedReasoningForTurn}\n\n${iterationReasoning}`
        : iterationReasoning
    }

    lastResponse = response.rawText || response.displayText
    const responseForDebug =
      rawResponse || response.rawText || response.displayText
    const responseForToolParsing = response.rawResponse ?? responseForDebug

    await emitDebugEvent({
      type: 'api_response_raw',
      turnId,
      timestamp: new Date().toISOString(),
      iteration,
      stage: 'loop',
      response: responseForDebug,
      rawResponse: response.rawResponse ?? responseForDebug,
      hasToolCall: hasToolCall(responseForToolParsing),
    })

    if (hasToolCall(responseForToolParsing)) {
      const parsedToolCalls = parseToolCallsToIR(responseForToolParsing, {
        turnId,
        metadata: {
          iteration,
          stage: 'loop',
        },
      }).map((call) => ({
        ir: call,
        tool: call.toolName,
        args: Object.entries(call.input || {}).reduce<Record<string, string>>(
          (acc, [key, value]) => {
            if (value === undefined || value === null) return acc
            if (typeof value === 'string') {
              acc[key] = value
              return acc
            }
            if (typeof value === 'number' || typeof value === 'boolean') {
              acc[key] = String(value)
              return acc
            }
            try {
              acc[key] = JSON.stringify(value)
            } catch {
              // ignore unserializable values
            }
            return acc
          },
          {},
        ),
        source: call.source,
      }))
      const { currentBatch: toolCalls, deferred: deferredToolCalls } =
        splitToolCallsForIteration(
          parsedToolCalls,
          maxToolCallsPerIteration,
        )

      toolLogger.logToolCallsParsed(
        parsedToolCalls.map((call) => ({
          tool: call.tool,
          args: { ...call.args },
        })),
      )

      await emitDebugEvent({
        type: 'tool_calls_parsed',
        turnId,
        timestamp: new Date().toISOString(),
        iteration,
        calls: parsedToolCalls.map((call) => ({
          tool: call.tool,
          args: { ...call.args },
          source: call.source,
        })),
      })

      if (deferredToolCalls.length > 0) {
        const deferReason = `deferred by controlled batching policy: max ${maxToolCallsPerIteration} tool call(s) per iteration`
        for (const deferredCall of deferredToolCalls) {
          await emitDebugEvent({
            type: 'tool_call_skipped',
            turnId,
            timestamp: new Date().toISOString(),
            iteration,
            tool: deferredCall.tool,
            args: { ...deferredCall.args },
            reason: deferReason,
          })
          callbacks?.onToolCallSkipped?.(
            deferredCall.tool,
            { ...deferredCall.args },
            deferReason,
          )
          nativeToolResultBindings.push({
            call: deferredCall.ir,
            result: {
              tool: deferredCall.tool,
              success: false,
              message: `skipped - ${deferReason}`,
            },
          })
        }
      }

      const parsedOutput = parseAssistantOutput(
        cleanModelOutput(responseForDebug),
      )
      if (parsedOutput.summary) {
        latestSummary = parsedOutput.summary
      }

      const thisRoundText = parsedOutput.displayText || ''
      if (thisRoundText) {
        callbacks?.onTextChunk?.(thisRoundText)
        accumulatedContent = thisRoundText
        streamPrefixRef.current = ''
        setStreamingContent('')
      }

      const nativeAssistantMessage =
        buildNativeAssistantMessage?.(responseForToolParsing) || null

      conversationMessages.push(
        nativeAssistantMessage || {
          role: 'assistant',
          content:
            responseForDebug ||
            parsedOutput.displayText ||
            '[assistant emitted native tool call]',
        },
      )

      const results: string[] = []
      const nativeToolResultBindings: Array<{ call: ToolCallIR; result: ToolResult }> = []
      let allSuccessful = true
      let hasReplaceInThisBatch = false
      let skippedCount = 0

      for (let ti = 0; ti < toolCalls.length; ti += 1) {
        const call = toolCalls[ti]
        const isEditTool = isReplaceLikeCall(call) || isReviewLikeCall(call)
        if (isReplaceLikeCall(call)) {
          hasReplaceInThisBatch = true
        }

        if (isEditTool) {
          const rawSearchText = call.args.search || ''
          const rawReplaceText = call.args.replace || ''
          const searchText = normalizeToolText(rawSearchText)
          const pairKey = buildEditKey(rawSearchText, rawReplaceText)

          if (successfulEditPairs.has(pairKey)) {
            toolLogger.logToolCallSkipped(
              call.tool,
              'duplicate search/replace pair',
              { ...call.args },
            )
            results.push(
              `[TOOL_RESULT]\ntool: ${call.tool}\nstatus: skipped - duplicate search/replace pair in this turn\n[/TOOL_RESULT]`,
            )
            skippedCount += 1
            await emitDebugEvent({
              type: 'tool_call_skipped',
              turnId,
              timestamp: new Date().toISOString(),
              iteration,
              tool: call.tool,
              args: { ...call.args },
              reason: 'duplicate search/replace pair in this turn',
            })
            callbacks?.onToolCallSkipped?.(
              call.tool,
              { ...call.args },
              'duplicate search/replace pair in this turn',
            )
            nativeToolResultBindings.push({
              call: call.ir,
              result: {
                tool: call.tool,
                success: false,
                message: 'skipped - duplicate search/replace pair in this turn',
              },
            })
            continue
          }

          if (searchText && modifiedReplaceTexts.has(searchText)) {
            toolLogger.logToolCallSkipped(
              call.tool,
              'search hits text already produced',
              { ...call.args },
            )
            results.push(
              `[TOOL_RESULT]\ntool: ${call.tool}\nstatus: skipped - search text has already been produced by an earlier edit\n[/TOOL_RESULT]`,
            )
            skippedCount += 1
            await emitDebugEvent({
              type: 'tool_call_skipped',
              turnId,
              timestamp: new Date().toISOString(),
              iteration,
              tool: call.tool,
              args: { ...call.args },
              reason:
                'search text has already been produced by an earlier edit in this turn',
            })
            callbacks?.onToolCallSkipped?.(
              call.tool,
              { ...call.args },
              'search text has already been produced by an earlier edit in this turn',
            )
            nativeToolResultBindings.push({
              call: call.ir,
              result: {
                tool: call.tool,
                success: false,
                message:
                  'skipped - search text has already been produced by an earlier edit in this turn',
              },
            })
            continue
          }

          if (searchText && modifiedSearchTexts.has(searchText)) {
            toolLogger.logToolCallSkipped(
              call.tool,
              'original search text already processed',
              { ...call.args },
            )
            results.push(
              `[TOOL_RESULT]\ntool: ${call.tool}\nstatus: skipped - original search text has already been processed in this turn\n[/TOOL_RESULT]`,
            )
            skippedCount += 1
            await emitDebugEvent({
              type: 'tool_call_skipped',
              turnId,
              timestamp: new Date().toISOString(),
              iteration,
              tool: call.tool,
              args: { ...call.args },
              reason:
                'original search text has already been processed in this turn',
            })
            callbacks?.onToolCallSkipped?.(
              call.tool,
              { ...call.args },
              'original search text has already been processed in this turn',
            )
            nativeToolResultBindings.push({
              call: call.ir,
              result: {
                tool: call.tool,
                success: false,
                message:
                  'skipped - original search text has already been processed in this turn',
              },
            })
            continue
          }
        }

        if (callbacks?.onToolCall) {
          toolLogger.logToolExecStart(call.tool, { ...call.args })
          const execStartTime = Date.now()
          const result = await callbacks.onToolCall(call.tool, call.args)
          allToolResults.push(result)
          if (!result.success) allSuccessful = false
          nativeToolResultBindings.push({
            call: call.ir,
            result,
          })

          if (yieldAfterToolCall) {
            await yieldAfterToolCall()
          }

          toolLogger.logToolExecResult(
            call.tool,
            result,
            Date.now() - execStartTime,
          )

          await emitDebugEvent({
            type: 'tool_result',
            turnId,
            timestamp: new Date().toISOString(),
            iteration,
            index: ti + 1,
            total: toolCalls.length,
            tool: call.tool,
            args: { ...call.args },
            result,
          })

          if (isEditTool && result.success) {
            const rawSearchText = call.args.search || ''
            const rawReplaceText = call.args.replace || ''
            const searchText = normalizeToolText(rawSearchText)
            const replaceText = normalizeToolText(rawReplaceText)

            if (searchText) modifiedSearchTexts.add(searchText)
            if (replaceText) modifiedReplaceTexts.add(replaceText)
            successfulEditPairs.add(buildEditKey(rawSearchText, rawReplaceText))

            if (isReplaceLikeCall(call)) {
              totalReplaceCount += 1
            }
          }

          let failureHint = ''
          if (!result.success && callbacks.getLatestDocument && isEditTool) {
            const latestDoc = callbacks.getLatestDocument()
            failureHint = buildToolFailureHint(
              latestDoc || '',
              call.args.search || '',
            )
          }

          const statusText = result.success
            ? 'success'
            : `failed: ${result.message}${failureHint ? `\n${failureHint}` : ''}`

          const progressInfo =
            isReplaceLikeCall(call) && result.success
              ? `\ncompleted replacements: ${totalReplaceCount}`
              : ''

          results.push(
            `[TOOL_RESULT]\ntool: ${call.tool}\nstatus: ${statusText}${progressInfo}\n[/TOOL_RESULT]`,
          )
        }
      }

      if (hasReplaceInThisBatch) {
        consecutiveReplaceCount += 1
        if (consecutiveReplaceCount >= maxConsecutiveReplace) {
          shouldForceStop = true
          results.push(
            `\n[SYSTEM] Reached consecutive replace safety limit (${maxConsecutiveReplace}). Stop tool calling and summarize completed work.`,
          )
        }
      } else {
        consecutiveReplaceCount = 0
      }

      if (skippedCount > 0 && skippedCount === toolCalls.length) {
        results.push(
          `\n[SYSTEM] All requested edits in this batch were skipped because they appear already completed. Summarize and stop.`,
        )
        shouldForceStop = true
      }

      if (deferredToolCalls.length > 0) {
        results.push(
          `\n[SYSTEM] This round executed ${toolCalls.length} tool call(s) under controlled batching. ${deferredToolCalls.length} deferred call(s) should be reconsidered in the next round based on latest results.`,
        )
      }

      let completionHint = ''
      if (allSuccessful && toolCalls.length > 0 && !shouldForceStop) {
        completionHint =
          '\n\n[SYSTEM] Tool calls finished for this round. If further edits are required, emit next [TOOL_CALL]. If done, provide a brief final summary.'
      }

      const toolResultContent = (
        results.join('\n\n') + completionHint
      ).trim()
      const nativeToolResultMessages =
        buildNativeToolResultMessages?.(nativeToolResultBindings) || []

      if (nativeToolResultMessages.length > 0) {
        conversationMessages.push(...nativeToolResultMessages)
        if (completionHint) {
          conversationMessages.push({
            role: 'user',
            content: completionHint.trim(),
          })
        }
      } else {
        conversationMessages.push({
          role: 'user',
          content: toolResultContent || '[Tool calls completed with no extra output]',
        })
      }

      toolLogger.logToolResultsSent(results, totalReplaceCount)

      if (shouldForceStop) {
        const forcedStopContent =
          parsedOutput.summary ||
          parsedOutput.displayText ||
          'Tool loop stopped by safety guard.'
        finalContentForMemory = forcedStopContent

        await emitDebugEvent({
          type: 'final_summary',
          turnId,
          timestamp: new Date().toISOString(),
          iteration,
          source: 'forced_stop',
          content: forcedStopContent,
        })

        await emitDebugEvent({
          type: 'turn_complete',
          turnId,
          timestamp: new Date().toISOString(),
          totalIterations: iteration,
          finalContent: forcedStopContent,
          toolResults: allToolResults.map((item) => ({ ...item })),
        })

        toolLogger.logTurnComplete({
          totalIterations: iteration,
          totalToolCalls: allToolResults.length,
          totalSkipped: skippedCount,
          stopReason: 'forced_stop',
          finalResponseLength: forcedStopContent.length,
        })

        callbacks?.onContent?.(forcedStopContent)
        callbacks?.onComplete?.(
          forcedStopContent,
          allToolResults,
          accumulatedReasoningForTurn.trim() || undefined,
        )
        break
      }

      continue
    }

    const hasOpenToolCall = hasIncompleteLegacyToolCall(responseForDebug)
    const hasOpenLegacyTag = (() => {
      if (!responseForDebug.includes('<')) return false
      const lowered = responseForDebug.toLowerCase()
      if (lowered.includes('<tool_use>') && !lowered.includes('</tool_use>')) {
        return true
      }
      for (const tool of LEGACY_TOOL_NAMES) {
        if (
          lowered.includes(`<${tool}>`) &&
          !lowered.includes(`</${tool}>`)
        ) {
          return true
        }
      }
      return false
    })()

    if (hasOpenToolCall || hasOpenLegacyTag) {
      conversationMessages.push({
        role: 'assistant',
        content: responseForDebug,
      })
      conversationMessages.push({
        role: 'user',
        content:
          '[SYSTEM] Your previous response was truncated mid-tool-call. Please continue from where you left off and complete the tool call block.',
      })
      continue
    }

    const parsedFinal = parseAssistantOutput(
      cleanModelOutput(responseForDebug),
    )
    if (parsedFinal.summary) {
      latestSummary = parsedFinal.summary
    }

    const finalText =
      parsedFinal.summary || parsedFinal.displayText || ''
    if (finalText) {
      callbacks?.onTextChunk?.(finalText)
    }
    finalContentForMemory = finalText

    await emitDebugEvent({
      type: 'final_summary',
      turnId,
      timestamp: new Date().toISOString(),
      iteration,
      source: 'normal',
      content: finalText,
    })

    await emitDebugEvent({
      type: 'turn_complete',
      turnId,
      timestamp: new Date().toISOString(),
      totalIterations: iteration,
      finalContent: finalText,
      toolResults: allToolResults.map((item) => ({ ...item })),
    })

    toolLogger.logTurnComplete({
      totalIterations: iteration,
      totalToolCalls: allToolResults.length,
      totalSkipped: 0,
      stopReason: 'normal',
      finalResponseLength: finalText.length,
    })

    callbacks?.onContent?.(finalText)
    callbacks?.onComplete?.(
      finalText,
      allToolResults,
      accumulatedReasoningForTurn.trim() || undefined,
    )
    break
  }

  if (iteration >= maxIterations && !shouldForceStop) {
    const maxIterSummary =
      latestSummary ||
      accumulatedContent ||
      lastResponse ||
      'Task completed (max iterations reached)'
    const maxIterSteps = streamPrefixRef.current.trim()
    const finalContent =
      maxIterSteps && maxIterSummary
        ? `${maxIterSteps}\n\n---\n\n${maxIterSummary}`
        : maxIterSummary || maxIterSteps
    finalContentForMemory = maxIterSummary

    await emitDebugEvent({
      type: 'final_summary',
      turnId,
      timestamp: new Date().toISOString(),
      iteration,
      source: 'max_iterations',
      content: finalContent,
    })

    await emitDebugEvent({
      type: 'turn_complete',
      turnId,
      timestamp: new Date().toISOString(),
      totalIterations: iteration,
      finalContent,
      toolResults: allToolResults.map((item) => ({ ...item })),
    })

    toolLogger.logTurnComplete({
      totalIterations: iteration,
      totalToolCalls: allToolResults.length,
      totalSkipped: 0,
      stopReason: 'max_iterations',
      finalResponseLength: finalContent.length,
    })

    callbacks?.onComplete?.(
      finalContent,
      allToolResults,
      accumulatedReasoningForTurn.trim() || undefined,
    )
  }

  return {
    iteration,
    toolResults: allToolResults,
    accumulatedReasoningForTurn,
    finalContentForMemory,
  }
}
