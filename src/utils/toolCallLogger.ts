/**
 * 工具调用日志记录器
 *
 * 记录模型的完整工具调用链路（请求 → 解析 → 执行 → 结果），
 * 以 JSON Lines 格式写入文件，方便后续分析模型的工具调用能力。
 *
 * 日志文件位置：工作目录/logs/tool-calls-{timestamp}.jsonl
 */

// ─── 类型定义 ───

export interface ToolCallLogEntry {
  /** 日志类型 */
  type:
    | 'session_start'       // 会话开始
    | 'request_context'     // 发给模型的上下文（文档内容摘要、消息数等）
    | 'api_request'         // API 请求（模型、消息数、token 估算）
    | 'api_response'        // API 响应（原始文本片段、停止原因）
    | 'tool_calls_parsed'   // 从响应中解析出的工具调用
    | 'tool_exec_start'     // 单个工具执行开始
    | 'tool_exec_result'    // 单个工具执行结果
    | 'tool_call_skipped'   // 工具调用被跳过（去重等）
    | 'tool_results_sent'   // 发回模型的工具结果汇总
    | 'turn_complete'       // 一轮对话结束
    | 'error'               // 错误
  /** ISO 时间戳 */
  timestamp: string
  /** 会话 ID */
  sessionId: string
  /** 当前迭代轮次 */
  iteration?: number
  /** 具体数据 */
  data: Record<string, unknown>
}

// ─── 日志记录器 ───

class ToolCallLogger {
  private sessionId = ''
  private logFilePath = ''
  private enabled = false
  private buffer: string[] = []
  private flushTimer: ReturnType<typeof setTimeout> | null = null
  private iteration = 0
  private workDir = ''

  /** 设置工作目录（在 startSession 之前调用） */
  setWorkDir(dir: string) {
    this.workDir = dir
  }

  /** 开始新会话 */
  startSession(workDir?: string) {
    this.sessionId = `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`
    this.iteration = 0
    this.enabled = !!window.electronAPI?.appendFile

    if (!this.enabled) return

    const ts = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19)
    const dir = workDir || this.workDir || '.'
    const sep = dir.includes('\\') ? '\\' : '/'
    this.logFilePath = `${dir}${dir.endsWith(sep) ? '' : sep}logs${sep}tool-calls-${ts}.jsonl`

    this.log({
      type: 'session_start',
      data: { workDir: dir, logFile: this.logFilePath }
    })

    console.log(`[ToolCallLogger] 日志文件: ${this.logFilePath}`)
  }

  /** 设置当前迭代轮次 */
  setIteration(n: number) {
    this.iteration = n
  }

  /** 记录一条日志 */
  log(entry: Omit<ToolCallLogEntry, 'timestamp' | 'sessionId' | 'iteration'> & { iteration?: number }) {
    if (!this.enabled) return

    const full: ToolCallLogEntry = {
      ...entry,
      timestamp: new Date().toISOString(),
      sessionId: this.sessionId,
      iteration: entry.iteration ?? this.iteration,
    }

    this.buffer.push(JSON.stringify(full))
    this.scheduleFlush()
  }

  /** 记录发给模型的上下文 */
  logRequestContext(info: {
    documentContentLength?: number
    documentContentTruncated?: boolean
    attachedFilesCount?: number
    messageCount: number
    totalCharsEstimate: number
  }) {
    this.log({ type: 'request_context', data: info })
  }

  /** 记录 API 请求 */
  logApiRequest(info: {
    model: string
    messageCount: number
    systemPromptLength: number
    temperature?: number
    maxTokens?: number
    nativeToolsCount?: number
    nativeToolProvider?: string
  }) {
    this.log({ type: 'api_request', data: info })
  }

  /** 记录解析出的工具调用 */
  logToolCallsParsed(calls: Array<{ tool: string; args: Record<string, unknown> }>) {
    this.log({
      type: 'tool_calls_parsed',
      data: {
        count: calls.length,
        calls: calls.map(c => ({
          tool: c.tool,
          args: truncateArgs(c.args),
        })),
      },
    })
  }

  /** 记录单个工具执行开始 */
  logToolExecStart(tool: string, args: Record<string, unknown>) {
    this.log({
      type: 'tool_exec_start',
      data: { tool, args: truncateArgs(args) },
    })
  }

  /** 记录单个工具执行结果 */
  logToolExecResult(tool: string, result: { success: boolean; message?: string }, durationMs: number) {
    this.log({
      type: 'tool_exec_result',
      data: { tool, success: result.success, message: result.message?.slice(0, 500), durationMs },
    })
  }

  /** 记录工具调用被跳过 */
  logToolCallSkipped(tool: string, reason: string, args?: Record<string, unknown>) {
    this.log({
      type: 'tool_call_skipped',
      data: { tool, reason, args: args ? truncateArgs(args) : undefined },
    })
  }

  /** 记录发回模型的工具结果 */
  logToolResultsSent(results: string[], totalReplaceCount: number) {
    this.log({
      type: 'tool_results_sent',
      data: {
        count: results.length,
        totalReplaceCount,
        resultsPreview: results.map(r => r.slice(0, 300)),
      },
    })
  }

  /** 记录一轮结束 */
  logTurnComplete(info: {
    totalIterations: number
    totalToolCalls: number
    totalSkipped: number
    stopReason?: string
    finalResponseLength?: number
  }) {
    this.log({ type: 'turn_complete', data: info })
    this.flush()
  }

  /** 记录错误 */
  logError(error: string, context?: Record<string, unknown>) {
    this.log({ type: 'error', data: { error, ...context } })
    this.flush()
  }

  /** 获取当前日志文件路径 */
  getLogFilePath() {
    return this.logFilePath
  }

  /** 获取会话 ID */
  getSessionId() {
    return this.sessionId
  }

  // ─── 内部方法 ───

  private scheduleFlush() {
    if (this.flushTimer) return
    this.flushTimer = setTimeout(() => {
      this.flush()
    }, 1000)
  }

  private flush() {
    if (this.flushTimer) {
      clearTimeout(this.flushTimer)
      this.flushTimer = null
    }
    if (!this.buffer.length || !this.logFilePath) return

    const content = this.buffer.join('\n') + '\n'
    this.buffer = []

    window.electronAPI?.appendFile(this.logFilePath, content).catch((err: unknown) => {
      console.warn('[ToolCallLogger] flush failed:', err)
    })
  }
}

/** 截断过长的参数值，避免日志膨胀 */
function truncateArgs(args: Record<string, unknown>): Record<string, unknown> {
  const result: Record<string, unknown> = {}
  for (const [k, v] of Object.entries(args)) {
    if (typeof v === 'string' && v.length > 500) {
      result[k] = v.slice(0, 500) + `...[${v.length} chars]`
    } else {
      result[k] = v
    }
  }
  return result
}

/** 全局单例 */
export const toolCallLogger = new ToolCallLogger()
