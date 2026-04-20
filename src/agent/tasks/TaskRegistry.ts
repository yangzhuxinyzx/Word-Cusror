export type RuntimeTaskStatus =
  | 'queued'
  | 'running'
  | 'completed'
  | 'failed'
  | 'cancelled'

export type RuntimeTaskKind = 'task' | 'subagent'

export type RuntimeTaskMode = 'foreground' | 'background'

export interface RuntimeTask {
  taskId: string
  label: string
  status: RuntimeTaskStatus
  kind?: RuntimeTaskKind
  mode?: RuntimeTaskMode
  ownerId?: string
  parentTaskId?: string
  outputPath?: string
  transcriptPath?: string
  createdAt: string
  updatedAt?: string
  startedAt?: string
  completedAt?: string
  summary?: string
  error?: string
}

export class TaskRegistry {
  private tasks = new Map<string, RuntimeTask>()

  upsert(task: RuntimeTask): void {
    this.tasks.set(task.taskId, {
      ...task,
      kind: task.kind || 'task',
      mode: task.mode || 'foreground',
      updatedAt: task.updatedAt || new Date().toISOString(),
    })
  }

  get(taskId: string): RuntimeTask | null {
    return this.tasks.get(taskId) || null
  }

  update(
    taskId: string,
    patch:
      | Partial<RuntimeTask>
      | ((current: RuntimeTask | null) => RuntimeTask | null),
  ): RuntimeTask | null {
    const current = this.get(taskId)
    const next =
      typeof patch === 'function'
        ? patch(current)
        : current
          ? {
              ...current,
              ...patch,
            }
          : null
    if (!next) return null
    this.upsert({
      ...next,
      updatedAt: new Date().toISOString(),
    })
    return this.get(taskId)
  }

  list(): RuntimeTask[] {
    return Array.from(this.tasks.values())
  }

  listByKind(kind: RuntimeTaskKind): RuntimeTask[] {
    return this.list().filter((task) => (task.kind || 'task') === kind)
  }

  snapshot() {
    const tasks = this.list()
    return {
      count: tasks.length,
      running: tasks.filter((task) => task.status === 'running').length,
      background: tasks.filter((task) => (task.mode || 'foreground') === 'background')
        .length,
      subagents: tasks.filter((task) => (task.kind || 'task') === 'subagent').length,
    }
  }
}
