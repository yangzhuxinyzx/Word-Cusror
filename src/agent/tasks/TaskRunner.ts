import {
  TaskRegistry,
  type RuntimeTask,
  type RuntimeTaskStatus,
} from './TaskRegistry'
import { TaskNotificationCenter } from './TaskNotifications'

export class TaskRunner {
  constructor(
    private readonly registry = new TaskRegistry(),
    private readonly notifications = new TaskNotificationCenter(),
  ) {}

  async run(
    task: RuntimeTask,
    execute: () => Promise<void>,
  ): Promise<RuntimeTask> {
    const startedAt = new Date().toISOString()
    this.registry.upsert({
      ...task,
      status: 'running',
      startedAt,
      updatedAt: startedAt,
    })

    try {
      await execute()
      const completedAt = new Date().toISOString()
      const completed = {
        ...task,
        status: 'completed' as RuntimeTaskStatus,
        startedAt,
        completedAt,
        updatedAt: completedAt,
      }
      this.registry.upsert(completed)
      if ((task.mode || 'foreground') === 'background') {
        this.notifications.push({
          taskId: task.taskId,
          message: `${task.label} completed`,
          createdAt: completedAt,
        })
      }
      return completed
    } catch (error) {
      const failedAt = new Date().toISOString()
      const failed = {
        ...task,
        status: 'failed' as RuntimeTaskStatus,
        startedAt,
        completedAt: failedAt,
        updatedAt: failedAt,
        error: (error as Error).message || String(error),
      }
      this.registry.upsert(failed)
      if ((task.mode || 'foreground') === 'background') {
        this.notifications.push({
          taskId: task.taskId,
          message: `${task.label} failed: ${failed.error}`,
          createdAt: failedAt,
        })
      }
      throw error
    }
  }

  cancel(taskId: string): RuntimeTask | null {
    const cancelledAt = new Date().toISOString()
    const updated = this.registry.update(taskId, (current) => {
      if (!current) return null
      return {
        ...current,
        status: 'cancelled',
        completedAt: cancelledAt,
        updatedAt: cancelledAt,
      }
    })
    if (updated && (updated.mode || 'foreground') === 'background') {
      this.notifications.push({
        taskId: updated.taskId,
        message: `${updated.label} cancelled`,
        createdAt: cancelledAt,
      })
    }
    return updated
  }
}
