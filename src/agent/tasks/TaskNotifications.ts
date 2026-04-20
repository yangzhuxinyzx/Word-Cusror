export interface TaskNotification {
  taskId: string
  message: string
  createdAt: string
}

export class TaskNotificationCenter {
  private notifications: TaskNotification[] = []

  push(notification: TaskNotification): void {
    this.notifications = [...this.notifications, { ...notification }]
  }

  list(): TaskNotification[] {
    return [...this.notifications]
  }
}

