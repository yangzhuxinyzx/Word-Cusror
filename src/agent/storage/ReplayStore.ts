export interface ReplayCheckpoint {
  sessionId: string
  lastMessageId: string | null
  lastToolCallId: string | null
  updatedAt: string
}

export class ReplayStore {
  private checkpoints = new Map<string, ReplayCheckpoint>()

  save(checkpoint: ReplayCheckpoint): void {
    this.checkpoints.set(checkpoint.sessionId, { ...checkpoint })
  }

  load(sessionId: string): ReplayCheckpoint | null {
    return this.checkpoints.get(sessionId) || null
  }
}

