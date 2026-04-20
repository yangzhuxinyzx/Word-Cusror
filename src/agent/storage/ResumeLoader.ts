import type { ReplayCheckpoint } from './ReplayStore'

export class ResumeLoader {
  loadSessionCheckpoint(
    sessionId: string,
    getCheckpoint: (sessionId: string) => ReplayCheckpoint | null,
  ): ReplayCheckpoint | null {
    return getCheckpoint(sessionId)
  }
}

