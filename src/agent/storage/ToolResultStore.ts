export interface ToolResultFileReference {
  toolCallId: string
  path: string
  sizeBytes: number
  preview: string
  createdAt: string
}

export class ToolResultStore {
  private results = new Map<string, ToolResultFileReference>()

  set(reference: ToolResultFileReference): void {
    this.results.set(reference.toolCallId, { ...reference })
  }

  get(toolCallId: string): ToolResultFileReference | null {
    return this.results.get(toolCallId) || null
  }
}

