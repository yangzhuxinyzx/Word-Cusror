export type MemorySearchResult = {
  path: string
  source: string
  workspaceKey?: string
  sessionId?: string
  startLine?: number
  endLine?: number
  score: number
  snippet: string
}

export type MemorySearchResponse = {
  success: boolean
  results: MemorySearchResult[]
  error?: string
}

export type MemoryStatusResponse = {
  success: boolean
  memoryDir?: string
  fileCount?: number
  chunkCount?: number
  lastIndexedAt?: string | null
  message?: string
  error?: string
}

export type MemoryStatusDetailResponse = {
  success: boolean
  memoryDir?: string
  lastIndexedAt?: string | null
  chunkSources?: Array<{ source: string; count: number }>
  fileSources?: Array<{ source: string; count: number }>
  message?: string
  error?: string
}
