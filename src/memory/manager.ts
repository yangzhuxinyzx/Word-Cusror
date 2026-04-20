import type { MemorySearchOptions } from '../types/electron'
import type { MemorySearchResponse, MemoryStatusResponse, MemoryStatusDetailResponse } from './schema'

export const memorySearch = async (options: MemorySearchOptions): Promise<MemorySearchResponse> => {
  if (typeof window === 'undefined' || !window.electronAPI?.memorySearch) {
    return { success: false, results: [], error: 'memorySearch 不可用' }
  }
  return window.electronAPI.memorySearch(options)
}

export const memoryAppend = async (payload: { text: string; source?: string; tags?: string[] }) => {
  if (typeof window === 'undefined' || !window.electronAPI?.memoryAppend) {
    return { success: false, error: 'memoryAppend 不可用' }
  }
  return window.electronAPI.memoryAppend(payload)
}

export const memoryStatus = async (): Promise<MemoryStatusResponse> => {
  if (typeof window === 'undefined' || !window.electronAPI?.memoryStatus) {
    return { success: false, error: 'memoryStatus 不可用' }
  }
  return window.electronAPI.memoryStatus()
}

export const memoryStatusDetail = async (): Promise<MemoryStatusDetailResponse> => {
  if (typeof window === 'undefined' || !window.electronAPI?.memoryStatusDetail) {
    return { success: false, error: 'memoryStatusDetail 不可用' }
  }
  return window.electronAPI.memoryStatusDetail()
}

export const memoryClear = async (scope: 'all' | 'daily' | 'long' | 'sessions' = 'all') => {
  if (typeof window === 'undefined' || !window.electronAPI?.memoryClear) {
    return { success: false, error: 'memoryClear 不可用' }
  }
  return window.electronAPI.memoryClear({ scope })
}

export const memoryAppendSession = async (payload: { sessionId: string; text: string; meta?: Record<string, unknown> }) => {
  if (typeof window === 'undefined' || !window.electronAPI?.memoryAppendSession) {
    return { success: false, error: 'memoryAppendSession 不可用' }
  }
  return window.electronAPI.memoryAppendSession(payload)
}

export const memoryRebuildIndex = async () => {
  if (typeof window === 'undefined' || !window.electronAPI?.memoryRebuildIndex) {
    return { success: false, error: 'memoryRebuildIndex 不可用' }
  }
  return window.electronAPI.memoryRebuildIndex()
}
