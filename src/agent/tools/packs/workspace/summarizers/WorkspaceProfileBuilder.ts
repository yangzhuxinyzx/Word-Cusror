import type { FileItem } from '../../../../../types'

export interface WorkspaceProfilePayload {
  folderPath: string
  totalFiles: number
  fileTypes: Record<string, number>
  topFiles: Array<{
    name: string
    path: string
    relativePath?: string
    extension: string
  }>
  summary?: string
}

export function buildWorkspaceProfilePayload(params: {
  folderPath: string
  flatFiles: FileItem[]
  summaryText?: string
}): WorkspaceProfilePayload {
  const extensionCounts = new Map<string, number>()
  params.flatFiles.forEach((file) => {
    const ext = (file.extension || file.name.split('.').pop() || '').toLowerCase()
    const key = ext || '(none)'
    extensionCounts.set(key, (extensionCounts.get(key) || 0) + 1)
  })

  const topFiles = [...params.flatFiles]
    .sort((left, right) => {
      const leftExt = (left.extension || left.name.split('.').pop() || '').toLowerCase()
      const rightExt = (right.extension || right.name.split('.').pop() || '').toLowerCase()
      const priority = (ext: string) => {
        if (ext === 'md' || ext === 'txt') return 0
        if (ext === 'docx') return 1
        if (ext === 'pptx' || ext === 'ppt') return 2
        if (ext === 'xlsx' || ext === 'xls') return 3
        if (ext === 'pdf') return 4
        return 9
      }
      return priority(leftExt) - priority(rightExt) || left.name.localeCompare(right.name)
    })
    .slice(0, 12)
    .map((file) => ({
      name: file.name,
      path: file.path,
      relativePath: file.relativePath,
      extension: file.extension || file.name.split('.').pop() || '',
    }))

  return {
    folderPath: params.folderPath,
    totalFiles: params.flatFiles.length,
    fileTypes: Object.fromEntries(
      [...extensionCounts.entries()].sort((left, right) => right[1] - left[1]),
    ),
    topFiles,
    summary: params.summaryText || undefined,
  }
}

export function stringifyWorkspaceProfile(
  payload: WorkspaceProfilePayload,
): string {
  return JSON.stringify(payload, null, 2)
}
