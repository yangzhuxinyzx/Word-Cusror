import type { FileItem } from '../../../types'
import {
  buildWorkspaceProfilePayload,
  stringifyWorkspaceProfile,
} from '../../tools/packs/workspace/summarizers/WorkspaceProfileBuilder'

export interface WorkspaceIndexSnapshot {
  folderPath: string
  flatFiles: FileItem[]
  updatedAt: number
}

export interface WorkspaceDomainAdapterOptions {
  isElectron: boolean
  workspacePath: string | null
  currentFilePath: string | null
  workspaceContextMaxChars?: number
  truncateWithNote?: (text: string, maxLen: number, note: string) => string
  readFolder?: (folderPath: string) => Promise<{
    success: boolean
    data?: FileItem[]
  }>
  buildWorkspaceAutoSummaries?: (flatFiles: FileItem[]) => Promise<string>
}

function normalizePath(value: string): string {
  return value.replace(/\\/g, '/').toLowerCase()
}

function getParentDir(filePath: string): string {
  const normalized = String(filePath || '')
  const idx = Math.max(normalized.lastIndexOf('/'), normalized.lastIndexOf('\\'))
  return idx >= 0 ? normalized.slice(0, idx) : ''
}

function flattenFileTree(items: FileItem[]): FileItem[] {
  const out: FileItem[] = []
  const walk = (nodes: FileItem[]) => {
    for (const node of nodes) {
      if (node.type === 'file') {
        out.push(node)
      } else if (node.children?.length) {
        walk(node.children)
      }
    }
  }
  walk(items)
  return out
}

export class WorkspaceDomainAdapter {
  private indexCache: WorkspaceIndexSnapshot | null = null

  constructor(private readonly options: WorkspaceDomainAdapterOptions) {}

  getWorkspaceFolderPath(): string | null {
    if (this.options.currentFilePath) {
      return getParentDir(this.options.currentFilePath)
    }
    return this.options.workspacePath || null
  }

  async buildWorkspaceIndex(
    folderPath: string,
    refresh = false,
  ): Promise<WorkspaceIndexSnapshot | null> {
    if (!this.options.isElectron || !this.options.readFolder) return null
    const cached = this.indexCache
    if (
      !refresh &&
      cached &&
      normalizePath(cached.folderPath) === normalizePath(folderPath)
    ) {
      return cached
    }

    const result = await this.options.readFolder(folderPath)
    if (!result?.success || !result.data) return null

    const next = {
      folderPath,
      flatFiles: flattenFileTree(result.data),
      updatedAt: Date.now(),
    }
    this.indexCache = next
    return next
  }

  formatWorkspaceIndex(flatFiles: FileItem[], folderPath: string): string {
    const counts = new Map<string, number>()
    for (const file of flatFiles) {
      const ext = (file.extension || file.name.split('.').pop() || '').toLowerCase()
      const key = ext || 'unknown'
      counts.set(key, (counts.get(key) || 0) + 1)
    }

    const countLines = Array.from(counts.entries())
      .sort((left, right) => left[0].localeCompare(right[0]))
      .map(([ext, count]) => `${ext}: ${count}`)

    const displayFiles = flatFiles.slice(0, 200)
    const fileLines = displayFiles.map((file) => {
      const rel = file.relativePath
        ? file.relativePath
        : file.path && normalizePath(file.path).startsWith(normalizePath(folderPath))
          ? file.path.slice(folderPath.length).replace(/^[/\\]/, '')
          : file.name
      const ext = (file.extension || file.name.split('.').pop() || '').toLowerCase() || 'file'
      return `- [${ext}] ${rel}`
    })

    if (flatFiles.length > displayFiles.length) {
      fileLines.push(`... (${flatFiles.length - displayFiles.length} 个文件未显示，使用 workspace_list 查看更多)`)
    }

    const sections: string[] = [`【文件统计】${flatFiles.length} 个文件`]
    if (countLines.length) {
      sections.push(countLines.join(', '))
    }
    sections.push('')
    sections.push('【文件清单】')
    sections.push(fileLines.join('\n'))
    return sections.join('\n')
  }

  async resolveWorkspaceFile(args: {
    path?: string
    name?: string
    relativePath?: string
  }): Promise<FileItem | null> {
    const folderPath = this.getWorkspaceFolderPath()
    if (!folderPath) return null
    const index = await this.buildWorkspaceIndex(folderPath)
    if (!index) return null

    const targetPath = (args.path || args.relativePath || '').trim()
    const targetName = (args.name || '').trim()

    if (targetPath) {
      const normalizedTarget = normalizePath(targetPath)
      const matched = index.flatFiles.find((file) => {
        const filePath = file.path ? normalizePath(file.path) : ''
        const rel = file.relativePath ? normalizePath(file.relativePath) : ''
        return filePath === normalizedTarget || rel === normalizedTarget
      })
      if (matched) return matched

      if (!normalizedTarget.includes('/') && !normalizedTarget.includes('\\')) {
        const byName = index.flatFiles.find((file) => file.name === targetPath)
        if (byName) return byName
      }
    }

    if (targetName) {
      const byName = index.flatFiles.find((file) => file.name === targetName)
      if (byName) return byName
    }

    return null
  }

  async buildWorkspaceProfile(
    folderPath?: string,
    refresh = false,
  ): Promise<string> {
    const resolvedFolderPath = folderPath || this.getWorkspaceFolderPath()
    if (!resolvedFolderPath) return ''
    const index = await this.buildWorkspaceIndex(resolvedFolderPath, refresh)
    if (!index) return ''

    const summaryText = this.options.buildWorkspaceAutoSummaries
      ? await this.options.buildWorkspaceAutoSummaries(index.flatFiles)
      : ''

    return stringifyWorkspaceProfile(
      buildWorkspaceProfilePayload({
        folderPath: resolvedFolderPath,
        flatFiles: index.flatFiles,
        summaryText,
      }),
    )
  }

  async buildWorkspaceContext(): Promise<string> {
    const folderPath = this.getWorkspaceFolderPath()
    if (!folderPath) return ''
    const index = await this.buildWorkspaceIndex(folderPath)
    if (!index) return ''
    const indexText = this.formatWorkspaceIndex(index.flatFiles, folderPath)
    const summaryText = this.options.buildWorkspaceAutoSummaries
      ? await this.options.buildWorkspaceAutoSummaries(index.flatFiles)
      : ''

    const blocks = [`=== 工作夹目录（${folderPath}）===`, indexText]
    if (summaryText) {
      blocks.push('', summaryText)
    }
    const content = blocks.join('\n')
    if (!this.options.truncateWithNote || !this.options.workspaceContextMaxChars) {
      return content
    }
    return this.options.truncateWithNote(
      content,
      this.options.workspaceContextMaxChars,
      '工作夹上下文',
    )
  }
}
