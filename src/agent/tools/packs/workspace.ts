import type { FileItem } from '../../../types'
import { defineAgentTool } from '../contracts'
import type { ExecutableAgentTool } from '../executor'

export interface WorkspaceIndexSnapshot {
  folderPath: string
  flatFiles: FileItem[]
  updatedAt: number
}

export interface WorkspaceSummaryOptions {
  maxChars?: number
  maxSlides?: number
  format?: string
}

export interface WorkspaceToolPackDeps {
  workspaceReadMaxChars: number
  workspaceSummaryMaxChars: number
  workspacePptxMaxSlides: number
  getWorkspaceFolderPath: () => string | null
  registerToolActivity: (tool: string, label: string) => string
  completeToolActivity: (
    activityId: string,
    status: 'success' | 'error' | 'skipped',
    detail?: string,
  ) => void
  updateAgentAction: (action: string) => void
  truncateLabel: (text: string, limit?: number) => string
  buildWorkspaceIndex: (
    folderPath: string,
    refresh?: boolean,
  ) => Promise<WorkspaceIndexSnapshot | null>
  formatWorkspaceIndex: (flatFiles: FileItem[], folderPath: string) => string
  resolveWorkspaceFile: (args: {
    path?: string
    name?: string
    relativePath?: string
  }) => Promise<FileItem | null>
  summarizeWorkspaceFile: (
    file: FileItem,
    options?: WorkspaceSummaryOptions,
  ) => Promise<string>
  buildWorkspaceProfile: (
    folderPath?: string,
    refresh?: boolean,
  ) => Promise<string>
  openFile: (file: FileItem) => Promise<void>
}

export function createWorkspaceToolPack(
  deps: WorkspaceToolPackDeps,
): ExecutableAgentTool[] {
  return [
    defineAgentTool({
      id: 'workspace_list',
      displayName: 'Workspace List',
      description: 'List files under the workspace or a specific folder',
      domain: 'workspace',
      mutation: 'read',
      concurrency: 'parallel_safe',
      tags: ['phase1', 'workspace', 'files'],
      inputKeys: ['folder', 'path', 'refresh'],
      async handler(args) {
        const folderArg = (args.folder || args.path || '').trim()
        const refresh = (args.refresh || '').toLowerCase() === 'true'
        const folderPath = folderArg || deps.getWorkspaceFolderPath()
        if (!folderPath) {
          return {
            tool: 'workspace_list',
            success: false,
            message: '无法确定工作夹目录，请先打开一个文件',
          }
        }

        const activityId = deps.registerToolActivity(
          'workspace_list',
          `索引：${deps.truncateLabel(folderPath, 24)}`,
        )
        deps.updateAgentAction('正在读取工作夹文件清单')
        const index = await deps.buildWorkspaceIndex(folderPath, refresh)
        if (!index) {
          deps.completeToolActivity(activityId, 'error', '读取失败')
          return {
            tool: 'workspace_list',
            success: false,
            message: '读取工作夹失败，请稍后重试',
          }
        }

        const indexText = deps.formatWorkspaceIndex(index.flatFiles, folderPath)
        deps.completeToolActivity(
          activityId,
          'success',
          `${index.flatFiles.length} 个文件`,
        )
        return {
          tool: 'workspace_list',
          success: true,
          message: `=== 工作夹目录（${folderPath}）===\n${indexText}`,
          data: { folderPath, total: index.flatFiles.length },
        }
      },
    }),
    defineAgentTool({
      id: 'workspace_profile',
      displayName: 'Workspace Profile',
      description: 'Build a structured workspace profile for the current folder',
      domain: 'workspace',
      mutation: 'read',
      concurrency: 'parallel_safe',
      tags: ['phase5', 'workspace', 'profile'],
      inputKeys: ['folder', 'path', 'refresh'],
      async handler(args) {
        const folderArg = (args.folder || args.path || '').trim()
        const refresh = (args.refresh || '').toLowerCase() === 'true'
        const folderPath = folderArg || deps.getWorkspaceFolderPath()
        if (!folderPath) {
          return {
            tool: 'workspace_profile',
            success: false,
            message: '未找到工作区目录，请先打开一个文件或文件夹。',
          }
        }

        const activityId = deps.registerToolActivity(
          'workspace_profile',
          `画像：${deps.truncateLabel(folderPath, 24)}`,
        )
        deps.updateAgentAction(`正在构建工作区画像：${deps.truncateLabel(folderPath, 24)}`)
        const profile = await deps.buildWorkspaceProfile(folderPath, refresh)
        if (!profile) {
          deps.completeToolActivity(activityId, 'error', '构建失败')
          return {
            tool: 'workspace_profile',
            success: false,
            message: '工作区画像构建失败。',
          }
        }
        deps.completeToolActivity(activityId, 'success')
        return {
          tool: 'workspace_profile',
          success: true,
          message: profile,
          data: { folderPath, profile },
        }
      },
    }),
    defineAgentTool({
      id: 'workspace_open',
      displayName: 'Workspace Open',
      description: 'Open a workspace file into the editor',
      domain: 'workspace',
      mutation: 'read',
      concurrency: 'serial',
      tags: ['phase1', 'workspace', 'open'],
      inputKeys: ['path', 'file', 'filePath', 'name', 'relativePath'],
      async handler(args) {
        const targetPath = (args.path || args.file || args.filePath || '').trim()
        const targetName = (args.name || '').trim()
        const targetRel = (args.relativePath || '').trim()
        const file = await deps.resolveWorkspaceFile({
          path: targetPath || targetRel,
          name: targetName,
          relativePath: targetRel,
        })
        if (!file) {
          return {
            tool: 'workspace_open',
            success: false,
            message: '未找到指定文件，请先使用 workspace_list 查看路径',
          }
        }

        const activityId = deps.registerToolActivity(
          'workspace_open',
          `打开：${deps.truncateLabel(file.name, 24)}`,
        )
        deps.updateAgentAction(`正在打开 ${file.name}`)
        await deps.openFile(file)
        deps.completeToolActivity(activityId, 'success')
        return {
          tool: 'workspace_open',
          success: true,
          message: `已打开文件：${file.name}`,
          data: { filePath: file.path },
        }
      },
    }),
    defineAgentTool({
      id: 'workspace_summarize',
      displayName: 'Workspace Summarize',
      description: 'Summarize a file from the workspace',
      domain: 'workspace',
      mutation: 'read',
      concurrency: 'parallel_safe',
      tags: ['phase1', 'workspace', 'summary'],
      inputKeys: ['path', 'file', 'filePath', 'name', 'relativePath', 'maxChars', 'maxSlides', 'format'],
      async handler(args) {
        return executeWorkspaceReadLikeTool('workspace_summarize', args, deps)
      },
    }),
    defineAgentTool({
      id: 'workspace_read',
      displayName: 'Workspace Read',
      description: 'Read a file from the workspace',
      domain: 'workspace',
      mutation: 'read',
      concurrency: 'parallel_safe',
      tags: ['phase1', 'workspace', 'read'],
      inputKeys: ['path', 'file', 'filePath', 'name', 'relativePath', 'maxChars', 'maxSlides', 'format'],
      async handler(args) {
        return executeWorkspaceReadLikeTool('workspace_read', args, deps)
      },
    }),
  ]
}

async function executeWorkspaceReadLikeTool(
  tool: 'workspace_read' | 'workspace_summarize',
  args: Record<string, string>,
  deps: WorkspaceToolPackDeps,
) {
  const targetPath = (args.path || args.file || args.filePath || '').trim()
  const targetName = (args.name || '').trim()
  const targetRel = (args.relativePath || '').trim()
  const format = (args.format || '').trim().toLowerCase()
  const file = await deps.resolveWorkspaceFile({
    path: targetPath || targetRel,
    name: targetName,
    relativePath: targetRel,
  })
  if (!file) {
    return {
      tool,
      success: false,
      message: '未找到指定文件，请先使用 workspace_list 查看路径',
    }
  }

  const maxCharsArg = args.maxChars ? parseInt(args.maxChars, 10) : undefined
  const maxSlidesArg = args.maxSlides ? parseInt(args.maxSlides, 10) : undefined
  const maxChars = Math.min(
    Math.max(
      maxCharsArg ||
        (tool === 'workspace_read'
          ? deps.workspaceReadMaxChars
          : deps.workspaceSummaryMaxChars),
      500,
    ),
    20000,
  )
  const maxSlides = Math.min(
    Math.max(maxSlidesArg || deps.workspacePptxMaxSlides, 1),
    20,
  )
  const activityId = deps.registerToolActivity(
    tool,
    `读取：${deps.truncateLabel(file.name, 24)}`,
  )
  deps.updateAgentAction(`正在读取 ${file.name}`)
  const content = await deps.summarizeWorkspaceFile(file, {
    maxChars,
    maxSlides,
    format: format || undefined,
  })
  deps.completeToolActivity(activityId, 'success')
  return {
    tool,
    success: true,
    message: content,
    data: {
      filePath: file.path,
      fileName: file.name,
      maxChars,
      format,
    },
  }
}
