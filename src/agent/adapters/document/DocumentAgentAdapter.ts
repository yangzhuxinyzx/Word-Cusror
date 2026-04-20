import type { FileItem } from '../../../types'
import type { DocDsl, DslBlock, DslRun } from '../../../types/docDsl'

export interface DocumentReplaceResult {
  success: boolean
  count: number
  message: string
  diffId?: string
  searchText?: string
  replaceText?: string
  positions?: number[]
  debug?: Record<string, unknown>
}

export interface DocumentWordOpsResult {
  success: boolean
  message: string
  data?: Record<string, unknown>
}

export interface DocumentSelectionSnapshot {
  text: string
  html?: string
}

export interface DocumentPendingChangeSnapshot {
  id: string
  summary: string
  scope: string
  beforePreview?: string
  afterPreview?: string
  reviewReason?: string
  reviewType?: string
  stats?: { matches: number }
}

export interface DocumentAgentAdapter {
  getFileName(): string
  getCurrentFile(): FileItem | null
  getCurrentFilePath(): string | null
  getWorkspacePath(): string | null
  isElectron(): boolean
  ensureTiptapEditor(): Promise<void>
  autoSave(): Promise<{ success: boolean; error?: string } | null>
  resolveReferencePath(maybePathOrName: string): string
  readOutline(): string
  readSelection(): DocumentSelectionSnapshot
  readDocumentDsl(): DocDsl | null
  listPendingChanges(): DocumentPendingChangeSnapshot[]
  acceptChange(id: string): void
  rejectChange(id: string): void
  acceptAllChanges(): void
  rejectAllChanges(): void
  replaceViaDsl(
    search: string,
    replace: string,
    options?: { blockIndex?: number; format?: Partial<DslRun> },
  ): DocumentReplaceResult
  replaceInDocument(
    search: string,
    replace: string,
    reviewMeta?: { reason?: string; type?: string },
  ): DocumentReplaceResult
  replaceWithFormat(
    search: string,
    replace: string,
    format?: {
      bold?: boolean
      italic?: boolean
      underline?: boolean
      color?: string
      backgroundColor?: string
      fontSize?: string
      fontFamily?: string
    },
    reviewMeta?: { reason?: string; type?: string },
  ): DocumentReplaceResult
  insertViaDsl(
    position: string,
    blocks: DslBlock[],
  ): { success: boolean; message: string }
  insertInDocument(
    position: string,
    content: string,
  ): { success: boolean; message: string }
  deleteViaDsl(
    target: string,
    options?: { blockIndex?: number },
  ): { success: boolean; count: number; message: string }
  deleteInDocument(
    target: string,
  ): { success: boolean; count: number; message: string }
  previewOps(ops: any[]): DocumentWordOpsResult
  applyOps(ops: any[]): DocumentWordOpsResult
  prepareTemplateFillOutput(
    ops: any[],
  ): Promise<{ success: boolean; message?: string }>
  createDocument(
    title: string,
    content: string,
    elements?: Array<{
      type: 'heading' | 'paragraph' | 'table'
      content?: string
      level?: number
      bold?: boolean
      fontSize?: number
      fontFamily?: string
      alignment?: 'left' | 'center' | 'right' | 'justify'
      rows?: number
      cols?: number
      data?: string[][]
    }>,
    styleRefPath?: string,
  ): Promise<void> | void
  createDocumentFromDsl(
    title: string,
    dsl: DocDsl,
  ): Promise<{ success: boolean; message: string; filePath?: string }>
}

export interface CreateDocumentAgentAdapterOptions {
  fileName: string
  currentFile: FileItem | null
  attachedFiles: FileItem[]
  workspacePath: string | null
  editorMode: 'tiptap' | 'onlyoffice'
  setEditorMode: (mode: 'tiptap' | 'onlyoffice') => void
  isElectron: boolean
  currentFilePath: string | null
  silentSaveToFile: () => Promise<{ success: boolean; error?: string }>
  replaceViaDsl: (
    search: string,
    replace: string,
    options?: { blockIndex?: number; format?: Partial<DslRun> },
  ) => DocumentReplaceResult
  replaceWithFormat: (
    search: string,
    replace: string,
    format?: {
      bold?: boolean
      italic?: boolean
      underline?: boolean
      color?: string
      backgroundColor?: string
      fontSize?: string
      fontFamily?: string
    },
    reviewMeta?: { reason?: string; type?: string },
  ) => DocumentReplaceResult
  replaceInDocument: (
    search: string,
    replace: string,
    reviewMeta?: { reason?: string; type?: string },
  ) => DocumentReplaceResult
  insertViaDsl: (
    position: string,
    blocks: DslBlock[],
  ) => { success: boolean; message: string }
  insertInDocument: (
    position: string,
    content: string,
  ) => { success: boolean; message: string }
  deleteViaDsl: (
    target: string,
    options?: { blockIndex?: number },
  ) => { success: boolean; count: number; message: string }
  deleteInDocument: (
    target: string,
  ) => { success: boolean; count: number; message: string }
  previewWordOps: (ops: any[]) => DocumentWordOpsResult
  applyWordOps: (ops: any[]) => DocumentWordOpsResult
  prepareTemplateFillOutput: (
    ops: any[],
  ) => Promise<{ success: boolean; message?: string }>
  createNewDocument: (
    title: string,
    content: string,
    elements?: Array<{
      type: 'heading' | 'paragraph' | 'table'
      content?: string
      level?: number
      bold?: boolean
      fontSize?: number
      fontFamily?: string
      alignment?: 'left' | 'center' | 'right' | 'justify'
      rows?: number
      cols?: number
      data?: string[][]
    }>,
    styleRefPath?: string,
  ) => Promise<void> | void
  createDocumentFromDsl: (
    title: string,
    dsl: DocDsl,
  ) => Promise<{ success: boolean; message: string; filePath?: string }>
  getTiptapDocumentStructure: () => string
  getSelectionSnapshot?: () => DocumentSelectionSnapshot
  readDocumentDsl: () => DocDsl | null
  listPendingChanges: () => DocumentPendingChangeSnapshot[]
  acceptChange: (id: string) => void
  rejectChange: (id: string) => void
  acceptAllChanges: () => void
  rejectAllChanges: () => void
}

export function createDocumentAgentAdapter(
  options: CreateDocumentAgentAdapterOptions,
): DocumentAgentAdapter {
  return {
    getFileName: () => options.fileName,
    getCurrentFile: () => options.currentFile,
    getCurrentFilePath: () => options.currentFilePath,
    getWorkspacePath: () => options.workspacePath,
    isElectron: () => options.isElectron,
    async ensureTiptapEditor() {
      if (options.editorMode !== 'onlyoffice') return
      options.setEditorMode('tiptap')
      await new Promise((resolve) => setTimeout(resolve, 100))
    },
    async autoSave() {
      if (!options.isElectron || !options.currentFilePath) return null
      return options.silentSaveToFile().catch(() => null)
    },
    resolveReferencePath(maybePathOrName: string) {
      if (!maybePathOrName) return ''
      if (
        /[A-Za-z]:\\/.test(maybePathOrName) ||
        maybePathOrName.startsWith('\\\\')
      ) {
        return maybePathOrName
      }
      const byAttached = options.attachedFiles.find(
        (file) => file.name === maybePathOrName,
      )
      if (byAttached?.path) return byAttached.path
      if (options.currentFile?.name === maybePathOrName) {
        return options.currentFile.path
      }
      return ''
    },
    readOutline() {
      return options.getTiptapDocumentStructure()
    },
    readSelection() {
      return options.getSelectionSnapshot?.() || { text: '' }
    },
    readDocumentDsl: options.readDocumentDsl,
    listPendingChanges: options.listPendingChanges,
    acceptChange: options.acceptChange,
    rejectChange: options.rejectChange,
    acceptAllChanges: options.acceptAllChanges,
    rejectAllChanges: options.rejectAllChanges,
    replaceViaDsl: options.replaceViaDsl,
    replaceInDocument: options.replaceInDocument,
    replaceWithFormat: options.replaceWithFormat,
    insertViaDsl: options.insertViaDsl,
    insertInDocument: options.insertInDocument,
    deleteViaDsl: options.deleteViaDsl,
    deleteInDocument: options.deleteInDocument,
    previewOps: options.previewWordOps,
    applyOps: options.applyWordOps,
    prepareTemplateFillOutput: options.prepareTemplateFillOutput,
    createDocument: options.createNewDocument,
    createDocumentFromDsl: options.createDocumentFromDsl,
  }
}
