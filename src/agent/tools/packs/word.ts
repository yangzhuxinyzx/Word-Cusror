import type { DocDsl } from '../../../types/docDsl'
import { validateDocDsl, dslToHtml } from '../../../utils/docDsl'
import { parseDocxToHtmlForAgent } from '../../../utils/docxParser'
import { docxHtmlToElements, elementsToHtmlPreview } from '../../../utils/docxHtmlToElements'
import { serializeDslForAI } from '../../../utils/dslSerializer'
import type { DocumentAgentAdapter } from '../../adapters/document/DocumentAgentAdapter'
import { defineAgentTool, type AgentToolExecutionContext } from '../contracts'
import type { ExecutableAgentTool } from '../executor'
import type { FileItem } from '../../../types'

type ReplaceResult = {
  success: boolean
  count: number
  message: string
  diffId?: string
  searchText?: string
  replaceText?: string
  positions?: number[]
  debug?: Record<string, unknown>
}

type WordOpsPreviewResult = {
  success: boolean
  message: string
  data?: Record<string, unknown>
}

type WordOpsApplyResult = {
  success: boolean
  message: string
  data?: Record<string, unknown>
}

export interface WordToolPackDeps {
  adapter: DocumentAgentAdapter
  fileName: string
  currentFile: FileItem | null
  attachedFiles: FileItem[]
  workspacePath: string | null
  editorMode: 'tiptap' | 'onlyoffice'
  setEditorMode: (mode: 'tiptap' | 'onlyoffice') => void
  flushUiFrame: () => Promise<void>
  isElectron: boolean
  currentFilePath: string | null
  silentSaveToFile: () => Promise<{ success: boolean; error?: string }>
  claimOrRegisterToolActivity: (
    tool: string,
    args: Record<string, string>,
    label: string,
    extra?: { searchText?: string; replaceText?: string },
  ) => string
  completeToolActivity: (
    activityId: string,
    status: 'success' | 'error' | 'skipped',
    detail?: string,
  ) => void
  patchToolActivity: (
    activityId: string,
    updates: Partial<{
      label: string
      detail?: string
      status: 'running' | 'success' | 'error' | 'skipped'
      searchText?: string
      replaceText?: string
      tool: string
    }>,
  ) => void
  updateAgentAction: (action: string) => void
  completeAgentStep: () => void
  updateAgentFile: (
    updates: Partial<{
      name: string
      additions: number
      deletions: number
      status: 'pending' | 'writing' | 'done'
    }>,
  ) => void
  addAgentFileOperation: (operation: string) => void
  buildReplaceSearchCandidates: (search: string) => string[]
  replaceViaDsl: (
    search: string,
    replace: string,
    options?: { blockIndex?: number; format?: Record<string, unknown> },
  ) => ReplaceResult
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
  ) => ReplaceResult
  replaceInDocument: (
    search: string,
    replace: string,
    reviewMeta?: { reason?: string; type?: string },
  ) => ReplaceResult
  insertViaDsl: (
    position: string,
    blocks: DocDsl['blocks'],
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
  previewWordOps: (ops: any[]) => WordOpsPreviewResult
  applyWordOps: (ops: any[]) => WordOpsApplyResult
  prepareTemplateFillOutput: (
    ops: any[],
  ) => Promise<{ success: boolean; message?: string }>
  setPendingWordOps: (value: {
    ops: any[]
    previewMessage: string
    previewLines: string[]
  }) => void
  refreshFiles: () => Promise<void>
  openFile: (file: FileItem) => Promise<void>
  resolveWorkspaceFile?: (args: {
    path?: string
    name?: string
    relativePath?: string
  }) => Promise<FileItem | null>
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
}

async function ensureTiptapEditor(deps: WordToolPackDeps) {
  await deps.adapter.ensureTiptapEditor()
}

async function autoSaveIfNeeded(deps: WordToolPackDeps) {
  return deps.adapter.autoSave()
}

function parseDslFormat(args: Record<string, string>) {
  const hasDslFormat =
    args.bold === 'true' ||
    args.italic === 'true' ||
    args.underline === 'true' ||
    !!args.color ||
    !!args.fontSize

  if (!hasDslFormat) return undefined

  return {
    bold: args.bold === 'true' || undefined,
    italic: args.italic === 'true' || undefined,
    underline: args.underline === 'true' || undefined,
    color: args.color || undefined,
    fontSize: args.fontSize ? parseFloat(args.fontSize) : undefined,
  }
}

function parseTextFormat(args: Record<string, string>) {
  return {
    bold: args.bold === 'true',
    italic: args.italic === 'true',
    underline: args.underline === 'true',
    color: args.color || undefined,
    backgroundColor: args.backgroundColor || undefined,
    fontSize: args.fontSize || undefined,
    fontFamily: args.fontFamily || undefined,
  }
}

function resolveRefPath(
  deps: WordToolPackDeps,
  maybePathOrName: string,
): string {
  return deps.adapter.resolveReferencePath(maybePathOrName)
}

function parseDocDslArg(raw: string): { dsl: DocDsl | null; error?: string } {
  try {
    const parsed = JSON.parse(raw)
    const validation = validateDocDsl(parsed)
    if (!validation.valid) {
      return {
        dsl: null,
        error: validation.errors.map((error) => `${error.path}: ${error.message}`).join('; '),
      }
    }
    return { dsl: parsed }
  } catch (error) {
    return { dsl: null, error: (error as Error).message }
  }
}

async function resolveWordReferencePath(
  deps: WordToolPackDeps,
  maybePathOrName: string,
): Promise<string> {
  const normalized = (maybePathOrName || '').trim()
  if (!normalized) return ''

  const direct = resolveRefPath(deps, normalized)
  if (direct) return direct

  if (!deps.resolveWorkspaceFile) return ''
  const matched = await deps.resolveWorkspaceFile({
    path: normalized,
    relativePath: normalized,
    name: normalized,
  })
  return matched?.path || ''
}

function getBlockText(block: DocDsl['blocks'][number]): string {
  switch (block.type) {
    case 'heading':
    case 'paragraph':
      if (typeof block.content === 'string') return block.content
      return block.content
        .map((item) => (typeof item === 'string' ? item : item.text || ''))
        .join('')
        .trim()
    case 'list':
      return block.items
        .map((item) =>
          item.content
            .map((run) => (typeof run === 'string' ? run : run.text || ''))
            .join(''),
        )
        .join(' | ')
        .trim()
    case 'table':
      return block.rows
        .flatMap((row) =>
          row.cells.map((cell) => {
            if (typeof cell.content === 'string') return cell.content
            if (Array.isArray(cell.content)) {
              return cell.content
                .map((item) => {
                  if (typeof item === 'string') return item
                  if ('type' in item) return getBlockText(item as DocDsl['blocks'][number])
                  return item.text || ''
                })
                .join('')
            }
            return ''
          }),
        )
        .join(' | ')
        .trim()
    case 'image':
      return block.caption || block.alt || '[image]'
    case 'blockquote':
      return block.content.map((item) => getBlockText(item)).join(' | ').trim()
    case 'pageBreak':
      return '[page break]'
    case 'sectionBreak':
      return `[section break${block.breakType ? `: ${block.breakType}` : ''}]`
    case 'horizontalRule':
      return '[horizontal rule]'
    default:
      return ''
  }
}

function buildDslBlockMap(dsl: DocDsl) {
  return dsl.blocks.map((block, index) => ({
    index,
    type: block.type,
    level: block.type === 'heading' ? block.level : undefined,
    preview: getBlockText(block).slice(0, 120),
  }))
}

function collectRunsFromBlock(block: DocDsl['blocks'][number]): Array<{
  fontFamily?: string
  fontSize?: string | number
}> {
  switch (block.type) {
    case 'heading':
    case 'paragraph':
      if (typeof block.content === 'string') return []
      return block.content
        .filter((item): item is Exclude<typeof item, string> => typeof item !== 'string')
        .map((run) => ({
          fontFamily: run.fontFamily,
          fontSize: run.fontSize,
        }))
    case 'list':
      return block.items.flatMap((item) =>
        item.content
          .filter((run): run is Exclude<typeof run, string> => typeof run !== 'string')
          .map((run) => ({
            fontFamily: run.fontFamily,
            fontSize: run.fontSize,
          })),
      )
    case 'blockquote':
      return block.content.flatMap((item) => collectRunsFromBlock(item))
    default:
      return []
  }
}

function incrementCount(map: Map<string, number>, key: string | null | undefined) {
  const normalized = (key || '').trim()
  if (!normalized) return
  map.set(normalized, (map.get(normalized) || 0) + 1)
}

function buildDslStyleProfile(dsl: DocDsl): string {
  const blockTypes = new Map<string, number>()
  const headingLevels = new Map<string, number>()
  const alignments = new Map<string, number>()
  const fontFamilies = new Map<string, number>()
  const fontSizes = new Map<string, number>()

  for (const block of dsl.blocks) {
    incrementCount(blockTypes, block.type)
    if (block.type === 'heading') {
      incrementCount(headingLevels, `h${block.level}`)
    }

    if ('format' in block && block.format?.alignment) {
      incrementCount(alignments, block.format.alignment)
    }

    for (const run of collectRunsFromBlock(block)) {
      incrementCount(fontFamilies, run.fontFamily)
      incrementCount(fontSizes, run.fontSize ? String(run.fontSize) : '')
    }
  }

  const formatMap = (map: Map<string, number>) =>
    Array.from(map.entries())
      .sort((left, right) => right[1] - left[1])
      .slice(0, 8)
      .map(([key, count]) => `${key}: ${count}`)
      .join(', ')

  return [
    `Blocks: ${dsl.blocks.length}`,
    blockTypes.size ? `Block types: ${formatMap(blockTypes)}` : '',
    headingLevels.size ? `Heading levels: ${formatMap(headingLevels)}` : '',
    alignments.size ? `Alignments: ${formatMap(alignments)}` : '',
    fontFamilies.size ? `Font families: ${formatMap(fontFamilies)}` : '',
    fontSizes.size ? `Font sizes: ${formatMap(fontSizes)}` : '',
  ]
    .filter(Boolean)
    .join('\n')
}

export function createWordToolPack(
  deps: WordToolPackDeps,
): ExecutableAgentTool[] {
  const tools: ExecutableAgentTool[] = []

  const invokeTool = async (
    toolId: string,
    args: Record<string, string>,
    context: AgentToolExecutionContext,
  ) => {
    const tool = tools.find((candidate) => candidate.id === toolId)
    if (!tool?.handler) {
      return {
        tool: toolId,
        success: false,
        message: `Word compatibility handler not found: ${toolId}`,
      }
    }
    return tool.handler(args, context)
  }

  const inferCreateMode = (args: Record<string, string>) => {
    const explicitMode = (args.mode || '').trim().toLowerCase()
    if (explicitMode === 'template') return 'template'
    if (explicitMode === 'document' || explicitMode === 'create') return 'document'
    if (args.templatePath || args.templateFileName || args.templateName || args.replacements) return 'template'
    return 'document'
  }

  const inferEditOperation = (args: Record<string, string>) => {
    const explicitOperation = (args.operation || args.mode || args.action || '')
      .trim()
      .toLowerCase()
    if (
      explicitOperation === 'replace' ||
      explicitOperation === 'review' ||
      explicitOperation === 'insert' ||
      explicitOperation === 'delete'
    ) {
      return explicitOperation
    }

    if (args.reason || args.type) return 'review'
    if (args.position || args.content || args.dsl) return 'insert'
    if (args.target && !args.search && !args.replace) return 'delete'
    return 'replace'
  }

  tools.push(
    defineAgentTool({
      id: 'word.read',
      displayName: 'Word Read',
      description: 'Read the current Word document selection or outline',
      prompt:
        'Use target=selection, outline, document, pending_changes, dsl, block_map, or style_profile to inspect structured Word state.',
      domain: 'word',
      mutation: 'read',
      concurrency: 'parallel_safe',
      tags: ['phase10', 'word', 'canonical', 'read'],
      inputKeys: ['target'],
      legacyAliases: ['word.read_selection', 'word.read_outline'],
      async handler(args, context) {
        const target = (args.target || 'selection').trim().toLowerCase()

        if (target === 'selection' || target === 'current_selection') {
          return invokeTool('word.read_selection', {}, context)
        }

        if (target === 'outline' || target === 'document_outline') {
          return invokeTool('word.read_outline', {}, context)
        }

        if (target === 'document' || target === 'current_document') {
          const outline = deps.adapter.readOutline().trim()
          const selection = deps.adapter.readSelection()
          const lines = [
            outline ? `Outline:\n${outline}` : '',
            selection.text.trim() ? `Selection:\n${selection.text}` : '',
          ].filter(Boolean)

          if (lines.length === 0) {
            return {
              tool: 'word.read',
              success: false,
              message: 'No readable document context is currently available.',
            }
          }

          return {
            tool: 'word.read',
            success: true,
            message: lines.join('\n\n'),
            data: {
              outline: outline || undefined,
              selection: selection.text || undefined,
              selectionHtml: selection.html || undefined,
            },
          }
        }

        if (target === 'pending_changes' || target === 'changes' || target === 'review_queue') {
          const changes = deps.adapter.listPendingChanges()
          if (changes.length === 0) {
            return {
              tool: 'word.read',
              success: true,
              message: 'No pending changes.',
              data: { count: 0, changes: [] },
            }
          }

          const lines = changes.map((change) => {
            const meta = [
              change.reviewType ? `type=${change.reviewType}` : '',
              change.reviewReason ? `reason=${change.reviewReason}` : '',
              change.scope ? `scope=${change.scope}` : '',
            ]
              .filter(Boolean)
              .join(', ')
            return `- ${change.id}: ${change.summary}${meta ? ` (${meta})` : ''}`
          })

          return {
            tool: 'word.read',
            success: true,
            message: `Pending changes (${changes.length}):\n${lines.join('\n')}`,
            data: {
              count: changes.length,
              changes,
            },
          }
        }

        if (target === 'dsl' || target === 'document_dsl') {
          const dsl = deps.adapter.readDocumentDsl()
          if (!dsl) {
            return {
              tool: 'word.read',
              success: false,
              message: 'Document DSL is not currently available.',
            }
          }

          const serialized = serializeDslForAI(dsl, { maxLength: 60000 })
          return {
            tool: 'word.read',
            success: true,
            message: serialized,
            data: {
              blockCount: dsl.blocks.length,
              dsl,
            },
          }
        }

        if (target === 'block_map' || target === 'blocks') {
          const dsl = deps.adapter.readDocumentDsl()
          if (!dsl) {
            return {
              tool: 'word.read',
              success: false,
              message: 'Document block map is not currently available.',
            }
          }

          const blocks = buildDslBlockMap(dsl)
          return {
            tool: 'word.read',
            success: true,
            message: blocks
              .map((block) =>
                `- [${block.index}] ${block.type}${block.level ? ` level=${block.level}` : ''}: ${block.preview}`,
              )
              .join('\n'),
            data: {
              count: blocks.length,
              blocks,
            },
          }
        }

        if (target === 'style_profile' || target === 'styles') {
          const dsl = deps.adapter.readDocumentDsl()
          if (!dsl) {
            return {
              tool: 'word.read',
              success: false,
              message: 'Document style profile is not currently available.',
            }
          }

          return {
            tool: 'word.read',
            success: true,
            message: buildDslStyleProfile(dsl),
            data: {
              blockCount: dsl.blocks.length,
            },
          }
        }

        return {
          tool: 'word.read',
          success: false,
          message:
            `Unsupported target: ${target}. Use selection, outline, document, pending_changes, dsl, block_map, or style_profile.`,
        }
      },
    }),
    defineAgentTool({
      id: 'word.create',
      displayName: 'Word Create',
      description: 'Create a Word document, optionally from a template, with DSL-first content',
      prompt:
        'Prefer DSL input for new content. Use mode=template when cloning or filling a template document.',
      domain: 'word',
      mutation: 'create',
      concurrency: 'serial',
      tags: ['phase10', 'word', 'canonical', 'create', 'dsl'],
      inputKeys: [
        'mode',
        'title',
        'newTitle',
        'content',
        'dsl',
        'elements',
        'styleRefPath',
        'styleRefFileName',
        'styleRefName',
        'contentRefPath',
        'contentRefFileName',
        'contentRefName',
        'templatePath',
        'templateFileName',
        'templateName',
        'replacements',
      ],
      legacyAliases: ['create', 'create_from_template', 'copy_template'],
      async handler(args, context) {
        if (inferCreateMode(args) === 'template') {
          return invokeTool(
            'create_from_template',
            {
              ...args,
              newTitle: args.newTitle || args.title || '',
            },
            context,
          )
        }

        return invokeTool('create', args, context)
      },
    }),
    defineAgentTool({
      id: 'word.edit',
      displayName: 'Word Edit',
      description: 'DSL-first Word editing for replace, review, insert, and delete operations',
      prompt:
        'Use operation=replace|review|insert|delete. Prefer DSL/blockIndex-scoped edits whenever the target structure is known.',
      domain: 'word',
      mutation: 'write',
      concurrency: 'serial',
      tags: ['phase10', 'word', 'canonical', 'edit', 'dsl'],
      inputKeys: [
        'operation',
        'strategy',
        'search',
        'replace',
        'position',
        'content',
        'dsl',
        'target',
        'blockIndex',
        'reason',
        'type',
        'bold',
        'italic',
        'underline',
        'color',
        'backgroundColor',
        'fontSize',
        'fontFamily',
      ],
      legacyAliases: [
        'replace',
        'review',
        'insert',
        'delete',
        'word.replace_via_dsl',
        'word.insert_via_dsl',
        'word.delete_via_dsl',
      ],
      async handler(args) {
        const operation = inferEditOperation(args)
        const strategy = (args.strategy || '').trim().toLowerCase()
        const requiresDsl = strategy === 'dsl' || strategy === 'structured'
        const parsedBlockIndex = args.blockIndex ? parseInt(args.blockIndex, 10) : NaN
        const blockIndex = Number.isFinite(parsedBlockIndex)
          ? parsedBlockIndex
          : undefined

        if (operation === 'insert') {
          const activityId = deps.claimOrRegisterToolActivity(
            'word.edit',
            args,
            `Insert: ${args.position || 'end'}`,
            { searchText: args.target || '', replaceText: args.content || '' },
          )
          await deps.flushUiFrame()
          await ensureTiptapEditor(deps)

          if (args.dsl) {
            const parsedDsl = parseDocDslArg(args.dsl)
            if (!parsedDsl.dsl) {
              deps.completeToolActivity(activityId, 'error', 'dsl invalid')
              return {
                tool: 'word.edit',
                success: false,
                message: parsedDsl.error || 'Invalid DSL payload.',
                data: { operation: 'insert', strategy: 'dsl' },
              }
            }
            const result = deps.adapter.insertViaDsl(
              args.position || 'end',
              parsedDsl.dsl.blocks || [],
            )
            if (result.success) {
              deps.completeToolActivity(
                activityId,
                'success',
                `${parsedDsl.dsl.blocks?.length || 0} blocks`,
              )
            } else {
              deps.completeToolActivity(activityId, 'error', result.message)
            }
            return {
              tool: 'word.edit',
              success: result.success,
              message: result.message,
              data: {
                operation: 'insert',
                strategy: 'dsl',
                position: args.position || 'end',
                blockCount: parsedDsl.dsl.blocks?.length || 0,
              },
            }
          }

          if (requiresDsl) {
            deps.completeToolActivity(activityId, 'error', 'dsl required')
            return {
              tool: 'word.edit',
              success: false,
              message: 'DSL strategy requires a dsl payload for insert operations.',
              data: { operation: 'insert', strategy: 'dsl' },
            }
          }

          const content = args.content || ''
          if (!content) {
            deps.completeToolActivity(activityId, 'error', 'missing content')
            return {
              tool: 'word.edit',
              success: false,
              message: 'Missing content or dsl parameter.',
              data: { operation: 'insert', strategy: 'text' },
            }
          }

          const result = deps.adapter.insertInDocument(args.position || 'end', content)
          if (result.success) {
            deps.completeToolActivity(activityId, 'success', 'inserted')
          } else {
            deps.completeToolActivity(activityId, 'error', result.message)
          }
          return {
            tool: 'word.edit',
            success: result.success,
            message: result.message,
            data: {
              operation: 'insert',
              strategy: 'text',
              position: args.position || 'end',
            },
          }
        }

        if (operation === 'delete') {
          const activityId = deps.claimOrRegisterToolActivity(
            'word.edit',
            args,
            `Delete: ${(args.target || args.blockIndex || '').toString().slice(0, 24)}`,
            { searchText: args.target || '' },
          )
          await deps.flushUiFrame()
          await ensureTiptapEditor(deps)

          if (blockIndex !== undefined || requiresDsl) {
            const dslResult = deps.adapter.deleteViaDsl(args.target || '', {
              blockIndex,
            })
            if (dslResult.success) {
              deps.completeToolActivity(activityId, 'success', `${dslResult.count} items`)
              return {
                tool: 'word.edit',
                success: true,
                message: dslResult.message,
                data: {
                  operation: 'delete',
                  strategy: 'dsl',
                  count: dslResult.count,
                  blockIndex,
                  target: args.target || '',
                },
              }
            }

            if (requiresDsl || !args.target) {
              deps.completeToolActivity(activityId, 'error', dslResult.message)
              return {
                tool: 'word.edit',
                success: false,
                message: dslResult.message,
                data: {
                  operation: 'delete',
                  strategy: 'dsl',
                  blockIndex,
                  target: args.target || '',
                },
              }
            }
          }

          if (!args.target) {
            deps.completeToolActivity(activityId, 'error', 'missing target')
            return {
              tool: 'word.edit',
              success: false,
              message: 'Missing target or blockIndex parameter.',
              data: { operation: 'delete', strategy: 'text' },
            }
          }

          const result = deps.adapter.deleteInDocument(args.target)
          if (result.success) {
            deps.completeToolActivity(activityId, 'success', `${result.count} items`)
          } else {
            deps.completeToolActivity(activityId, 'error', result.message)
          }
          return {
            tool: 'word.edit',
            success: result.success,
            message: result.message,
            data: {
              operation: 'delete',
              strategy: 'text',
              count: result.count,
              target: args.target,
            },
          }
        }

        if (operation === 'review') {
          const search = args.search || ''
          const replaceText = args.replace || ''
          if (!search) {
            return {
              tool: 'word.edit',
              success: false,
              message: 'Missing search parameter.',
              data: { operation: 'review' },
            }
          }

          const activityId = deps.claimOrRegisterToolActivity(
            'word.edit',
            args,
            `[review] ${(args.reason || search).slice(0, 24)}`,
            { searchText: search, replaceText },
          )
          await deps.flushUiFrame()
          await ensureTiptapEditor(deps)

          const format = parseTextFormat(args)
          const hasFormat =
            format.bold ||
            format.italic ||
            format.underline ||
            !!format.color ||
            !!format.backgroundColor ||
            !!format.fontSize ||
            !!format.fontFamily
          const reviewMeta = {
            reason: args.reason || '',
            type: args.type || 'style',
          }
          const result = hasFormat
            ? deps.adapter.replaceWithFormat(search, replaceText, format, reviewMeta)
            : deps.adapter.replaceInDocument(search, replaceText, reviewMeta)

          if (result.success && result.count > 0) {
            deps.completeToolActivity(activityId, 'success', `${result.count} edits`)
          } else {
            deps.completeToolActivity(activityId, 'error', result.message)
          }
          return {
            tool: 'word.edit',
            success: result.success,
            message: result.message,
            data: {
              operation: 'review',
              strategy: 'text',
              count: result.count,
              diffId: result.diffId,
              searchText: result.searchText || search,
              replaceText,
              reason: reviewMeta.reason,
              type: reviewMeta.type,
            },
          }
        }

        const search = args.search || ''
        const replaceText = args.replace || ''
        if (!search) {
          return {
            tool: 'word.edit',
            success: false,
            message: 'Missing search parameter.',
            data: { operation: 'replace' },
          }
        }

        const activityId = deps.claimOrRegisterToolActivity(
          'word.edit',
          args,
          `Replace: ${search.slice(0, 24)}`,
          { searchText: search, replaceText },
        )
        await deps.flushUiFrame()
        await ensureTiptapEditor(deps)

        const attemptedSearches = deps.buildReplaceSearchCandidates(search)
        const maxAttempts = Math.min(8, attemptedSearches.length)
        const dslFormat = parseDslFormat(args)
        let dslResult: ReplaceResult | null = null
        let usedSearch = search

        for (let i = 0; i < maxAttempts; i += 1) {
          const candidateSearch = attemptedSearches[i]
          const result = deps.adapter.replaceViaDsl(candidateSearch, replaceText, {
            blockIndex,
            format: dslFormat,
          })
          usedSearch = candidateSearch
          if (result.success && result.count > 0) {
            dslResult = result
            break
          }
        }

        if (dslResult?.success && dslResult.count > 0) {
          deps.completeToolActivity(activityId, 'success', `${dslResult.count} edits`)
          return {
            tool: 'word.edit',
            success: true,
            message: dslResult.message,
            data: {
              operation: 'replace',
              strategy: 'dsl',
              count: dslResult.count,
              diffId: dslResult.diffId,
              searchText: usedSearch,
              originalSearchText: search,
              replaceText,
              blockIndex,
            },
          }
        }

        if (requiresDsl) {
          deps.completeToolActivity(
            activityId,
            'error',
            dslResult?.message || 'dsl replace failed',
          )
          return {
            tool: 'word.edit',
            success: false,
            message: dslResult?.message || 'DSL replace failed.',
            data: {
              operation: 'replace',
              strategy: 'dsl',
              searchText: search,
              replaceText,
              blockIndex,
            },
          }
        }

        const format = parseTextFormat(args)
        const hasFormat =
          format.bold ||
          format.italic ||
          format.underline ||
          !!format.color ||
          !!format.backgroundColor ||
          !!format.fontSize ||
          !!format.fontFamily
        let textResult: ReplaceResult | null = null
        usedSearch = search
        for (let i = 0; i < maxAttempts; i += 1) {
          const candidateSearch = attemptedSearches[i]
          textResult = hasFormat
            ? deps.adapter.replaceWithFormat(candidateSearch, replaceText, format)
            : deps.adapter.replaceInDocument(candidateSearch, replaceText)
          usedSearch = candidateSearch
          if (textResult.success && textResult.count > 0) break
        }

        if (textResult?.success && textResult.count > 0) {
          deps.completeToolActivity(activityId, 'success', `${textResult.count} edits`)
          return {
            tool: 'word.edit',
            success: true,
            message: textResult.message,
            data: {
              operation: 'replace',
              strategy: 'text',
              count: textResult.count,
              diffId: textResult.diffId,
              searchText: usedSearch,
              originalSearchText: search,
              replaceText,
              blockIndex,
            },
          }
        }

        deps.completeToolActivity(activityId, 'error', textResult?.message || 'replace failed')
        return {
          tool: 'word.edit',
          success: false,
          message: textResult?.message || dslResult?.message || 'Replace failed.',
          data: {
            operation: 'replace',
            strategy: dslResult ? 'text_fallback' : 'text',
            searchText: usedSearch,
            originalSearchText: search,
            replaceText,
            blockIndex,
          },
        }
      },
    }),
    defineAgentTool({
      id: 'word.format',
      displayName: 'Word Format',
      description: 'Preview or apply structured Word ops for formatting and layout changes',
      prompt:
        'Provide ops as a JSON array. Use mode=preview for dry runs and mode=apply to execute the structured ops.',
      domain: 'word',
      mutation: 'transform',
      concurrency: 'serial',
      tags: ['phase10', 'word', 'canonical', 'ops'],
      inputKeys: ['mode', 'ops', 'dryRun'],
      legacyAliases: ['word_edit_ops', 'word.preview_ops', 'word.apply_ops'],
      async handler(args) {
        const mode = (args.mode || args.action || '').trim().toLowerCase()
        let ops: any[] = []
        try {
          ops = JSON.parse(args.ops || '[]')
        } catch {
          return {
            tool: 'word.format',
            success: false,
            message: 'ops parse failed: expected a JSON array',
          }
        }

        if (!Array.isArray(ops) || ops.length === 0) {
          return {
            tool: 'word.format',
            success: false,
            message: 'Missing ops: expected a non-empty JSON array.',
          }
        }

        if (mode === 'preview' || (args.dryRun || '').toLowerCase() === 'true') {
          const preview = deps.adapter.previewOps(ops)
          const lines = (preview.data?.lines as string[] | undefined) || []
          deps.setPendingWordOps({
            ops,
            previewMessage: preview.message,
            previewLines: lines,
          })
          return {
            tool: 'word.format',
            success: preview.success,
            message: preview.message,
            data: {
              mode: 'preview',
              ...(preview.data || {}),
            },
          }
        }

        const prep = await deps.adapter.prepareTemplateFillOutput(ops)
        if (!prep.success) {
          return {
            tool: 'word.format',
            success: false,
            message: prep.message || 'template fill preparation failed',
          }
        }

        const result = deps.adapter.applyOps(ops)
        return {
          tool: 'word.format',
          success: result.success,
          message: result.message,
          data: {
            mode: mode === 'execute' ? 'apply' : mode || 'apply',
            ...(result.data || {}),
          },
        }
      },
    }),
    defineAgentTool({
      id: 'word.resolve_change',
      displayName: 'Word Resolve Change',
      description: 'Accept or reject pending review changes in the active Word document',
      prompt:
        'Use action=list to inspect pending changes, or action=accept/action=reject to resolve them. Supply changeId for a single change or all=true for bulk resolution.',
      domain: 'word',
      mutation: 'write',
      concurrency: 'serial',
      tags: ['phase10', 'word', 'canonical', 'review'],
      inputKeys: ['action', 'changeId', 'id', 'all'],
      legacyAliases: ['word.accept_change', 'word.reject_change'],
      async handler(args) {
        const action = (args.action || '').trim().toLowerCase()

        if (action === 'list') {
          const changes = deps.adapter.listPendingChanges()
          return {
            tool: 'word.resolve_change',
            success: true,
            message:
              changes.length === 0
                ? 'No pending changes.'
                : changes.map((change) => `- ${change.id}: ${change.summary}`).join('\n'),
            data: {
              count: changes.length,
              changes,
            },
          }
        }

        if (action === 'accept') {
          const acceptAll = (args.all || '').toLowerCase() === 'true'
          if (acceptAll) {
            deps.adapter.acceptAllChanges()
            await autoSaveIfNeeded(deps)
            return {
              tool: 'word.resolve_change',
              success: true,
              message: 'Accepted all pending changes.',
              data: { action: 'accept', all: true },
            }
          }

          const changeId = args.changeId || args.id || ''
          if (!changeId) {
            return {
              tool: 'word.resolve_change',
              success: false,
              message: 'Missing changeId parameter.',
            }
          }

          const pending = deps.adapter.listPendingChanges().find((item) => item.id === changeId)
          if (!pending) {
            return {
              tool: 'word.resolve_change',
              success: false,
              message: `Pending change not found: ${changeId}`,
            }
          }

          deps.adapter.acceptChange(changeId)
          await autoSaveIfNeeded(deps)
          return {
            tool: 'word.resolve_change',
            success: true,
            message: `Accepted change: ${pending.summary}`,
            data: { action: 'accept', changeId, summary: pending.summary },
          }
        }

        if (action === 'reject') {
          const rejectAll = (args.all || '').toLowerCase() === 'true'
          if (rejectAll) {
            deps.adapter.rejectAllChanges()
            await autoSaveIfNeeded(deps)
            return {
              tool: 'word.resolve_change',
              success: true,
              message: 'Rejected all pending changes.',
              data: { action: 'reject', all: true },
            }
          }

          const changeId = args.changeId || args.id || ''
          if (!changeId) {
            return {
              tool: 'word.resolve_change',
              success: false,
              message: 'Missing changeId parameter.',
            }
          }

          const pending = deps.adapter.listPendingChanges().find((item) => item.id === changeId)
          if (!pending) {
            return {
              tool: 'word.resolve_change',
              success: false,
              message: `Pending change not found: ${changeId}`,
            }
          }

          deps.adapter.rejectChange(changeId)
          await autoSaveIfNeeded(deps)
          return {
            tool: 'word.resolve_change',
            success: true,
            message: `Rejected change: ${pending.summary}`,
            data: { action: 'reject', changeId, summary: pending.summary },
          }
        }

        return {
          tool: 'word.resolve_change',
          success: false,
          message: 'Missing action parameter. Use list, accept, or reject.',
        }
      },
    }),
    defineAgentTool({
      id: 'word.chart',
      displayName: 'Word Chart',
      description: 'Insert a chart into the active Word document',
      prompt:
        'Provide categories and series as structured arrays, then insert the resulting chart at the requested position.',
      domain: 'word',
      mutation: 'create',
      concurrency: 'serial',
      tags: ['phase10', 'word', 'canonical', 'chart'],
      inputKeys: ['type', 'title', 'categories', 'series', 'position', 'width', 'height'],
      legacyAliases: ['word_chart'],
      async handler(args, context) {
        return invokeTool('word_chart', args, context)
      },
    }),
    defineAgentTool({
      id: 'word.read_outline',
      displayName: 'Word Read Outline',
      description: 'Read the current document outline from the Word domain adapter',
      domain: 'word',
      mutation: 'read',
      concurrency: 'parallel_safe',
      tags: ['phase4', 'word', 'read'],
      async handler() {
        const outline = deps.adapter.readOutline().trim()
        if (!outline) {
          return {
            tool: 'word.read_outline',
            success: false,
            message: 'No outline is currently available.',
          }
        }
        return {
          tool: 'word.read_outline',
          success: true,
          message: outline,
          data: { outline },
        }
      },
    }),
    defineAgentTool({
      id: 'word.read_selection',
      displayName: 'Word Read Selection',
      description: 'Read the current text selection from the Word domain adapter',
      domain: 'word',
      mutation: 'read',
      concurrency: 'parallel_safe',
      tags: ['phase4', 'word', 'read'],
      async handler() {
        const selection = deps.adapter.readSelection()
        if (!selection.text.trim()) {
          return {
            tool: 'word.read_selection',
            success: false,
            message: 'No current selection is available.',
          }
        }
        return {
          tool: 'word.read_selection',
          success: true,
          message: selection.text,
          data: selection.html
            ? { text: selection.text, html: selection.html }
            : { text: selection.text },
        }
      },
    }),
    defineAgentTool({
      id: 'create',
      displayName: 'Create Document',
      description: 'Create a new Word document',
      domain: 'word',
      mutation: 'create',
      concurrency: 'serial',
      async handler(args) {
        const title = args.title || '新文档'
        const content = args.content || ''
        const activityId = deps.claimOrRegisterToolActivity(
          'create',
          args,
          `Create: ${title.slice(0, 24)}`,
        )

        const styleRefPathArg =
          args.styleRefPath || args.styleRefFileName || args.styleRefName || ''
        const contentRefPathArg =
          args.contentRefPath || args.contentRefFileName || args.contentRefName || ''

        let elements: Array<{
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
        }> = []

        const dslProvided = !!args.dsl
        let parsedDsl: DocDsl | null = null
        let dslError: string | null = null
        if (args.dsl) {
          const parsed = parseDocDslArg(args.dsl)
          if (parsed.dsl) {
            parsedDsl = parsed.dsl
          } else {
            dslError = `DSL parse failed: ${parsed.error || 'unknown error'}`
          }
        }

        if (args.elements) {
          try {
            elements = JSON.parse(args.elements)
          } catch {
            elements = []
          }
        }

        deps.updateAgentAction(`Creating ${title}.docx...`)
        deps.completeAgentStep()
        deps.updateAgentFile({ status: 'writing', name: `${title}.docx` })

        try {
          const styleRefPath = await resolveWordReferencePath(
            deps,
            String(styleRefPathArg || ''),
          )
          const contentRefPath = await resolveWordReferencePath(
            deps,
            String(contentRefPathArg || ''),
          )

          if (styleRefPathArg && !styleRefPath) {
            deps.completeToolActivity(activityId, 'error', 'missing style reference')
            return {
              tool: 'create',
              success: false,
              message: `Style reference file not found: ${styleRefPathArg}`,
            }
          }

          if (contentRefPathArg && !contentRefPath) {
            deps.completeToolActivity(activityId, 'error', 'missing content reference')
            return {
              tool: 'create',
              success: false,
              message: `Content reference file not found: ${contentRefPathArg}`,
            }
          }

          let previewHtml = content
          if (elements.length === 0 && contentRefPath) {
            if (!deps.isElectron || !window.electronAPI?.readFile) {
              return {
                tool: 'create',
                success: false,
                message: 'contentRefPath requires Electron file access.',
              }
            }

            const read = await window.electronAPI.readFile(contentRefPath)
            if (!read.success || !read.data) {
              return {
                tool: 'create',
                success: false,
                message: `Failed to read content reference: ${read.error || 'unknown error'}`,
              }
            }
            if (read.type !== 'docx') {
              return {
                tool: 'create',
                success: false,
                message: 'contentRefPath must point to a .docx file.',
              }
            }

            const parsed = await parseDocxToHtmlForAgent(read.data)
            const parsedElements = docxHtmlToElements(parsed.html || '')
            if (parsedElements.length > 0) {
              elements = parsedElements as typeof elements
              previewHtml = elementsToHtmlPreview(parsedElements)
            } else {
              previewHtml = parsed.html || content
            }
          }

          if (dslProvided && !parsedDsl) {
            deps.completeToolActivity(activityId, 'error', 'dsl invalid')
            return {
              tool: 'create',
              success: false,
              message: dslError || 'DSL validation failed.',
            }
          }

          if (parsedDsl) {
          const result = await deps.adapter.createDocumentFromDsl(title, parsedDsl)
            deps.completeToolActivity(
              activityId,
              result.success ? 'success' : 'error',
              `${parsedDsl.blocks?.length || 0} blocks`,
            )
            return {
              tool: 'create',
              success: result.success,
              message: result.message,
              data: {
                fileName: `${title}.docx`,
                blockCount: parsedDsl.blocks?.length,
                filePath: result.filePath,
              },
            }
          }

          await Promise.resolve(
            deps.adapter.createDocument(
              title,
              elements.length > 0 ? previewHtml || content : content,
              elements.length > 0 ? elements : undefined,
              styleRefPath || undefined,
            ),
          )

          deps.completeToolActivity(
            activityId,
            'success',
            elements.length > 0 ? `${elements.length} elements` : `${content.split('\n').length} lines`,
          )

          return {
            tool: 'create',
            success: true,
            message: elements.length > 0
              ? `Created ${title}.docx with ${elements.length} formatted element(s).`
              : `Created ${title}.docx.`,
            data: {
              fileName: `${title}.docx`,
              elementCount: elements.length || undefined,
              styleRefPath: styleRefPath || undefined,
              contentRefPath: contentRefPath || undefined,
            },
          }
        } catch (error) {
          deps.completeToolActivity(activityId, 'error', 'create failed')
          return {
            tool: 'create',
            success: false,
            message: `Create failed: ${error}`,
          }
        }
      },
    }),
    defineAgentTool({
      id: 'create_from_template',
      displayName: 'Create From Template',
      description: 'Create a new document from the current template',
      domain: 'word',
      mutation: 'create',
      concurrency: 'serial',
      legacyAliases: ['copy_template'],
      async handler(args) {
        const templateRefArg =
          args.templatePath || args.templateFileName || args.templateName || ''
        const newTitle = args.newTitle || '新文档'
        let replacements: Array<{ search: string; replace: string }> = []
        if (args.replacements) {
          try {
            replacements = JSON.parse(args.replacements)
          } catch {
            replacements = []
          }
        }

        const activityId = deps.claimOrRegisterToolActivity(
          'create_from_template',
          args,
          `Template: ${newTitle.slice(0, 24)}`,
        )

        if (!deps.currentFile && !templateRefArg) {
          deps.completeToolActivity(activityId, 'error', 'missing template')
          return {
            tool: 'create_from_template',
            success: false,
            message: 'No template was provided and no currently opened document is available as template.',
          }
        }

        if (!deps.isElectron || !window.electronAPI) {
          deps.completeToolActivity(activityId, 'error', 'desktop only')
          return {
            tool: 'create_from_template',
            success: false,
            message: 'Template-based creation requires the desktop app.',
          }
        }

        deps.updateAgentAction(`Creating ${newTitle}.docx from template...`)
        deps.completeAgentStep()

        try {
          const sourcePath = templateRefArg
            ? await resolveWordReferencePath(deps, templateRefArg)
            : deps.currentFile?.path || ''
          if (templateRefArg && !sourcePath) {
            deps.completeToolActivity(activityId, 'error', 'template not found')
            return {
              tool: 'create_from_template',
              success: false,
              message: `Template file not found: ${templateRefArg}`,
            }
          }
          if (!sourcePath) {
            deps.completeToolActivity(activityId, 'error', 'missing template')
            return {
              tool: 'create_from_template',
              success: false,
              message: 'No template source path is available.',
            }
          }
          const dir = sourcePath.substring(0, sourcePath.lastIndexOf('\\'))
          const newPath = `${dir}\\${newTitle}.docx`

          if (replacements.length > 0 && window.electronAPI.docxSearchReplace) {
            const result = await window.electronAPI.docxSearchReplace({
              sourcePath,
              outputPath: newPath,
              replacements,
            })
            if (!result.success) {
              deps.completeToolActivity(activityId, 'error', 'template replace failed')
              return {
                tool: 'create_from_template',
                success: false,
                message: result.error || 'Template replacement failed.',
              }
            }
          } else {
            const sourceContent = await window.electronAPI.readFile(sourcePath)
            if (!sourceContent.success || !sourceContent.data) {
              deps.completeToolActivity(activityId, 'error', 'template read failed')
              return {
                tool: 'create_from_template',
                success: false,
                message: 'Failed to read template document.',
              }
            }

            if (sourceContent.type === 'docx') {
              await window.electronAPI.writeBinaryFile(newPath, sourceContent.data)
            } else {
              await window.electronAPI.writeFile(newPath, sourceContent.data)
            }
          }

          await deps.refreshFiles()
          await deps.openFile({
            name: `${newTitle}.docx`,
            path: newPath,
            type: 'file',
          })

          deps.completeToolActivity(
            activityId,
            'success',
            replacements.length > 0 ? `${replacements.length} replacements` : 'copied',
          )
          return {
            tool: 'create_from_template',
            success: true,
            message: replacements.length > 0
              ? `Created ${newTitle}.docx from template and applied ${replacements.length} replacement(s).`
              : `Created ${newTitle}.docx from template.`,
            data: {
              fileName: `${newTitle}.docx`,
              filePath: newPath,
              totalReplacements: replacements.length,
            },
          }
        } catch (error) {
          deps.completeToolActivity(activityId, 'error', 'template create failed')
          return {
            tool: 'create_from_template',
            success: false,
            message: `Template creation failed: ${error}`,
          }
        }
      },
    }),
    defineAgentTool({
      id: 'word.replace_via_dsl',
      displayName: 'Word Replace Via DSL',
      description: 'Apply a DSL-scoped replace against the active Word document',
      domain: 'word',
      mutation: 'write',
      concurrency: 'serial',
      tags: ['phase4', 'word', 'dsl'],
      inputKeys: ['search', 'replace', 'blockIndex', 'bold', 'italic', 'underline', 'color', 'fontSize'],
      async handler(args) {
        const search = args.search || ''
        const replaceText = args.replace || ''
        const blockIndex = args.blockIndex ? parseInt(args.blockIndex, 10) : undefined
        if (!search) {
          return {
            tool: 'word.replace_via_dsl',
            success: false,
            message: 'Missing search parameter.',
          }
        }

        const result = deps.adapter.replaceViaDsl(search, replaceText, {
          blockIndex,
          format: parseDslFormat(args),
        })
        return {
          tool: 'word.replace_via_dsl',
          success: result.success,
          message: result.message,
          data: {
            count: result.count,
            diffId: result.diffId,
            searchText: result.searchText || search,
            replaceText,
            positions: result.positions,
            debug: result.debug,
          },
        }
      },
    }),
    defineAgentTool({
      id: 'word.insert_via_dsl',
      displayName: 'Word Insert Via DSL',
      description: 'Insert DSL blocks into the active Word document',
      domain: 'word',
      mutation: 'write',
      concurrency: 'serial',
      tags: ['phase4', 'word', 'dsl'],
      inputKeys: ['position', 'dsl'],
      async handler(args) {
        const position = args.position || 'end'
        if (!args.dsl) {
          return {
            tool: 'word.insert_via_dsl',
            success: false,
            message: 'Missing dsl parameter.',
          }
        }
        try {
          const parsedDsl = parseDocDslArg(args.dsl)
          if (!parsedDsl.dsl) {
            return {
              tool: 'word.insert_via_dsl',
              success: false,
              message: parsedDsl.error || 'Invalid DSL payload.',
            }
          }
          const result = deps.adapter.insertViaDsl(position, parsedDsl.dsl.blocks || [])
          return {
            tool: 'word.insert_via_dsl',
            success: result.success,
            message: result.message,
            data: { position, blockCount: parsedDsl.dsl.blocks?.length || 0 },
          }
        } catch (error) {
          return {
            tool: 'word.insert_via_dsl',
            success: false,
            message: `DSL insert failed: ${error}`,
          }
        }
      },
    }),
    defineAgentTool({
      id: 'word.delete_via_dsl',
      displayName: 'Word Delete Via DSL',
      description: 'Delete content via DSL block targeting',
      domain: 'word',
      mutation: 'delete',
      concurrency: 'serial',
      tags: ['phase4', 'word', 'dsl'],
      inputKeys: ['target', 'blockIndex'],
      async handler(args) {
        const target = args.target || ''
        const blockIndex = args.blockIndex ? parseInt(args.blockIndex, 10) : undefined
        const result = deps.adapter.deleteViaDsl(target, { blockIndex })
        return {
          tool: 'word.delete_via_dsl',
          success: result.success,
          message: result.message,
          data: { count: result.count, target, blockIndex },
        }
      },
    }),
    defineAgentTool({
      id: 'word.preview_ops',
      displayName: 'Word Preview Ops',
      description: 'Preview structured Word ops without applying them',
      domain: 'word',
      mutation: 'read',
      concurrency: 'serial',
      tags: ['phase4', 'word', 'ops'],
      inputKeys: ['ops'],
      async handler(args) {
        let ops: any[] = []
        try {
          ops = JSON.parse(args.ops || '[]')
        } catch {
          return {
            tool: 'word.preview_ops',
            success: false,
            message: 'ops parse failed: expected a JSON array',
          }
        }
        if (!Array.isArray(ops) || ops.length === 0) {
          return {
            tool: 'word.preview_ops',
            success: false,
            message: 'Missing ops: expected a non-empty JSON array.',
          }
        }
        const preview = deps.adapter.previewOps(ops)
        const lines = (preview.data?.lines as string[] | undefined) || []
        deps.setPendingWordOps({
          ops,
          previewMessage: preview.message,
          previewLines: lines,
        })
        return {
          tool: 'word.preview_ops',
          success: preview.success,
          message: preview.message,
          data: preview.data,
        }
      },
    }),
    defineAgentTool({
      id: 'word.apply_ops',
      displayName: 'Word Apply Ops',
      description: 'Apply structured Word ops through the document adapter',
      domain: 'word',
      mutation: 'transform',
      concurrency: 'serial',
      tags: ['phase4', 'word', 'ops'],
      inputKeys: ['ops'],
      async handler(args) {
        let ops: any[] = []
        try {
          ops = JSON.parse(args.ops || '[]')
        } catch {
          return {
            tool: 'word.apply_ops',
            success: false,
            message: 'ops parse failed: expected a JSON array',
          }
        }
        if (!Array.isArray(ops) || ops.length === 0) {
          return {
            tool: 'word.apply_ops',
            success: false,
            message: 'Missing ops: expected a non-empty JSON array.',
          }
        }
        const prep = await deps.adapter.prepareTemplateFillOutput(ops)
        if (!prep.success) {
          return {
            tool: 'word.apply_ops',
            success: false,
            message: prep.message || 'template fill preparation failed',
          }
        }
        const result = deps.adapter.applyOps(ops)
        return {
          tool: 'word.apply_ops',
          success: result.success,
          message: result.message,
          data: result.data,
        }
      },
    }),
    defineAgentTool({
      id: 'word.accept_change',
      displayName: 'Word Accept Change',
      description: 'Accept a pending review/diff change in the active document',
      domain: 'word',
      mutation: 'write',
      concurrency: 'serial',
      tags: ['phase4', 'word', 'review'],
      inputKeys: ['changeId', 'id', 'all'],
      async handler(args) {
        const acceptAll = (args.all || '').toLowerCase() === 'true'
        if (acceptAll) {
          deps.adapter.acceptAllChanges()
          await autoSaveIfNeeded(deps)
          return {
            tool: 'word.accept_change',
            success: true,
            message: 'Accepted all pending changes.',
          }
        }

        const changeId = args.changeId || args.id || ''
        if (!changeId) {
          return {
            tool: 'word.accept_change',
            success: false,
            message: 'Missing changeId parameter.',
          }
        }

        const pending = deps.adapter.listPendingChanges().find((item) => item.id === changeId)
        if (!pending) {
          return {
            tool: 'word.accept_change',
            success: false,
            message: `Pending change not found: ${changeId}`,
          }
        }

        deps.adapter.acceptChange(changeId)
        await autoSaveIfNeeded(deps)
        return {
          tool: 'word.accept_change',
          success: true,
          message: `Accepted change: ${pending.summary}`,
          data: { changeId, summary: pending.summary },
        }
      },
    }),
    defineAgentTool({
      id: 'word.reject_change',
      displayName: 'Word Reject Change',
      description: 'Reject a pending review/diff change in the active document',
      domain: 'word',
      mutation: 'write',
      concurrency: 'serial',
      tags: ['phase4', 'word', 'review'],
      inputKeys: ['changeId', 'id', 'all'],
      async handler(args) {
        const rejectAll = (args.all || '').toLowerCase() === 'true'
        if (rejectAll) {
          deps.adapter.rejectAllChanges()
          await autoSaveIfNeeded(deps)
          return {
            tool: 'word.reject_change',
            success: true,
            message: 'Rejected all pending changes.',
          }
        }

        const changeId = args.changeId || args.id || ''
        if (!changeId) {
          return {
            tool: 'word.reject_change',
            success: false,
            message: 'Missing changeId parameter.',
          }
        }

        const pending = deps.adapter.listPendingChanges().find((item) => item.id === changeId)
        if (!pending) {
          return {
            tool: 'word.reject_change',
            success: false,
            message: `Pending change not found: ${changeId}`,
          }
        }

        deps.adapter.rejectChange(changeId)
        await autoSaveIfNeeded(deps)
        return {
          tool: 'word.reject_change',
          success: true,
          message: `Rejected change: ${pending.summary}`,
          data: { changeId, summary: pending.summary },
        }
      },
    }),
    defineAgentTool({
      id: 'replace',
      displayName: 'Replace',
      description: 'Replace text inside the active Word document',
      domain: 'word',
      mutation: 'write',
      concurrency: 'serial',
      tags: ['phase1', 'word', 'edit'],
      inputKeys: ['search', 'replace', 'blockIndex', 'bold', 'italic', 'underline', 'color', 'backgroundColor', 'fontSize'],
      async handler(args) {
        const search = args.search || ''
        const replaceText = args.replace || ''
        const blockIndex = args.blockIndex ? parseInt(args.blockIndex, 10) : undefined

        if (!search) {
          return { tool: 'replace', success: false, message: '缺少 search 参数' }
        }

        const activityId = deps.claimOrRegisterToolActivity(
          'replace',
          args,
          `Replace: ${search.slice(0, 24)}`,
          { searchText: search, replaceText },
        )
        await deps.flushUiFrame()
        await ensureTiptapEditor(deps)

        if (blockIndex !== undefined && !isNaN(blockIndex)) {
          const dslResult = deps.adapter.replaceViaDsl(search, replaceText, {
            blockIndex,
            format: parseDslFormat(args),
          })

          if (dslResult.success) {
            deps.updateAgentAction(`正在替换 [块${blockIndex}] ${search.slice(0, 20)}...`)
            deps.completeAgentStep()
            deps.updateAgentFile({
              status: 'writing',
              additions: dslResult.count,
              name: deps.fileName,
            })
            deps.completeToolActivity(activityId, 'success', `${dslResult.count} 处`)
            return {
              tool: 'replace',
              success: true,
              message: dslResult.message,
              data: {
                count: dslResult.count,
                blockIndex,
                diffId: dslResult.diffId,
                searchText: dslResult.searchText || search,
                replaceText,
              },
            }
          }
        }

        const format = parseTextFormat(args)
        const hasFormat =
          format.bold ||
          format.italic ||
          format.underline ||
          !!format.color ||
          !!format.backgroundColor ||
          !!format.fontSize

        deps.updateAgentAction(
          `正在替换 ${search.slice(0, 20)}${search.length > 20 ? '...' : ''}${hasFormat ? ' (带格式)' : ''}`,
        )
        deps.completeAgentStep()
        deps.updateAgentFile({ status: 'writing', name: deps.fileName })
        deps.addAgentFileOperation(
          `替换: "${search.slice(0, 15)}..." -> "${replaceText.slice(0, 15)}..."`,
        )

        const attemptedSearches = deps.buildReplaceSearchCandidates(search)
        const maxAttempts = Math.min(8, attemptedSearches.length)

        let result: ReplaceResult | null = null
        let usedSearch = search
        for (let i = 0; i < maxAttempts; i += 1) {
          const candidateSearch = attemptedSearches[i]
          result = hasFormat
            ? deps.adapter.replaceWithFormat(candidateSearch, replaceText, format)
            : deps.adapter.replaceInDocument(candidateSearch, replaceText)
          usedSearch = candidateSearch
          if (result.success && result.count > 0) break
        }

        const fallbackCount = attemptedSearches.slice(0, maxAttempts).length

        if (result && result.success && result.count > 0) {
          deps.updateAgentFile({
            additions: result.count,
            status: 'writing',
            name: deps.fileName,
          })
          deps.completeToolActivity(activityId, 'success', `${result.count} 处`)
          const fallbackHint = usedSearch !== search ? '（自动修正 search 后命中）' : ''
          return {
            tool: 'replace',
            success: true,
            message: `成功替换 ${result.count} 处${fallbackHint}：“${search}”→“${replaceText}”`,
            data: {
              count: result.count,
              diffId: result.diffId,
              searchText: usedSearch,
              originalSearchText: search,
              replaceText,
              positions: result.positions,
              attemptedSearches: attemptedSearches.slice(0, maxAttempts),
            },
          }
        }

        const failMessageCore =
          result?.message || `未找到“${search}”，请检查是否与文档内容完全匹配`
        const failMessage = `${failMessageCore}；已自动尝试 ${fallbackCount} 种 search 变体`
        const shortReason =
          failMessage.length > 26 ? `${failMessage.slice(0, 26)}...` : failMessage
        deps.completeToolActivity(activityId, 'error', shortReason)
        return {
          tool: 'replace',
          success: false,
          message: failMessage,
          data: {
            count: result?.count || 0,
            searchText: result?.searchText || search,
            originalSearchText: search,
            usedSearch,
            replaceText,
            attemptedSearches: attemptedSearches.slice(0, maxAttempts),
            debug: result?.debug,
          },
        }
      },
    }),
    defineAgentTool({
      id: 'review',
      displayName: 'Review',
      description: 'Apply review-style edits to the active Word document',
      domain: 'word',
      mutation: 'write',
      concurrency: 'serial',
      tags: ['phase1', 'word', 'review'],
      inputKeys: ['search', 'replace', 'reason', 'type', 'bold', 'italic', 'underline', 'color', 'backgroundColor', 'fontSize', 'fontFamily'],
      async handler(args) {
        const search = args.search || ''
        const replaceText = args.replace || ''
        const reason = args.reason || ''
        const reviewType = args.type || 'style'

        if (!search) {
          return { tool: 'review', success: false, message: '缺少 search 参数' }
        }

        const reviewTypeLabels: Record<string, string> = {
          grammar: '语法',
          logic: '逻辑',
          style: '措辞',
          typo: '错别字',
          format: '格式',
        }
        const typeLabel = reviewTypeLabels[reviewType] || reviewType

        const activityId = deps.claimOrRegisterToolActivity(
          'review',
          args,
          `[${typeLabel}] ${reason || search}`.slice(0, 36),
          { searchText: search, replaceText },
        )
        await deps.flushUiFrame()
        await ensureTiptapEditor(deps)

        const format = parseTextFormat(args)
        const hasFormat =
          format.bold ||
          format.italic ||
          format.underline ||
          !!format.color ||
          !!format.backgroundColor ||
          !!format.fontSize ||
          !!format.fontFamily

        deps.updateAgentAction(
          `[${typeLabel}] 审查：${search.slice(0, 20)}${search.length > 20 ? '...' : ''}`,
        )
        deps.completeAgentStep()
        deps.updateAgentFile({ status: 'writing', name: deps.fileName })
        deps.addAgentFileOperation(
          `审查: [${typeLabel}] "${search.slice(0, 15)}..." -> "${replaceText.slice(0, 15)}..."`,
        )

        const reviewMeta = { reason, type: reviewType }
        const result = hasFormat
          ? deps.adapter.replaceWithFormat(search, replaceText, format, reviewMeta)
          : deps.adapter.replaceInDocument(search, replaceText, reviewMeta)

        if (result.success && result.count > 0) {
          deps.updateAgentFile({
            additions: result.count,
            status: 'writing',
            name: deps.fileName,
          })
          deps.completeToolActivity(activityId, 'success', `[${typeLabel}] ${result.count} 处`)
          return {
            tool: 'review',
            success: true,
            message: `[${typeLabel}] ${reason}\n替换 ${result.count} 处：“${search}”→“${replaceText}”`,
            data: {
              count: result.count,
              diffId: result.diffId,
              reason,
              type: reviewType,
              searchText: result.searchText || search,
              replaceText,
            },
          }
        }

        const failMessage =
          result.message || `未找到“${search}”，请检查是否与文档内容完全匹配`
        const shortReason =
          failMessage.length > 26 ? `${failMessage.slice(0, 26)}...` : failMessage
        deps.completeToolActivity(activityId, 'error', shortReason)
        return {
          tool: 'review',
          success: false,
          message: failMessage,
          data: {
            count: result.count,
            reason,
            type: reviewType,
            searchText: result.searchText || search,
            replaceText,
            debug: result.debug,
          },
        }
      },
    }),
    defineAgentTool({
      id: 'word_edit_ops',
      displayName: 'Word Edit Ops',
      description: 'Apply structured formatting and layout operations',
      domain: 'word',
      mutation: 'transform',
      concurrency: 'serial',
      tags: ['phase1', 'word', 'ops'],
      inputKeys: ['ops', 'dryRun'],
      async handler(args) {
        const rawOps = args.ops || ''
        const dryRunTop = (args.dryRun || '').toLowerCase() === 'true'
        const activityId = deps.claimOrRegisterToolActivity(
          'word_edit_ops',
          args,
          'WordOps: running',
        )
        await deps.flushUiFrame()

        let ops: any[] = []
        if (rawOps) {
          try {
            ops = JSON.parse(rawOps)
          } catch {
            deps.completeToolActivity(activityId, 'error', 'ops JSON invalid')
            return {
              tool: 'word_edit_ops',
              success: false,
              message: 'ops parse failed: expected a JSON array',
            }
          }
        }

        if (!Array.isArray(ops) || ops.length === 0) {
          deps.completeToolActivity(activityId, 'error', 'ops empty')
          return {
            tool: 'word_edit_ops',
            success: false,
            message: 'missing ops: expected a non-empty JSON array',
          }
        }

        deps.patchToolActivity(activityId, { label: `WordOps: ${ops.length} ops` })

        const inferredDryRun = ops.some((op) => op?.dryRun === true)
        const isDryRun = dryRunTop || inferredDryRun

        if (isDryRun) {
          const preview = deps.adapter.previewOps(ops)
          const lines = (preview.data?.lines as string[] | undefined) || []
          deps.setPendingWordOps({
            ops,
            previewMessage: preview.message,
            previewLines: lines,
          })
          deps.completeToolActivity(
            activityId,
            preview.success ? 'success' : 'error',
            preview.success ? `${ops.length} ops preview` : 'preview failed',
          )
          return {
            tool: 'word_edit_ops',
            success: preview.success,
            message: preview.success
              ? `${preview.message}${lines.length ? `\n- ${lines.join('\n- ')}` : ''}\n\nClick Apply Revisions below to execute.`
              : preview.message,
            data: preview.data,
          }
        }

        const prep = await deps.adapter.prepareTemplateFillOutput(ops)
        if (!prep.success) {
          deps.completeToolActivity(activityId, 'error', 'prepare failed')
          return {
            tool: 'word_edit_ops',
            success: false,
            message: prep.message || 'template fill preparation failed',
          }
        }

        const result = deps.adapter.applyOps(ops)
        const createdCount = Number(
          (result.data as { created?: number } | undefined)?.created ?? -1,
        )
        const activityDetail = result.success
          ? createdCount >= 0
            ? `${createdCount} changes`
            : `${ops.length} ops`
          : createdCount === 0
            ? '0 changes'
            : 'apply failed'
        deps.completeToolActivity(
          activityId,
          result.success ? 'success' : 'error',
          activityDetail,
        )
        return {
          tool: 'word_edit_ops',
          success: result.success,
          message: result.message,
          data: result.data,
        }
      },
    }),
    defineAgentTool({
      id: 'insert',
      displayName: 'Insert',
      description: 'Insert content into the active document',
      domain: 'word',
      mutation: 'write',
      concurrency: 'serial',
      tags: ['phase1', 'word', 'insert'],
      inputKeys: ['position', 'content', 'dsl', 'target'],
      async handler(args) {
        const position = args.position || 'end'
        let content = args.content || ''
        let dslBlockCount = 0

        if (args.dsl) {
          try {
            const dslObj: DocDsl = JSON.parse(args.dsl)
            const validation = validateDocDsl(dslObj)
            if (!validation.valid) {
              const errMsg = validation.errors
                .map((e) => `${e.path}: ${e.message}`)
                .join('; ')
              return { tool: 'insert', success: false, message: `DSL 校验失败: ${errMsg}` }
            }
            content = dslToHtml(dslObj)
            dslBlockCount = dslObj.blocks?.length || 0
          } catch (e) {
            return {
              tool: 'insert',
              success: false,
              message: `DSL 解析失败: ${(e as Error).message}`,
            }
          }
        }

        if (!content) {
          return { tool: 'insert', success: false, message: '缺少 content 或 dsl 参数' }
        }

        const activityId = deps.claimOrRegisterToolActivity(
          'insert',
          args,
          `Insert: ${position}`,
          { searchText: args.target || position, replaceText: content },
        )
        await deps.flushUiFrame()
        await ensureTiptapEditor(deps)

        deps.updateAgentAction(
          `正在插入内容到${position === 'start' ? '开头' : position === 'end' ? '末尾' : position}`,
        )
        deps.completeAgentStep()
        deps.addAgentFileOperation(
          `插入: ${dslBlockCount ? `${dslBlockCount} 个 DSL 块` : content.slice(0, 30) + '...'}`,
        )

        let result: { success: boolean; message: string }
        if (
          args.dsl &&
          (position.startsWith('blockIndex:') ||
            position === 'start' ||
            position === 'end' ||
            position.startsWith('after:') ||
            position.startsWith('before:'))
        ) {
          try {
            const dslObj: DocDsl = JSON.parse(args.dsl)
            result = deps.adapter.insertViaDsl(position, dslObj.blocks || [])
          } catch {
            result = deps.adapter.insertInDocument(position, content)
          }
        } else {
          result = deps.adapter.insertInDocument(position, content)
        }

        if (result.success) {
          deps.updateAgentFile({
            additions: 1,
            status: 'writing',
            name: deps.fileName,
          })
          deps.completeToolActivity(
            activityId,
            'success',
            dslBlockCount ? `${dslBlockCount} 块` : undefined,
          )
          return {
            tool: 'insert',
            success: true,
            message:
              (dslBlockCount
                ? `已插入 ${dslBlockCount} 个 DSL 格式块到${position === 'end' ? '末尾' : position === 'start' ? '开头' : position}`
                : result.message),
            data: {
              position,
              contentLength: content.length,
              dslBlocks: dslBlockCount || undefined,
            },
          }
        }

        deps.completeToolActivity(activityId, 'error', result.message)
        return { tool: 'insert', success: false, message: result.message }
      },
    }),
    defineAgentTool({
      id: 'delete',
      displayName: 'Delete',
      description: 'Delete content from the active document',
      domain: 'word',
      mutation: 'delete',
      concurrency: 'serial',
      tags: ['phase1', 'word', 'delete'],
      inputKeys: ['target', 'blockIndex'],
      async handler(args) {
        const target = args.target || ''
        const blockIndex = args.blockIndex ? parseInt(args.blockIndex, 10) : undefined

        if (!target && blockIndex === undefined) {
          return {
            tool: 'delete',
            success: false,
            message: '缺少 target 或 blockIndex 参数',
          }
        }

        const activityId = deps.claimOrRegisterToolActivity(
          'delete',
          args,
          `Delete: ${target.slice(0, 24)}`,
          { searchText: target },
        )
        await deps.flushUiFrame()
        await ensureTiptapEditor(deps)

        deps.updateAgentAction(
          `正在删除 ${target.slice(0, 20)}${target.length > 20 ? '...' : ''}`,
        )
        deps.completeAgentStep()
        deps.addAgentFileOperation(`删除: "${target.slice(0, 30)}..."`)

        const result =
          blockIndex !== undefined && !isNaN(blockIndex)
            ? deps.adapter.deleteViaDsl(target, { blockIndex })
            : deps.adapter.deleteInDocument(target)

        if (result.success) {
          deps.updateAgentFile({
            deletions: result.count,
            status: 'writing',
            name: deps.fileName,
          })
          deps.completeToolActivity(activityId, 'success', `${result.count} 处`)
          return {
            tool: 'delete',
            success: true,
            message: result.message,
            data: { count: result.count, target },
          }
        }

        deps.completeToolActivity(activityId, 'error', result.message)
        return { tool: 'delete', success: false, message: result.message }
      },
    }),
    defineAgentTool({
      id: 'word_chart',
      displayName: 'Word Chart',
      description: 'Insert a chart placeholder into the active Word document',
      domain: 'word',
      mutation: 'create',
      concurrency: 'serial',
      async handler(args) {
        const chartType = args.type || 'bar'
        const title = args.title || ''
        const position = args.position || 'end'
        const width = args.width ? parseInt(args.width, 10) : 500
        const height = args.height ? parseInt(args.height, 10) : 300

        let categories: string[]
        let series: Array<{ name?: string; values?: number[] }>
        try {
          categories = JSON.parse(args.categories || '[]')
          series = JSON.parse(args.series || '[]')
        } catch (error) {
          return {
            tool: 'word_chart',
            success: false,
            message: `Chart parameter parse failed: ${error}`,
          }
        }

        if (!categories.length || !series.length) {
          return {
            tool: 'word_chart',
            success: false,
            message: 'Missing categories or series data.',
          }
        }

        const chartConfig = {
          type: chartType,
          title: title || undefined,
          categories,
          series,
          widthPx: width,
          heightPx: height,
          stacking: args.stacking,
          legendPosition: args.legendPosition || 'bottom',
        }

        const encoded = encodeURIComponent(JSON.stringify(chartConfig))
        const safeTitleAttr = title ? ` data-chart-title="${String(title).replace(/"/g, '&quot;')}"` : ''
        const chartHtml = `<div data-type="docx-chart" data-chart-config="${encoded}"${safeTitleAttr} style="width:${width}px;height:${height}px"></div>`

        const activityId = deps.claimOrRegisterToolActivity(
          'word_chart',
          args,
          `Chart: ${chartType}`,
          { searchText: '', replaceText: title || chartType },
        )
        await deps.flushUiFrame()
        await ensureTiptapEditor(deps)

        deps.updateAgentAction(`Inserting ${title || chartType} chart...`)
        deps.completeAgentStep()
        const result = deps.adapter.insertInDocument(position, chartHtml)

        if (result.success) {
          deps.updateAgentFile({
            additions: 1,
            status: 'writing',
            name: deps.fileName,
          })
          deps.completeToolActivity(activityId, 'success', `${chartType} chart`)
          return {
            tool: 'word_chart',
            success: true,
            message: `Inserted ${title ? `"${title}" ` : ''}${chartType} chart at ${position}.`,
            data: { chartType, position, title },
          }
        }

        deps.completeToolActivity(activityId, 'error', result.message)
        return {
          tool: 'word_chart',
          success: false,
          message: result.message,
        }
      },
    }),
  )

  return tools
}
