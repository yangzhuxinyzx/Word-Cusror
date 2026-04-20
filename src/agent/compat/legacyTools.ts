import { defineAgentTool, type AgentToolDefinition } from '../tools/contracts'

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

const DOC_DSL_SCHEMA = {
  type: 'object',
  properties: {
    blocks: {
      type: 'array',
      items: {
        type: 'object',
        additionalProperties: true,
      },
    },
  },
  required: ['blocks'],
  additionalProperties: true,
} as const

const WORD_OPS_SCHEMA = {
  type: 'array',
  items: {
    type: 'object',
    additionalProperties: true,
  },
} as const

const REPLACEMENTS_SCHEMA = {
  type: 'array',
  items: {
    type: 'object',
    properties: {
      search: { type: 'string' },
      replace: { type: 'string' },
    },
    required: ['search', 'replace'],
    additionalProperties: true,
  },
} as const

const WORD_CHART_SERIES_SCHEMA = {
  type: 'array',
  items: {
    type: 'object',
    properties: {
      name: { type: 'string' },
      values: {
        type: 'array',
        items: { type: 'number' },
      },
    },
    required: ['values'],
    additionalProperties: true,
  },
} as const

export const LEGACY_TOOL_DEFINITIONS = [
  defineAgentTool({
    id: 'word.read',
    displayName: 'Word Read',
    description: 'Read the current Word document context such as the selection or outline',
    domain: 'word',
    mutation: 'read',
    concurrency: 'parallel_safe',
    tags: ['legacy', 'word', 'read', 'canonical'],
    inputKeys: ['target'],
    inputSchema: {
      type: 'object',
      properties: {
        target: {
          type: 'string',
          enum: [
            'selection',
            'outline',
            'document',
            'pending_changes',
            'dsl',
            'block_map',
            'style_profile',
          ],
        },
      },
      additionalProperties: true,
    },
    legacyAliases: ['word.read_selection', 'word.read_outline'],
    prompt:
      'Use target=selection, outline, document, pending_changes, dsl, block_map, or style_profile to read structured Word state.',
  }),
  defineAgentTool({
    id: 'word.create',
    displayName: 'Word Create',
    description: 'Create a Word document, optionally from a template, with DSL-first structured content',
    domain: 'word',
    mutation: 'create',
    concurrency: 'serial',
    tags: ['legacy', 'word', 'create', 'canonical', 'dsl'],
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
    inputSchema: {
      type: 'object',
      properties: {
        mode: {
          type: 'string',
          enum: ['document', 'template'],
        },
        title: { type: 'string' },
        newTitle: { type: 'string' },
        content: { type: 'string' },
        dsl: DOC_DSL_SCHEMA,
        elements: {
          type: 'array',
          items: {
            type: 'object',
            additionalProperties: true,
          },
        },
        styleRefPath: { type: 'string' },
        styleRefFileName: { type: 'string' },
        styleRefName: { type: 'string' },
        contentRefPath: { type: 'string' },
        contentRefFileName: { type: 'string' },
        contentRefName: { type: 'string' },
        templatePath: { type: 'string' },
        templateFileName: { type: 'string' },
        templateName: { type: 'string' },
        replacements: REPLACEMENTS_SCHEMA,
      },
      additionalProperties: true,
    },
    legacyAliases: ['create', 'create_from_template', 'copy_template'],
    prompt:
      'Prefer DSL input when creating new content. Use mode=template when creating from a template with replacements.',
  }),
  defineAgentTool({
    id: 'word.edit',
    displayName: 'Word Edit',
    description: 'Edit the active Word document with DSL-first replace, insert, delete, or review operations',
    domain: 'word',
    mutation: 'write',
    concurrency: 'serial',
    tags: ['legacy', 'word', 'edit', 'canonical', 'dsl'],
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
    inputSchema: {
      type: 'object',
      properties: {
        operation: {
          type: 'string',
          enum: ['replace', 'review', 'insert', 'delete'],
        },
        strategy: {
          type: 'string',
          enum: ['auto', 'dsl', 'structured', 'text'],
        },
        search: { type: 'string' },
        replace: { type: 'string' },
        position: { type: 'string' },
        content: { type: 'string' },
        dsl: DOC_DSL_SCHEMA,
        target: { type: 'string' },
        blockIndex: { type: 'integer' },
        reason: { type: 'string' },
        type: { type: 'string' },
        bold: { type: 'boolean' },
        italic: { type: 'boolean' },
        underline: { type: 'boolean' },
        color: { type: 'string' },
        backgroundColor: { type: 'string' },
        fontSize: {
          anyOf: [{ type: 'number' }, { type: 'string' }],
        },
        fontFamily: { type: 'string' },
      },
      additionalProperties: true,
    },
    legacyAliases: [
      'replace',
      'review',
      'insert',
      'delete',
      'word.replace_via_dsl',
      'word.insert_via_dsl',
      'word.delete_via_dsl',
    ],
    prompt:
      'Use operation=replace|review|insert|delete. Prefer strategy=dsl or blockIndex-scoped edits for structured document updates.',
  }),
  defineAgentTool({
    id: 'word.format',
    displayName: 'Word Format',
    description: 'Preview or apply structured Word formatting/layout ops',
    domain: 'word',
    mutation: 'transform',
    concurrency: 'serial',
    tags: ['legacy', 'word', 'format', 'canonical', 'ops'],
    inputKeys: ['mode', 'ops', 'dryRun'],
    inputSchema: {
      type: 'object',
      properties: {
        mode: {
          type: 'string',
          enum: ['preview', 'apply', 'execute'],
        },
        ops: WORD_OPS_SCHEMA,
        dryRun: { type: 'boolean' },
      },
      required: ['ops'],
      additionalProperties: true,
    },
    legacyAliases: ['word_edit_ops', 'word.preview_ops', 'word.apply_ops'],
    prompt:
      'Provide ops as a JSON array. Use mode=preview for dry runs and mode=apply to execute the ops.',
  }),
  defineAgentTool({
    id: 'word.resolve_change',
    displayName: 'Word Resolve Change',
    description: 'Accept or reject pending Word review changes',
    domain: 'word',
    mutation: 'write',
    concurrency: 'serial',
    tags: ['legacy', 'word', 'review', 'canonical'],
    inputKeys: ['action', 'changeId', 'id', 'all'],
    inputSchema: {
      type: 'object',
      properties: {
        action: {
          type: 'string',
          enum: ['list', 'accept', 'reject'],
        },
        changeId: { type: 'string' },
        id: { type: 'string' },
        all: { type: 'boolean' },
      },
      additionalProperties: true,
    },
    legacyAliases: ['word.accept_change', 'word.reject_change'],
    prompt:
      'Use action=list to inspect pending changes, or action=accept/action=reject to resolve them. Supply changeId for a single pending change or all=true for all changes.',
  }),
  defineAgentTool({
    id: 'word.chart',
    displayName: 'Word Chart',
    description: 'Insert a generated chart into the active Word document',
    domain: 'word',
    mutation: 'create',
    concurrency: 'serial',
    tags: ['legacy', 'word', 'chart', 'canonical'],
    inputKeys: ['type', 'title', 'categories', 'series', 'position', 'width', 'height'],
    inputSchema: {
      type: 'object',
      properties: {
        type: {
          type: 'string',
          enum: ['bar', 'column', 'line', 'pie', 'area'],
        },
        title: { type: 'string' },
        categories: {
          type: 'array',
          items: { type: 'string' },
        },
        series: WORD_CHART_SERIES_SCHEMA,
        position: { type: 'string' },
        width: { type: 'integer' },
        height: { type: 'integer' },
      },
      required: ['categories', 'series'],
      additionalProperties: true,
    },
    legacyAliases: ['word_chart'],
  }),
  defineAgentTool({
    id: 'ppt_create',
    displayName: 'Create PPT',
    description: 'Generate a new PPTX deck',
    domain: 'ppt',
    mutation: 'create',
    concurrency: 'serial',
    tags: ['legacy', 'ppt', 'create'],
    inputKeys: ['title', 'outline', 'theme', 'style'],
  }),
  defineAgentTool({
    id: 'ppt_edit',
    displayName: 'Edit PPT',
    description: 'Edit an existing PPT deck or selected slides',
    domain: 'ppt',
    mutation: 'write',
    concurrency: 'serial',
    tags: ['legacy', 'ppt', 'edit'],
    inputKeys: ['instruction', 'pages', 'mode'],
  }),
  defineAgentTool({
    id: 'ppt_text_edit',
    displayName: 'Edit PPT Text',
    description: 'Detect and replace text inside an image-only PPT slide',
    domain: 'ppt',
    mutation: 'write',
    concurrency: 'serial',
    tags: ['legacy', 'ppt', 'text-edit'],
    inputKeys: ['pptxPath', 'pageNumber', 'edits'],
  }),
  defineAgentTool({
    id: 'workspace_list',
    displayName: 'Workspace List',
    description: 'List files under the workspace or a specific folder',
    domain: 'workspace',
    mutation: 'read',
    concurrency: 'parallel_safe',
    tags: ['legacy', 'workspace', 'files'],
    inputKeys: ['folder', 'path'],
  }),
  defineAgentTool({
    id: 'workspace_open',
    displayName: 'Workspace Open',
    description: 'Open a workspace file into the editor',
    domain: 'workspace',
    mutation: 'read',
    concurrency: 'serial',
    tags: ['legacy', 'workspace', 'open'],
    inputKeys: ['path', 'filePath', 'name'],
  }),
  defineAgentTool({
    id: 'workspace_summarize',
    displayName: 'Workspace Summarize',
    description: 'Summarize a file or folder from the workspace',
    domain: 'workspace',
    mutation: 'read',
    concurrency: 'parallel_safe',
    tags: ['legacy', 'workspace', 'summary'],
    inputKeys: ['path', 'filePath', 'name'],
  }),
  defineAgentTool({
    id: 'workspace_read',
    displayName: 'Workspace Read',
    description: 'Read a file from the workspace',
    domain: 'workspace',
    mutation: 'read',
    concurrency: 'parallel_safe',
    tags: ['legacy', 'workspace', 'read'],
    inputKeys: ['path', 'filePath', 'name', 'relativePath'],
  }),
  defineAgentTool({
    id: 'web_search',
    displayName: 'Web Search',
    description: 'Search the web for external information',
    domain: 'web',
    mutation: 'external',
    concurrency: 'parallel_safe',
    tags: ['legacy', 'web', 'search'],
    inputKeys: ['query'],
  }),
  defineAgentTool({
    id: 'knowledge_search',
    displayName: 'Knowledge Search',
    description: 'Search local workspace/global knowledge and profile memory',
    domain: 'knowledge',
    mutation: 'read',
    concurrency: 'parallel_safe',
    tags: ['legacy', 'knowledge', 'search'],
    inputKeys: ['query', 'topK'],
  }),
  defineAgentTool({
    id: 'excel_read',
    displayName: 'Excel Read',
    description: 'Read cell ranges from the active Excel workbook',
    domain: 'excel',
    mutation: 'read',
    concurrency: 'parallel_safe',
    tags: ['legacy', 'excel', 'read'],
    inputKeys: ['sheet', 'range'],
  }),
  defineAgentTool({
    id: 'excel_search',
    displayName: 'Excel Search',
    description: 'Search within the active Excel workbook',
    domain: 'excel',
    mutation: 'read',
    concurrency: 'parallel_safe',
    tags: ['legacy', 'excel', 'search'],
    inputKeys: ['query', 'sheet'],
  }),
  defineAgentTool({
    id: 'excel_write',
    displayName: 'Excel Write',
    description: 'Write cell values into the active Excel workbook',
    domain: 'excel',
    mutation: 'write',
    concurrency: 'serial',
    tags: ['legacy', 'excel', 'write'],
    inputKeys: ['sheet', 'range', 'value', 'values'],
  }),
  defineAgentTool({
    id: 'excel_insert_rows',
    displayName: 'Excel Insert Rows',
    description: 'Insert rows into the active Excel workbook',
    domain: 'excel',
    mutation: 'write',
    concurrency: 'serial',
    tags: ['legacy', 'excel', 'rows'],
    inputKeys: ['sheet', 'startRow', 'count', 'data'],
  }),
  defineAgentTool({
    id: 'excel_insert_columns',
    displayName: 'Excel Insert Columns',
    description: 'Insert columns into the active Excel workbook',
    domain: 'excel',
    mutation: 'write',
    concurrency: 'serial',
    tags: ['legacy', 'excel', 'columns'],
    inputKeys: ['sheet', 'startCol', 'count'],
  }),
  defineAgentTool({
    id: 'excel_delete_rows',
    displayName: 'Excel Delete Rows',
    description: 'Delete rows from the active Excel workbook',
    domain: 'excel',
    mutation: 'delete',
    concurrency: 'serial',
    tags: ['legacy', 'excel', 'rows'],
    inputKeys: ['sheet', 'startRow', 'count'],
  }),
  defineAgentTool({
    id: 'excel_delete_columns',
    displayName: 'Excel Delete Columns',
    description: 'Delete columns from the active Excel workbook',
    domain: 'excel',
    mutation: 'delete',
    concurrency: 'serial',
    tags: ['legacy', 'excel', 'columns'],
    inputKeys: ['sheet', 'startCol', 'count'],
  }),
  defineAgentTool({
    id: 'excel_add_sheet',
    displayName: 'Excel Add Sheet',
    description: 'Add a worksheet to the active Excel workbook',
    domain: 'excel',
    mutation: 'create',
    concurrency: 'serial',
    tags: ['legacy', 'excel', 'sheet'],
    inputKeys: ['name'],
  }),
  defineAgentTool({
    id: 'excel_delete_sheet',
    displayName: 'Excel Delete Sheet',
    description: 'Delete a worksheet from the active Excel workbook',
    domain: 'excel',
    mutation: 'delete',
    concurrency: 'serial',
    tags: ['legacy', 'excel', 'sheet'],
    inputKeys: ['name'],
  }),
  defineAgentTool({
    id: 'excel_merge',
    displayName: 'Excel Merge',
    description: 'Merge cells in the active Excel workbook',
    domain: 'excel',
    mutation: 'write',
    concurrency: 'serial',
    tags: ['legacy', 'excel', 'merge'],
    inputKeys: ['sheet', 'range'],
  }),
  defineAgentTool({
    id: 'excel_unmerge',
    displayName: 'Excel Unmerge',
    description: 'Unmerge cells in the active Excel workbook',
    domain: 'excel',
    mutation: 'write',
    concurrency: 'serial',
    tags: ['legacy', 'excel', 'merge'],
    inputKeys: ['sheet', 'range'],
  }),
  defineAgentTool({
    id: 'excel_create',
    displayName: 'Excel Create',
    description: 'Create a new Excel workbook',
    domain: 'excel',
    mutation: 'create',
    concurrency: 'serial',
    tags: ['legacy', 'excel', 'create'],
    inputKeys: ['title', 'sheets'],
  }),
  defineAgentTool({
    id: 'excel_formula',
    displayName: 'Excel Formula',
    description: 'Apply formulas in the active Excel workbook',
    domain: 'excel',
    mutation: 'write',
    concurrency: 'serial',
    tags: ['legacy', 'excel', 'formula'],
    inputKeys: ['sheet', 'range', 'formula'],
  }),
  defineAgentTool({
    id: 'excel_sort',
    displayName: 'Excel Sort',
    description: 'Sort a range in the active Excel workbook',
    domain: 'excel',
    mutation: 'write',
    concurrency: 'serial',
    tags: ['legacy', 'excel', 'sort'],
    inputKeys: ['sheet', 'range', 'column', 'direction'],
  }),
  defineAgentTool({
    id: 'excel_autofill',
    displayName: 'Excel Autofill',
    description: 'Autofill a range in the active Excel workbook',
    domain: 'excel',
    mutation: 'write',
    concurrency: 'serial',
    tags: ['legacy', 'excel', 'autofill'],
    inputKeys: ['sheet', 'sourceRange', 'targetRange'],
  }),
  defineAgentTool({
    id: 'excel_dimensions',
    displayName: 'Excel Dimensions',
    description: 'Update row or column dimensions in the active Excel workbook',
    domain: 'excel',
    mutation: 'write',
    concurrency: 'serial',
    tags: ['legacy', 'excel', 'dimensions'],
    inputKeys: ['sheet', 'type', 'index', 'size'],
  }),
  defineAgentTool({
    id: 'excel_conditional_format',
    displayName: 'Excel Conditional Format',
    description: 'Apply conditional formatting rules in the active Excel workbook',
    domain: 'excel',
    mutation: 'write',
    concurrency: 'serial',
    tags: ['legacy', 'excel', 'format'],
    inputKeys: ['sheet', 'range', 'rule'],
  }),
  defineAgentTool({
    id: 'excel_calculate',
    displayName: 'Excel Calculate',
    description: 'Recalculate formulas in the active Excel workbook',
    domain: 'excel',
    mutation: 'transform',
    concurrency: 'serial',
    tags: ['legacy', 'excel', 'calculate'],
    inputKeys: ['sheet'],
  }),
  defineAgentTool({
    id: 'excel_filter',
    displayName: 'Excel Filter',
    description: 'Apply filters in the active Excel workbook',
    domain: 'excel',
    mutation: 'write',
    concurrency: 'serial',
    tags: ['legacy', 'excel', 'filter'],
    inputKeys: ['sheet', 'range', 'criteria'],
  }),
  defineAgentTool({
    id: 'excel_validation',
    displayName: 'Excel Validation',
    description: 'Apply data validation rules in the active Excel workbook',
    domain: 'excel',
    mutation: 'write',
    concurrency: 'serial',
    tags: ['legacy', 'excel', 'validation'],
    inputKeys: ['sheet', 'range', 'rule'],
  }),
  defineAgentTool({
    id: 'excel_hyperlink',
    displayName: 'Excel Hyperlink',
    description: 'Add hyperlinks in the active Excel workbook',
    domain: 'excel',
    mutation: 'write',
    concurrency: 'serial',
    tags: ['legacy', 'excel', 'hyperlink'],
    inputKeys: ['sheet', 'cell', 'url', 'text'],
  }),
  defineAgentTool({
    id: 'excel_find_replace',
    displayName: 'Excel Find Replace',
    description: 'Find and replace values in the active Excel workbook',
    domain: 'excel',
    mutation: 'write',
    concurrency: 'serial',
    tags: ['legacy', 'excel', 'find-replace'],
    inputKeys: ['sheet', 'search', 'replace'],
  }),
  defineAgentTool({
    id: 'excel_chart',
    displayName: 'Excel Chart',
    description: 'Create a chart from the active Excel workbook',
    domain: 'excel',
    mutation: 'create',
    concurrency: 'serial',
    tags: ['legacy', 'excel', 'chart'],
    inputKeys: ['sheet', 'range', 'chartType'],
  }),
] as const satisfies readonly AgentToolDefinition[]

export type LegacyToolId = (typeof LEGACY_TOOL_DEFINITIONS)[number]['id']

export const LEGACY_TOOL_IDS = LEGACY_TOOL_DEFINITIONS.map((definition) => definition.id)

export const LEGACY_TOOL_NAMES = Array.from(
  new Set(
    LEGACY_TOOL_DEFINITIONS.flatMap((definition) => [
      definition.id,
      ...(definition.legacyAliases || []),
    ]),
  ),
)

const LEGACY_TOOL_ID_SET = new Set<string>(LEGACY_TOOL_IDS)

export const LEGACY_PREVIEW_TRACKED_TOOL_IDS = [
  'word.edit',
  'word.format',
  'word.chart',
  'word.resolve_change',
] as const satisfies readonly LegacyToolId[]

const LEGACY_PREVIEW_TRACKED_TOOL_ID_SET = new Set<string>(LEGACY_PREVIEW_TRACKED_TOOL_IDS)

export const LEGACY_XML_TOOL_PATTERN_SOURCE = LEGACY_TOOL_NAMES
  .map((toolId) => escapeRegExp(toolId))
  .join('|')

export function createLegacyXmlToolBlockRegex(): RegExp {
  return new RegExp(`<(${LEGACY_XML_TOOL_PATTERN_SOURCE})>([\\s\\S]*?)<\\/\\1>`, 'gi')
}

export function createLegacyXmlToolOpenTagRegex(): RegExp {
  return new RegExp(`<(${LEGACY_XML_TOOL_PATTERN_SOURCE})>`, 'gi')
}

export function isLegacyToolId(value: string): value is LegacyToolId {
  return LEGACY_TOOL_ID_SET.has(value)
}

export function isLegacyPreviewTrackedTool(value: string): value is (typeof LEGACY_PREVIEW_TRACKED_TOOL_IDS)[number] {
  return LEGACY_PREVIEW_TRACKED_TOOL_ID_SET.has(value)
}

export function getLegacyToolDefinition(id: string): AgentToolDefinition | undefined {
  return LEGACY_TOOL_DEFINITIONS.find((definition) => definition.id === id)
}
