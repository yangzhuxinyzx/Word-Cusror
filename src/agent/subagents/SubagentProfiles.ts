export type SubagentSafety =
  | 'read_only'
  | 'planning'
  | 'mutating'
  | 'verification'

export type SubagentExecutionMode = 'sync' | 'background'

export interface SubagentProfileDefinition {
  id: string
  displayName: string
  description: string
  safety: SubagentSafety
  defaultMode: SubagentExecutionMode
  canRunInBackground?: boolean
  prompt: string
  allowedToolIds?: string[]
  tags?: string[]
}

export const BUILTIN_SUBAGENT_PROFILES = [
  {
    id: 'doc-explore',
    displayName: 'Doc Explore',
    description: 'Read and analyze the current document without modifying it.',
    safety: 'read_only',
    defaultMode: 'sync',
    canRunInBackground: true,
    prompt:
      'Focus on reading, outlining, and extracting insights from the current document. Do not perform mutating operations.',
    allowedToolIds: [
      'word.read',
      'workspace_read',
      'workspace_profile',
      'knowledge_search',
      'memory_search',
      'web_search',
    ],
    tags: ['document', 'read_only', 'analysis'],
  },
  {
    id: 'workspace-explore',
    displayName: 'Workspace Explore',
    description: 'Survey the workspace, identify high-value files, and summarize findings.',
    safety: 'read_only',
    defaultMode: 'background',
    canRunInBackground: true,
    prompt:
      'Inspect the workspace, map relevant files, summarize source material, and return a concise evidence-based report.',
    allowedToolIds: [
      'workspace_list',
      'workspace_profile',
      'workspace_read',
      'workspace_summarize',
      'knowledge_search',
      'memory_search',
      'web_search',
    ],
    tags: ['workspace', 'read_only', 'background'],
  },
  {
    id: 'doc-editor',
    displayName: 'Doc Editor',
    description: 'Execute focused document editing tasks against the active document.',
    safety: 'mutating',
    defaultMode: 'sync',
    canRunInBackground: false,
    prompt:
      'Edit the target document directly, keep changes bounded to the assigned task, and return a concise change summary.',
    allowedToolIds: [
      'word.edit',
      'word.format',
      'word.resolve_change',
    ],
    tags: ['document', 'editing'],
  },
  {
    id: 'ppt-builder',
    displayName: 'PPT Builder',
    description: 'Create or refine PPT output from outline or feedback.',
    safety: 'mutating',
    defaultMode: 'background',
    canRunInBackground: true,
    prompt:
      'Build or revise presentation output, keep slide structure coherent, and report the resulting deck path or slide summary.',
    allowedToolIds: ['ppt_create', 'ppt_edit', 'workspace_read', 'workspace_profile'],
    tags: ['ppt', 'background'],
  },
  {
    id: 'excel-operator',
    displayName: 'Excel Operator',
    description: 'Inspect and modify workbook content for bounded spreadsheet tasks.',
    safety: 'mutating',
    defaultMode: 'sync',
    canRunInBackground: false,
    prompt:
      'Operate on workbook content carefully, validate target sheet/range first, and summarize the exact spreadsheet changes.',
    allowedToolIds: [
      'excel_read',
      'excel_search',
      'excel_write',
      'excel_formula',
      'excel_sort',
      'excel_find_replace',
      'excel_filter',
      'excel_validation',
      'excel_chart',
    ],
    tags: ['excel', 'spreadsheet'],
  },
  {
    id: 'verification',
    displayName: 'Verification',
    description: 'Critically verify outputs, find risks, and challenge weak assumptions.',
    safety: 'verification',
    defaultMode: 'sync',
    canRunInBackground: true,
    prompt:
      'Act as a verification specialist. Prefer evidence gathering, challenge unsupported claims, and report concrete findings first.',
    allowedToolIds: [
      'word.read',
      'workspace_read',
      'workspace_profile',
      'workspace_list',
      'knowledge_search',
      'memory_search',
      'web_search',
    ],
    tags: ['verification', 'read_only'],
  },
] as const satisfies readonly SubagentProfileDefinition[]
