import type { AgentSkillDefinition } from './SkillRegistry'

export const BUILTIN_SKILLS = [
  {
    id: 'rewrite-formal',
    displayName: 'Rewrite Formal',
    description: 'Rewrite the target content into formal written Chinese.',
    source: 'builtin',
    executionKind: 'prompt_transform',
    safety: 'mutating',
    prompt:
      'Use a formal, concise, professional tone. Preserve facts, structure, and named entities. Prefer editing the current document or selection rather than generating unrelated prose.',
    toolIds: [
      'word.read',
      'word.edit',
      'word.format',
    ],
    tags: ['rewrite', 'document', 'style'],
    invocation: {
      slashCommands: ['rewrite-formal', 'formal'],
    },
  },
  {
    id: 'rewrite-report',
    displayName: 'Rewrite Report',
    description: 'Rewrite content into report style with headings and clear structure.',
    source: 'builtin',
    executionKind: 'prompt_transform',
    safety: 'mutating',
    prompt:
      'Rewrite the target content into a report-style document. Use explicit headings, grouped bullet points when needed, and preserve all important details.',
    toolIds: [
      'word.read',
      'word.edit',
      'word.format',
    ],
    tags: ['rewrite', 'report', 'document'],
    invocation: {
      slashCommands: ['rewrite-report', 'report'],
    },
  },
  {
    id: 'summarize-to-slides',
    displayName: 'Summarize To Slides',
    description: 'Turn the source content into a slide-ready outline.',
    source: 'builtin',
    executionKind: 'prompt_transform',
    safety: 'planning',
    prompt:
      'Extract the core ideas and produce a slide-ready outline. Focus on page structure, key message per slide, and concise bullet points.',
    toolIds: ['workspace_read', 'word.read', 'ppt_create'],
    tags: ['summary', 'ppt', 'outline'],
    invocation: {
      slashCommands: ['summarize-to-slides', 'slides'],
    },
  },
  {
    id: 'generate-meeting-minutes',
    displayName: 'Generate Meeting Minutes',
    description: 'Generate structured meeting minutes from notes or source materials.',
    source: 'builtin',
    executionKind: 'prompt_transform',
    safety: 'mutating',
    prompt:
      'Generate meeting minutes with attendees, agenda, decisions, actions, and follow-ups. Use concise, traceable wording.',
    toolIds: ['workspace_read', 'word.create', 'word.edit'],
    tags: ['minutes', 'meeting', 'document'],
    invocation: {
      slashCommands: ['generate-meeting-minutes', 'meeting-minutes', 'minutes'],
    },
  },
  {
    id: 'document-proofread',
    displayName: 'Document Proofread',
    description: 'Proofread and fix wording, grammar, and consistency issues.',
    source: 'builtin',
    executionKind: 'prompt_transform',
    safety: 'verification',
    prompt:
      'Proofread the target document. Fix grammar, punctuation, consistency, and obvious style problems without changing meaning.',
    toolIds: [
      'word.read',
      'word.edit',
      'word.format',
      'word.resolve_change',
    ],
    tags: ['proofread', 'verification', 'document'],
    invocation: {
      slashCommands: ['document-proofread', 'proofread'],
    },
  },
  {
    id: 'format-normalization',
    displayName: 'Format Normalization',
    description: 'Normalize headings, lists, spacing, punctuation, and document formatting.',
    source: 'builtin',
    executionKind: 'prompt_transform',
    safety: 'mutating',
    prompt:
      'Normalize document formatting, heading levels, list style, spacing, punctuation, and structural consistency.',
    toolIds: [
      'word.read',
      'word.format',
      'word.edit',
    ],
    tags: ['format', 'document', 'normalization'],
    invocation: {
      slashCommands: ['format-normalization', 'normalize-format'],
    },
  },
  {
    id: 'template-based-doc',
    displayName: 'Template Based Doc',
    description: 'Create or fill a document using templates in the workspace.',
    source: 'builtin',
    executionKind: 'prompt_transform',
    safety: 'mutating',
    prompt:
      'Look for relevant templates in the workspace, reuse their structure and placeholders, then create or fill the target document.',
    toolIds: ['workspace_profile', 'workspace_read', 'word.create'],
    tags: ['template', 'document', 'workspace'],
    invocation: {
      slashCommands: ['template-based-doc', 'template-doc'],
    },
  },
  {
    id: 'ppt-from-outline',
    displayName: 'PPT From Outline',
    description: 'Create or refine a PPT deck from a structured outline.',
    source: 'builtin',
    executionKind: 'prompt_transform',
    safety: 'mutating',
    prompt:
      'Use a structured outline to build a PPT. If the request is to generate a deck, produce a clean outline first and then use PPT tools to create or edit slides.',
    toolIds: ['workspace_read', 'ppt_create', 'ppt_edit'],
    tags: ['ppt', 'outline', 'deck'],
    invocation: {
      slashCommands: ['ppt-from-outline', 'ppt'],
    },
  },
  {
    id: 'excel-cleanup',
    displayName: 'Excel Cleanup',
    description: 'Clean workbook structure, normalize columns, and fix obvious data quality issues.',
    source: 'builtin',
    executionKind: 'prompt_transform',
    safety: 'mutating',
    prompt:
      'Inspect the workbook, identify obvious cleanup actions, and then update sheets carefully. Prefer reversible or explicit operations.',
    toolIds: [
      'excel_read',
      'excel_search',
      'excel_write',
      'excel_sort',
      'excel_find_replace',
    ],
    tags: ['excel', 'cleanup', 'data'],
    invocation: {
      slashCommands: ['excel-cleanup', 'excel'],
    },
  },
  {
    id: 'skills',
    displayName: 'Skill List',
    description: 'List available builtin and workspace skills for the current session.',
    source: 'builtin',
    executionKind: 'workflow',
    safety: 'read_only',
    prompt:
      'List currently available skills, their intent, safety level, and activation commands.',
    tags: ['skills', 'discovery', 'read_only'],
    invocation: {
      slashCommands: ['skills', '技能'],
      aliases: ['skill-list'],
    },
  },
  {
    id: 'init',
    displayName: 'Workspace Init',
    description: 'Build a workspace profile before starting document or file operations.',
    source: 'builtin',
    executionKind: 'workflow',
    safety: 'read_only',
    prompt:
      'Scan the current workspace, identify high-value files, summarize them, and produce a structured workspace profile with next actions.',
    toolIds: ['workspace_profile', 'workspace_list', 'workspace_read'],
    tags: ['workspace', 'init', 'read_only'],
    invocation: {
      slashCommands: ['init', '初始化', '项目理解'],
      aliases: ['workspace-init', 'workspace-profile'],
    },
  },
] as const satisfies readonly AgentSkillDefinition[]
