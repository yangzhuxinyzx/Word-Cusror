import { Mark, mergeAttributes } from '@tiptap/core'

function parseIds(raw: string | null): string[] {
  if (!raw) return []
  return raw
    .split(',')
    .map(s => s.trim())
    .filter(Boolean)
}

/**
 * CommentMark - Word 批注范围标记（仅用于锚定范围 + 高亮）
 * HTML shape: <span class="docx-comment" data-comment-ids="1,2">...</span>
 */
export const CommentMark = Mark.create({
  name: 'comment',

  addAttributes() {
    return {
      commentIds: {
        default: [] as string[],
        parseHTML: (el: HTMLElement) => parseIds(el.getAttribute('data-comment-ids') || el.getAttribute('data-comment-id')),
        renderHTML: (attrs: { commentIds?: string[] }) => {
          const ids = Array.isArray(attrs.commentIds) ? attrs.commentIds.filter(Boolean) : []
          if (ids.length === 0) return {}
          return { 'data-comment-ids': ids.join(',') }
        },
      },
    }
  },

  parseHTML() {
    return [
      { tag: 'span.docx-comment[data-comment-ids]' },
      { tag: 'span[data-comment-ids]' },
      { tag: 'span[data-comment-id]' },
    ]
  },

  renderHTML({ HTMLAttributes }) {
    return [
      'span',
      mergeAttributes(HTMLAttributes, {
        class: 'docx-comment',
        style: 'background-color: rgba(254, 249, 195, 0.9); border-bottom: 1px dotted rgba(245, 158, 11, 0.9);',
      }),
      0,
    ]
  },
})











