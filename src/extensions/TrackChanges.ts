import { Mark, mergeAttributes } from '@tiptap/core'

type TrackAttrs = {
  trackId?: string | null
  author?: string | null
  date?: string | null
}

function parseStrAttr(el: HTMLElement, name: string): string | null {
  const v = el.getAttribute(name)
  return v && v.trim() ? v.trim() : null
}

/**
 * TrackInsert - Word 修订“插入”标记（绿色高亮）
 * HTML shape: <span class="docx-track" data-track-type="insert" data-track-id=".." ...>...</span>
 */
export const TrackInsert = Mark.create({
  name: 'trackInsert',

  addAttributes() {
    return {
      trackId: {
        default: null,
        parseHTML: (el: HTMLElement) => parseStrAttr(el, 'data-track-id'),
        renderHTML: (attrs: TrackAttrs) => (attrs.trackId ? { 'data-track-id': attrs.trackId } : {}),
      },
      author: {
        default: null,
        parseHTML: (el: HTMLElement) => parseStrAttr(el, 'data-track-author'),
        renderHTML: (attrs: TrackAttrs) => (attrs.author ? { 'data-track-author': attrs.author } : {}),
      },
      date: {
        default: null,
        parseHTML: (el: HTMLElement) => parseStrAttr(el, 'data-track-date'),
        renderHTML: (attrs: TrackAttrs) => (attrs.date ? { 'data-track-date': attrs.date } : {}),
      },
    }
  },

  parseHTML() {
    return [
      { tag: 'span.docx-track[data-track-type="insert"]' },
      { tag: 'span[data-track-type="insert"]' },
    ]
  },

  renderHTML({ HTMLAttributes }) {
    return [
      'span',
      mergeAttributes(HTMLAttributes, {
        class: 'docx-track docx-track-insert',
        'data-track-type': 'insert',
      }),
      0,
    ]
  },
})

/**
 * TrackDelete - Word 修订“删除”标记（红色删除线，内容保留以便接受/拒绝）
 * HTML shape: <span class="docx-track" data-track-type="delete" ...>...</span>
 */
export const TrackDelete = Mark.create({
  name: 'trackDelete',

  addAttributes() {
    return {
      trackId: {
        default: null,
        parseHTML: (el: HTMLElement) => parseStrAttr(el, 'data-track-id'),
        renderHTML: (attrs: TrackAttrs) => (attrs.trackId ? { 'data-track-id': attrs.trackId } : {}),
      },
      author: {
        default: null,
        parseHTML: (el: HTMLElement) => parseStrAttr(el, 'data-track-author'),
        renderHTML: (attrs: TrackAttrs) => (attrs.author ? { 'data-track-author': attrs.author } : {}),
      },
      date: {
        default: null,
        parseHTML: (el: HTMLElement) => parseStrAttr(el, 'data-track-date'),
        renderHTML: (attrs: TrackAttrs) => (attrs.date ? { 'data-track-date': attrs.date } : {}),
      },
    }
  },

  parseHTML() {
    return [
      { tag: 'span.docx-track[data-track-type="delete"]' },
      { tag: 'span[data-track-type="delete"]' },
    ]
  },

  renderHTML({ HTMLAttributes }) {
    return [
      'span',
      mergeAttributes(HTMLAttributes, {
        class: 'docx-track docx-track-delete',
        'data-track-type': 'delete',
      }),
      0,
    ]
  },
})











