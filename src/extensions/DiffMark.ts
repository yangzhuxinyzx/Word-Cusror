import { Mark, mergeAttributes } from '@tiptap/core'

/**
 * DiffOld 标记 - 用于显示被删除的旧内容（红色删除线）
 */
export const DiffOld = Mark.create({
  name: 'diffOld',

  // 定义标记的属性
  addAttributes() {
    return {
      'data-diff-id': {
        default: null,
        parseHTML: element => element.getAttribute('data-diff-id'),
        renderHTML: attributes => {
          if (!attributes['data-diff-id']) {
            return {}
          }
          return { 'data-diff-id': attributes['data-diff-id'] }
        },
      },
    }
  },

  // 定义如何从 HTML 解析这个标记
  parseHTML() {
    return [
      {
        tag: 'span.diff-old',
      },
      {
        tag: 'span[class="diff-old"]',
      },
      {
        style: 'text-decoration: line-through; background-color: #fecaca',
      },
    ]
  },

  // 定义如何将这个标记渲染为 HTML
  renderHTML({ HTMLAttributes }) {
    return [
      'span',
      mergeAttributes(HTMLAttributes, {
        class: 'diff-old',
        style: 'background-color: #fecaca; color: #b91c1c; text-decoration: line-through; padding: 1px 2px; border-radius: 2px;',
      }),
      0, // 0 表示内容将被放在这里
    ]
  },
})

/**
 * DiffNew 标记 - 用于显示新增的内容（绿色高亮）
 */
export const DiffNew = Mark.create({
  name: 'diffNew',

  // 定义标记的属性
  addAttributes() {
    return {
      'data-diff-id': {
        default: null,
        parseHTML: element => element.getAttribute('data-diff-id'),
        renderHTML: attributes => {
          if (!attributes['data-diff-id']) {
            return {}
          }
          return { 'data-diff-id': attributes['data-diff-id'] }
        },
      },
    }
  },

  // 定义如何从 HTML 解析这个标记
  parseHTML() {
    return [
      {
        tag: 'span.diff-new',
      },
      {
        tag: 'span[class="diff-new"]',
      },
      {
        style: 'background-color: #bbf7d0',
      },
    ]
  },

  // 定义如何将这个标记渲染为 HTML
  renderHTML({ HTMLAttributes }) {
    return [
      'span',
      mergeAttributes(HTMLAttributes, {
        class: 'diff-new',
        style: 'background-color: #bbf7d0; color: #15803d; padding: 1px 2px; border-radius: 2px;',
      }),
      0,
    ]
  },
})




