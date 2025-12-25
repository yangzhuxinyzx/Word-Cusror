import { Extension } from '@tiptap/core'

/**
 * DiffBlock 扩展
 * - 让 paragraph/heading 节点能持久化 data-diff-id / data-diff-role 等属性
 * - 用于“段落/样式”级别的修订（块级 old/new 两份并存）
 */
export const DiffBlock = Extension.create({
  name: 'diffBlock',

  addGlobalAttributes() {
    return [
      {
        types: ['paragraph', 'heading'],
        attributes: {
          diffId: {
            default: null,
            parseHTML: (element) => element.getAttribute('data-diff-id'),
            renderHTML: (attributes) => {
              if (!attributes.diffId) return {}
              return { 'data-diff-id': attributes.diffId }
            },
          },
          diffRole: {
            default: null,
            parseHTML: (element) => element.getAttribute('data-diff-role'),
            renderHTML: (attributes) => {
              if (!attributes.diffRole) return {}
              return { 'data-diff-role': attributes.diffRole }
            },
          },
          diffKind: {
            default: null,
            parseHTML: (element) => element.getAttribute('data-diff-kind'),
            renderHTML: (attributes) => {
              if (!attributes.diffKind) return {}
              return { 'data-diff-kind': attributes.diffKind }
            },
          },
        },
      },
    ]
  },
})



