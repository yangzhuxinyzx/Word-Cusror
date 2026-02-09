import { Image } from '@tiptap/extension-image'

type AnyRecord = Record<string, any>

export const DocxImage = Image.extend({
  name: 'docxImage',

  addOptions() {
    const parent = this.parent?.()
    return {
      ...(parent || {}),
      // ImageOptions requires HTMLAttributes; keep it always defined
      HTMLAttributes: parent?.HTMLAttributes || {},
      // ImageOptions.resize does not allow undefined
      resize: parent?.resize ?? false,
      // Allow data: URLs (fallback) – blob: URLs are allowed by default
      allowBase64: true,
      // Make images inline so they can stay in the same paragraph flow (Word-like)
      inline: true,
    }
  },

  addAttributes() {
    return {
      ...this.parent?.(),

      widthPx: {
        default: null,
        parseHTML: (element: HTMLElement) => {
          const w = element.getAttribute('data-w')
          return w ? Number(w) || null : null
        },
        renderHTML: (attributes: AnyRecord) => {
          if (!attributes.widthPx) return {}
          return { 'data-w': String(attributes.widthPx) }
        },
      },

      heightPx: {
        default: null,
        parseHTML: (element: HTMLElement) => {
          const h = element.getAttribute('data-h')
          return h ? Number(h) || null : null
        },
        renderHTML: (attributes: AnyRecord) => {
          if (!attributes.heightPx) return {}
          return { 'data-h': String(attributes.heightPx) }
        },
      },

      rid: {
        default: null,
        parseHTML: (element: HTMLElement) => element.getAttribute('data-rid'),
        renderHTML: (attributes: AnyRecord) => (attributes.rid ? { 'data-rid': String(attributes.rid) } : {}),
      },

      target: {
        default: null,
        parseHTML: (element: HTMLElement) => element.getAttribute('data-target'),
        renderHTML: (attributes: AnyRecord) => (attributes.target ? { 'data-target': String(attributes.target) } : {}),
      },

      floating: {
        default: null,
        parseHTML: (element: HTMLElement) => {
          const v = element.getAttribute('data-floating')
          return v === '1' ? 1 : null
        },
        renderHTML: (attributes: AnyRecord) => (attributes.floating ? { 'data-floating': '1' } : {}),
      },

      relH: {
        default: null,
        parseHTML: (element: HTMLElement) => element.getAttribute('data-rel-h'),
        renderHTML: (attributes: AnyRecord) => (attributes.relH ? { 'data-rel-h': String(attributes.relH) } : {}),
      },

      relV: {
        default: null,
        parseHTML: (element: HTMLElement) => element.getAttribute('data-rel-v'),
        renderHTML: (attributes: AnyRecord) => (attributes.relV ? { 'data-rel-v': String(attributes.relV) } : {}),
      },

      xEmu: {
        default: null,
        parseHTML: (element: HTMLElement) => {
          const v = element.getAttribute('data-x-emu')
          return v ? Number(v) || null : null
        },
        renderHTML: (attributes: AnyRecord) => (attributes.xEmu != null ? { 'data-x-emu': String(attributes.xEmu) } : {}),
      },

      yEmu: {
        default: null,
        parseHTML: (element: HTMLElement) => {
          const v = element.getAttribute('data-y-emu')
          return v ? Number(v) || null : null
        },
        renderHTML: (attributes: AnyRecord) => (attributes.yEmu != null ? { 'data-y-emu': String(attributes.yEmu) } : {}),
      },

      distL: {
        default: null,
        parseHTML: (element: HTMLElement) => element.getAttribute('data-dist-l'),
        renderHTML: (attributes: AnyRecord) => (attributes.distL ? { 'data-dist-l': String(attributes.distL) } : {}),
      },
      distR: {
        default: null,
        parseHTML: (element: HTMLElement) => element.getAttribute('data-dist-r'),
        renderHTML: (attributes: AnyRecord) => (attributes.distR ? { 'data-dist-r': String(attributes.distR) } : {}),
      },
      distT: {
        default: null,
        parseHTML: (element: HTMLElement) => element.getAttribute('data-dist-t'),
        renderHTML: (attributes: AnyRecord) => (attributes.distT ? { 'data-dist-t': String(attributes.distT) } : {}),
      },
      distB: {
        default: null,
        parseHTML: (element: HTMLElement) => element.getAttribute('data-dist-b'),
        renderHTML: (attributes: AnyRecord) => (attributes.distB ? { 'data-dist-b': String(attributes.distB) } : {}),
      },

      wrap: {
        default: null,
        parseHTML: (element: HTMLElement) => element.getAttribute('data-wrap'),
        renderHTML: (attributes: AnyRecord) => (attributes.wrap ? { 'data-wrap': String(attributes.wrap) } : {}),
      },
    }
  },

  renderHTML({ HTMLAttributes }) {
    // Ensure width/height are applied for layout stability (when present)
    const attrs = { ...HTMLAttributes } as AnyRecord
    const width = attrs['data-w'] ? Number(attrs['data-w']) : null
    const height = attrs['data-h'] ? Number(attrs['data-h']) : null
    const floating = attrs['data-floating'] === '1'

    const styleParts: string[] = []
    if (typeof attrs.style === 'string' && attrs.style.trim()) {
      styleParts.push(attrs.style.trim().replace(/;$/, ''))
    }
    if (width && width > 0) styleParts.push(`width:${width}px`)
    if (height && height > 0) styleParts.push(`height:${height}px`)
    if (floating) {
      // Hide until WordEditor positions it (avoid flashing at wrong place)
      styleParts.push('position:absolute', 'left:0', 'top:0', 'visibility:hidden', 'z-index:5')
    } else {
      styleParts.push('display:inline-block', 'vertical-align:baseline')
    }

    if (styleParts.length) attrs.style = styleParts.join('; ')

    return ['img', attrs]
  },
})


