export type ElementAlignment = 'left' | 'center' | 'right' | 'justify'

// Structural match for DocumentContext.tsx's internal FormattedElement interface.
// (We intentionally keep this file decoupled; TS structural typing makes it assignable.)
export type FormattedElementLike = {
  type: 'heading' | 'paragraph' | 'table'
  content?: string
  level?: number
  bold?: boolean
  fontSize?: number
  fontFamily?: string
  alignment?: ElementAlignment
  rows?: number
  cols?: number
  data?: string[][]
}

function normalizeText(s: string): string {
  return (s || '').replace(/\u00A0/g, ' ').replace(/[ \t]+\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim()
}

function parseTextAlignFromStyle(styleText: string | null | undefined): ElementAlignment | undefined {
  if (!styleText) return undefined
  const m = styleText.match(/text-align\s*:\s*(left|center|right|justify)\s*;?/i)
  if (!m?.[1]) return undefined
  const v = m[1].toLowerCase()
  if (v === 'left' || v === 'center' || v === 'right' || v === 'justify') return v
  return undefined
}

function getAlignment(el: Element): ElementAlignment | undefined {
  const style = el.getAttribute('style')
  return parseTextAlignFromStyle(style)
}

function parseHeading(el: Element): FormattedElementLike | null {
  const tag = (el.tagName || '').toLowerCase()
  const m = tag.match(/^h([1-6])$/)
  if (!m) return null
  const level = Number(m[1])
  const text = normalizeText(el.textContent || '')
  if (!text) return null
  return { type: 'heading', level, content: text, alignment: getAlignment(el) }
}

function parseParagraph(el: Element, prefix?: string): FormattedElementLike | null {
  const text = normalizeText(el.textContent || '')
  if (!text) return null
  const content = prefix ? `${prefix}${text}` : text
  return { type: 'paragraph', content, alignment: getAlignment(el) }
}

function parseTable(el: Element): FormattedElementLike | null {
  const rows: string[][] = []
  const trList = Array.from(el.querySelectorAll('tr'))
  let maxCols = 0
  for (const tr of trList) {
    const cells = Array.from(tr.querySelectorAll('th,td'))
    const row: string[] = []
    for (const cell of cells) {
      row.push(normalizeText(cell.textContent || ''))
    }
    maxCols = Math.max(maxCols, row.length)
    // Keep empty rows if table exists; they may be meaningful for layout.
    rows.push(row)
  }
  if (rows.length === 0) return null
  const cols = maxCols || 1
  // Normalize jagged rows
  const data = rows.map((r) => {
    const rr = r.slice(0, cols)
    while (rr.length < cols) rr.push('')
    return rr
  })
  return { type: 'table', rows: data.length, cols, data }
}

function parseList(el: Element, ordered: boolean, depth: number): FormattedElementLike[] {
  const out: FormattedElementLike[] = []
  const items = Array.from(el.children).filter((c) => c.tagName.toLowerCase() === 'li')
  let idx = 1
  for (const li of items) {
    // Extract text excluding nested lists (we will process nested lists separately)
    const cloned = li.cloneNode(true) as Element
    const nestedLists = Array.from(cloned.querySelectorAll('ul,ol'))
    for (const nl of nestedLists) nl.remove()

    const bullet = ordered ? `${idx}. ` : '• '
    const indent = depth > 0 ? '  '.repeat(Math.min(depth, 6)) : ''
    const prefix = `${indent}${bullet}`
    const para = parseParagraph(cloned, prefix)
    if (para) out.push(para)

    // Nested lists
    const nested = Array.from(li.children).filter((c) => {
      const t = c.tagName.toLowerCase()
      return t === 'ul' || t === 'ol'
    })
    for (const nl of nested) {
      out.push(...parseList(nl, nl.tagName.toLowerCase() === 'ol', depth + 1))
    }

    idx++
  }
  return out
}

function walkBlocks(root: Element, out: FormattedElementLike[]) {
  const children = Array.from(root.children)
  for (const el of children) {
    const tag = (el.tagName || '').toLowerCase()

    // headings
    const heading = parseHeading(el)
    if (heading) {
      out.push(heading)
      continue
    }

    // table
    if (tag === 'table') {
      const t = parseTable(el)
      if (t) out.push(t)
      continue
    }

    // list
    if (tag === 'ul' || tag === 'ol') {
      out.push(...parseList(el, tag === 'ol', 0))
      continue
    }

    // paragraph
    if (tag === 'p') {
      const p = parseParagraph(el)
      if (p) out.push(p)
      continue
    }

    // containers / unknown: recurse
    walkBlocks(el, out)
  }
}

export function docxHtmlToElements(html: string): FormattedElementLike[] {
  const out: FormattedElementLike[] = []
  if (!html || !html.trim()) return out

  const parser = new DOMParser()
  const doc = parser.parseFromString(`<div id="__docx_root__">${html}</div>`, 'text/html')
  const root = doc.getElementById('__docx_root__')
  if (!root) return out

  walkBlocks(root, out)

  // Avoid returning empty list if there is non-empty text but no recognizable tags
  if (out.length === 0) {
    const text = normalizeText(root.textContent || '')
    if (text) out.push({ type: 'paragraph', content: text })
  }

  return out
}

export function elementsToHtmlPreview(elements: FormattedElementLike[]): string {
  const esc = (s: string) =>
    (s || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/\"/g, '&quot;')
      .replace(/'/g, '&#39;')

  const parts: string[] = []
  for (const el of elements || []) {
    if (el.type === 'heading' && el.content) {
      const level = Math.min(Math.max(el.level || 1, 1), 6)
      parts.push(`<h${level}>${esc(el.content)}</h${level}>`)
    } else if (el.type === 'paragraph' && el.content) {
      parts.push(`<p>${esc(el.content)}</p>`)
    } else if (el.type === 'table' && el.data && el.data.length) {
      const rows = el.data
      const trs: string[] = []
      for (const r of rows) {
        const tds = (r || []).map((c) => `<td>${esc(c || '')}</td>`).join('')
        trs.push(`<tr>${tds}</tr>`)
      }
      parts.push(`<table><tbody>${trs.join('')}</tbody></table>`)
    }
  }
  return parts.join('\n')
}


