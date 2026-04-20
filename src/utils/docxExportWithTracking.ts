import JSZip from 'jszip'

type ExportComment = {
  id: string
  author?: string
  date?: string
  text: string
}

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

function b64uDecodeUtf8(input: string): string {
  try {
    if (!input) return ''
    let b64 = input.replace(/-/g, '+').replace(/_/g, '/')
    while (b64.length % 4 !== 0) b64 += '='
    return decodeURIComponent(escape(atob(b64)))
  } catch {
    return ''
  }
}

function parseMarker(raw: string): { kind: string; kv: Record<string, string> } | null {
  const s = raw.trim()
  if (!s.startsWith('[[[WC_') || !s.endsWith(']]]')) return null
  const body = s.slice(3, -3) // remove [[[ and ]]]
  // body example: WC_TC_START|t=ins|id=1|a=..|d=..
  const parts = body.split('|').filter(Boolean)
  const kind = parts.shift() || ''
  const kv: Record<string, string> = {}
  for (const p of parts) {
    const idx = p.indexOf('=')
    if (idx <= 0) continue
    const k = p.slice(0, idx).trim()
    const v = p.slice(idx + 1).trim()
    if (k) kv[k] = v
  }
  return { kind, kv }
}

function getRunMarkerText(runEl: Element): string | null {
  const ts = runEl.getElementsByTagName('w:t')
  for (let i = 0; i < ts.length; i++) {
    const t = ts[i]
    const txt = (t.textContent || '').trim()
    if (txt.startsWith('[[[WC_') && txt.endsWith(']]]')) return txt
  }
  return null
}

function createW(doc: Document, local: string) {
  return doc.createElementNS(W_NS, `w:${local}`)
}

function renameTextNodesToDelText(doc: Document, root: Element) {
  const ts = Array.from(root.getElementsByTagName('w:t'))
  for (const t of ts) {
    const delText = createW(doc, 'delText')
    // copy attributes (xml:space etc)
    for (let i = 0; i < t.attributes.length; i++) {
      const a = t.attributes[i]
      delText.setAttribute(a.name, a.value)
    }
    while (t.firstChild) delText.appendChild(t.firstChild)
    t.parentNode?.replaceChild(delText, t)
  }
}

function ensureCommentsRelationship(doc: Document, relsXml: string): string {
  try {
    const parser = new DOMParser()
    const relsDoc = parser.parseFromString(relsXml, 'application/xml')
    const parseError = relsDoc.querySelector('parsererror')
    if (parseError) return relsXml

    const relsRoot = relsDoc.documentElement
    const existing = Array.from(relsDoc.getElementsByTagName('Relationship')).find((r) => {
      return (r.getAttribute('Type') || '').includes('/comments')
    })
    if (existing) return new XMLSerializer().serializeToString(relsDoc)

    let maxRid = 0
    Array.from(relsDoc.getElementsByTagName('Relationship')).forEach((r) => {
      const id = r.getAttribute('Id') || ''
      const m = id.match(/^rId(\d+)$/)
      if (m) maxRid = Math.max(maxRid, Number(m[1]))
    })
    const nextId = `rId${maxRid + 1}`
    const rel = relsDoc.createElement('Relationship')
    rel.setAttribute('Id', nextId)
    rel.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments')
    rel.setAttribute('Target', 'comments.xml')
    relsRoot.appendChild(rel)
    return new XMLSerializer().serializeToString(relsDoc)
  } catch {
    return relsXml
  }
}

function ensureCommentsContentType(ctXml: string): string {
  try {
    const parser = new DOMParser()
    const ctDoc = parser.parseFromString(ctXml, 'application/xml')
    const parseError = ctDoc.querySelector('parsererror')
    if (parseError) return ctXml
    const has = Array.from(ctDoc.getElementsByTagName('Override')).some((o) => o.getAttribute('PartName') === '/word/comments.xml')
    if (has) return new XMLSerializer().serializeToString(ctDoc)
    const root = ctDoc.documentElement
    const ov = ctDoc.createElement('Override')
    ov.setAttribute('PartName', '/word/comments.xml')
    ov.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml')
    root.appendChild(ov)
    return new XMLSerializer().serializeToString(ctDoc)
  } catch {
    return ctXml
  }
}

function ensureTrackRevisions(settingsXml: string): string {
  try {
    const parser = new DOMParser()
    const sDoc = parser.parseFromString(settingsXml, 'application/xml')
    const parseError = sDoc.querySelector('parsererror')
    if (parseError) return settingsXml
    const exists = Array.from(sDoc.getElementsByTagName('w:trackRevisions')).length > 0
    if (exists) return new XMLSerializer().serializeToString(sDoc)
    const root = sDoc.getElementsByTagName('w:settings')[0] || sDoc.documentElement
    const tr = sDoc.createElementNS(W_NS, 'w:trackRevisions')
    // insert near top
    root.insertBefore(tr, root.firstChild)
    return new XMLSerializer().serializeToString(sDoc)
  } catch {
    return settingsXml
  }
}

function buildCommentsXml(comments: ExportComment[], usedIds: Set<string>): string {
  const byId = new Map<string, ExportComment>()
  for (const c of comments) byId.set(String(c.id), c)

  const ids = Array.from(usedIds).sort((a, b) => Number(a) - Number(b))
  const esc = (t: string) =>
    (t || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;')

  const parts: string[] = []
  parts.push(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
  parts.push(`<w:comments xmlns:w="${W_NS}">`)
  for (const id of ids) {
    const c = byId.get(id)
    const author = esc(c?.author || 'User')
    const date = esc(c?.date || new Date().toISOString())
    const text = (c?.text || '').trim()
    const lines = text ? text.split(/\r?\n/) : ['']
    parts.push(`<w:comment w:id="${esc(id)}" w:author="${author}" w:date="${date}">`)
    for (const line of lines) {
      parts.push(`<w:p><w:r><w:t>${esc(line)}</w:t></w:r></w:p>`)
    }
    parts.push(`</w:comment>`)
  }
  parts.push(`</w:comments>`)
  return parts.join('')
}

/**
 * Post-process a docx generated by our `docx` library:
 * - Replaces sentinel runs `[[[WC_*]]]` with real OOXML tracked changes / comment ranges.
 * - Adds `word/comments.xml` + rels + content types when comments are present.
 */
export async function postProcessDocxWithAnnotations(
  baseDocxArrayBuffer: ArrayBuffer,
  options: { comments: ExportComment[] }
): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(baseDocxArrayBuffer)
  const documentXml = await zip.file('word/document.xml')?.async('string')
  if (!documentXml) return baseDocxArrayBuffer

  const parser = new DOMParser()
  const doc = parser.parseFromString(documentXml, 'application/xml')
  const parseError = doc.querySelector('parsererror')
  if (parseError) return baseDocxArrayBuffer

  const usedCommentIds = new Set<string>()
  let hasTrackChanges = false

  const paragraphs = Array.from(doc.getElementsByTagName('w:p'))
  for (const p of paragraphs) {
    const nodes = Array.from(p.childNodes).filter((n) => n.nodeType === Node.ELEMENT_NODE) as Element[]

    // pass 1: comments markers (start/end)
    for (let i = 0; i < nodes.length; i++) {
      const el = nodes[i]
      if (el.tagName !== 'w:r') continue
      const markerText = getRunMarkerText(el)
      if (!markerText) continue
      const m = parseMarker(markerText)
      if (!m) continue

      if (m.kind === 'WC_CM_START') {
        const id = m.kv['id'] || ''
        if (id) usedCommentIds.add(id)
        const start = createW(doc, 'commentRangeStart')
        if (id) start.setAttribute('w:id', id)
        p.insertBefore(start, el)
        p.removeChild(el)
      } else if (m.kind === 'WC_CM_END') {
        const id = m.kv['id'] || ''
        if (id) usedCommentIds.add(id)
        const end = createW(doc, 'commentRangeEnd')
        if (id) end.setAttribute('w:id', id)
        p.insertBefore(end, el)

        // add comment reference run
        const r = createW(doc, 'r')
        const rPr = createW(doc, 'rPr')
        const rStyle = createW(doc, 'rStyle')
        rStyle.setAttribute('w:val', 'CommentReference')
        rPr.appendChild(rStyle)
        const ref = createW(doc, 'commentReference')
        if (id) ref.setAttribute('w:id', id)
        r.appendChild(rPr)
        r.appendChild(ref)
        p.insertBefore(r, el)

        p.removeChild(el)
      }
    }

    // refresh node list after comment pass
    const nodes2 = Array.from(p.childNodes).filter((n) => n.nodeType === Node.ELEMENT_NODE) as Element[]

    // pass 2: track change wrappers
    for (let i = 0; i < nodes2.length; i++) {
      const el = nodes2[i]
      if (el.tagName !== 'w:r') continue
      const markerText = getRunMarkerText(el)
      if (!markerText) continue
      const m = parseMarker(markerText)
      if (!m) continue

      if (m.kind === 'WC_TC_START') {
        const t = (m.kv['t'] || 'ins').toLowerCase()
        const id = m.kv['id'] || '0'
        const author = b64uDecodeUtf8(m.kv['a'] || '') || 'User'
        const date = b64uDecodeUtf8(m.kv['d'] || '') || new Date().toISOString()

        // find end marker
        let endIdx = -1
        for (let j = i + 1; j < nodes2.length; j++) {
          const r = nodes2[j]
          if (r.tagName !== 'w:r') continue
          const mt = getRunMarkerText(r)
          if (mt && parseMarker(mt)?.kind === 'WC_TC_END') {
            endIdx = j
            break
          }
        }
        if (endIdx === -1) continue

        const wrapper = createW(doc, t === 'del' ? 'del' : 'ins')
        wrapper.setAttribute('w:id', id)
        wrapper.setAttribute('w:author', author)
        wrapper.setAttribute('w:date', date)

        // move nodes between start/end into wrapper (exclude markers)
        const toMove: Element[] = []
        for (let j = i + 1; j < endIdx; j++) {
          const n = nodes2[j]
          if (n && n.parentNode === p) toMove.push(n)
        }
        for (const n of toMove) {
          wrapper.appendChild(n)
        }

        // convert deletion text nodes
        if (t === 'del') {
          renameTextNodesToDelText(doc, wrapper)
        }

        // insert wrapper at start position
        p.insertBefore(wrapper, el)

        // remove start marker run
        if (el.parentNode === p) p.removeChild(el)
        // remove end marker run
        const endRun = nodes2[endIdx]
        if (endRun && endRun.parentNode === p) p.removeChild(endRun)

        hasTrackChanges = true
      }
    }
  }

  const outDocXml = new XMLSerializer().serializeToString(doc)
  zip.file('word/document.xml', outDocXml)

  // comments.xml + rels + content types
  if (usedCommentIds.size > 0) {
    const commentsXml = buildCommentsXml(options.comments || [], usedCommentIds)
    zip.file('word/comments.xml', commentsXml)

    const relsPath = 'word/_rels/document.xml.rels'
    const relsXml = await zip.file(relsPath)?.async('string')
    if (relsXml) {
      zip.file(relsPath, ensureCommentsRelationship(doc, relsXml))
    }

    const ctPath = '[Content_Types].xml'
    const ctXml = await zip.file(ctPath)?.async('string')
    if (ctXml) {
      zip.file(ctPath, ensureCommentsContentType(ctXml))
    }
  }

  if (hasTrackChanges) {
    const settingsPath = 'word/settings.xml'
    const settingsXml = await zip.file(settingsPath)?.async('string')
    if (settingsXml) {
      zip.file(settingsPath, ensureTrackRevisions(settingsXml))
    }
  }

  return await zip.generateAsync({ type: 'arraybuffer' })
}











