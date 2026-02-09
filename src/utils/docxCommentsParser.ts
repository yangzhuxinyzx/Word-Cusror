import JSZip from 'jszip'

export type DocxComment = {
  id: string
  author?: string
  date?: string
  text: string
  /** replies (best-effort) */
  parentId?: string
}

function base64ToBytes(base64: string): Uint8Array {
  const bin = atob(base64)
  const bytes = new Uint8Array(bin.length)
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i)
  return bytes
}

function extractCommentText(commentEl: Element): string {
  const parts: string[] = []
  const paras = commentEl.getElementsByTagName('w:p')
  if (paras && paras.length > 0) {
    for (let i = 0; i < paras.length; i++) {
      const p = paras[i]
      const runs = p.getElementsByTagName('w:t')
      const line: string[] = []
      for (let j = 0; j < runs.length; j++) {
        line.push(runs[j].textContent || '')
      }
      const s = line.join('')
      if (s.trim()) parts.push(s)
    }
    return parts.join('\n')
  }
  // fallback: all w:t
  const ts = commentEl.getElementsByTagName('w:t')
  const line: string[] = []
  for (let j = 0; j < ts.length; j++) line.push(ts[j].textContent || '')
  return line.join('')
}

async function parseCommentsExtendedParentMap(zip: JSZip): Promise<Record<string, string>> {
  const map: Record<string, string> = {}
  const xml = await zip.file('word/commentsExtended.xml')?.async('string')
  if (!xml) return map

  const parser = new DOMParser()
  const doc = parser.parseFromString(xml, 'application/xml')
  const parseError = doc.querySelector('parsererror')
  if (parseError) return map

  // w15:commentEx w15:paraId="..." w15:paraIdParent="..."
  const all = doc.getElementsByTagName('*')
  for (let i = 0; i < all.length; i++) {
    const el = all[i]
    if (el.localName !== 'commentEx') continue
    const paraId = el.getAttribute('w15:paraId') || el.getAttribute('paraId') || ''
    const parent = el.getAttribute('w15:paraIdParent') || el.getAttribute('paraIdParent') || ''
    if (paraId && parent) {
      map[paraId] = parent
    }
  }
  return map
}

/**
 * Parse DOCX comments (best-effort):
 * - extracts text from word/comments.xml
 * - links replies using commentsExtended.xml (paraIdParent) when available
 */
export async function parseDocxComments(base64Data: string): Promise<DocxComment[]> {
  try {
    const bytes = base64ToBytes(base64Data)
    const ab = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer
    const zip = await JSZip.loadAsync(ab)

    const commentsXml = await zip.file('word/comments.xml')?.async('string')
    if (!commentsXml) return []

    const parser = new DOMParser()
    const doc = parser.parseFromString(commentsXml, 'application/xml')
    const parseError = doc.querySelector('parsererror')
    if (parseError) return []

    const parentMap = await parseCommentsExtendedParentMap(zip)

    // Build paraId -> commentId map (for replies linkage)
    const paraIdToCommentId: Record<string, string> = {}
    const commentNodes = doc.getElementsByTagName('w:comment')
    for (let i = 0; i < commentNodes.length; i++) {
      const c = commentNodes[i]
      const cid = c.getAttribute('w:id') || ''
      const paras = c.getElementsByTagName('w:p')
      let paraId = ''
      for (let j = 0; j < paras.length; j++) {
        const p = paras[j]
        const pid = p.getAttribute('w14:paraId')
        if (pid) {
          paraId = pid
          break
        }
      }
      if (cid && paraId) paraIdToCommentId[paraId] = cid
    }

    const out: DocxComment[] = []
    for (let i = 0; i < commentNodes.length; i++) {
      const c = commentNodes[i]
      const id = c.getAttribute('w:id') || ''
      if (!id) continue
      const author = c.getAttribute('w:author') || undefined
      const date = c.getAttribute('w:date') || undefined
      const text = extractCommentText(c)

      // infer parentId via paraIdParent mapping if possible
      let parentId: string | undefined
      const paras = c.getElementsByTagName('w:p')
      let paraId = ''
      for (let j = 0; j < paras.length; j++) {
        const p = paras[j]
        const pid = p.getAttribute('w14:paraId')
        if (pid) {
          paraId = pid
          break
        }
      }
      if (paraId && parentMap[paraId]) {
        parentId = paraIdToCommentId[parentMap[paraId]]
      }

      out.push({ id, author, date, text, parentId })
    }

    return out
  } catch {
    return []
  }
}











