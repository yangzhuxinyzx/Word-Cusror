import JSZip from 'jszip'
import mammoth from 'mammoth'

type PageSettings = {
  width: number
  height: number
  marginTop: number
  marginBottom: number
  marginLeft: number
  marginRight: number
  headerHeight: number
  footerHeight: number
  orientation?: 'portrait' | 'landscape'
}

type WorkerRequest = {
  id: string
  arrayBuffer: ArrayBuffer
}

type WorkerResponse =
  | {
      id: string
      ok: true
      html: string
      rawText: string
      pageSettings: PageSettings
      headerText?: string
      footerText?: string
    }
  | {
      id: string
      ok: false
      error: string
    }

function parsePageSettingsFromDocXml(docXml: string): PageSettings {
  // Default A4 in pt
  const defaults: PageSettings = {
    width: 595,
    height: 842,
    marginTop: 72,
    marginBottom: 72,
    marginLeft: 90,
    marginRight: 90,
    headerHeight: 36,
    footerHeight: 36,
    orientation: 'portrait',
  }

  const sectPrMatches = docXml.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/g)
  const sectPr = sectPrMatches?.[sectPrMatches.length - 1]
  if (!sectPr) return defaults

  const pgSzMatch = sectPr.match(/<w:pgSz\b[^>]*?>/i)?.[0] || ''
  const wMatch = pgSzMatch.match(/\bw:w="(\d+)"/i)
  const hMatch = pgSzMatch.match(/\bw:h="(\d+)"/i)
  const orientMatch = pgSzMatch.match(/\bw:orient="(landscape)"/i)
  if (wMatch?.[1]) defaults.width = Math.round(Number(wMatch[1]) / 20)
  if (hMatch?.[1]) defaults.height = Math.round(Number(hMatch[1]) / 20)
  if (orientMatch?.[1] === 'landscape') defaults.orientation = 'landscape'

  const pgMarMatch = sectPr.match(/<w:pgMar\b[^>]*?>/i)?.[0] || ''
  const getMar = (attr: string) => pgMarMatch.match(new RegExp(`\\bw:${attr}="(\\d+)"`, 'i'))?.[1]
  const top = getMar('top')
  const bottom = getMar('bottom')
  const left = getMar('left')
  const right = getMar('right')
  const header = getMar('header')
  const footer = getMar('footer')
  if (top) defaults.marginTop = Math.round(Number(top) / 20)
  if (bottom) defaults.marginBottom = Math.round(Number(bottom) / 20)
  if (left) defaults.marginLeft = Math.round(Number(left) / 20)
  if (right) defaults.marginRight = Math.round(Number(right) / 20)
  if (header) defaults.headerHeight = Math.round(Number(header) / 20)
  if (footer) defaults.footerHeight = Math.round(Number(footer) / 20)

  return defaults
}

function extractHeaderFooterRefs(docXml: string): { headerRid?: string; footerRid?: string } {
  const sectPrMatches = docXml.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/g)
  const sectPr = sectPrMatches?.[sectPrMatches.length - 1]
  if (!sectPr) return {}

  // Prefer default type; fallback to any
  const headerDefault =
    sectPr.match(/<w:headerReference\b[^>]*w:type="default"[^>]*r:id="([^"]+)"/i)?.[1] ||
    sectPr.match(/<w:headerReference\b[^>]*r:id="([^"]+)"/i)?.[1]
  const footerDefault =
    sectPr.match(/<w:footerReference\b[^>]*w:type="default"[^>]*r:id="([^"]+)"/i)?.[1] ||
    sectPr.match(/<w:footerReference\b[^>]*r:id="([^"]+)"/i)?.[1]
  return { headerRid: headerDefault, footerRid: footerDefault }
}

function parseRelsMap(relsXml: string): Record<string, string> {
  const map: Record<string, string> = {}
  // Very lightweight XML parsing via regex
  const relRe = /<Relationship\b[^>]*\bId="([^"]+)"[^>]*\bTarget="([^"]+)"[^>]*\/?>/gi
  let m: RegExpExecArray | null
  while ((m = relRe.exec(relsXml)) !== null) {
    const id = m[1]
    const target = m[2]
    if (!id || !target) continue
    map[id] = target
  }
  return map
}

function extractTextFromWordXml(xml: string): string {
  if (!xml) return ''
  const texts: string[] = []
  const re = /<w:t\b[^>]*>([\s\S]*?)<\/w:t>/gi
  let m: RegExpExecArray | null
  while ((m = re.exec(xml)) !== null) {
    const t = m[1]
    if (t) texts.push(t)
  }
  return texts.join('').replace(/\s+/g, ' ').trim()
}

async function safeMammothHtml(arrayBuffer: ArrayBuffer): Promise<string> {
  const result = await mammoth.convertToHtml(
    { arrayBuffer },
    {
      // CRITICAL: do NOT inline images as base64 (prevents huge strings and UI stalls)
      convertImage: mammoth.images.imgElement(async () => ({ src: 'about:blank' })),
    }
  )
  return result.value || ''
}

async function safeMammothRawText(arrayBuffer: ArrayBuffer): Promise<string> {
  const result = await mammoth.extractRawText({ arrayBuffer })
  return result.value || ''
}

self.onmessage = async (event: MessageEvent<WorkerRequest>) => {
  const { id, arrayBuffer } = event.data
  try {
    // Parse minimal OOXML metadata (page settings + header/footer text) without touching images
    const zip = await JSZip.loadAsync(arrayBuffer)
    const docXml = (await zip.file('word/document.xml')?.async('string')) || ''
    const pageSettings = parsePageSettingsFromDocXml(docXml)

    let headerText: string | undefined
    let footerText: string | undefined
    try {
      const relsXml = (await zip.file('word/_rels/document.xml.rels')?.async('string')) || ''
      const relsMap = parseRelsMap(relsXml)
      const { headerRid, footerRid } = extractHeaderFooterRefs(docXml)

      const resolveTarget = (rid?: string) => {
        if (!rid) return undefined
        const target = relsMap[rid]
        if (!target) return undefined
        if (target.startsWith('/')) return target.slice(1)
        if (target.startsWith('word/')) return target
        if (target.startsWith('../')) return target.replace(/^\.\.\//, '')
        return `word/${target}`
      }

      const headerPath = resolveTarget(headerRid)
      const footerPath = resolveTarget(footerRid)
      if (headerPath) {
        const hx = await zip.file(headerPath)?.async('string')
        if (hx) headerText = extractTextFromWordXml(hx)
      }
      if (footerPath) {
        const fx = await zip.file(footerPath)?.async('string')
        if (fx) footerText = extractTextFromWordXml(fx)
      }
    } catch {
      // ignore header/footer failures
    }

    // Content extraction (no inline images)
    const [html, rawText] = await Promise.all([
      safeMammothHtml(arrayBuffer),
      safeMammothRawText(arrayBuffer),
    ])

    const response: WorkerResponse = {
      id,
      ok: true,
      html,
      rawText,
      pageSettings,
      headerText,
      footerText,
    }
    ;(self as any).postMessage(response)
  } catch (e) {
    const response: WorkerResponse = {
      id,
      ok: false,
      error: (e as Error)?.message || String(e),
    }
    ;(self as any).postMessage(response)
  }
}




