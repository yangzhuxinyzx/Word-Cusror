import JSZip from 'jszip'

export type DocxLineRule = 'auto' | 'exact' | 'atLeast'

export interface DocxTypographyProfile {
  page?: {
    size?: { wTwips?: number; hTwips?: number; orientation?: 'portrait' | 'landscape' }
    margin?: { topTwips?: number; rightTwips?: number; bottomTwips?: number; leftTwips?: number }
  }
  heading1?: DocxTypographyProfile['normal']
  heading2?: DocxTypographyProfile['normal']
  heading3?: DocxTypographyProfile['normal']
  normal?: {
    fontAscii?: string
    fontEastAsia?: string
    fontHAnsi?: string
    fontSizeHalfPoints?: number
    alignment?: 'left' | 'center' | 'right' | 'justify'
    spacing?: {
      beforeTwips?: number
      afterTwips?: number
      lineTwips?: number
      lineRule?: DocxLineRule
    }
    indent?: {
      firstLineTwips?: number
      leftTwips?: number
      rightTwips?: number
    }
  }
}

export interface DocxOutlineStats {
  heading1Count: number
  heading2Count: number
  heading3Count: number
  tableCount: number
  imageCount: number
}

type NormalStyle = NonNullable<DocxTypographyProfile['normal']>
type Alignment = 'left' | 'center' | 'right' | 'justify'

type ThemeFontGroup = {
  latin?: string
  ea?: string
  cs?: string
  script?: Record<string, string> // e.g. Hans/Hant/Jpan/Kore
}

type ThemeFonts = {
  major: ThemeFontGroup
  minor: ThemeFontGroup
}

function safeNum(v: string | null | undefined): number | undefined {
  if (v == null) return undefined
  const n = Number(v)
  return Number.isFinite(n) ? n : undefined
}

function charsHundredthToTwips(charsHundredth: number, fontSizeHalfPoints?: number): number {
  const hps = typeof fontSizeHalfPoints === 'number' && Number.isFinite(fontSizeHalfPoints) && fontSizeHalfPoints > 0
    ? fontSizeHalfPoints
    : 24 // 12pt default
  // 1 char ~= 1em ~= fontSizePt; 1pt = 20 twips; fontSizePt = hps/2
  // twipsPerChar = (hps/2)*20 = hps*10
  // charsHundredth is 1/100 char
  // twips = (charsHundredth/100) * (hps*10) = charsHundredth*hps/10
  return Math.round((charsHundredth * hps) / 10)
}

function mapJc(v?: string | null): Alignment {
  const val = (v || '').toLowerCase()
  if (val === 'center') return 'center'
  if (val === 'right') return 'right'
  if (val === 'both' || val === 'justify') return 'justify'
  return 'left'
}

function mapLineRule(v?: string | null): DocxLineRule | undefined {
  const val = (v || '').toLowerCase()
  if (val === 'exact') return 'exact'
  if (val === 'atleast') return 'atLeast'
  if (val === 'auto') return 'auto'
  return undefined
}

function parseXml(xml: string): Document | null {
  try {
    const parser = new DOMParser()
    return parser.parseFromString(xml, 'application/xml')
  } catch {
    return null
  }
}

function getAttr(el: Element, names: string[]): string | undefined {
  for (const n of names) {
    const v = el.getAttribute(n)
    if (v) return v
  }
  return undefined
}

function getFirstByLocalName(root: ParentNode, localName: string): Element | null {
  const all = (root as any).getElementsByTagName?.('*') as HTMLCollectionOf<Element> | undefined
  if (!all) return null
  for (let i = 0; i < all.length; i++) {
    const el = all.item(i)
    if (!el) continue
    if ((el as any).localName === localName) return el
  }
  return null
}

function findNormalStyle(doc: Document): Element | null {
  const styles = Array.from(doc.getElementsByTagName('w:style'))
  // Prefer explicit Normal
  const byId = styles.find((s) => (s.getAttribute('w:styleId') || '').toLowerCase() === 'normal')
  if (byId) return byId
  // Fallback: default paragraph style
  const byDefault = styles.find((s) => s.getAttribute('w:type') === 'paragraph' && s.getAttribute('w:default') === '1')
  return byDefault || null
}

function extractThemeFonts(themeXml: string): ThemeFonts | undefined {
  const doc = parseXml(themeXml)
  if (!doc) return undefined
  const fontScheme = getFirstByLocalName(doc, 'fontScheme')
  if (!fontScheme) return undefined

  const majorFont = getFirstByLocalName(fontScheme, 'majorFont')
  const minorFont = getFirstByLocalName(fontScheme, 'minorFont')

  const pick = (fontEl: Element | null): ThemeFontGroup => {
    if (!fontEl) return {}
    const latin = getFirstByLocalName(fontEl, 'latin')?.getAttribute('typeface') || undefined
    const ea = getFirstByLocalName(fontEl, 'ea')?.getAttribute('typeface') || undefined
    const cs = getFirstByLocalName(fontEl, 'cs')?.getAttribute('typeface') || undefined
    // Script-specific fonts: <a:font script="Hans" typeface="..."/>
    const script: Record<string, string> = {}
    const all = (fontEl as any).getElementsByTagName?.('*') as HTMLCollectionOf<Element> | undefined
    if (all) {
      for (let i = 0; i < all.length; i++) {
        const el = all.item(i)
        if (!el) continue
        if ((el as any).localName !== 'font') continue
        const scr = el.getAttribute('script')
        const tf = el.getAttribute('typeface')
        if (scr && tf) script[scr] = tf
      }
    }
    return { latin, ea, cs, script: Object.keys(script).length ? script : undefined }
  }

  return { major: pick(majorFont), minor: pick(minorFont) }
}

function resolveThemeFont(themeKey: string | undefined, themeFonts?: ThemeFonts): string | undefined {
  if (!themeKey || !themeFonts) return undefined
  const fontMap = buildThemeFontMap(themeFonts)
  return checkThemeFont(themeKey, fontMap)
}

type ThemeFontMap = Record<string, string>

function pickEastAsiaFont(group: ThemeFontGroup): string | undefined {
  // theme can store ea as "+mn-ea/+mj-ea" placeholder; real font is often in script fonts (Hans/Hant...)
  const ea = group.ea
  if (ea && !ea.startsWith('+')) return ea
  return group.script?.Hans || group.script?.Hant || ea || group.latin
}

function buildThemeFontMap(themeFonts: ThemeFonts): ThemeFontMap {
  const map: ThemeFontMap = {}

  const majorLatin = themeFonts.major.latin || 'Arial'
  const minorLatin = themeFonts.minor.latin || majorLatin || 'Arial'

  const majorEA = pickEastAsiaFont(themeFonts.major) || majorLatin
  const minorEA = pickEastAsiaFont(themeFonts.minor) || minorLatin

  const majorCS = themeFonts.major.cs || majorLatin
  const minorCS = themeFonts.minor.cs || minorLatin

  // Mimic OnlyOffice FontScheme.checkFromFontCollection mappings (see DocumentServer/sdkjs-src/common/Drawings/Format/Format.js)
  map['+mj-lt'] = majorLatin
  map['majorAscii'] = majorLatin
  map['majorHAnsi'] = majorLatin

  map['+mj-ea'] = majorEA
  map['majorEastAsia'] = majorEA

  map['+mj-cs'] = majorCS
  map['majorBidi'] = majorCS

  map['+mn-lt'] = minorLatin
  map['minorAscii'] = minorLatin
  map['minorHAnsi'] = minorLatin

  map['+mn-ea'] = minorEA
  map['minorEastAsia'] = minorEA

  map['+mn-cs'] = minorCS
  map['minorBidi'] = minorCS

  return map
}

const THEME_FONT_TOKENS = new Set([
  '+mj-lt', '+mj-ea', '+mj-cs',
  '+mn-lt', '+mn-ea', '+mn-cs',
  'majorAscii', 'majorHAnsi', 'majorEastAsia', 'majorBidi',
  'minorAscii', 'minorHAnsi', 'minorEastAsia', 'minorBidi',
])

function checkThemeFont(font: string | undefined, fontMap: ThemeFontMap): string | undefined {
  if (!font) return undefined
  const key = font
  if (THEME_FONT_TOKENS.has(key)) {
    return fontMap[key] || fontMap['+mn-lt'] || 'Arial'
  }
  return font
}

function extractRFontsFromRPr(rPr: Element | undefined | null, themeFonts?: ThemeFonts): {
  ascii?: string
  hAnsi?: string
  eastAsia?: string
  asciiTheme?: string
  hAnsiTheme?: string
  eastAsiaTheme?: string
} {
  if (!rPr) return {}
  const rFonts =
    rPr.getElementsByTagName('w:rFonts')[0] ||
    (getFirstByLocalName(rPr, 'rFonts') as Element | null) ||
    undefined

  if (!rFonts) return {}

  const fontMap = themeFonts ? buildThemeFontMap(themeFonts) : {}

  const asciiRaw = getAttr(rFonts, ['w:ascii', 'ascii'])
  const hAnsiRaw = getAttr(rFonts, ['w:hAnsi', 'hAnsi'])
  const eastAsiaRaw = getAttr(rFonts, ['w:eastAsia', 'eastAsia'])

  const asciiTheme = getAttr(rFonts, ['w:asciiTheme', 'asciiTheme'])
  const hAnsiTheme = getAttr(rFonts, ['w:hAnsiTheme', 'hAnsiTheme'])
  const eastAsiaTheme = getAttr(rFonts, ['w:eastAsiaTheme', 'eastAsiaTheme'])

  // Resolve theme -> actual font if direct font absent
  return {
    ascii: checkThemeFont(asciiRaw, fontMap) || resolveThemeFont(asciiTheme, themeFonts),
    hAnsi: checkThemeFont(hAnsiRaw, fontMap) || resolveThemeFont(hAnsiTheme, themeFonts),
    eastAsia: checkThemeFont(eastAsiaRaw, fontMap) || resolveThemeFont(eastAsiaTheme, themeFonts),
    asciiTheme,
    hAnsiTheme,
    eastAsiaTheme,
  }
}

function extractDocDefaultsNormal(stylesDoc: Document, themeFonts?: ThemeFonts): Partial<NormalStyle> {
  const result: Partial<NormalStyle> = {}

  const docDefaults = stylesDoc.getElementsByTagName('w:docDefaults')[0] || (getFirstByLocalName(stylesDoc, 'docDefaults') as Element | null) || undefined
  if (!docDefaults) return result

  const rPrDefault =
    docDefaults.getElementsByTagName('w:rPrDefault')[0] ||
    (getFirstByLocalName(docDefaults, 'rPrDefault') as Element | null) ||
    undefined
  const rPr =
    rPrDefault?.getElementsByTagName('w:rPr')[0] ||
    (rPrDefault ? (getFirstByLocalName(rPrDefault, 'rPr') as Element | null) : null) ||
    undefined

  const rf = extractRFontsFromRPr(rPr, themeFonts)
  if (rf.ascii) result.fontAscii = rf.ascii
  if (rf.hAnsi) result.fontHAnsi = rf.hAnsi
  if (rf.eastAsia) result.fontEastAsia = rf.eastAsia

  const sz = rPr?.getElementsByTagName('w:sz')[0] || (rPr ? (getFirstByLocalName(rPr, 'sz') as Element | null) : null) || undefined
  const szVal = sz?.getAttribute('w:val') || sz?.getAttribute('val') || undefined
  const sizeHalfPoints = safeNum(szVal)
  if (sizeHalfPoints) result.fontSizeHalfPoints = sizeHalfPoints

  return result
}

function findStyleByIds(doc: Document, styleIds: string[]): Element | null {
  const styles = Array.from(doc.getElementsByTagName('w:style'))
  const lowered = styleIds.map((s) => s.toLowerCase())
  const found = styles.find((s) => lowered.includes((s.getAttribute('w:styleId') || '').toLowerCase()))
  if (found) return found
  // fallback: match on w:name/@w:val (local name "name" attr "val")
  for (const st of styles) {
    const nameEl = st.getElementsByTagName('w:name')[0] || (getFirstByLocalName(st, 'name') as Element | null) || undefined
    const nm = nameEl?.getAttribute('w:val') || nameEl?.getAttribute('val') || ''
    if (nm && lowered.includes(nm.toLowerCase())) return st
  }
  return null
}

function extractParagraphStyleFromStyleEl(styleEl: Element, stylesDoc: Document, themeFonts?: ThemeFonts): NormalStyle {
  const pPr = styleEl.getElementsByTagName('w:pPr')[0]
  const rPr = styleEl.getElementsByTagName('w:rPr')[0]
  const out: NormalStyle = {}

  if (rPr) {
    const rf = extractRFontsFromRPr(rPr, themeFonts)
    out.fontAscii = rf.ascii
    out.fontHAnsi = rf.hAnsi
    out.fontEastAsia = rf.eastAsia

    const sz = rPr.getElementsByTagName('w:sz')[0] || (getFirstByLocalName(rPr, 'sz') as Element | null) || undefined
    const szVal = sz?.getAttribute('w:val') || sz?.getAttribute('val') || undefined
    const sizeHalfPoints = safeNum(szVal)
    if (sizeHalfPoints) out.fontSizeHalfPoints = sizeHalfPoints
  }

  if (pPr) {
    const jc = pPr.getElementsByTagName('w:jc')[0]
    out.alignment = mapJc(jc?.getAttribute('w:val'))

    const spacing = pPr.getElementsByTagName('w:spacing')[0]
    if (spacing) {
      out.spacing = {
        beforeTwips: safeNum(spacing.getAttribute('w:before')),
        afterTwips: safeNum(spacing.getAttribute('w:after')),
        lineTwips: safeNum(spacing.getAttribute('w:line')),
        lineRule: mapLineRule(spacing.getAttribute('w:lineRule')),
      }
    }

    const ind = pPr.getElementsByTagName('w:ind')[0]
    if (ind) {
      const firstLine = safeNum(ind.getAttribute('w:firstLine'))
      const left = safeNum(ind.getAttribute('w:left') || ind.getAttribute('w:start'))
      const right = safeNum(ind.getAttribute('w:right') || ind.getAttribute('w:end'))

      const firstLineChars = safeNum(ind.getAttribute('w:firstLineChars'))
      const leftChars = safeNum(ind.getAttribute('w:leftChars') || ind.getAttribute('w:startChars'))
      const rightChars = safeNum(ind.getAttribute('w:rightChars') || ind.getAttribute('w:endChars'))

      const docDefaults = extractDocDefaultsNormal(stylesDoc, themeFonts)
      const sizeForIndent = out.fontSizeHalfPoints || docDefaults.fontSizeHalfPoints || 24

      out.indent = {
        firstLineTwips:
          firstLine ??
          (typeof firstLineChars === 'number' ? charsHundredthToTwips(firstLineChars, sizeForIndent) : undefined),
        leftTwips:
          left ??
          (typeof leftChars === 'number' ? charsHundredthToTwips(leftChars, sizeForIndent) : undefined),
        rightTwips:
          right ??
          (typeof rightChars === 'number' ? charsHundredthToTwips(rightChars, sizeForIndent) : undefined),
      }
    }
  }

  // fallback to docDefaults if still empty
  if (!out.fontAscii && !out.fontHAnsi && !out.fontEastAsia) {
    const d = extractDocDefaultsNormal(stylesDoc, themeFonts)
    out.fontAscii = d.fontAscii || out.fontAscii
    out.fontHAnsi = d.fontHAnsi || out.fontHAnsi
    out.fontEastAsia = d.fontEastAsia || out.fontEastAsia
    out.fontSizeHalfPoints = d.fontSizeHalfPoints || out.fontSizeHalfPoints
  }

  return out
}

function extractNormalFromStyles(stylesXml: string, themeXml?: string): NormalStyle | undefined {
  const doc = parseXml(stylesXml)
  if (!doc) return undefined
  const themeFonts = themeXml ? extractThemeFonts(themeXml) : undefined

  const normalStyle = findNormalStyle(doc)
  if (!normalStyle) return undefined

  const pPr = normalStyle.getElementsByTagName('w:pPr')[0]
  const rPr = normalStyle.getElementsByTagName('w:rPr')[0]

  const normal: NormalStyle = {}

  if (rPr) {
    const rf = extractRFontsFromRPr(rPr, themeFonts)
    normal.fontAscii = rf.ascii
    normal.fontHAnsi = rf.hAnsi
    normal.fontEastAsia = rf.eastAsia

    const sz = rPr.getElementsByTagName('w:sz')[0]
    const szVal = sz?.getAttribute('w:val')
    const sizeHalfPoints = safeNum(szVal)
    if (sizeHalfPoints) normal.fontSizeHalfPoints = sizeHalfPoints
  }

  if (pPr) {
    const jc = pPr.getElementsByTagName('w:jc')[0]
    normal.alignment = mapJc(jc?.getAttribute('w:val'))

    const spacing = pPr.getElementsByTagName('w:spacing')[0]
    if (spacing) {
      normal.spacing = {
        beforeTwips: safeNum(spacing.getAttribute('w:before')),
        afterTwips: safeNum(spacing.getAttribute('w:after')),
        lineTwips: safeNum(spacing.getAttribute('w:line')),
        lineRule: mapLineRule(spacing.getAttribute('w:lineRule')),
      }
    }

    const ind = pPr.getElementsByTagName('w:ind')[0]
    if (ind) {
      const firstLine = safeNum(ind.getAttribute('w:firstLine'))
      const left = safeNum(ind.getAttribute('w:left') || ind.getAttribute('w:start'))
      const right = safeNum(ind.getAttribute('w:right') || ind.getAttribute('w:end'))

      const firstLineChars = safeNum(ind.getAttribute('w:firstLineChars'))
      const leftChars = safeNum(ind.getAttribute('w:leftChars') || ind.getAttribute('w:startChars'))
      const rightChars = safeNum(ind.getAttribute('w:rightChars') || ind.getAttribute('w:endChars'))

      const docDefaults = extractDocDefaultsNormal(doc, themeFonts)
      const sizeForIndent = normal.fontSizeHalfPoints || docDefaults.fontSizeHalfPoints || 24

      normal.indent = {
        firstLineTwips:
          firstLine ??
          (typeof firstLineChars === 'number' ? charsHundredthToTwips(firstLineChars, sizeForIndent) : undefined),
        leftTwips:
          left ??
          (typeof leftChars === 'number' ? charsHundredthToTwips(leftChars, sizeForIndent) : undefined),
        rightTwips:
          right ??
          (typeof rightChars === 'number' ? charsHundredthToTwips(rightChars, sizeForIndent) : undefined),
      }
    }
  }

  // Fallback to docDefaults if Normal doesn't define fonts/size
  if (!normal.fontAscii && !normal.fontHAnsi && !normal.fontEastAsia) {
    const d = extractDocDefaultsNormal(doc, themeFonts)
    normal.fontAscii = d.fontAscii || normal.fontAscii
    normal.fontHAnsi = d.fontHAnsi || normal.fontHAnsi
    normal.fontEastAsia = d.fontEastAsia || normal.fontEastAsia
    normal.fontSizeHalfPoints = d.fontSizeHalfPoints || normal.fontSizeHalfPoints
  }

  // Last resort fallback (avoid empty)
  if (!normal.fontEastAsia && !normal.fontAscii && !normal.fontHAnsi) {
    normal.fontEastAsia = '等线'
    normal.fontAscii = 'Calibri'
    normal.fontHAnsi = 'Calibri'
  }

  return normal
}

function extractPageFromDocument(documentXml: string): DocxTypographyProfile['page'] | undefined {
  const doc = parseXml(documentXml)
  if (!doc) return undefined
  const sectPr = doc.getElementsByTagName('w:sectPr')[0]
  if (!sectPr) return undefined

  const page: DocxTypographyProfile['page'] = {}

  const pgMar = sectPr.getElementsByTagName('w:pgMar')[0]
  if (pgMar) {
    page.margin = {
      topTwips: safeNum(pgMar.getAttribute('w:top')),
      rightTwips: safeNum(pgMar.getAttribute('w:right')),
      bottomTwips: safeNum(pgMar.getAttribute('w:bottom')),
      leftTwips: safeNum(pgMar.getAttribute('w:left')),
    }
  }

  const pgSz = sectPr.getElementsByTagName('w:pgSz')[0]
  if (pgSz) {
    const orient = (pgSz.getAttribute('w:orient') || '').toLowerCase()
    page.size = {
      wTwips: safeNum(pgSz.getAttribute('w:w')),
      hTwips: safeNum(pgSz.getAttribute('w:h')),
      orientation: orient === 'landscape' ? 'landscape' : 'portrait',
    }
  }

  return page
}

function extractOutlineStats(documentXml: string): DocxOutlineStats {
  const xml = documentXml || ''
  const heading1Count =
    (xml.match(/<w:pStyle[^>]*w:val="Heading1"[^>]*>/g) || []).length +
    (xml.match(/<w:pStyle[^>]*w:val="标题1"[^>]*>/g) || []).length +
    (xml.match(/<w:pStyle[^>]*w:val="标题 1"[^>]*>/g) || []).length
  const heading2Count =
    (xml.match(/<w:pStyle[^>]*w:val="Heading2"[^>]*>/g) || []).length +
    (xml.match(/<w:pStyle[^>]*w:val="标题2"[^>]*>/g) || []).length +
    (xml.match(/<w:pStyle[^>]*w:val="标题 2"[^>]*>/g) || []).length
  const heading3Count =
    (xml.match(/<w:pStyle[^>]*w:val="Heading3"[^>]*>/g) || []).length +
    (xml.match(/<w:pStyle[^>]*w:val="标题3"[^>]*>/g) || []).length +
    (xml.match(/<w:pStyle[^>]*w:val="标题 3"[^>]*>/g) || []).length
  const tableCount = (xml.match(/<w:tbl[\s>]/g) || []).length
  const imageCount = (xml.match(/<w:drawing[\s>]/g) || []).length + (xml.match(/<w:pict[\s>]/g) || []).length
  return { heading1Count, heading2Count, heading3Count, tableCount, imageCount }
}

export async function extractTypographyProfileFromArrayBuffer(arrayBuffer: ArrayBuffer): Promise<{
  profile: DocxTypographyProfile
  outline: DocxOutlineStats
}> {
  const zip = await JSZip.loadAsync(arrayBuffer)
  const stylesXml = (await zip.file('word/styles.xml')?.async('string').catch(() => undefined)) || ''
  const documentXml = (await zip.file('word/document.xml')?.async('string').catch(() => undefined)) || ''
  const themeXml =
    (await zip.file('word/theme/theme1.xml')?.async('string').catch(() => undefined)) ||
    (await zip.file('word/theme/theme.xml')?.async('string').catch(() => undefined)) ||
    ''

  const profile: DocxTypographyProfile = {
    page: documentXml ? extractPageFromDocument(documentXml) : undefined,
    normal: stylesXml ? extractNormalFromStyles(stylesXml, themeXml || undefined) : undefined,
  }

  // Extract heading styles (fonts/size/spacing/indent) so headings won't look like body text.
  const stylesDoc = stylesXml ? parseXml(stylesXml) : null
  if (stylesDoc) {
    const themeFonts = themeXml ? extractThemeFonts(themeXml) : undefined
    const h1 = findStyleByIds(stylesDoc, ['Heading1', '标题1', '标题 1'])
    const h2 = findStyleByIds(stylesDoc, ['Heading2', '标题2', '标题 2'])
    const h3 = findStyleByIds(stylesDoc, ['Heading3', '标题3', '标题 3'])
    if (h1) profile.heading1 = extractParagraphStyleFromStyleEl(h1, stylesDoc, themeFonts)
    if (h2) profile.heading2 = extractParagraphStyleFromStyleEl(h2, stylesDoc, themeFonts)
    if (h3) profile.heading3 = extractParagraphStyleFromStyleEl(h3, stylesDoc, themeFonts)
  }

  const outline = extractOutlineStats(documentXml)
  return { profile, outline }
}

export function formatTypographyProfileForAgent(profile: DocxTypographyProfile, outline?: DocxOutlineStats): string {
  const lines: string[] = []
  const margin = profile.page?.margin
  if (margin) {
    lines.push(
      `页边距(twips): top=${margin.topTwips ?? '-'} right=${margin.rightTwips ?? '-'} bottom=${margin.bottomTwips ?? '-'} left=${margin.leftTwips ?? '-'}`
    )
  }
  const normal = profile.normal
  if (normal) {
    lines.push(
      `正文Normal: 字体(中)=${normal.fontEastAsia ?? '-'} 字体(英)=${normal.fontAscii ?? normal.fontHAnsi ?? '-'} 字号(half-points)=${normal.fontSizeHalfPoints ?? '-'} 对齐=${normal.alignment ?? '-'}`
    )
    const sp = normal.spacing
    if (sp) {
      lines.push(
        `段落间距/行距: before=${sp.beforeTwips ?? '-'} after=${sp.afterTwips ?? '-'} line=${sp.lineTwips ?? '-'} lineRule=${sp.lineRule ?? '-'}`
      )
    }
    const ind = normal.indent
    if (ind) {
      lines.push(
        `缩进: firstLine=${ind.firstLineTwips ?? '-'} left=${ind.leftTwips ?? '-'} right=${ind.rightTwips ?? '-'}`
      )
    }
  }
  if (profile.heading1) {
    lines.push(`Heading1: 字体(中)=${profile.heading1.fontEastAsia ?? '-'} 字体(英)=${profile.heading1.fontAscii ?? profile.heading1.fontHAnsi ?? '-'} 字号(half-points)=${profile.heading1.fontSizeHalfPoints ?? '-'}`)
  }
  if (profile.heading2) {
    lines.push(`Heading2: 字体(中)=${profile.heading2.fontEastAsia ?? '-'} 字体(英)=${profile.heading2.fontAscii ?? profile.heading2.fontHAnsi ?? '-'} 字号(half-points)=${profile.heading2.fontSizeHalfPoints ?? '-'}`)
  }
  if (profile.heading3) {
    lines.push(`Heading3: 字体(中)=${profile.heading3.fontEastAsia ?? '-'} 字体(英)=${profile.heading3.fontAscii ?? profile.heading3.fontHAnsi ?? '-'} 字号(half-points)=${profile.heading3.fontSizeHalfPoints ?? '-'}`)
  }
  if (outline) {
    lines.push(
      `结构统计: H1=${outline.heading1Count} H2=${outline.heading2Count} H3=${outline.heading3Count} tables=${outline.tableCount} images=${outline.imageCount}`
    )
  }
  return lines.join('\n')
}


