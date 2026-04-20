const JSZip = require('jszip')
const { XMLParser } = require('fast-xml-parser')
const { ensureDir, stableHash } = require('./macos-utils.cjs')

function createDocxInspectorService(options = {}) {
  const { fs, path, app } = options
  const xmlParser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: '@_',
    removeNSPrefix: false,
    parseTagValue: false,
    trimValues: false,
  })

  const toArray = (value) => {
    if (Array.isArray(value)) return value
    if (value == null) return []
    return [value]
  }

  const pickAttr = (node, names = []) => {
    if (!node || typeof node !== 'object') return undefined
    for (const name of names) {
      if (node[name] != null) return node[name]
    }
    return undefined
  }

  const textValue = (value) => {
    if (typeof value === 'string') return value
    if (value == null) return ''
    if (typeof value === 'number') return String(value)
    if (typeof value === 'object') {
      if (typeof value['#text'] === 'string') return value['#text']
      if (typeof value['__text'] === 'string') return value['__text']
    }
    return ''
  }

  const walk = (node, visitor, currentPath = []) => {
    if (node == null) return
    if (Array.isArray(node)) {
      node.forEach((item, index) => walk(item, visitor, currentPath.concat(index)))
      return
    }
    if (typeof node !== 'object') return

    visitor(node, currentPath)
    for (const [key, value] of Object.entries(node)) {
      if (key.startsWith('@_')) continue
      walk(value, visitor, currentPath.concat(key))
    }
  }

  const collectNodes = (root, nodeName) => {
    const result = []
    walk(root, (node) => {
      if (!node || typeof node !== 'object') return
      if (node[nodeName] != null) {
        result.push(...toArray(node[nodeName]))
      }
    })
    return result
  }

  const readXml = async (zip, innerPath) => {
    const file = zip.file(innerPath)
    if (!file) return null
    const xml = await file.async('string')
    return {
      xml,
      parsed: xmlParser.parse(xml),
    }
  }

  const extractZipToDir = async (zip, destDir) => {
    const extractedFiles = []
    const tasks = []

    zip.forEach((relativePath, file) => {
      const absolutePath = path.join(destDir, relativePath)
      if (file.dir) {
        ensureDir(fs, absolutePath)
        return
      }
      ensureDir(fs, path.dirname(absolutePath))
      extractedFiles.push(relativePath)
      tasks.push(
        file.async('nodebuffer').then((buffer) => {
          fs.writeFileSync(absolutePath, buffer)
        }),
      )
    })

    await Promise.all(tasks)
    return extractedFiles.sort()
  }

  const buildExtractDir = (filePath) => {
    const stats = fs.statSync(filePath)
    const baseDir = path.join(app.getPath('temp'), 'word-cursor', 'docx-inspector')
    ensureDir(fs, baseDir)
    const fingerprint = stableHash(`${filePath}:${stats.size}:${stats.mtimeMs}`)
    const destDir = path.join(baseDir, fingerprint)
    ensureDir(fs, destDir)
    return destDir
  }

  const parseRelationships = (parsedRels) => {
    const relsRoot = parsedRels?.Relationships || parsedRels?.['Relationships']
    const relationships = toArray(relsRoot?.Relationship || relsRoot?.['Relationship']).map((item) => ({
      id: pickAttr(item, ['@_Id']),
      type: pickAttr(item, ['@_Type']),
      target: pickAttr(item, ['@_Target']),
    }))
    return relationships
  }

  const summarizeFontTable = (parsedFontTable) => {
    const fontRoot = parsedFontTable?.['w:fonts']
    return toArray(fontRoot?.['w:font']).map((font) => ({
      name: pickAttr(font, ['@_w:name']),
      altName: pickAttr(font?.['w:altName'], ['@_w:val']),
      family: pickAttr(font?.['w:family'], ['@_w:val']),
      charset: pickAttr(font?.['w:charset'], ['@_w:val']),
      pitch: pickAttr(font?.['w:pitch'], ['@_w:val']),
      panose1: pickAttr(font?.['w:panose1'], ['@_w:val']),
    })).filter((font) => font.name)
  }

  const summarizePageSettings = (parsedDocument) => {
    const body = parsedDocument?.['w:document']?.['w:body']
    const sectPr = body?.['w:sectPr']
    if (!sectPr) return undefined

    return {
      widthTwips: Number(pickAttr(sectPr?.['w:pgSz'], ['@_w:w'])) || undefined,
      heightTwips: Number(pickAttr(sectPr?.['w:pgSz'], ['@_w:h'])) || undefined,
      marginTopTwips: Number(pickAttr(sectPr?.['w:pgMar'], ['@_w:top'])) || undefined,
      marginRightTwips: Number(pickAttr(sectPr?.['w:pgMar'], ['@_w:right'])) || undefined,
      marginBottomTwips: Number(pickAttr(sectPr?.['w:pgMar'], ['@_w:bottom'])) || undefined,
      marginLeftTwips: Number(pickAttr(sectPr?.['w:pgMar'], ['@_w:left'])) || undefined,
      headerTwips: Number(pickAttr(sectPr?.['w:pgMar'], ['@_w:header'])) || undefined,
      footerTwips: Number(pickAttr(sectPr?.['w:pgMar'], ['@_w:footer'])) || undefined,
      columns: Number(pickAttr(sectPr?.['w:cols'], ['@_w:num'])) || undefined,
      columnSpacingTwips: Number(pickAttr(sectPr?.['w:cols'], ['@_w:space'])) || undefined,
      docGridLinePitch: Number(pickAttr(sectPr?.['w:docGrid'], ['@_w:linePitch'])) || undefined,
      docGridCharSpace: Number(pickAttr(sectPr?.['w:docGrid'], ['@_w:charSpace'])) || undefined,
    }
  }

  const summarizeCompat = (parsedSettings) => {
    const compat = parsedSettings?.['w:settings']?.['w:compat'] || {}
    const compatSettings = toArray(compat?.['w:compatSetting'])
    const compatibilityMode = compatSettings.find((item) => pickAttr(item, ['@_w:name']) === 'compatibilityMode')

    const rawCharacterSpacing = pickAttr(parsedSettings?.['w:settings']?.['w:characterSpacingControl'], ['@_w:val'])
    return {
      compatibilityMode: pickAttr(compatibilityMode, ['@_w:val']),
      characterSpacingControl: rawCharacterSpacing,
      noPunctuationKerning: !!compat?.['w:noPunctuationKerning'],
      useFELayout: !!compat?.['w:useFELayout'],
      doNotUseEastAsianBreakRules: !!compat?.['w:doNotUseEastAsianBreakRules'],
      compressPunctuation: rawCharacterSpacing === 'compressPunctuation',
    }
  }

  const summarizeStyles = (parsedStyles) => {
    const stylesRoot = parsedStyles?.['w:styles']
    return toArray(stylesRoot?.['w:style']).map((style) => ({
      styleId: pickAttr(style, ['@_w:styleId']),
      type: pickAttr(style, ['@_w:type']),
      name: pickAttr(style?.['w:name'], ['@_w:val']),
      basedOn: pickAttr(style?.['w:basedOn'], ['@_w:val']),
      next: pickAttr(style?.['w:next'], ['@_w:val']),
      link: pickAttr(style?.['w:link'], ['@_w:val']),
      isDefault: String(pickAttr(style, ['@_w:default']) || '') === '1',
    })).filter((style) => style.styleId)
  }

  const summarizeReferencedStyleIds = (parsedDocument) => {
    const styleIds = new Set()
    walk(parsedDocument, (node) => {
      const pStyle = node?.['w:pStyle']
      if (pStyle) {
        toArray(pStyle).forEach((entry) => {
          const styleId = pickAttr(entry, ['@_w:val'])
          if (styleId) styleIds.add(styleId)
        })
      }
    })
    return Array.from(styleIds).sort()
  }

  const summarizeReferencedFonts = (parsedDocument, parsedStyles, fontTable) => {
    const altNameMap = new Map()
    fontTable.forEach((font) => {
      if (!altNameMap.has(font.name)) {
        altNameMap.set(font.name, new Set())
      }
      if (font.altName) altNameMap.get(font.name).add(font.altName)
    })

    const fontMap = new Map()
    const collectRFonts = (root) => {
      walk(root, (node) => {
        const rFonts = node?.['w:rFonts']
        if (!rFonts) return
        toArray(rFonts).forEach((entry) => {
          const candidates = [
            pickAttr(entry, ['@_w:ascii']),
            pickAttr(entry, ['@_w:hAnsi']),
            pickAttr(entry, ['@_w:eastAsia']),
            pickAttr(entry, ['@_w:cs']),
          ].filter(Boolean)
          candidates.forEach((name) => {
            if (!fontMap.has(name)) {
              fontMap.set(name, new Set())
            }
            const alternates = altNameMap.get(name)
            if (alternates) {
              alternates.forEach((altName) => fontMap.get(name).add(altName))
            }
          })
        })
      })
    }

    collectRFonts(parsedDocument)
    collectRFonts(parsedStyles)

    return Array.from(fontMap.entries())
      .map(([name, alternates]) => ({
        name,
        alternates: Array.from(alternates).sort(),
      }))
      .sort((a, b) => a.name.localeCompare(b.name))
  }

  const summarizeTocFields = (parsedDocument) => {
    const tocFields = new Set()
    walk(parsedDocument, (node) => {
      const instrText = node?.['w:instrText']
      if (!instrText) return
      toArray(instrText).forEach((entry) => {
        const text = textValue(entry).trim()
        if (text.toUpperCase().includes('TOC')) {
          tocFields.add(text)
        }
      })
    })
    return Array.from(tocFields)
  }

  const summarizeTables = (parsedDocument) => {
    const tables = []
    const tableNodes = collectNodes(parsedDocument, 'w:tbl')
    tableNodes.forEach((table, index) => {
      const rows = toArray(table?.['w:tr'])
      let maxCols = 0
      rows.forEach((row) => {
        let cols = 0
        toArray(row?.['w:tc']).forEach((cell) => {
          const span = Number(pickAttr(cell?.['w:tcPr']?.['w:gridSpan'], ['@_w:val'])) || 1
          cols += span
        })
        maxCols = Math.max(maxCols, cols)
      })

      tables.push({
        index: index + 1,
        rows: rows.length,
        columns: maxCols,
        widthTwips: Number(pickAttr(table?.['w:tblPr']?.['w:tblW'], ['@_w:w'])) || undefined,
        layout: pickAttr(table?.['w:tblPr']?.['w:tblLayout'], ['@_w:type']),
        floating: !!table?.['w:tblPr']?.['w:tblpPr'],
      })
    })
    return tables
  }

  const summarizeImages = (relationships, zip) => {
    return relationships
      .filter((item) => String(item.type || '').includes('/image'))
      .map((item) => {
        const normalizedTarget = item.target?.startsWith('word/')
          ? item.target
          : `word/${String(item.target || '').replace(/^\/+/, '')}`
        const file = normalizedTarget ? zip.file(normalizedTarget) : null
        return {
          relId: item.id,
          target: normalizedTarget,
          size: file?._data?.uncompressedSize || 0,
        }
      })
      .filter((image) => image.relId && image.target)
  }

  return {
    async inspect(filePath) {
      try {
        if (!filePath || !fs.existsSync(filePath)) {
          return { success: false, error: '文件不存在' }
        }
        if (path.extname(filePath).toLowerCase() !== '.docx') {
          return { success: false, error: 'docx-inspect 仅支持 .docx 文件' }
        }

        const buffer = fs.readFileSync(filePath)
        const zip = await JSZip.loadAsync(buffer)
        const extractedDir = buildExtractDir(filePath)
        const extractedFiles = await extractZipToDir(zip, extractedDir)

        const documentData = await readXml(zip, 'word/document.xml')
        const stylesData = await readXml(zip, 'word/styles.xml')
        const settingsData = await readXml(zip, 'word/settings.xml')
        const fontTableData = await readXml(zip, 'word/fontTable.xml')
        const relationshipsData = await readXml(zip, 'word/_rels/document.xml.rels')

        const fontTable = summarizeFontTable(fontTableData?.parsed)
        const relationships = parseRelationships(relationshipsData?.parsed)
        const footerTargets = relationships
          .filter((item) => String(item.type || '').includes('/footer'))
          .map((item) => item.target)
          .filter(Boolean)

        const report = {
          sourcePath: filePath,
          extractedDir,
          extractedFiles,
          xmlPaths: {
            document: documentData ? path.join(extractedDir, 'word', 'document.xml') : undefined,
            styles: stylesData ? path.join(extractedDir, 'word', 'styles.xml') : undefined,
            settings: settingsData ? path.join(extractedDir, 'word', 'settings.xml') : undefined,
            fontTable: fontTableData ? path.join(extractedDir, 'word', 'fontTable.xml') : undefined,
            numbering: zip.file('word/numbering.xml') ? path.join(extractedDir, 'word', 'numbering.xml') : undefined,
            theme: zip.file('word/theme/theme1.xml') ? path.join(extractedDir, 'word', 'theme', 'theme1.xml') : undefined,
            footers: footerTargets.map((target) => path.join(extractedDir, 'word', String(target).replace(/^word\//, '').replace(/^\/+/, ''))),
            rels: zip.file('_rels/.rels')
              ? [path.join(extractedDir, '_rels', '.rels'), path.join(extractedDir, 'word', '_rels', 'document.xml.rels')]
              : [path.join(extractedDir, 'word', '_rels', 'document.xml.rels')],
          },
          createdAt: new Date().toISOString(),
          summary: {
            pageSettings: summarizePageSettings(documentData?.parsed),
            compat: summarizeCompat(settingsData?.parsed),
            fontTable,
            referencedFonts: summarizeReferencedFonts(documentData?.parsed, stylesData?.parsed, fontTable),
            styleGraph: summarizeStyles(stylesData?.parsed),
            relationships,
            images: summarizeImages(relationships, zip),
            tables: summarizeTables(documentData?.parsed),
            tocFields: summarizeTocFields(documentData?.parsed),
            footerTargets,
            referencedStyleIds: summarizeReferencedStyleIds(documentData?.parsed),
          },
        }

        return { success: true, report }
      } catch (error) {
        return { success: false, error: error?.message || String(error) }
      }
    },
  }
}

module.exports = {
  createDocxInspectorService,
}
