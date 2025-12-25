import JSZip from 'jszip'
import mammoth from 'mammoth'

// Word 字号到 pt 的映射
const WORD_FONT_SIZE_MAP: Record<number, string> = {
  // 半点值 (half-points) 到 pt
  20: '10pt',   // 五号
  21: '10.5pt', // 五号半
  24: '12pt',   // 小四
  28: '14pt',   // 四号
  30: '15pt',   // 小三
  32: '16pt',   // 三号
  36: '18pt',   // 小二
  44: '22pt',   // 二号
  52: '26pt',   // 小一
  72: '36pt',   // 一号
}

// 将 Word 的 half-points 转换为 pt
function halfPointsToPt(halfPoints: number): string {
  if (WORD_FONT_SIZE_MAP[halfPoints]) {
    return WORD_FONT_SIZE_MAP[halfPoints]
  }
  return `${halfPoints / 2}pt`
}

interface RunStyle {
  bold?: boolean
  italic?: boolean
  underline?: boolean
  underlineStyle?: 'solid' | 'dotted' | 'dashed' | 'double' | 'wavy'
  strike?: boolean
  fontSize?: string
  fontFamily?: string
  color?: string
}

interface ParagraphStyle {
  alignment?: 'left' | 'center' | 'right' | 'justify'
  indent?: number
  heading?: number
  fontSize?: string
  fontFamily?: string
  color?: string
  lineHeight?: string
  marginTop?: string
  marginBottom?: string
}

// 字体映射 - 将 Word 字体映射到系统可用字体
const FONT_FALLBACK_MAP: Record<string, string> = {
  // 宋体系列
  '宋体': '"宋体", "SimSun", "Songti SC", serif',
  'SimSun': '"宋体", "SimSun", "Songti SC", serif',
  '新宋体': '"新宋体", "NSimSun", "宋体", serif',
  'NSimSun': '"新宋体", "NSimSun", "宋体", serif',
  '华文宋体': '"华文宋体", "STSong", "宋体", serif',
  'STSong': '"华文宋体", "STSong", "宋体", serif',
  '方正小标宋简体': '"方正小标宋简体", "宋体", serif',
  '方正小标宋_GBK': '"方正小标宋_GBK", "宋体", serif',
  
  // 黑体系列
  '黑体': '"黑体", "SimHei", "Heiti SC", sans-serif',
  'SimHei': '"黑体", "SimHei", "Heiti SC", sans-serif',
  '华文黑体': '"华文黑体", "STHeiti", "黑体", sans-serif',
  'STHeiti': '"华文黑体", "STHeiti", "黑体", sans-serif',
  '微软雅黑': '"微软雅黑", "Microsoft YaHei", "黑体", sans-serif',
  'Microsoft YaHei': '"微软雅黑", "Microsoft YaHei", "黑体", sans-serif',
  
  // 楷体系列
  '楷体': '"楷体", "KaiTi", "Kaiti SC", serif',
  'KaiTi': '"楷体", "KaiTi", "Kaiti SC", serif',
  '楷体_GB2312': '"楷体_GB2312", "楷体", "KaiTi", serif',
  '华文楷体': '"华文楷体", "STKaiti", "楷体", serif',
  'STKaiti': '"华文楷体", "STKaiti", "楷体", serif',
  
  // 仿宋系列
  '仿宋': '"仿宋", "FangSong", "Fangsong SC", serif',
  'FangSong': '"仿宋", "FangSong", "Fangsong SC", serif',
  '仿宋_GB2312': '"仿宋_GB2312", "仿宋", "FangSong", serif',
  '华文仿宋': '"华文仿宋", "STFangsong", "仿宋", serif',
  'STFangsong': '"华文仿宋", "STFangsong", "仿宋", serif',
  
  // 其他常用字体
  '华文中宋': '"华文中宋", "STZhongsong", "宋体", serif',
  'STZhongsong': '"华文中宋", "STZhongsong", "宋体", serif',
  '华文细黑': '"华文细黑", "STXihei", "黑体", sans-serif',
  '等线': '"等线", "DengXian", "微软雅黑", sans-serif',
  'DengXian': '"等线", "DengXian", "微软雅黑", sans-serif',
  
  // 英文字体
  'Times New Roman': '"Times New Roman", "宋体", serif',
  'Arial': '"Arial", "黑体", sans-serif',
  'Calibri': '"Calibri", "等线", sans-serif',
}

// 获取安全的字体族
function getSafeFontFamily(fontName: string | null | undefined): string {
  if (!fontName) return ''
  
  if (FONT_FALLBACK_MAP[fontName]) {
    return FONT_FALLBACK_MAP[fontName]
  }
  
  const isChinese = /[\u4e00-\u9fa5]/.test(fontName) || 
                    fontName.includes('Song') || 
                    fontName.includes('Hei') || 
                    fontName.includes('Kai') ||
                    fontName.includes('Fang')
  
  if (isChinese) {
    return `"${fontName}", "宋体", "SimSun", serif`
  }
  
  return `"${fontName}", "Arial", sans-serif`
}

// 检查文件是否是有效的 ZIP/DOCX 格式
function isValidZip(bytes: Uint8Array): boolean {
  return bytes.length >= 4 && bytes[0] === 0x50 && bytes[1] === 0x4B
}

// 检查是否是旧版 .doc 格式 (OLE Compound Document)
function isOldDocFormat(bytes: Uint8Array): boolean {
  return bytes.length >= 4 && 
         bytes[0] === 0xD0 && bytes[1] === 0xCF && 
         bytes[2] === 0x11 && bytes[3] === 0xE0
}

// 使用 mammoth 解析 Word 文档（支持 .doc 和 .docx）
async function parseWithMammoth(arrayBuffer: ArrayBuffer): Promise<string> {
  try {
    const result = await mammoth.convertToHtml({ arrayBuffer })
    
    if (result.messages.length > 0) {
      console.log('Mammoth 解析消息:', result.messages)
    }
    
    let html = result.value
    
    // 如果没有内容，返回空段落
    if (!html || html.trim() === '') {
      return '<p></p>'
    }
    
    // 添加一些基本样式处理
    // 将 mammoth 生成的简单 HTML 转换为更适合显示的格式
    html = html
      // 段落添加缩进样式
      .replace(/<p>/g, '<p style="text-indent: 2em; margin: 0.5em 0;">')
      // 表格添加边框
      .replace(/<table>/g, '<table style="border-collapse: collapse; width: 100%; margin: 1em 0;">')
      .replace(/<td>/g, '<td style="border: 1px solid #000; padding: 8px;">')
      .replace(/<th>/g, '<th style="border: 1px solid #000; padding: 8px; background: #f5f5f5;">')
    
    return html
  } catch (error) {
    console.error('Mammoth 解析失败:', error)
    throw error
  }
}

// 解析 docx 文件并转换为 HTML
export async function parseDocxToHtml(base64Data: string): Promise<string> {
  try {
    console.log('开始解析 Word 文档，数据长度:', base64Data.length)
    
    // 解码 base64
    const binaryString = atob(base64Data)
    const bytes = new Uint8Array(binaryString.length)
    for (let i = 0; i < binaryString.length; i++) {
      bytes[i] = binaryString.charCodeAt(i)
    }
    
    console.log('文件头字节:', bytes[0], bytes[1], bytes[2], bytes[3])
    
    // 获取 ArrayBuffer 用于 mammoth
    const arrayBuffer = bytes.buffer

    // 检查文件格式并选择解析方式
    if (isOldDocFormat(bytes)) {
      console.log('检测到旧版 .doc 格式，使用 mammoth 解析')
      return await parseWithMammoth(arrayBuffer)
    }

    if (isValidZip(bytes)) {
      console.log('检测到 .docx 格式')
      
      // 先尝试用我们的自定义解析器（保留更多样式）
      try {
        const customResult = await parseDocxCustom(bytes)
        if (customResult && customResult.trim() && !customResult.includes('□')) {
          console.log('自定义解析器成功')
          return customResult
        }
      } catch (e) {
        console.log('自定义解析器失败，回退到 mammoth:', e)
      }
      
      // 回退到 mammoth
      console.log('使用 mammoth 解析 .docx')
      return await parseWithMammoth(arrayBuffer)
    }

    // 未知格式，尝试用 mammoth
    console.log('未知格式，尝试用 mammoth 解析')
    try {
      return await parseWithMammoth(arrayBuffer)
    } catch (e) {
      console.error('mammoth 也无法解析:', e)
      return `<div style="padding: 40px; text-align: center; color: #888;">
        <p style="font-size: 18px; margin-bottom: 10px;">⚠️ 无法识别的文件格式</p>
        <p style="font-size: 14px;">请确保文件是有效的 Word 文档 (.doc 或 .docx)</p>
      </div>`
    }
  } catch (error) {
    console.error('Word 文档解析错误:', error)
    return `<div style="padding: 40px; text-align: center; color: #888;">
      <p style="font-size: 18px; margin-bottom: 10px;">⚠️ 文档解析失败</p>
      <p style="font-size: 14px;">${(error as Error).message}</p>
    </div>`
  }
}

// 自定义 docx 解析器（保留更多样式信息）
async function parseDocxCustom(bytes: Uint8Array): Promise<string> {
  // JSZip.loadAsync 需要 ArrayBuffer；Uint8Array.buffer 可能是 SharedArrayBuffer
  const ab = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer
  const zip = await JSZip.loadAsync(ab)
  
  const documentXml = await zip.file('word/document.xml')?.async('string')
  if (!documentXml) {
    throw new Error('找不到 document.xml')
  }

  const stylesXml = await zip.file('word/styles.xml')?.async('string')
  const styles = stylesXml ? parseStyles(stylesXml) : {}

  const parser = new DOMParser()
  const doc = parser.parseFromString(documentXml, 'application/xml')

  const parseError = doc.querySelector('parsererror')
  if (parseError) {
    throw new Error('XML 解析失败')
  }

  // 获取 body 元素
  const body = doc.getElementsByTagName('w:body')[0]
  if (!body) {
    throw new Error('找不到文档主体')
  }

  let html = ''
  
  // 遍历 body 的直接子元素，处理段落和表格
  const children = body.childNodes
  for (let i = 0; i < children.length; i++) {
    const child = children[i] as Element
    if (child.nodeName === 'w:p') {
      html += parseParagraph(child, styles)
    } else if (child.nodeName === 'w:tbl') {
      html += parseTable(child, styles)
    } else if (child.nodeName === 'w:sdt') {
      // 结构化文档标签，递归处理其内容
      const sdtContent = child.getElementsByTagName('w:sdtContent')[0]
      if (sdtContent) {
        const sdtChildren = sdtContent.childNodes
        for (let j = 0; j < sdtChildren.length; j++) {
          const sdtChild = sdtChildren[j] as Element
          if (sdtChild.nodeName === 'w:p') {
            html += parseParagraph(sdtChild, styles)
          } else if (sdtChild.nodeName === 'w:tbl') {
            html += parseTable(sdtChild, styles)
          }
        }
      }
    }
  }

  return html || '<p></p>'
}

// 解析表格
function parseTable(tbl: Element, styles: Record<string, any>): string {
  let tableStyle = 'border-collapse: collapse; margin: 1em 0; table-layout: fixed;'
  
  // 获取表格网格定义（列宽）
  const tblGrid = tbl.getElementsByTagName('w:tblGrid')[0]
  const columnWidths: number[] = []
  let totalWidth = 0
  
  if (tblGrid) {
    const gridCols = tblGrid.getElementsByTagName('w:gridCol')
    for (let i = 0; i < gridCols.length; i++) {
      const w = gridCols[i].getAttribute('w:w')
      if (w) {
        const width = parseInt(w)
        columnWidths.push(width)
        totalWidth += width
      }
    }
  }
  
  // 解析表格属性
  const tblPr = tbl.getElementsByTagName('w:tblPr')[0]
  if (tblPr) {
    // 表格对齐
    const jc = tblPr.getElementsByTagName('w:jc')[0]
    if (jc) {
      const val = jc.getAttribute('w:val')
      if (val === 'center') {
        tableStyle += ' margin-left: auto; margin-right: auto;'
      }
    }
    
    // 表格宽度
    const tblW = tblPr.getElementsByTagName('w:tblW')[0]
    if (tblW) {
      const w = tblW.getAttribute('w:w')
      const type = tblW.getAttribute('w:type')
      if (w && type === 'pct') {
        tableStyle += ` width: ${parseInt(w) / 50}%;`
      } else if (w && type === 'dxa') {
        tableStyle += ` width: ${parseInt(w) / 20}pt;`
      } else if (w && type === 'auto') {
        tableStyle += ' width: 100%;'
      }
    } else if (totalWidth > 0) {
      // 使用计算出的总宽度
      tableStyle += ` width: ${totalWidth / 20}pt;`
    } else {
      tableStyle += ' width: 100%;'
    }
  } else {
    tableStyle += ' width: 100%;'
  }
  
  let html = `<table style="${tableStyle}">`
  
  // 如果有列宽定义，添加 colgroup
  if (columnWidths.length > 0) {
    html += '<colgroup>'
    for (const width of columnWidths) {
      // 转换为百分比或固定宽度
      if (totalWidth > 0) {
        const pct = (width / totalWidth * 100).toFixed(2)
        html += `<col style="width: ${pct}%;">`
      } else {
        html += `<col style="width: ${width / 20}pt;">`
      }
    }
    html += '</colgroup>'
  }
  
  // 解析表格行（直接子元素，避免嵌套表格问题）
  const children = tbl.childNodes
  for (let i = 0; i < children.length; i++) {
    const child = children[i] as Element
    if (child.nodeName === 'w:tr') {
      html += parseTableRow(child, styles, false, columnWidths, totalWidth)
    }
  }
  
  html += '</table>'
  return html
}

// 解析表格行
function parseTableRow(tr: Element, styles: Record<string, any>, isFirstRow: boolean, columnWidths: number[] = [], totalWidth: number = 0): string {
  let rowStyle = ''
  
  // 解析行属性
  const trPr = tr.getElementsByTagName('w:trPr')[0]
  if (trPr) {
    // 行高
    const trHeight = trPr.getElementsByTagName('w:trHeight')[0]
    if (trHeight) {
      const val = trHeight.getAttribute('w:val')
      if (val) {
        rowStyle += `height: ${parseInt(val) / 20}pt;`
      }
    }
  }
  
  let html = rowStyle ? `<tr style="${rowStyle}">` : '<tr>'
  
  // 解析单元格（只处理直接子元素）
  let colIndex = 0
  const children = tr.childNodes
  for (let i = 0; i < children.length; i++) {
    const child = children[i] as Element
    if (child.nodeName === 'w:tc') {
      const cellResult = parseTableCell(child, styles, isFirstRow, columnWidths, totalWidth, colIndex)
      html += cellResult.html
      colIndex += cellResult.colSpan
    }
  }
  
  html += '</tr>'
  return html
}

// 解析表格单元格
function parseTableCell(tc: Element, styles: Record<string, any>, isHeader: boolean, columnWidths: number[] = [], totalWidth: number = 0, colIndex: number = 0): { html: string; colSpan: number } {
  let cellStyle = 'border: 1px solid #000; padding: 8px 12px; vertical-align: middle;'
  let colspan = ''
  let colSpan = 1
  
  // 解析单元格属性
  const tcPr = tc.getElementsByTagName('w:tcPr')[0]
  if (tcPr) {
    // 合并列
    const gridSpan = tcPr.getElementsByTagName('w:gridSpan')[0]
    if (gridSpan) {
      const val = gridSpan.getAttribute('w:val')
      if (val && parseInt(val) > 1) {
        colSpan = parseInt(val)
        colspan = ` colspan="${val}"`
      }
    }
    
    // 合并行 - 检查是否是被合并的单元格
    const vMerge = tcPr.getElementsByTagName('w:vMerge')[0]
    if (vMerge) {
      const val = vMerge.getAttribute('w:val')
      // 如果没有 val 属性或 val="continue"，说明是被合并的单元格，跳过
      if (!val || val === 'continue') {
        return { html: '', colSpan }
      }
      // val="restart" 表示这是合并的起始单元格
    }
    
    // 单元格宽度 - 优先使用 tcW，否则从 columnWidths 计算
    const tcW = tcPr.getElementsByTagName('w:tcW')[0]
    if (tcW) {
      const w = tcW.getAttribute('w:w')
      const type = tcW.getAttribute('w:type')
      if (w && type === 'dxa') {
        // 不设置固定宽度，让 colgroup 控制
      } else if (w && type === 'pct') {
        cellStyle += ` width: ${parseInt(w) / 50}%;`
      }
    }
    
    // 单元格背景色
    const shd = tcPr.getElementsByTagName('w:shd')[0]
    if (shd) {
      const fill = shd.getAttribute('w:fill')
      if (fill && fill !== 'auto' && fill !== 'FFFFFF') {
        cellStyle += ` background-color: #${fill};`
      }
    }
    
    // 垂直对齐
    const vAlign = tcPr.getElementsByTagName('w:vAlign')[0]
    if (vAlign) {
      const val = vAlign.getAttribute('w:val')
      if (val === 'center') {
        cellStyle += ' vertical-align: middle;'
      } else if (val === 'bottom') {
        cellStyle += ' vertical-align: bottom;'
      } else if (val === 'top') {
        cellStyle += ' vertical-align: top;'
      }
    }
  }
  
  // 解析单元格内容（只处理直接子段落）
  let content = ''
  const children = tc.childNodes
  let paragraphCount = 0
  
  for (let i = 0; i < children.length; i++) {
    const child = children[i] as Element
    if (child.nodeName === 'w:p') {
      if (paragraphCount > 0) {
        content += '<br>'
      }
      content += parseParagraphContent(child, styles, true)
      paragraphCount++
    }
  }
  
  const tag = isHeader ? 'th' : 'td'
  if (isHeader) {
    cellStyle += ' font-weight: bold; background-color: #f5f5f5;'
  }
  
  return { 
    html: `<${tag} style="${cellStyle}"${colspan}>${content || '&nbsp;'}</${tag}>`,
    colSpan 
  }
}

// 解析段落内容（不含外层标签，用于表格单元格）
function parseParagraphContent(para: Element, styles: Record<string, any>, inTable: boolean = false): string {
  const pPr = para.getElementsByTagName('w:pPr')[0]
  let alignment = ''
  
  if (pPr) {
    const jc = pPr.getElementsByTagName('w:jc')[0]
    if (jc) {
      const val = jc.getAttribute('w:val')
      if (val === 'center') alignment = 'text-align: center;'
      else if (val === 'right') alignment = 'text-align: right;'
    }
  }
  
  let content = ''
  const childNodes = para.childNodes
  
  for (let i = 0; i < childNodes.length; i++) {
    const child = childNodes[i] as Element
    if (child.nodeName === 'w:r') {
      content += parseRun(child)
    } else if (child.nodeName === 'w:hyperlink') {
      const linkRuns = child.getElementsByTagName('w:r')
      let linkContent = ''
      for (let j = 0; j < linkRuns.length; j++) {
        linkContent += parseRun(linkRuns[j])
      }
      content += linkContent
    }
  }
  
  if (alignment && content) {
    return `<span style="${alignment}">${content}</span>`
  }
  
  return content
}

// 解析样式定义
function parseStyles(stylesXml: string): Record<string, any> {
  const parser = new DOMParser()
  const doc = parser.parseFromString(stylesXml, 'application/xml')
  const styles: Record<string, any> = {}

  const styleElements = doc.getElementsByTagName('w:style')
  for (let i = 0; i < styleElements.length; i++) {
    const style = styleElements[i]
    const styleId = style.getAttribute('w:styleId')
    if (styleId) {
      styles[styleId] = parseStyleElement(style)
    }
  }

  return styles
}

function parseStyleElement(style: Element): any {
  const result: any = {}
  
  const pPr = style.getElementsByTagName('w:pPr')[0]
  if (pPr) {
    const jc = pPr.getElementsByTagName('w:jc')[0]
    if (jc) {
      const val = jc.getAttribute('w:val')
      if (val === 'center') result.alignment = 'center'
      else if (val === 'right') result.alignment = 'right'
      else if (val === 'both') result.alignment = 'justify'
    }
  }

  const rPr = style.getElementsByTagName('w:rPr')[0]
  if (rPr) {
    const sz = rPr.getElementsByTagName('w:sz')[0]
    if (sz) {
      const val = sz.getAttribute('w:val')
      if (val) result.fontSize = halfPointsToPt(parseInt(val))
    }

    const rFonts = rPr.getElementsByTagName('w:rFonts')[0]
    if (rFonts) {
      const fontName = rFonts.getAttribute('w:eastAsia') || 
                       rFonts.getAttribute('w:ascii')
      result.fontFamily = getSafeFontFamily(fontName) || '"仿宋", "FangSong", serif'
    }
  }

  return result
}

// 解析段落
function parseParagraph(para: Element, styles: Record<string, any>): string {
  const pPr = para.getElementsByTagName('w:pPr')[0]
  let paraStyle: ParagraphStyle = {}
  let tag = 'p'
  const styleProps: string[] = []
  let hasIndent = false
  let paraFontSize = ''
  let paraFontFamily = ''
  let paraColor = ''

  if (pPr) {
    const pStyle = pPr.getElementsByTagName('w:pStyle')[0]
    if (pStyle) {
      const styleId = pStyle.getAttribute('w:val')
      if (styleId) {
        if (styleId.includes('Heading') || styleId.includes('标题')) {
          const level = styleId.match(/\d/)?.[0] || '1'
          tag = `h${level}`
          paraStyle.heading = parseInt(level)
        }
        if (styles[styleId]) {
          Object.assign(paraStyle, styles[styleId])
        }
      }
    }

    // 段落级别的文字样式 (rPr in pPr)
    const pRpr = pPr.getElementsByTagName('w:rPr')[0]
    if (pRpr) {
      const sz = pRpr.getElementsByTagName('w:sz')[0]
      if (sz) {
        const val = sz.getAttribute('w:val')
        if (val) {
          paraFontSize = halfPointsToPt(parseInt(val))
        }
      }
      const rFonts = pRpr.getElementsByTagName('w:rFonts')[0]
      if (rFonts) {
        const fontName = rFonts.getAttribute('w:eastAsia') || 
                         rFonts.getAttribute('w:ascii') ||
                         rFonts.getAttribute('w:hAnsi')
        if (fontName) {
          paraFontFamily = getSafeFontFamily(fontName) || ''
        }
      }
      const color = pRpr.getElementsByTagName('w:color')[0]
      if (color) {
        const val = color.getAttribute('w:val')
        if (val && val !== 'auto') {
          paraColor = `#${val}`
        }
      }
    }

    // 对齐方式
    const jc = pPr.getElementsByTagName('w:jc')[0]
    if (jc) {
      const val = jc.getAttribute('w:val')
      if (val === 'center') {
        paraStyle.alignment = 'center'
        styleProps.push('text-align: center')
      } else if (val === 'right') {
        paraStyle.alignment = 'right'
        styleProps.push('text-align: right')
      } else if (val === 'both' || val === 'distribute') {
        paraStyle.alignment = 'justify'
        styleProps.push('text-align: justify')
      }
    }

    // 缩进
    const ind = pPr.getElementsByTagName('w:ind')[0]
    if (ind) {
      const firstLineChars = ind.getAttribute('w:firstLineChars')
      const firstLine = ind.getAttribute('w:firstLine')
      const left = ind.getAttribute('w:left') || ind.getAttribute('w:start')
      const leftChars = ind.getAttribute('w:leftChars') || ind.getAttribute('w:startChars')
      const hanging = ind.getAttribute('w:hanging')
      
      // 首行缩进
      if (firstLineChars) {
        const chars = parseInt(firstLineChars) / 100
        if (chars > 0) {
          styleProps.push(`text-indent: ${chars}em`)
          hasIndent = true
        }
      } else if (firstLine) {
        const twips = parseInt(firstLine)
        if (twips > 0) {
          const em = twips / 240
          styleProps.push(`text-indent: ${em.toFixed(2)}em`)
          hasIndent = true
        }
      }
      
      // 悬挂缩进（负缩进）
      if (hanging) {
        const twips = parseInt(hanging)
        if (twips > 0) {
          const em = twips / 240
          styleProps.push(`text-indent: -${em.toFixed(2)}em`)
          hasIndent = true
        }
      }
      
      // 左缩进
      if (leftChars) {
        const chars = parseInt(leftChars) / 100
        if (chars > 0) {
          styleProps.push(`padding-left: ${chars}em`)
        }
      } else if (left) {
        const twips = parseInt(left)
        if (twips > 0) {
          const em = twips / 240
          styleProps.push(`padding-left: ${em.toFixed(2)}em`)
        }
      }
    }

    // 段落间距和行距
    const spacing = pPr.getElementsByTagName('w:spacing')[0]
    if (spacing) {
      const before = spacing.getAttribute('w:before')
      const after = spacing.getAttribute('w:after')
      const line = spacing.getAttribute('w:line')
      const lineRule = spacing.getAttribute('w:lineRule')
      
      if (before) {
        const twips = parseInt(before)
        if (twips > 0) {
          styleProps.push(`margin-top: ${(twips / 20).toFixed(1)}pt`)
        }
      }
      if (after) {
        const twips = parseInt(after)
        if (twips > 0) {
          styleProps.push(`margin-bottom: ${(twips / 20).toFixed(1)}pt`)
        }
      }
      
      // 行距处理
      if (line) {
        const lineVal = parseInt(line)
        if (lineRule === 'exact') {
          // 固定值行距
          styleProps.push(`line-height: ${(lineVal / 20).toFixed(1)}pt`)
        } else if (lineRule === 'atLeast') {
          // 最小值行距
          styleProps.push(`line-height: ${(lineVal / 20).toFixed(1)}pt`)
        } else if (!lineRule || lineRule === 'auto') {
          // 倍数行距：240 = 单倍行距
          const multiplier = lineVal / 240
          styleProps.push(`line-height: ${multiplier.toFixed(2)}`)
        }
      }
    }
  }

  // 添加段落级别的字体样式
  if (paraFontSize) {
    styleProps.push(`font-size: ${paraFontSize}`)
  }
  if (paraFontFamily) {
    styleProps.push(`font-family: ${paraFontFamily}`)
  }
  if (paraColor) {
    styleProps.push(`color: ${paraColor}`)
  }

  const styleAttr = styleProps.length > 0 ? ` style="${styleProps.join('; ')}"` : ''

  let content = ''
  const childNodes = para.childNodes
  
  for (let i = 0; i < childNodes.length; i++) {
    const child = childNodes[i] as Element
    if (child.nodeName === 'w:r') {
      content += parseRun(child)
    } else if (child.nodeName === 'w:hyperlink') {
      const linkRuns = child.getElementsByTagName('w:r')
      let linkContent = ''
      for (let j = 0; j < linkRuns.length; j++) {
        linkContent += parseRun(linkRuns[j])
      }
      const rId = child.getAttribute('r:id')
      if (rId) {
        content += `<a href="#${rId}">${linkContent}</a>`
      } else {
        content += linkContent
      }
    } else if (child.nodeName === 'w:fldSimple' || child.nodeName === 'w:smartTag' || child.nodeName === 'w:sdt') {
      const innerRuns = child.getElementsByTagName('w:r')
      for (let j = 0; j < innerRuns.length; j++) {
        content += parseRun(innerRuns[j])
      }
    }
  }

  if (!content.trim()) {
    return `<${tag}${styleAttr}><br></${tag}>`
  }

  return `<${tag}${styleAttr}>${content}</${tag}>`
}

// 解析 run（文本块）
function parseRun(run: Element): string {
  const rPr = run.getElementsByTagName('w:rPr')[0]
  let style: RunStyle = {}

  if (rPr) {
    if (rPr.getElementsByTagName('w:b').length > 0) {
      style.bold = true
    }

    if (rPr.getElementsByTagName('w:i').length > 0) {
      style.italic = true
    }

    const u = rPr.getElementsByTagName('w:u')[0]
    if (u) {
      const uVal = u.getAttribute('w:val')
      if (uVal && uVal !== 'none') {
        style.underline = true
        if (uVal === 'dotted' || uVal === 'dottedHeavy') {
          style.underlineStyle = 'dotted'
        } else if (uVal === 'dash' || uVal === 'dashLong') {
          style.underlineStyle = 'dashed'
        } else if (uVal === 'double') {
          style.underlineStyle = 'double'
        } else if (uVal === 'wave' || uVal === 'wavyDouble') {
          style.underlineStyle = 'wavy'
        } else {
          style.underlineStyle = 'solid'
        }
      }
    }

    if (rPr.getElementsByTagName('w:strike').length > 0) {
      style.strike = true
    }

    const sz = rPr.getElementsByTagName('w:sz')[0]
    if (sz) {
      const val = sz.getAttribute('w:val')
      if (val) {
        style.fontSize = halfPointsToPt(parseInt(val))
      }
    }

    const rFonts = rPr.getElementsByTagName('w:rFonts')[0]
    if (rFonts) {
      const fontName = rFonts.getAttribute('w:eastAsia') || 
                       rFonts.getAttribute('w:ascii') ||
                       rFonts.getAttribute('w:hAnsi')
      style.fontFamily = getSafeFontFamily(fontName)
    }

    const color = rPr.getElementsByTagName('w:color')[0]
    if (color) {
      const val = color.getAttribute('w:val')
      if (val && val !== 'auto') {
        style.color = `#${val}`
      }
    }
  }

  let text = ''
  let hasSpecialChars = false
  const children = run.childNodes
  for (let i = 0; i < children.length; i++) {
    const child = children[i] as Element
    if (child.nodeName === 'w:t') {
      text += child.textContent || ''
    } else if (child.nodeName === 'w:tab') {
      // 使用实际的 tab 空格，而不是 &nbsp; 实体
      text += '        ' // 8个普通空格
      hasSpecialChars = true
    } else if (child.nodeName === 'w:br' || child.nodeName === 'w:cr') {
      text += '[[BR]]' // 使用占位符
      hasSpecialChars = true
    } else if (child.nodeName === 'w:sym') {
      const char = child.getAttribute('w:char')
      if (char) {
        text += String.fromCharCode(parseInt(char, 16))
      }
    } else if (child.nodeName === 'w:ptab') {
      text += '    ' // 4个普通空格
      hasSpecialChars = true
    }
  }

  if (!text) return ''

  // 先转义 HTML，然后处理占位符
  let html = escapeHtml(text)
  
  // 将占位符替换回 HTML 标签
  html = html.replace(/\[\[BR\]\]/g, '<br>')

  const styleProps: string[] = []
  if (style.fontSize) {
    styleProps.push(`font-size: ${style.fontSize}`)
  }
  if (style.fontFamily) {
    styleProps.push(`font-family: ${style.fontFamily}`)
  }
  if (style.color) {
    styleProps.push(`color: ${style.color}`)
  }

  if (style.underline && style.underlineStyle && style.underlineStyle !== 'solid') {
    styleProps.push('text-decoration: underline')
    styleProps.push(`text-decoration-style: ${style.underlineStyle}`)
  }

  if (styleProps.length > 0) {
    html = `<span style="${styleProps.join('; ')}">${html}</span>`
  }

  if (style.bold) html = `<strong>${html}</strong>`
  if (style.italic) html = `<em>${html}</em>`
  if (style.underline && (!style.underlineStyle || style.underlineStyle === 'solid')) {
    html = `<u>${html}</u>`
  }
  if (style.strike) html = `<s>${html}</s>`

  return html
}

// HTML 转义
function escapeHtml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;')
}
