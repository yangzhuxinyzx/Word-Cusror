function createExcelService(options) {
  const {
    fs,
    path,
    app,
    XLSX,
    ExcelJS,
  } = options
async function readExcelWithSheetJS(filePath) {
  try {
    const XLSX = require('xlsx')
    const buffer = fs.readFileSync(filePath)
    
    console.log('[Excel] 开始读取 .xls 文件:', filePath)
    
    // 读取 xls 文件，启用样式选项
    const workbook = XLSX.read(buffer, { 
      type: 'buffer', 
      cellStyles: true, 
      cellFormula: true,
      cellNF: true,
      cellDates: true,
    })
    
    // 获取样式表
    const styles = workbook.Styles || {}
    const cellXfs = styles.CellXf || []
    const fonts = styles.Fonts || []
    const fills = styles.Fills || []
    const borders = styles.Borders || []
    const numFmts = styles.NumberFmt || {}
    
    console.log('[Excel] 样式表信息:', {
      cellXfsCount: cellXfs.length,
      fontsCount: fonts.length,
      fillsCount: fills.length,
      bordersCount: borders.length,
    })
    
    const sheets = []
    
    for (const sheetName of workbook.SheetNames) {
      const worksheet = workbook.Sheets[sheetName]
      const range = worksheet['!ref'] ? XLSX.utils.decode_range(worksheet['!ref']) : { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } }
      
      const sheetData = {
        name: sheetName,
        range: range,
        merges: [],
        colWidths: [],
        rowHeights: [],
        cells: []
      }
      
      // 合并单元格
      if (worksheet['!merges']) {
        sheetData.merges = worksheet['!merges'].map(m => ({
          s: { r: m.s.r, c: m.s.c },
          e: { r: m.e.r, c: m.e.c }
        }))
      }
      
      // 列宽
      if (worksheet['!cols']) {
        worksheet['!cols'].forEach((col, idx) => {
          if (col && col.wpx) {
            sheetData.colWidths[idx] = col.wpx
          } else if (col && col.wch) {
            sheetData.colWidths[idx] = Math.round(col.wch * 7 + 5)
          }
        })
      }
      
      // 行高
      if (worksheet['!rows']) {
        worksheet['!rows'].forEach((row, idx) => {
          if (row && row.hpx) {
            sheetData.rowHeights[idx] = row.hpx
          } else if (row && row.hpt) {
            sheetData.rowHeights[idx] = Math.round(row.hpt * 1.333)
          }
        })
      }
      
      // 遍历单元格
      let debugCount = 0
      const keys = Object.keys(worksheet).filter(k => !k.startsWith('!'))
      
      for (const addr of keys) {
        const cell = worksheet[addr]
        if (!cell) continue
        
        const decoded = XLSX.utils.decode_cell(addr)
        const r = decoded.r
        const c = decoded.c
        
        // 调试：打印前3个单元格的完整信息
        if (debugCount < 3) {
          console.log('[Excel XLS] 单元格完整数据:', {
            address: addr,
            cell: JSON.stringify(cell, null, 2)
          })
          debugCount++
        }
        
        // 解析样式
        const styleObj = {}
        
        // 方法1: 直接从 cell.s 获取样式对象
        if (cell.s && typeof cell.s === 'object') {
          console.log('[Excel XLS] 发现样式对象 cell.s:', cell.s)
          
          // 字体
          if (cell.s.font) {
            styleObj.font = {
              name: cell.s.font.name,
              sz: cell.s.font.sz,
              bold: cell.s.font.bold,
              italic: cell.s.font.italic,
              underline: cell.s.font.underline,
              strike: cell.s.font.strike,
              color: cell.s.font.color
            }
          }
          
          // 填充
          if (cell.s.fill || cell.s.fgColor || cell.s.bgColor) {
            styleObj.fill = {
              fgColor: cell.s.fgColor || cell.s.fill?.fgColor,
              bgColor: cell.s.bgColor || cell.s.fill?.bgColor
            }
          }
          
          // 对齐
          if (cell.s.alignment) {
            styleObj.alignment = cell.s.alignment
          }
          
          // 边框
          if (cell.s.border) {
            styleObj.border = cell.s.border
          }
        }
        // 方法2: 通过样式索引获取
        else if (typeof cell.s === 'number' && cellXfs[cell.s]) {
          const xf = cellXfs[cell.s]
          
          if (!debuggedFirstCell) {
            console.log('[Excel XLS] 单元格样式示例 (通过索引):', {
              address: addr,
              value: cell.v,
              styleIndex: cell.s,
              xf: xf,
              font: fonts[xf.fontId],
              fill: fills[xf.fillId]
            })
            debuggedFirstCell = true
          }
          
          // 字体
          if (xf.fontId !== undefined && fonts[xf.fontId]) {
            const font = fonts[xf.fontId]
            styleObj.font = {
              name: font.name,
              sz: font.sz,
              bold: font.bold,
              italic: font.italic,
              underline: font.underline,
              strike: font.strike,
              color: font.color
            }
          }
          
          // 填充
          if (xf.fillId !== undefined && fills[xf.fillId]) {
            const fill = fills[xf.fillId]
            styleObj.fill = {
              fgColor: fill.fgColor,
              bgColor: fill.bgColor
            }
          }
          
          // 对齐
          if (xf.alignment) {
            styleObj.alignment = xf.alignment
          }
          
          // 边框
          if (xf.borderId !== undefined && borders[xf.borderId]) {
            styleObj.border = borders[xf.borderId]
          }
          
          // 数字格式
          if (xf.numFmtId !== undefined) {
            styleObj.numFmt = numFmts[xf.numFmtId] || xf.numFmtId
          }
        }
        
        const cellData = {
          r,
          c,
          v: cell.v,
          t: cell.t,
          f: cell.f,
          s: styleObj,
          w: cell.w,
          display: cell.w || (cell.v != null ? String(cell.v) : '')
        }
        
        sheetData.cells.push(cellData)
      }
      
      sheets.push(sheetData)
    }
    
    console.log('[Excel] .xls 文件读取成功，工作表数:', sheets.length)
    return { success: true, sheets }
  } catch (error) {
    console.error('读取 .xls 文件失败:', error)
    return { success: false, error: error.message }
  }
}

// 检查 LibreOffice 是否安装
function findLibreOffice() {
  const possiblePaths = [
    // Windows 常见路径
    'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
    'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe',
    // 应用内置便携版（如果打包）
    path.join(__dirname, '..', 'libreoffice', 'program', 'soffice.exe'),
    path.join(__dirname, 'libreoffice', 'program', 'soffice.exe'),
    // 环境变量
    process.env.LIBREOFFICE_PATH,
  ].filter(Boolean)
  
  for (const p of possiblePaths) {
    if (fs.existsSync(p)) {
      console.log('[Excel] 找到 LibreOffice:', p)
      return p
    }
  }
  console.log('[Excel] LibreOffice 未找到')
  return null
}

// 获取 LibreOffice 下载链接
function getLibreOfficeDownloadUrl() {
  if (process.platform === 'win32') {
    // LibreOffice 便携版 (约 300MB)
    return 'https://download.documentfoundation.org/libreoffice/portable/7.6.4/LibreOfficePortable_7.6.4_MultilingualStandard.paf.exe'
  }
  return null
}

// ==================== PPTX 预览渲染（LibreOffice → PNG） ====================

function hashForFileCache(filePath) {
  const st = fs.statSync(filePath)
  const key = `${filePath}|${st.size}|${st.mtimeMs}`
  return crypto.createHash('sha1').update(key).digest('hex')
}

function getPptxPreviewCacheDir(filePath) {
  const hash = hashForFileCache(filePath)
  const tempDir = app.getPath('temp')
  return path.join(tempDir, 'word-cursor-ppt-preview', hash)
}

function listPngFilesSorted(dir) {
  const files = fs.readdirSync(dir).filter((f) => f.toLowerCase().endsWith('.png'))
  const withMeta = files.map((name) => {
    const m = name.match(/(\d+)(?=\.png$)/)
    const idx = m ? parseInt(m[1], 10) : 0
    return { name, idx }
  })
  withMeta.sort((a, b) => (a.idx - b.idx) || a.name.localeCompare(b.name))
  return withMeta.map((x) => path.join(dir, x.name))
}

async function renderPptxToPngsWithLibreOffice(pptxPath, outDir) {
  const libreOfficePath = findLibreOffice()
  if (!libreOfficePath) {
    return { success: false, error: 'LibreOffice 未安装', downloadUrl: getLibreOfficeDownloadUrl() }
  }

  if (!fs.existsSync(outDir)) {
    fs.mkdirSync(outDir, { recursive: true })
  }

  const { execFile } = require('child_process')
  // LibreOffice 将每页导出为 PNG（文件名规则依版本不同，导出后我们扫描目录排序）
  return new Promise((resolve) => {
    execFile(
      libreOfficePath,
      ['--headless', '--nologo', '--nolockcheck', '--norestore', '--convert-to', 'png', '--outdir', outDir, pptxPath],
      { timeout: 180000 },
      (error, stdout, stderr) => {
        if (error) {
          console.error('[PPTX] LibreOffice 转换失败:', error)
          resolve({ success: false, error: 'LibreOffice 转换失败', details: stderr || stdout })
          return
        }
        const pngs = listPngFilesSorted(outDir)
        if (!pngs.length) {
          resolve({ success: false, error: 'LibreOffice 转换未生成 PNG' })
          return
        }
        resolve({ success: true, images: pngs })
      }
    )
  })
}

async function checkLibreOffice() {
  const path = findLibreOffice()
  return {
    installed: !!path,
    path: path,
    downloadUrl: !path ? getLibreOfficeDownloadUrl() : null
  }
}

// 使用 LibreOffice 进行无损转换（开源方案）
async function convertWithLibreOffice(xlsPath) {
  const libreOfficePath = findLibreOffice()
  if (!libreOfficePath) {
    return { success: false, error: 'LibreOffice 未安装' }
  }
  
  const xlsxPath = xlsPath.replace(/\.xls$/i, '.xlsx')
  const outputDir = path.dirname(xlsPath)
  
  if (fs.existsSync(xlsxPath)) {
    return { 
      success: false, 
      error: `文件 ${path.basename(xlsxPath)} 已存在。请先删除或重命名现有文件。` 
    }
  }
  
  const { execFile } = require('child_process')
  
  return new Promise((resolve) => {
    // LibreOffice 命令行转换
    execFile(libreOfficePath, [
      '--headless',
      '--convert-to', 'xlsx',
      '--outdir', outputDir,
      xlsPath
    ], { timeout: 60000 }, (error, stdout, stderr) => {
      if (error) {
        console.error('[Excel] LibreOffice 转换失败:', error)
        resolve({ success: false, error: 'LibreOffice 转换失败', details: stderr })
      } else if (fs.existsSync(xlsxPath)) {
        console.log('[Excel] LibreOffice 转换成功:', xlsxPath)
        resolve({ 
          success: true, 
          xlsxPath,
          message: `已使用 LibreOffice 转换为 ${path.basename(xlsxPath)}，所有样式已完整保留！`
        })
      } else {
        resolve({ success: false, error: 'LibreOffice 转换后文件不存在' })
      }
    })
  })
}

// 使用系统安装的 Excel 进行无损转换（保留所有样式）
async function convertWithExcel(xlsPath) {
  const xlsxPath = xlsPath.replace(/\.xls$/i, '.xlsx')
  
  // 检查输出文件是否已存在
  if (fs.existsSync(xlsxPath)) {
    return { 
      success: false, 
      error: `文件 ${path.basename(xlsxPath)} 已存在。请先删除或重命名现有文件。` 
    }
  }
  
  // 使用 PowerShell 调用 Excel COM 对象
  const { exec } = require('child_process')
  
  // 转义路径中的特殊字符
  const escapedXlsPath = xlsPath.replace(/'/g, "''")
  const escapedXlsxPath = xlsxPath.replace(/'/g, "''")
  
  const psScript = `
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try {
      $workbook = $excel.Workbooks.Open('${escapedXlsPath}')
      $workbook.SaveAs('${escapedXlsxPath}', 51)
      $workbook.Close($false)
      Write-Output "SUCCESS"
    } catch {
      Write-Output "ERROR: $($_.Exception.Message)"
    } finally {
      $excel.Quit()
      [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
  `
  
  return new Promise((resolve) => {
    exec(`powershell -Command "${psScript.replace(/"/g, '\\"').replace(/\n/g, ' ')}"`, 
      { encoding: 'utf8', maxBuffer: 1024 * 1024, timeout: 60000 },
      (error, stdout, stderr) => {
        if (error || !stdout.includes('SUCCESS')) {
          console.error('[Excel] PowerShell 转换失败:', error || stderr || stdout)
          resolve({ 
            success: false, 
            error: '调用 Excel 失败',
            details: stderr || stdout
          })
        } else {
          console.log('[Excel] Excel COM 转换成功:', xlsxPath)
          resolve({ 
            success: true, 
            xlsxPath,
            message: `已使用 Microsoft Excel 转换为 ${path.basename(xlsxPath)}，所有样式已完整保留！`
          })
        }
      }
    )
  })
}

// 使用 SheetJS 转换（数据转换，样式可能丢失）
async function convertWithSheetJS(xlsPath) {
  const XLSX = require('xlsx')
  const xlsxPath = xlsPath.replace(/\.xls$/i, '.xlsx')
  
  if (fs.existsSync(xlsxPath)) {
    return { 
      success: false, 
      error: `文件 ${path.basename(xlsxPath)} 已存在。请先删除或重命名现有文件。` 
    }
  }
  
  const buffer = fs.readFileSync(xlsPath)
  const workbook = XLSX.read(buffer, { 
    type: 'buffer',
    cellFormula: true,
    cellNF: true,
    cellDates: true
  })
  
  XLSX.writeFile(workbook, xlsxPath, { bookType: 'xlsx' })
  
  return { 
    success: true, 
    xlsxPath,
    message: `已转换为 ${path.basename(xlsxPath)}。注意：由于技术限制，样式信息可能丢失。`
  }
}

// 将 xls 转换为 xlsx（优先级：LibreOffice > Excel > SheetJS）
async function excelConvertXlsToXlsx(xlsPath) {
  try {
    console.log('[Excel] 开始转换 xls 到 xlsx:', xlsPath)
    
    // 1. 优先尝试 LibreOffice（开源，跨平台）
    console.log('[Excel] 尝试 LibreOffice...')
    const libreResult = await convertWithLibreOffice(xlsPath)
    if (libreResult.success) {
      return libreResult
    }
    console.log('[Excel] LibreOffice 不可用:', libreResult.error)
    
    // 2. Windows 上尝试 Excel COM
    if (process.platform === 'win32') {
      console.log('[Excel] 尝试 Microsoft Excel...')
      const excelResult = await convertWithExcel(xlsPath)
      if (excelResult.success) {
        return excelResult
      }
      console.log('[Excel] Excel COM 不可用:', excelResult.error)
    }
    
    // 3. 最后使用 SheetJS（数据转换，样式可能丢失）
    console.log('[Excel] 使用 SheetJS 进行基础转换（样式可能丢失）...')
    return await convertWithSheetJS(xlsPath)
  } catch (error) {
    console.error('xls 转 xlsx 失败:', error)
    return { success: false, error: error.message }
  }
}

// 读取 Excel（高保真只读预览数据）
// .xlsx 使用 ExcelJS（更好的样式支持），.xls 使用 SheetJS
async function excelOpen(filePath) {
  if (!filePath) {
    return { success: false, error: '缺少 filePath 参数' }
  }

  const ext = path.extname(filePath).toLowerCase()
  
  // .xls 文件使用 SheetJS 直接读取
  // 注意：SheetJS 免费版对 xls 样式支持有限
  if (ext === '.xls') {
    const result = await readExcelWithSheetJS(filePath)
    result.isXls = true  // 标记为 xls 文件
    result.originalPath = filePath
    // 添加警告信息，提示用户样式可能不完整
    result.warning = '提示：.xls 格式的样式支持有限。建议在 Microsoft Excel 中打开原文件，另存为 .xlsx 格式后重新打开，即可完整显示所有样式。'
    return result
  }
  
  // .xlsx 文件使用 ExcelJS 读取（更好的样式支持）
  try {
    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.readFile(filePath)
    
    const sheets = []
    const names = workbook.definedNames?.model || []
    
    workbook.eachSheet((worksheet, sheetId) => {
      // 注意：ExcelJS 的 worksheet.rowCount/columnCount 可能因为“整表格式/模板”变成 1048576/16384
      // 前端会按 range 构造矩阵，导致 OOM/白屏。这里按真实 used-range（非空单元格 + 合并区域）计算。
      let maxR = -1
      let maxC = -1

      const decodeCell = (addr) => {
        const match = String(addr || '').toUpperCase().match(/^(\$?)([A-Z]+)(\$?)(\d+)$/)
        if (!match) return null
        let col = 0
        for (let i = 0; i < match[2].length; i++) {
          col = col * 26 + (match[2].charCodeAt(i) - 64)
        }
        return { c: col - 1, r: parseInt(match[4], 10) - 1 }
      }

      const sheetData = {
        name: worksheet.name,
        range: { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } },
        merges: [],
        colWidths: [],
        rowHeights: [],
        autoFilter: worksheet.autoFilter || null,
        printArea: null,
        margins: null,
        dataValidations: null,
        cells: []
      }
      
      // 合并单元格
      if (worksheet.model && worksheet.model.merges) {
        worksheet.model.merges.forEach((mergeRange) => {
          const parts = String(mergeRange || '').split(':')
          if (parts.length !== 2) return
          const start = decodeCell(parts[0])
          const end = decodeCell(parts[1])
          if (!start || !end) return
          sheetData.merges.push({ s: { r: start.r, c: start.c }, e: { r: end.r, c: end.c } })
          maxR = Math.max(maxR, end.r)
          maxC = Math.max(maxC, end.c)
        })
      }
      
      // 列宽
      if (worksheet.columns) {
        worksheet.columns.forEach((col, idx) => {
          if (col && col.width) {
            // ExcelJS 列宽是字符数，转为像素（约 7px/字符 + 5px padding）
            sheetData.colWidths[idx] = Math.round(col.width * 7 + 5)
          }
        })
      }
      
      // 行高和单元格
      worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        // 行高（ExcelJS 返回 points，转为像素）
        if (row.height) {
          sheetData.rowHeights[rowNumber - 1] = Math.round(row.height * 1.333)
        }
        
        row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
          const r = rowNumber - 1
          const c = colNumber - 1
          maxR = Math.max(maxR, r)
          maxC = Math.max(maxC, c)
          
          // 提取样式
          const styleObj = {}
          
          // 字体
          if (cell.font) {
            styleObj.font = {
              name: cell.font.name,
              sz: cell.font.size,
              bold: cell.font.bold,
              italic: cell.font.italic,
              underline: cell.font.underline,
              strike: cell.font.strike,
              color: cell.font.color ? { argb: cell.font.color.argb, rgb: cell.font.color.argb?.slice(2) } : null
            }
          }
          
          // 填充/背景色
          if (cell.fill) {
            styleObj.fill = {}
            if (cell.fill.type === 'pattern' && cell.fill.fgColor) {
              styleObj.fill.fgColor = { argb: cell.fill.fgColor.argb, rgb: cell.fill.fgColor.argb?.slice(2) }
            }
            if (cell.fill.bgColor) {
              styleObj.fill.bgColor = { argb: cell.fill.bgColor.argb, rgb: cell.fill.bgColor.argb?.slice(2) }
            }
          }
          
          // 对齐
          if (cell.alignment) {
            styleObj.alignment = {
              horizontal: cell.alignment.horizontal,
              vertical: cell.alignment.vertical,
              wrapText: cell.alignment.wrapText,
              shrinkToFit: cell.alignment.shrinkToFit,
              indent: cell.alignment.indent,
              textRotation: cell.alignment.textRotation
            }
          }
          
          // 边框
          if (cell.border) {
            styleObj.border = {}
            ;['top', 'bottom', 'left', 'right'].forEach((side) => {
              if (cell.border[side]) {
                styleObj.border[side] = {
                  style: cell.border[side].style,
                  color: cell.border[side].color ? { argb: cell.border[side].color.argb, rgb: cell.border[side].color.argb?.slice(2) } : null
                }
              }
            })
          }
          
          // 数字格式
          if (cell.numFmt) {
            styleObj.numFmt = cell.numFmt
          }
          
          // 获取显示值（安全处理，避免 null 值和合并单元格错误）
          let display = ''
          try {
            // 先尝试获取 value，因为 text getter 在合并单元格时会报错
            const cellValue = cell.value
            if (cellValue != null) {
              if (typeof cellValue === 'object') {
                // 富文本 { richText: [...] }
                if (cellValue.richText && Array.isArray(cellValue.richText)) {
                  display = cellValue.richText.map(rt => rt.text || '').join('')
                }
                // 公式 { formula: '...', result: ... }
                else if (cellValue.formula) {
                  // 如果有计算结果，显示结果
                  if (cellValue.result != null) {
                    display = String(cellValue.result)
                  } else {
                    // 尝试计算公式（传入 workbook 支持跨工作表引用）
                    const calculated = evaluateSimpleFormula(cellValue.formula, worksheet, workbook)
                    if (calculated != null) {
                      display = String(calculated)
                    } else {
                      // 无法计算时显示公式本身
                      display = '=' + cellValue.formula
                    }
                  }
                }
                // 超链接 { text: '...', hyperlink: '...' }
                else if (cellValue.text != null) {
                  display = String(cellValue.text)
                }
                // 其他对象（可能有 result 但没有 formula）
                else if (cellValue.result != null) {
                  display = String(cellValue.result)
                }
                // 其他对象
                else {
                  display = String(cellValue)
                }
              } else {
                display = String(cellValue)
              }
            }
          } catch (e) {
            // 如果还是失败，返回空字符串
            console.warn(`[Excel Read] 单元格 ${colNumber}:${rowNumber} 读取失败:`, e.message)
            display = ''
          }
          
          // 公式
          const formula = cell.formula || (cell.value && cell.value.formula) || null
          
          // 超链接
          const hyperlink = cell.hyperlink || null
          
          // 批注
          let comment = null
          if (cell.note) {
            comment = typeof cell.note === 'string' ? cell.note : (cell.note.texts ? cell.note.texts.map(t => t.text || t).join('') : '')
          }
          
          sheetData.cells.push({
            r,
            c,
            v: cell.value,
            t: cell.type,
            w: display, // 使用安全计算的 display 值，避免 cell.text getter 错误
            f: formula,
            l: hyperlink,
            z: cell.numFmt,
            cmt: comment,
            display,
            s: styleObj
          })
        })
      })

      // 修正 range：使用真实 used-range，避免 rowCount/columnCount 造成超大范围
      if (maxR >= 0 && maxC >= 0) {
        sheetData.range.e = { r: maxR, c: maxC }
      } else {
        sheetData.range.e = { r: 0, c: 0 }
      }
      
      sheets.push(sheetData)
    })

    return { success: true, sheets, names }
  } catch (error) {
    console.error('读取 Excel 失败:', error)
    return { success: false, error: error.message || '读取 Excel 失败' }
  }
}

// ==================== Excel 增删查改操作 ====================

// 缓存打开的工作簿，避免每次操作都重新加载
const openWorkbooks = new Map()

// 获取或加载工作簿
async function getWorkbook(filePath) {
  if (openWorkbooks.has(filePath)) {
    return openWorkbooks.get(filePath)
  }
  
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(filePath)
  openWorkbooks.set(filePath, workbook)
  return workbook
}

// 保存工作簿
async function saveWorkbook(filePath) {
  const workbook = openWorkbooks.get(filePath)
  if (workbook) {
    await workbook.xlsx.writeFile(filePath)
    return true
  }
  return false
}

// 清除工作簿缓存
function clearWorkbookCache(filePath) {
  openWorkbooks.delete(filePath)
}

// ============================================================
// Excel 公式计算引擎 - 支持跨工作表引用和完整函数库
// ============================================================

/**
 * 创建一个公式计算器实例
 * @param {Object} workbook - ExcelJS 工作簿对象
 * @param {Object} currentWorksheet - 当前工作表
 */
function createFormulaEngine(workbook, currentWorksheet) {
  // 缓存已计算的单元格，防止循环引用
  const calculationCache = new Map()
  const calculationStack = new Set()
  
  // 解析单元格地址 (如 "A1" -> { r: 0, c: 0 })
  // 也支持纯列引用 "A" -> { r: null, c: 0, isColumn: true }
  const parseCellAddr = (address) => {
    const upperAddr = address.toUpperCase()
    
    // 尝试匹配带行号的地址 (如 A1, $B$2)
    const match = upperAddr.match(/^(\$?)([A-Z]+)(\$?)(\d+)$/)
    if (match) {
      let col = 0
      for (let i = 0; i < match[2].length; i++) {
        col = col * 26 + (match[2].charCodeAt(i) - 64)
      }
      return { r: parseInt(match[4], 10) - 1, c: col - 1 }
    }
    
    // 尝试匹配纯列引用 (如 A, B, $C)
    const colMatch = upperAddr.match(/^(\$?)([A-Z]+)$/)
    if (colMatch) {
      let col = 0
      for (let i = 0; i < colMatch[2].length; i++) {
        col = col * 26 + (colMatch[2].charCodeAt(i) - 64)
      }
      return { r: null, c: col - 1, isColumn: true }
    }
    
    return null
  }
  
  // 获取工作表（支持跨工作表引用）
  const getWorksheet = (sheetName) => {
    if (!sheetName) return currentWorksheet
    // 移除引号
    const cleanName = sheetName.replace(/^'|'$/g, '')
    const targetSheet = workbook.getWorksheet(cleanName)
    
    // 调试日志
    console.log(`[Formula Debug] getWorksheet: sheetName="${sheetName}", cleanName="${cleanName}", found=${!!targetSheet}`)
    if (!targetSheet) {
      // 列出所有可用的工作表名称
      const availableSheets = []
      workbook.eachSheet((ws) => availableSheets.push(ws.name))
      console.log(`[Formula Debug] 可用工作表: ${availableSheets.join(', ')}`)
    }
    
    return targetSheet || currentWorksheet
  }
  
  // 解析带工作表引用的单元格地址 (如 "'Sheet1'!A1" 或 "A1")
  const parseFullReference = (ref) => {
    const sheetMatch = ref.match(/^'?([^'!]+)'?!(.+)$/)
    if (sheetMatch) {
      return { sheetName: sheetMatch[1], cellRef: sheetMatch[2] }
    }
    return { sheetName: null, cellRef: ref }
  }
  
  // 获取单元格的原始值（不计算）
  const getRawCellValue = (ws, row, col) => {
    const cell = ws.getCell(row, col)
    return cell.value
  }
  
  // 当前计算上下文的工作表（用于嵌套公式计算）
  let activeWorksheet = currentWorksheet
  
  // 获取单元格的计算值
  const getCellValue = (ref, defaultWs = null) => {
    const { sheetName, cellRef } = parseFullReference(ref)
    // 优先级：1. ref 中指定的工作表 2. 传入的 defaultWs 3. activeWorksheet
    const ws = sheetName ? getWorksheet(sheetName) : (defaultWs || activeWorksheet)
    const addr = parseCellAddr(cellRef)
    if (!addr) return 0
    
    const cacheKey = `${ws.name || 'default'}!${cellRef}`
    
    // 检查循环引用
    if (calculationStack.has(cacheKey)) {
      console.warn(`[Formula] 检测到循环引用: ${cacheKey}`)
      return 0
    }
    
    // 检查缓存
    if (calculationCache.has(cacheKey)) {
      return calculationCache.get(cacheKey)
    }
    
    const cell = ws.getCell(addr.r + 1, addr.c + 1)
    const value = cell.value
    
    if (value == null) return 0
    if (typeof value === 'number') return value
    if (typeof value === 'string') {
      const num = parseFloat(value)
      return isNaN(num) ? value : num
    }
    if (typeof value === 'object') {
      if (value.result != null) return value.result
      if (value.formula) {
        calculationStack.add(cacheKey)
        // 关键修复：临时切换活动工作表上下文，确保嵌套公式在正确的工作表中计算
        const previousActiveWs = activeWorksheet
        activeWorksheet = ws
        const result = evaluateFormula(value.formula, ws)
        activeWorksheet = previousActiveWs  // 恢复之前的上下文
        calculationStack.delete(cacheKey)
        if (result != null) {
          calculationCache.set(cacheKey, result)
          return result
        }
      }
      if (value.richText) {
        return value.richText.map(t => t.text || '').join('')
      }
      if (value.text != null) return value.text
    }
    return 0
  }
  
  // 获取单元格的文本值（用于文本函数）
  const getCellText = (ref, defaultWs = currentWorksheet) => {
    const val = getCellValue(ref, defaultWs)
    return String(val)
  }
  
  // 解析范围并获取所有值
  const getRangeValues = (rangeStr, ws = currentWorksheet) => {
    const { sheetName, cellRef } = parseFullReference(rangeStr)
    const targetWs = getWorksheet(sheetName)
    
    console.log(`[Formula Debug] getRangeValues: rangeStr="${rangeStr}", sheetName="${sheetName}", cellRef="${cellRef}", targetWs="${targetWs?.name}"`)
    
    const parts = cellRef.split(':')
    if (parts.length !== 2) {
      // 单个单元格
      return [getCellValue(rangeStr, ws)]
    }
    
    const start = parseCellAddr(parts[0])
    const end = parseCellAddr(parts[1])
    if (!start || !end) return []
    
    // 处理整列范围（如 E:E）
    let startRow = start.r
    let endRow = end.r
    if (start.isColumn || end.isColumn) {
      // 整列范围：只遍历有数据的行
      startRow = 0
      endRow = Math.max((targetWs?.rowCount || 100) - 1, 0)
      // 限制最大行数，避免遍历太多空行
      endRow = Math.min(endRow, 999)
    }
    
    const values = []
    for (let r = startRow; r <= endRow; r++) {
      for (let c = start.c; c <= end.c; c++) {
        const val = getCellValue(`${getColumnLabel(c)}${r + 1}`, targetWs)
        values.push(val)
      }
    }
    
    // 显示前10个值用于调试
    console.log(`[Formula Debug] getRangeValues 结果: 共${values.length}个值, 非0值数量: ${values.filter(v => v !== 0 && v !== '').length}`)
    
    return values
  }
  
  // 解析范围并获取所有单元格信息（包含位置）
  const getRangeCells = (rangeStr, ws = currentWorksheet) => {
    const { sheetName, cellRef } = parseFullReference(rangeStr)
    const targetWs = getWorksheet(sheetName)
    
    const parts = cellRef.split(':')
    if (parts.length !== 2) return []
    
    const start = parseCellAddr(parts[0])
    const end = parseCellAddr(parts[1])
    if (!start || !end) return []
    
    // 处理整列范围（如 E:E, H:H）
    let startRow = start.r
    let endRow = end.r
    if (start.isColumn || end.isColumn) {
      startRow = 0
      endRow = Math.max((targetWs?.rowCount || 100) - 1, 0)
      endRow = Math.min(endRow, 999)
    }
    
    const cells = []
    for (let r = startRow; r <= endRow; r++) {
      for (let c = start.c; c <= end.c; c++) {
        const ref = `${getColumnLabel(c)}${r + 1}`
        cells.push({
          row: r,
          col: c,
          ref,
          value: getCellValue(ref, targetWs),
          rawValue: getRawCellValue(targetWs, r + 1, c + 1)
        })
      }
    }
    return cells
  }
  
  // 获取列标签
  const getColumnLabel = (colIndex) => {
    let label = ''
    let n = colIndex
    while (n >= 0) {
      label = String.fromCharCode(65 + (n % 26)) + label
      n = Math.floor(n / 26) - 1
    }
    return label
  }
  
  // 解析函数参数（处理嵌套括号和逗号）
  const parseFunctionArgs = (argsStr) => {
    const args = []
    let depth = 0
    let current = ''
    
    for (let i = 0; i < argsStr.length; i++) {
      const char = argsStr[i]
      if (char === '(') depth++
      else if (char === ')') depth--
      else if (char === ',' && depth === 0) {
        args.push(current.trim())
        current = ''
        continue
      }
      current += char
    }
    if (current.trim()) args.push(current.trim())
    return args
  }
  
  // ============================================================
  // Excel 函数实现
  // ============================================================
  
  const functions = {
    // -------------------- 基础数学函数 --------------------
    
    // SUM - 求和
    SUM: (args) => {
      let total = 0
      for (const arg of args) {
        if (arg.includes(':')) {
          const values = getRangeValues(arg)
          total += values.filter(v => typeof v === 'number').reduce((a, b) => a + b, 0)
        } else {
          const val = evaluateExpression(arg)
          if (typeof val === 'number') total += val
        }
      }
      return total
    },
    
    // SUMIF - 条件求和
    SUMIF: (args) => {
      if (args.length < 2) return 0
      const [rangeStr, criteria, sumRangeStr] = args
      const cells = getRangeCells(rangeStr)
      const sumCells = sumRangeStr ? getRangeCells(sumRangeStr) : cells
      
      const criteriaValue = evaluateExpression(criteria.replace(/^"|"$/g, ''))
      let total = 0
      
      cells.forEach((cell, idx) => {
        if (matchCriteria(cell.value, criteriaValue)) {
          const sumVal = sumCells[idx]?.value
          if (typeof sumVal === 'number') total += sumVal
        }
      })
      return total
    },
    
    // SUMIFS - 多条件求和
    SUMIFS: (args) => {
      if (args.length < 3) return 0
      const sumRangeStr = args[0]
      const sumCells = getRangeCells(sumRangeStr)
      
      // 解析条件对
      const conditions = []
      for (let i = 1; i < args.length; i += 2) {
        if (i + 1 < args.length) {
          conditions.push({
            cells: getRangeCells(args[i]),
            criteria: evaluateExpression(args[i + 1].replace(/^"|"$/g, ''))
          })
        }
      }
      
      let total = 0
      sumCells.forEach((sumCell, idx) => {
        const allMatch = conditions.every(cond => {
          const cell = cond.cells[idx]
          return cell && matchCriteria(cell.value, cond.criteria)
        })
        if (allMatch && typeof sumCell.value === 'number') {
          total += sumCell.value
        }
      })
      return total
    },
    
    // AVERAGE - 平均值（只计算非空单元格中的数字）
    AVERAGE: (args) => {
      const values = []
      for (const arg of args) {
        if (arg.includes(':')) {
          // 使用 getRangeCells 获取原始值，排除真正的空单元格
          const cells = getRangeCells(arg)
          cells.forEach(c => {
            if (c.rawValue != null && c.rawValue !== '' && typeof c.value === 'number') {
              values.push(c.value)
            }
          })
        } else {
          const val = evaluateExpression(arg)
          if (typeof val === 'number') values.push(val)
        }
      }
      return values.length > 0 ? values.reduce((a, b) => a + b, 0) / values.length : 0
    },
    
    // AVERAGEIF - 条件平均值
    AVERAGEIF: (args) => {
      if (args.length < 2) return 0
      const [rangeStr, criteria, avgRangeStr] = args
      const cells = getRangeCells(rangeStr)
      const avgCells = avgRangeStr ? getRangeCells(avgRangeStr) : cells
      
      const criteriaValue = evaluateExpression(criteria.replace(/^"|"$/g, ''))
      const values = []
      
      cells.forEach((cell, idx) => {
        if (matchCriteria(cell.value, criteriaValue)) {
          const avgVal = avgCells[idx]?.value
          if (typeof avgVal === 'number') values.push(avgVal)
        }
      })
      return values.length > 0 ? values.reduce((a, b) => a + b, 0) / values.length : 0
    },
    
    // MAX - 最大值
    MAX: (args) => {
      const values = []
      for (const arg of args) {
        if (arg.includes(':')) {
          values.push(...getRangeValues(arg).filter(v => typeof v === 'number'))
        } else {
          const val = evaluateExpression(arg)
          if (typeof val === 'number') values.push(val)
        }
      }
      return values.length > 0 ? Math.max(...values) : 0
    },
    
    // MIN - 最小值（只计算非空单元格中的数字）
    MIN: (args) => {
      const values = []
      for (const arg of args) {
        if (arg.includes(':')) {
          // 使用 getRangeCells 获取原始值，排除真正的空单元格
          const cells = getRangeCells(arg)
          cells.forEach(c => {
            if (c.rawValue != null && c.rawValue !== '' && typeof c.value === 'number') {
              values.push(c.value)
            }
          })
        } else {
          const val = evaluateExpression(arg)
          if (typeof val === 'number') values.push(val)
        }
      }
      return values.length > 0 ? Math.min(...values) : 0
    },
    
    // ROUND - 四舍五入
    ROUND: (args) => {
      const num = evaluateExpression(args[0])
      const digits = args[1] ? evaluateExpression(args[1]) : 0
      if (typeof num !== 'number') return 0
      const factor = Math.pow(10, digits)
      return Math.round(num * factor) / factor
    },
    
    // ABS - 绝对值
    ABS: (args) => Math.abs(evaluateExpression(args[0]) || 0),
    
    // SQRT - 平方根
    SQRT: (args) => Math.sqrt(evaluateExpression(args[0]) || 0),
    
    // POWER - 幂运算
    POWER: (args) => Math.pow(evaluateExpression(args[0]) || 0, evaluateExpression(args[1]) || 0),
    
    // MOD - 取余
    MOD: (args) => {
      const num = evaluateExpression(args[0])
      const divisor = evaluateExpression(args[1])
      if (divisor === 0) return 0
      return num % divisor
    },
    
    // -------------------- 统计函数 --------------------
    
    // COUNT - 计数（仅数字）
    COUNT: (args) => {
      let count = 0
      for (const arg of args) {
        if (arg.includes(':')) {
          count += getRangeValues(arg).filter(v => typeof v === 'number').length
        } else {
          const val = evaluateExpression(arg)
          if (typeof val === 'number') count++
        }
      }
      return count
    },
    
    // COUNTA - 计数（非空单元格）
    COUNTA: (args) => {
      let count = 0
      for (const arg of args) {
        if (arg.includes(':')) {
          // 使用 getRangeCells 获取原始值，正确判断空单元格
          const cells = getRangeCells(arg)
          count += cells.filter(c => c.rawValue != null && c.rawValue !== '').length
        } else {
          const val = evaluateExpression(arg)
          if (val != null && val !== '') count++
        }
      }
      return count
    },
    
    // COUNTBLANK - 计数空单元格
    COUNTBLANK: (args) => {
      let count = 0
      for (const arg of args) {
        if (arg.includes(':')) {
          const cells = getRangeCells(arg)
          count += cells.filter(c => c.rawValue == null || c.rawValue === '').length
        }
      }
      return count
    },
    
    // COUNTIF - 条件计数
    COUNTIF: (args) => {
      if (args.length < 2) return 0
      const [rangeStr, criteria] = args
      const cells = getRangeCells(rangeStr)
      const criteriaValue = criteria.replace(/^"|"$/g, '')
      
      return cells.filter(cell => matchCriteria(cell.value, criteriaValue)).length
    },
    
    // COUNTIFS - 多条件计数
    COUNTIFS: (args) => {
      if (args.length < 2) return 0
      
      // 获取第一个范围作为基准
      const baseCells = getRangeCells(args[0])
      
      // 解析所有条件对
      const conditions = []
      for (let i = 0; i < args.length; i += 2) {
        if (i + 1 < args.length) {
          conditions.push({
            cells: getRangeCells(args[i]),
            criteria: args[i + 1].replace(/^"|"$/g, '')
          })
        }
      }
      
      let count = 0
      for (let idx = 0; idx < baseCells.length; idx++) {
        const allMatch = conditions.every(cond => {
          const cell = cond.cells[idx]
          return cell && matchCriteria(cell.value, cond.criteria)
        })
        if (allMatch) count++
      }
      return count
    },
    
    // -------------------- 逻辑函数 --------------------
    
    // IF - 条件判断
    IF: (args) => {
      const condition = evaluateExpression(args[0])
      const trueValue = args[1] ? evaluateExpression(args[1]) : true
      const falseValue = args[2] ? evaluateExpression(args[2]) : false
      return condition ? trueValue : falseValue
    },
    
    // AND - 逻辑与
    AND: (args) => args.every(arg => !!evaluateExpression(arg)),
    
    // OR - 逻辑或
    OR: (args) => args.some(arg => !!evaluateExpression(arg)),
    
    // NOT - 逻辑非
    NOT: (args) => !evaluateExpression(args[0]),
    
    // IFERROR - 错误处理
    IFERROR: (args) => {
      try {
        const result = evaluateExpression(args[0])
        if (result == null || (typeof result === 'number' && isNaN(result))) {
          return evaluateExpression(args[1])
        }
        return result
      } catch {
        return evaluateExpression(args[1])
      }
    },
    
    // -------------------- 查找/引用函数 --------------------
    
    // VLOOKUP - 垂直查找
    VLOOKUP: (args) => {
      const lookupValue = evaluateExpression(args[0])
      const tableRangeStr = args[1]
      const colIndex = evaluateExpression(args[2])
      const exactMatch = args[3] ? evaluateExpression(args[3]) === false : true
      
      const cells = getRangeCells(tableRangeStr)
      if (cells.length === 0) return '#N/A'
      
      // 确定表格的列数
      const { sheetName, cellRef } = parseFullReference(tableRangeStr)
      const parts = cellRef.split(':')
      const start = parseCellAddr(parts[0])
      const end = parseCellAddr(parts[1])
      const numCols = end.c - start.c + 1
      const numRows = end.r - start.r + 1
      
      // 查找匹配行
      for (let r = 0; r < numRows; r++) {
        const firstColValue = cells[r * numCols]?.value
        
        if (exactMatch) {
          if (firstColValue === lookupValue || String(firstColValue) === String(lookupValue)) {
            const targetIdx = r * numCols + (colIndex - 1)
            return cells[targetIdx]?.value ?? '#N/A'
          }
        } else {
          // 近似匹配（假设已排序）
          if (firstColValue <= lookupValue) {
            const nextRowValue = cells[(r + 1) * numCols]?.value
            if (nextRowValue == null || nextRowValue > lookupValue) {
              const targetIdx = r * numCols + (colIndex - 1)
              return cells[targetIdx]?.value ?? '#N/A'
            }
          }
        }
      }
      return '#N/A'
    },
    
    // INDEX - 返回指定位置的值
    INDEX: (args) => {
      const rangeStr = args[0]
      const rowNum = evaluateExpression(args[1])
      const colNum = args[2] ? evaluateExpression(args[2]) : 1
      
      const { sheetName, cellRef } = parseFullReference(rangeStr)
      const parts = cellRef.split(':')
      const start = parseCellAddr(parts[0])
      const end = parseCellAddr(parts[1])
      const numCols = end.c - start.c + 1
      
      const cells = getRangeCells(rangeStr)
      const idx = (rowNum - 1) * numCols + (colNum - 1)
      return cells[idx]?.value ?? '#REF!'
    },
    
    // MATCH - 查找匹配位置
    MATCH: (args) => {
      const lookupValue = evaluateExpression(args[0])
      const rangeStr = args[1]
      const matchType = args[2] ? evaluateExpression(args[2]) : 1
      
      const values = getRangeValues(rangeStr)
      
      if (matchType === 0) {
        // 精确匹配
        const idx = values.findIndex(v => v === lookupValue || String(v) === String(lookupValue))
        return idx >= 0 ? idx + 1 : '#N/A'
      } else if (matchType === 1) {
        // 小于或等于
        let lastIdx = -1
        for (let i = 0; i < values.length; i++) {
          if (values[i] <= lookupValue) lastIdx = i
          else break
        }
        return lastIdx >= 0 ? lastIdx + 1 : '#N/A'
      } else {
        // 大于或等于
        for (let i = 0; i < values.length; i++) {
          if (values[i] >= lookupValue) return i + 1
        }
        return '#N/A'
      }
    },
    
    // OFFSET - 偏移引用
    OFFSET: (args) => {
      const refStr = args[0]
      const rowOffset = evaluateExpression(args[1])
      const colOffset = evaluateExpression(args[2])
      const height = args[3] ? evaluateExpression(args[3]) : 1
      const width = args[4] ? evaluateExpression(args[4]) : 1
      
      const { sheetName, cellRef } = parseFullReference(refStr)
      const addr = parseCellAddr(cellRef.split(':')[0])
      if (!addr) return '#REF!'
      
      const newRow = addr.r + rowOffset
      const newCol = addr.c + colOffset
      
      if (height === 1 && width === 1) {
        return getCellValue(`${getColumnLabel(newCol)}${newRow + 1}`)
      }
      
      // 返回范围的值（求和）
      const values = []
      for (let r = 0; r < height; r++) {
        for (let c = 0; c < width; c++) {
          values.push(getCellValue(`${getColumnLabel(newCol + c)}${newRow + r + 1}`))
        }
      }
      return values.filter(v => typeof v === 'number').reduce((a, b) => a + b, 0)
    },
    
    // -------------------- 文本函数 --------------------
    
    // LEFT - 左侧字符
    LEFT: (args) => {
      const text = String(evaluateExpression(args[0]) || '')
      const numChars = args[1] ? evaluateExpression(args[1]) : 1
      return text.substring(0, numChars)
    },
    
    // RIGHT - 右侧字符
    RIGHT: (args) => {
      const text = String(evaluateExpression(args[0]) || '')
      const numChars = args[1] ? evaluateExpression(args[1]) : 1
      return text.substring(text.length - numChars)
    },
    
    // MID - 中间字符
    MID: (args) => {
      const text = String(evaluateExpression(args[0]) || '')
      const startNum = evaluateExpression(args[1])
      const numChars = evaluateExpression(args[2])
      return text.substring(startNum - 1, startNum - 1 + numChars)
    },
    
    // LEN - 字符长度
    LEN: (args) => String(evaluateExpression(args[0]) || '').length,
    
    // EXACT - 精确比较
    EXACT: (args) => {
      const text1 = String(evaluateExpression(args[0]) || '')
      const text2 = String(evaluateExpression(args[1]) || '')
      return text1 === text2
    },
    
    // CONCATENATE / CONCAT - 连接文本
    CONCATENATE: (args) => args.map(a => String(evaluateExpression(a) || '')).join(''),
    CONCAT: (args) => args.map(a => String(evaluateExpression(a) || '')).join(''),
    
    // TEXT - 格式化文本
    TEXT: (args) => {
      const value = evaluateExpression(args[0])
      const format = String(args[1] || '').replace(/^"|"$/g, '')
      const valueStr = String(value)
      
      // 日期格式化：如 "0000-00-00" 将 "19950315" 转为 "1995-03-15"
      if (format.match(/^0+-0+-0+$/) && /^\d{8}$/.test(valueStr)) {
        return `${valueStr.substring(0, 4)}-${valueStr.substring(4, 6)}-${valueStr.substring(6, 8)}`
      }
      
      // 日期格式化：如 "yyyy-mm-dd" 将 "19950315" 转为 "1995-03-15"
      if (format.toLowerCase().match(/^y+-m+-d+$/) && /^\d{8}$/.test(valueStr)) {
        return `${valueStr.substring(0, 4)}-${valueStr.substring(4, 6)}-${valueStr.substring(6, 8)}`
      }
      
      // 日期格式化：如 "yyyy/mm/dd"
      if (format.toLowerCase().match(/^y+\/m+\/d+$/) && /^\d{8}$/.test(valueStr)) {
        return `${valueStr.substring(0, 4)}/${valueStr.substring(4, 6)}/${valueStr.substring(6, 8)}`
      }
      
      // 日期格式化：如 "yyyy年mm月dd日"
      if (format.includes('年') && format.includes('月') && /^\d{8}$/.test(valueStr)) {
        return `${valueStr.substring(0, 4)}年${valueStr.substring(4, 6)}月${valueStr.substring(6, 8)}日`
      }
      
      if (typeof value === 'number') {
        // 简单的数字格式化
        if (format.includes('0') && !format.includes('-')) {
          const decimals = (format.split('.')[1] || '').length
          return value.toFixed(decimals)
        }
        if (format.includes('%')) {
          return (value * 100).toFixed(0) + '%'
        }
        // 千位分隔符格式 #,##0
        if (format.includes(',')) {
          return value.toLocaleString('en-US')
        }
      }
      return String(value)
    },
    
    // TRIM - 去除空格
    TRIM: (args) => String(evaluateExpression(args[0]) || '').trim(),
    
    // UPPER - 转大写
    UPPER: (args) => String(evaluateExpression(args[0]) || '').toUpperCase(),
    
    // LOWER - 转小写
    LOWER: (args) => String(evaluateExpression(args[0]) || '').toLowerCase(),
    
    // -------------------- 日期函数 --------------------
    
    // TODAY - 今天日期（返回 Date 对象，便于 YEAR/MONTH/DAY 处理）
    TODAY: () => {
      const now = new Date()
      now.setHours(0, 0, 0, 0) // 只保留日期部分
      return now
    },
    
    // NOW - 当前日期时间
    NOW: () => new Date(),
    
    // YEAR - 获取年份
    YEAR: (args) => {
      const val = evaluateExpression(args[0])
      // 如果是 Date 对象
      if (val instanceof Date) return val.getFullYear()
      // 如果是字符串格式的日期 "2025-12-08"
      if (typeof val === 'string') {
        // 尝试 YYYY-MM-DD 格式
        const match = val.match(/^(\d{4})-(\d{2})-(\d{2})/)
        if (match) return parseInt(match[1], 10)
        // 尝试 Date 解析
        const date = new Date(val)
        if (!isNaN(date.getTime())) return date.getFullYear()
      }
      // 如果是 Excel 日期序列号
      if (typeof val === 'number' && val > 1000 && val < 100000) {
        // Excel 日期从 1900-01-01 开始
        const excelEpoch = new Date(1900, 0, 1)
        const date = new Date(excelEpoch.getTime() + (val - 1) * 24 * 60 * 60 * 1000)
        return date.getFullYear()
      }
      return new Date().getFullYear() // 默认返回当前年份
    },
    
    // MONTH - 获取月份
    MONTH: (args) => {
      const val = evaluateExpression(args[0])
      if (val instanceof Date) return val.getMonth() + 1
      if (typeof val === 'string') {
        const match = val.match(/^(\d{4})-(\d{2})-(\d{2})/)
        if (match) return parseInt(match[2], 10)
        const date = new Date(val)
        if (!isNaN(date.getTime())) return date.getMonth() + 1
      }
      if (typeof val === 'number' && val > 1000 && val < 100000) {
        const excelEpoch = new Date(1900, 0, 1)
        const date = new Date(excelEpoch.getTime() + (val - 1) * 24 * 60 * 60 * 1000)
        return date.getMonth() + 1
      }
      return new Date().getMonth() + 1
    },
    
    // DAY - 获取日期
    DAY: (args) => {
      const val = evaluateExpression(args[0])
      if (val instanceof Date) return val.getDate()
      if (typeof val === 'string') {
        const match = val.match(/^(\d{4})-(\d{2})-(\d{2})/)
        if (match) return parseInt(match[3], 10)
        const date = new Date(val)
        if (!isNaN(date.getTime())) return date.getDate()
      }
      if (typeof val === 'number' && val > 1000 && val < 100000) {
        const excelEpoch = new Date(1900, 0, 1)
        const date = new Date(excelEpoch.getTime() + (val - 1) * 24 * 60 * 60 * 1000)
        return date.getDate()
      }
      return new Date().getDate()
    },
    
    // -------------------- 信息函数 --------------------
    
    // ISBLANK - 是否为空
    ISBLANK: (args) => {
      const val = evaluateExpression(args[0])
      return val == null || val === ''
    },
    
    // ISNUMBER - 是否为数字
    ISNUMBER: (args) => typeof evaluateExpression(args[0]) === 'number',
    
    // ISTEXT - 是否为文本
    ISTEXT: (args) => typeof evaluateExpression(args[0]) === 'string'
  }
  
  // 条件匹配函数（支持通配符和比较运算符）
  const matchCriteria = (value, criteria) => {
    const criteriaStr = String(criteria)
    
    // 比较运算符
    if (criteriaStr.startsWith('>=')) {
      return value >= parseFloat(criteriaStr.slice(2))
    }
    if (criteriaStr.startsWith('<=')) {
      return value <= parseFloat(criteriaStr.slice(2))
    }
    if (criteriaStr.startsWith('<>')) {
      return String(value) !== criteriaStr.slice(2)
    }
    if (criteriaStr.startsWith('>')) {
      return value > parseFloat(criteriaStr.slice(1))
    }
    if (criteriaStr.startsWith('<')) {
      return value < parseFloat(criteriaStr.slice(1))
    }
    if (criteriaStr.startsWith('=')) {
      return String(value) === criteriaStr.slice(1)
    }
    
    // 通配符匹配
    if (criteriaStr.includes('*') || criteriaStr.includes('?')) {
      const regex = new RegExp('^' + criteriaStr.replace(/\*/g, '.*').replace(/\?/g, '.') + '$', 'i')
      return regex.test(String(value))
    }
    
    // 精确匹配
    return String(value) === criteriaStr || value === criteria
  }
  
  // 解析并计算表达式
  const evaluateExpression = (expr) => {
    if (expr == null) return 0
    expr = String(expr).trim()
    
    // 字符串字面量
    if ((expr.startsWith('"') && expr.endsWith('"')) || (expr.startsWith("'") && expr.endsWith("'"))) {
      return expr.slice(1, -1)
    }
    
    // 数字
    if (/^-?\d+\.?\d*$/.test(expr)) {
      return parseFloat(expr)
    }
    
    // 布尔值
    if (expr.toUpperCase() === 'TRUE') return true
    if (expr.toUpperCase() === 'FALSE') return false
    
    // 单元格引用（包括跨工作表）- 必须是完整的引用，不是表达式的一部分
    if (/^'?[^'!]*'?![A-Z]+\d+$/i.test(expr) || /^[A-Z]+\d+$/i.test(expr)) {
      return getCellValue(expr)
    }
    
    // ============================================================
    // 复合表达式处理 - 支持 FUNC1()-FUNC2()+... 格式
    // ============================================================
    
    // 将表达式分解为标记（函数调用、运算符、数字、单元格引用）
    const tokenizeExpression = (expression) => {
      const tokens = []
      let i = 0
      
      while (i < expression.length) {
        // 跳过空格
        if (expression[i] === ' ') {
          i++
          continue
        }
        
        // 运算符
        if ('+-*/'.includes(expression[i])) {
          tokens.push({ type: 'operator', value: expression[i] })
          i++
          continue
        }
        
        // 数字
        if (/\d/.test(expression[i]) || (expression[i] === '-' && i === 0)) {
          let numStr = ''
          if (expression[i] === '-') {
            numStr = '-'
            i++
          }
          while (i < expression.length && /[\d.]/.test(expression[i])) {
            numStr += expression[i]
            i++
          }
          tokens.push({ type: 'number', value: parseFloat(numStr) })
          continue
        }
        
        // 字符串
        if (expression[i] === '"') {
          let str = ''
          i++ // 跳过开始引号
          while (i < expression.length && expression[i] !== '"') {
            str += expression[i]
            i++
          }
          i++ // 跳过结束引号
          tokens.push({ type: 'string', value: str })
          continue
        }
        
        // 函数调用或单元格引用
        if (/[A-Z']/i.test(expression[i])) {
          let token = ''
          
          // 处理带引号的工作表名（如 'Sheet1'!A1）
          if (expression[i] === "'") {
            while (i < expression.length && expression[i] !== '!') {
              token += expression[i]
              i++
            }
            if (expression[i] === '!') {
              token += expression[i]
              i++
            }
          }
          
          // 继续读取字母/数字
          while (i < expression.length && /[A-Z0-9_]/i.test(expression[i])) {
            token += expression[i]
            i++
          }
          
          // 检查是否是函数调用
          if (i < expression.length && expression[i] === '(') {
            // 找到匹配的右括号
            let depth = 1
            i++ // 跳过开始括号
            let argsStr = ''
            while (i < expression.length && depth > 0) {
              if (expression[i] === '(') depth++
              else if (expression[i] === ')') depth--
              if (depth > 0) argsStr += expression[i]
              i++
            }
            
            // 调用函数
            const funcName = token.toUpperCase()
            if (functions[funcName]) {
              const args = parseFunctionArgs(argsStr)
              const result = functions[funcName](args)
              tokens.push({ type: 'value', value: result })
            } else {
              tokens.push({ type: 'value', value: 0 })
            }
          } else {
            // 单元格引用
            const cellValue = getCellValue(token)
            // 如果是字符串形式的数字，转换为数字用于计算
            if (typeof cellValue === 'string' && /^-?\d+\.?\d*$/.test(cellValue)) {
              tokens.push({ type: 'value', value: parseFloat(cellValue) })
            } else {
              tokens.push({ type: 'value', value: cellValue })
            }
          }
          continue
        }
        
        // 括号
        if (expression[i] === '(') {
          let depth = 1
          i++
          let subExpr = ''
          while (i < expression.length && depth > 0) {
            if (expression[i] === '(') depth++
            else if (expression[i] === ')') depth--
            if (depth > 0) subExpr += expression[i]
            i++
          }
          tokens.push({ type: 'value', value: evaluateExpression(subExpr) })
          continue
        }
        
        i++ // 跳过未知字符
      }
      
      return tokens
    }
    
    // 计算标记序列
    const calculateTokens = (tokens) => {
      if (tokens.length === 0) return 0
      if (tokens.length === 1) {
        const t = tokens[0]
        return t.type === 'value' || t.type === 'number' ? t.value : 0
      }
      
      // 先处理乘除
      let i = 0
      while (i < tokens.length) {
        if (tokens[i].type === 'operator' && (tokens[i].value === '*' || tokens[i].value === '/')) {
          const left = tokens[i - 1]?.value ?? 0
          const right = tokens[i + 1]?.value ?? 0
          const leftNum = typeof left === 'string' ? (parseFloat(left) || 0) : (left || 0)
          const rightNum = typeof right === 'string' ? (parseFloat(right) || 0) : (right || 0)
          
          let result
          if (tokens[i].value === '*') {
            result = leftNum * rightNum
          } else {
            result = rightNum !== 0 ? leftNum / rightNum : 0
          }
          tokens.splice(i - 1, 3, { type: 'value', value: result })
          i = Math.max(0, i - 1)
        } else {
          i++
        }
      }
      
      // 再处理加减
      i = 0
      while (i < tokens.length) {
        if (tokens[i].type === 'operator' && (tokens[i].value === '+' || tokens[i].value === '-')) {
          const left = tokens[i - 1]?.value ?? 0
          const right = tokens[i + 1]?.value ?? 0
          const leftNum = typeof left === 'string' ? (parseFloat(left) || 0) : (left || 0)
          const rightNum = typeof right === 'string' ? (parseFloat(right) || 0) : (right || 0)
          
          let result
          if (tokens[i].value === '+') {
            result = leftNum + rightNum
          } else {
            result = leftNum - rightNum
          }
          tokens.splice(i - 1, 3, { type: 'value', value: result })
          i = Math.max(0, i - 1)
        } else {
          i++
        }
      }
      
      return tokens[0]?.value ?? 0
    }
    
    // 检测是否是复合表达式（包含运算符或多个函数）
    const hasOperator = /[+\-*/]/.test(expr.replace(/'[^']+'/g, '')) // 排除工作表名中的引号
    const hasFunctionCall = /[A-Z]+\(/i.test(expr)
    
    if (hasOperator || hasFunctionCall) {
      try {
        const tokens = tokenizeExpression(expr)
        if (tokens.length > 0) {
          return calculateTokens(tokens)
        }
      } catch (e) {
        console.warn('[Formula] 表达式解析错误:', expr, e.message)
      }
    }
    
    // 比较表达式
    const compareMatch = expr.match(/^(.+)(>=|<=|<>|>|<|=)(.+)$/)
    if (compareMatch) {
      const left = evaluateExpression(compareMatch[1])
      const right = evaluateExpression(compareMatch[3])
      switch (compareMatch[2]) {
        case '>=': return left >= right
        case '<=': return left <= right
        case '<>': return left !== right
        case '>': return left > right
        case '<': return left < right
        case '=': return left === right
      }
    }
    
    return expr
  }
  
  // 主计算函数
  const evaluateFormula = (formula, ws = currentWorksheet) => {
    try {
      return evaluateExpression(formula)
    } catch (e) {
      console.warn('[Formula Engine] 计算失败:', formula, e.message)
      return null
    }
  }
  
  return { evaluateFormula, getCellValue, getRangeValues }
}

// 简单公式计算器 - 兼容旧接口
function evaluateSimpleFormula(formula, worksheet, workbook = null) {
  // 如果没有 workbook，创建一个简单的包装
  const wb = workbook || { 
    getWorksheet: () => worksheet,
    worksheets: [worksheet]
  }
  const engine = createFormulaEngine(wb, worksheet)
  return engine.evaluateFormula(formula)
}

// 解析单元格地址（如 "A1" -> { r: 0, c: 0 }）
function parseCellAddress(address) {
  const match = address.toUpperCase().match(/^([A-Z]+)(\d+)$/)
  if (!match) return null
  
  let col = 0
  for (let i = 0; i < match[1].length; i++) {
    col = col * 26 + (match[1].charCodeAt(i) - 64)
  }
  return { r: parseInt(match[2], 10) - 1, c: col - 1 }
}

// 生成列标（如 0 -> "A", 25 -> "Z", 26 -> "AA"）
function getColumnLabel(i) {
  let label = ''
  let n = i
  while (n >= 0) {
    label = String.fromCharCode((n % 26) + 65) + label
    n = Math.floor(n / 26) - 1
  }
  return label
}

// 格式化单元格地址
function formatCellAddress(r, c) {
  return `${getColumnLabel(c)}${r + 1}`
}

// 【查询】读取单元格/区域
async function excelReadCells(filePath, sheetName, rangeOrCell) {
  try {
    const workbook = await getWorkbook(filePath)
    const worksheet = workbook.getWorksheet(sheetName)
    if (!worksheet) {
      return { success: false, error: `工作表 "${sheetName}" 不存在` }
    }
    
    // 解析范围：可以是单个单元格 "A1" 或范围 "A1:C5"
    const parts = rangeOrCell.toUpperCase().split(':')
    const start = parseCellAddress(parts[0])
    const end = parts.length > 1 ? parseCellAddress(parts[1]) : start
    
    if (!start || !end) {
      return { success: false, error: `无效的单元格地址: ${rangeOrCell}` }
    }
    
    const cells = []
    for (let r = start.r; r <= end.r; r++) {
      for (let c = start.c; c <= end.c; c++) {
        const cell = worksheet.getCell(r + 1, c + 1)
        // 安全获取文本值
        let textValue = ''
        try {
          const v = cell.value
          if (v != null) {
            if (typeof v === 'object' && v.richText) {
              textValue = v.richText.map(rt => rt.text || '').join('')
            } else if (typeof v === 'object' && v.result != null) {
              textValue = String(v.result)
            } else if (typeof v === 'object' && v.text != null) {
              textValue = String(v.text)
            } else {
              textValue = String(v)
            }
          }
        } catch (e) {
          textValue = ''
        }
        cells.push({
          address: formatCellAddress(r, c),
          r, c,
          value: cell.value,
          text: textValue,
          formula: cell.formula,
          type: cell.type
        })
      }
    }
    
    return { success: true, cells, range: rangeOrCell }
  } catch (error) {
    console.error('[Excel Read] 失败:', error)
    return { success: false, error: error.message }
  }
}

// 【查询】搜索单元格内容
async function excelSearch(filePath, sheetName, searchText, options = {}) {
  try {
    const workbook = await getWorkbook(filePath)
    const worksheet = workbook.getWorksheet(sheetName)
    if (!worksheet) {
      return { success: false, error: `工作表 "${sheetName}" 不存在` }
    }
    
    const results = []
    const { caseSensitive = false, matchWholeCell = false } = options
    const searchLower = caseSensitive ? searchText : searchText.toLowerCase()
    
    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        // 安全获取单元格文本
        let cellText = ''
        try {
          const v = cell.value
          if (v != null) {
            if (typeof v === 'object' && v.richText) {
              cellText = v.richText.map(rt => rt.text || '').join('')
            } else if (typeof v === 'object' && v.result != null) {
              cellText = String(v.result)
            } else if (typeof v === 'object' && v.text != null) {
              cellText = String(v.text)
            } else {
              cellText = String(v)
            }
          }
        } catch (e) {
          cellText = ''
        }
        const compareText = caseSensitive ? cellText : cellText.toLowerCase()
        
        let match = false
        if (matchWholeCell) {
          match = compareText === searchLower
        } else {
          match = compareText.includes(searchLower)
        }
        
        if (match) {
          results.push({
            address: formatCellAddress(rowNumber - 1, colNumber - 1),
            r: rowNumber - 1,
            c: colNumber - 1,
            value: cell.value,
            text: cellText
          })
        }
      })
    })
    
    return { success: true, results, count: results.length }
  } catch (error) {
    console.error('[Excel Search] 失败:', error)
    return { success: false, error: error.message }
  }
}

// 【修改】写入单元格
async function excelWriteCells(filePath, sheetName, cellUpdates) {
  try {
    // 检查文件是否被锁定
    try {
      const fd = fs.openSync(filePath, 'r+')
      fs.closeSync(fd)
    } catch (lockErr) {
      if (lockErr.code === 'EBUSY' || lockErr.code === 'EACCES') {
        return { 
          success: false, 
          error: '文件被其他程序占用（可能是 Excel 正在打开此文件）。请关闭 Excel 后重试。' 
        }
      }
    }
    
    // 清除缓存，重新加载文件
    clearWorkbookCache(filePath)
    const workbook = await getWorkbook(filePath)
    const worksheet = workbook.getWorksheet(sheetName)
    if (!worksheet) {
      return { success: false, error: `工作表 "${sheetName}" 不存在` }
    }
    
    console.log(`[Excel Write] 写入 ${cellUpdates.length} 个单元格到 ${sheetName}`)
    
    // cellUpdates: [{ address: "A1", value: "new value", style?: {...} }, ...]
    const updatedCells = []
    for (const update of cellUpdates) {
      const addr = parseCellAddress(update.address)
      if (!addr) {
        console.warn(`[Excel Write] 跳过无效地址: ${update.address}`)
        continue
      }
      
      const cell = worksheet.getCell(addr.r + 1, addr.c + 1)
      
      // 设置值（支持公式）
      if (update.value !== undefined) {
        if (typeof update.value === 'string' && update.value.startsWith('=')) {
          cell.value = { formula: update.value.slice(1) }
        } else {
          cell.value = update.value
        }
      }
      
      // 设置样式
      if (update.style) {
        if (update.style.font) {
          cell.font = { ...cell.font, ...update.style.font }
        }
        if (update.style.fill) {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: update.style.fill.fgColor || update.style.fill
          }
        }
        if (update.style.alignment) {
          cell.alignment = { ...cell.alignment, ...update.style.alignment }
        }
        if (update.style.border) {
          cell.border = { ...cell.border, ...update.style.border }
        }
        if (update.style.numFmt) {
          cell.numFmt = update.style.numFmt
        }
      }
      
      updatedCells.push(update.address)
    }
    
    // 保存文件
    await saveWorkbook(filePath)
    clearWorkbookCache(filePath) // 清除缓存以便重新读取
    
    console.log(`[Excel Write] 成功写入 ${updatedCells.length} 个单元格`)
    return { success: true, updatedCells, count: updatedCells.length }
  } catch (error) {
    console.error('[Excel Write] 失败:', error)
    return { success: false, error: error.message }
  }
}

// 【新增】插入行
async function excelInsertRows(filePath, sheetName, startRow, count = 1, data = null) {
  try {
    // 清除缓存，重新加载文件
    clearWorkbookCache(filePath)
    const workbook = await getWorkbook(filePath)
    const worksheet = workbook.getWorksheet(sheetName)
    if (!worksheet) {
      return { success: false, error: `工作表 "${sheetName}" 不存在` }
    }
    
    console.log(`[Excel Insert Rows] 在第 ${startRow} 行插入 ${count} 行`)
    
    // 准备要插入的行数据
    let rowsToInsert = []
    if (data && Array.isArray(data) && data.length > 0) {
      // 使用提供的数据
      rowsToInsert = data.slice(0, count)
      // 如果数据不够，填充空行
      while (rowsToInsert.length < count) {
        rowsToInsert.push([])
      }
    } else {
      // 创建空行
      for (let i = 0; i < count; i++) {
        rowsToInsert.push([])
      }
    }
    
    // ExcelJS insertRows: 第二个参数是行数据数组
    worksheet.insertRows(startRow, rowsToInsert)
    
    await saveWorkbook(filePath)
    clearWorkbookCache(filePath) // 清除缓存以便重新读取
    
    console.log(`[Excel Insert Rows] 成功插入 ${count} 行`)
    return { success: true, insertedAt: startRow, count }
  } catch (error) {
    console.error('[Excel Insert Rows] 失败:', error)
    return { success: false, error: error.message }
  }
}

// 【新增】插入列
async function excelInsertColumns(filePath, sheetName, startCol, count = 1) {
  try {
    // 清除缓存，重新加载文件
    clearWorkbookCache(filePath)
    const workbook = await getWorkbook(filePath)
    const worksheet = workbook.getWorksheet(sheetName)
    if (!worksheet) {
      return { success: false, error: `工作表 "${sheetName}" 不存在` }
    }
    
    console.log(`[Excel Insert Columns] 在第 ${startCol} 列插入 ${count} 列`)
    
    // ExcelJS spliceColumns(start, deleteCount, ...insert)
    // 第二个参数 0 表示不删除，后面的参数是要插入的列数据
    // 每个列数据是一个数组，代表该列所有行的值
    const emptyColumns = []
    for (let i = 0; i < count; i++) {
      emptyColumns.push([]) // 空列
    }
    worksheet.spliceColumns(startCol, 0, ...emptyColumns)
    
    await saveWorkbook(filePath)
    clearWorkbookCache(filePath)
    
    console.log(`[Excel Insert Columns] 成功插入 ${count} 列`)
    return { success: true, insertedAt: startCol, count }
  } catch (error) {
    console.error('[Excel Insert Columns] 失败:', error)
    return { success: false, error: error.message }
  }
}

// 【新增】新建工作表
async function excelAddSheet(filePath, sheetName) {
  try {
    // 清除缓存，重新加载文件
    clearWorkbookCache(filePath)
    const workbook = await getWorkbook(filePath)
    
    // 检查是否已存在
    if (workbook.getWorksheet(sheetName)) {
      return { success: false, error: `工作表 "${sheetName}" 已存在` }
    }
    
    console.log(`[Excel Add Sheet] 新建工作表: ${sheetName}`)
    
    workbook.addWorksheet(sheetName)
    await saveWorkbook(filePath)
    clearWorkbookCache(filePath)
    
    console.log(`[Excel Add Sheet] 成功创建工作表: ${sheetName}`)
    return { success: true, sheetName }
  } catch (error) {
    console.error('[Excel Add Sheet] 失败:', error)
    return { success: false, error: error.message }
  }
}

// 【删除】删除行
async function excelDeleteRows(filePath, sheetName, startRow, count = 1) {
  try {
    // 清除缓存，重新加载文件
    clearWorkbookCache(filePath)
    const workbook = await getWorkbook(filePath)
    const worksheet = workbook.getWorksheet(sheetName)
    if (!worksheet) {
      return { success: false, error: `工作表 "${sheetName}" 不存在` }
    }
    
    console.log(`[Excel Delete Rows] 删除第 ${startRow} 行开始的 ${count} 行`)
    
    // ExcelJS spliceRows(start, count) - 从 start 行开始删除 count 行
    worksheet.spliceRows(startRow, count)
    
    await saveWorkbook(filePath)
    clearWorkbookCache(filePath) // 清除缓存以便重新读取
    
    console.log(`[Excel Delete Rows] 成功删除 ${count} 行`)
    return { success: true, deletedFrom: startRow, count }
  } catch (error) {
    console.error('[Excel Delete Rows] 失败:', error)
    return { success: false, error: error.message }
  }
}

// 【删除】删除列
async function excelDeleteColumns(filePath, sheetName, startCol, count = 1) {
  try {
    // 清除缓存，重新加载文件
    clearWorkbookCache(filePath)
    const workbook = await getWorkbook(filePath)
    const worksheet = workbook.getWorksheet(sheetName)
    if (!worksheet) {
      return { success: false, error: `工作表 "${sheetName}" 不存在` }
    }
    
    console.log(`[Excel Delete Columns] 删除第 ${startCol} 列开始的 ${count} 列`)
    
    // ExcelJS spliceColumns(start, deleteCount)
    worksheet.spliceColumns(startCol, count)
    
    await saveWorkbook(filePath)
    clearWorkbookCache(filePath)
    
    console.log(`[Excel Delete Columns] 成功删除 ${count} 列`)
    return { success: true, deletedFrom: startCol, count }
  } catch (error) {
    console.error('[Excel Delete Columns] 失败:', error)
    return { success: false, error: error.message }
  }
}

// 【删除】删除工作表
async function excelDeleteSheet(filePath, sheetName) {
  try {
    // 清除缓存，重新加载文件
    clearWorkbookCache(filePath)
    const workbook = await getWorkbook(filePath)
    const worksheet = workbook.getWorksheet(sheetName)
    if (!worksheet) {
      return { success: false, error: `工作表 "${sheetName}" 不存在` }
    }
    
    console.log(`[Excel Delete Sheet] 删除工作表: ${sheetName}, id: ${worksheet.id}`)
    
    workbook.removeWorksheet(worksheet.id)
    await saveWorkbook(filePath)
    clearWorkbookCache(filePath)
    
    console.log(`[Excel Delete Sheet] 成功删除工作表: ${sheetName}`)
    return { success: true, deletedSheet: sheetName }
  } catch (error) {
    console.error('[Excel Delete Sheet] 失败:', error)
    return { success: false, error: error.message }
  }
}

// 【查询】获取工作表列表
async function excelListSheets(filePath) {
  try {
    const workbook = await getWorkbook(filePath)
    const sheets = []
    
    workbook.eachSheet((worksheet) => {
      sheets.push({
        name: worksheet.name,
        rowCount: worksheet.rowCount,
        columnCount: worksheet.columnCount
      })
    })
    
    return { success: true, sheets }
  } catch (error) {
    console.error('[Excel List Sheets] 失败:', error)
    return { success: false, error: error.message }
  }
}

// 【修改】合并单元格
async function excelMergeCells(filePath, sheetName, range) {
  try {
    // 清除缓存，重新加载文件
    clearWorkbookCache(filePath)
    const workbook = await getWorkbook(filePath)
    const worksheet = workbook.getWorksheet(sheetName)
    if (!worksheet) {
      return { success: false, error: `工作表 "${sheetName}" 不存在` }
    }
    
    console.log(`[Excel Merge Cells] 合并单元格: ${range}`)
    
    worksheet.mergeCells(range)
    await saveWorkbook(filePath)
    clearWorkbookCache(filePath)
    
    console.log(`[Excel Merge Cells] 成功合并: ${range}`)
    return { success: true, mergedRange: range }
  } catch (error) {
    console.error('[Excel Merge Cells] 失败:', error)
    return { success: false, error: error.message }
  }
}

// 【修改】取消合并单元格
async function excelUnmergeCells(filePath, sheetName, range) {
  try {
    // 清除缓存，重新加载文件
    clearWorkbookCache(filePath)
    const workbook = await getWorkbook(filePath)
    const worksheet = workbook.getWorksheet(sheetName)
    if (!worksheet) {
      return { success: false, error: `工作表 "${sheetName}" 不存在` }
    }
    
    console.log(`[Excel Unmerge Cells] 取消合并: ${range}`)
    
    worksheet.unMergeCells(range)
    await saveWorkbook(filePath)
    clearWorkbookCache(filePath)
    
    console.log(`[Excel Unmerge Cells] 成功取消合并: ${range}`)
    return { success: true, unmergedRange: range }
  } catch (error) {
    console.error('[Excel Unmerge Cells] 失败:', error)
    return { success: false, error: error.message }
  }
}

// 【新增】批量设置公式
async function excelSetFormula(filePath, sheetName, formulas) {
  try {
    clearWorkbookCache(filePath)
    const workbook = await getWorkbook(filePath)
    const worksheet = workbook.getWorksheet(sheetName)
    if (!worksheet) {
      return { success: false, error: `工作表 "${sheetName}" 不存在` }
    }
    
    console.log(`[Excel Formula] 设置 ${formulas.length} 个公式到 ${sheetName}`)
    
    const setFormulas = []
    for (const item of formulas) {
      const { address, formula, numberFormat } = item
      const addr = parseCellAddress(address)
      if (!addr) continue
      
      const cell = worksheet.getCell(addr.r + 1, addr.c + 1)
      
      // 设置公式（去掉开头的 = 如果有的话）
      const formulaText = formula.startsWith('=') ? formula.slice(1) : formula
      cell.value = { formula: formulaText }
      
      // 设置数字格式（可选）
      if (numberFormat) {
        cell.numFmt = numberFormat
      }
      
      setFormulas.push({ address, formula: formulaText })
    }
    
    await saveWorkbook(filePath)
    clearWorkbookCache(filePath)
    
    console.log(`[Excel Formula] 成功设置 ${setFormulas.length} 个公式`)
    return { success: true, formulas: setFormulas, count: setFormulas.length }
  } catch (error) {
    console.error('[Excel Formula] 失败:', error)
    return { success: false, error: error.message }
  }
}

// 【新增】排序数据
async function excelSort(filePath, sheetName, options) {
  try {
    const { range, column, ascending = true, hasHeader = true } = options
    
    clearWorkbookCache(filePath)
    const workbook = await getWorkbook(filePath)
    const worksheet = workbook.getWorksheet(sheetName)
    if (!worksheet) {
      return { success: false, error: `工作表 "${sheetName}" 不存在` }
    }
    
    console.log(`[Excel Sort] 排序 ${sheetName} 范围 ${range} 按列 ${column}`)
    
    // 解析范围
    const rangeMatch = range.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/)
    if (!rangeMatch) {
      return { success: false, error: `无效的范围格式: ${range}` }
    }
    
    const startCol = columnToNumber(rangeMatch[1])
    const startRow = parseInt(rangeMatch[2])
    const endCol = columnToNumber(rangeMatch[3])
    const endRow = parseInt(rangeMatch[4])
    
    // 确定排序列的索引
    const sortColIndex = columnToNumber(column) - startCol
    
    // 收集数据
    const rows = []
    const dataStartRow = hasHeader ? startRow + 1 : startRow
    
    for (let r = dataStartRow; r <= endRow; r++) {
      const rowData = []
      for (let c = startCol; c <= endCol; c++) {
        const cell = worksheet.getCell(r, c)
        rowData.push({
          value: cell.value,
          style: {
            font: cell.font,
            fill: cell.fill,
            alignment: cell.alignment,
            border: cell.border,
            numFmt: cell.numFmt
          }
        })
      }
      rows.push(rowData)
    }
    
    // 排序
    rows.sort((a, b) => {
      let valA = a[sortColIndex]?.value
      let valB = b[sortColIndex]?.value
      
      // 处理公式结果
      if (valA && typeof valA === 'object' && valA.result !== undefined) valA = valA.result
      if (valB && typeof valB === 'object' && valB.result !== undefined) valB = valB.result
      
      // 处理 null/undefined
      if (valA == null && valB == null) return 0
      if (valA == null) return ascending ? 1 : -1
      if (valB == null) return ascending ? -1 : 1
      
      // 数字比较
      const numA = typeof valA === 'number' ? valA : parseFloat(valA)
      const numB = typeof valB === 'number' ? valB : parseFloat(valB)
      
      if (!isNaN(numA) && !isNaN(numB)) {
        return ascending ? numA - numB : numB - numA
      }
      
      // 字符串比较
      const strA = String(valA).toLowerCase()
      const strB = String(valB).toLowerCase()
      return ascending ? strA.localeCompare(strB, 'zh-CN') : strB.localeCompare(strA, 'zh-CN')
    })
    
    // 写回数据
    for (let i = 0; i < rows.length; i++) {
      const rowData = rows[i]
      const r = dataStartRow + i
      for (let j = 0; j < rowData.length; j++) {
        const c = startCol + j
        const cell = worksheet.getCell(r, c)
        const data = rowData[j]
        
        cell.value = data.value
        if (data.style.font) cell.font = data.style.font
        if (data.style.fill) cell.fill = data.style.fill
        if (data.style.alignment) cell.alignment = data.style.alignment
        if (data.style.border) cell.border = data.style.border
        if (data.style.numFmt) cell.numFmt = data.style.numFmt
      }
    }
    
    await saveWorkbook(filePath)
    clearWorkbookCache(filePath)
    
    console.log(`[Excel Sort] 成功排序 ${rows.length} 行`)
    return { success: true, sortedRows: rows.length, column, ascending }
  } catch (error) {
    console.error('[Excel Sort] 失败:', error)
    return { success: false, error: error.message }
  }
}

// 辅助函数：列字母转数字
function columnToNumber(col) {
  let result = 0
  for (let i = 0; i < col.length; i++) {
    result = result * 26 + (col.charCodeAt(i) - 64)
  }
  return result
}

// 【新增】设置条件格式
async function excelConditionalFormat(filePath, sheetName, options) {
  try {
    const { range, rules } = options
    
    clearWorkbookCache(filePath)
    const workbook = await getWorkbook(filePath)
    const worksheet = workbook.getWorksheet(sheetName)
    if (!worksheet) {
      return { success: false, error: `工作表 "${sheetName}" 不存在` }
    }
    
    console.log(`[Excel ConditionalFormat] 设置条件格式到 ${sheetName} 范围 ${range}`)
    
    // ExcelJS 支持的条件格式
    const conditionalFormattings = []
    
    for (const rule of rules) {
      const cfRule = {
        ref: range,
        rules: []
      }
      
      if (rule.type === 'cellIs') {
        // 单元格值条件
        cfRule.rules.push({
          type: 'cellIs',
          operator: rule.operator, // greaterThan, lessThan, equal, between, etc.
          formulae: Array.isArray(rule.value) ? rule.value : [rule.value],
          style: {
            fill: rule.fill ? {
              type: 'pattern',
              pattern: 'solid',
              bgColor: rule.fill.bgColor || rule.fill
            } : undefined,
            font: rule.font
          }
        })
      } else if (rule.type === 'colorScale') {
        // 色阶
        cfRule.rules.push({
          type: 'colorScale',
          cfvo: [
            { type: 'min' },
            { type: 'max' }
          ],
          color: [
            { argb: rule.minColor || 'FFF8696B' },
            { argb: rule.maxColor || 'FF63BE7B' }
          ]
        })
      } else if (rule.type === 'dataBar') {
        // 数据条
        cfRule.rules.push({
          type: 'dataBar',
          minLength: 0,
          maxLength: 100,
          showValue: true,
          gradient: true,
          color: { argb: rule.color || 'FF638EC6' }
        })
      } else if (rule.type === 'containsText') {
        // 包含文本
        cfRule.rules.push({
          type: 'containsText',
          operator: 'containsText',
          text: rule.text,
          style: {
            fill: rule.fill ? {
              type: 'pattern',
              pattern: 'solid',
              bgColor: rule.fill.bgColor || rule.fill
            } : undefined,
            font: rule.font
          }
        })
      }
      
      conditionalFormattings.push(cfRule)
    }
    
    // 添加条件格式
    worksheet.addConditionalFormatting(...conditionalFormattings)
    
    await saveWorkbook(filePath)
    clearWorkbookCache(filePath)
    
    console.log(`[Excel ConditionalFormat] 成功设置 ${rules.length} 条规则`)
    return { success: true, rulesApplied: rules.length }
  } catch (error) {
    console.error('[Excel ConditionalFormat] 失败:', error)
    return { success: false, error: error.message }
  }
}

// 【新增】自动填充/序列填充
async function excelAutoFill(filePath, sheetName, options) {
  try {
    const { sourceRange, targetRange, fillType = 'copy' } = options
    
    clearWorkbookCache(filePath)
    const workbook = await getWorkbook(filePath)
    const worksheet = workbook.getWorksheet(sheetName)
    if (!worksheet) {
      return { success: false, error: `工作表 "${sheetName}" 不存在` }
    }
    
    console.log(`[Excel AutoFill] 从 ${sourceRange} 填充到 ${targetRange}`)
    
    // 解析源范围
    const srcMatch = sourceRange.match(/([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?/)
    if (!srcMatch) {
      return { success: false, error: `无效的源范围: ${sourceRange}` }
    }
    
    const srcStartCol = columnToNumber(srcMatch[1])
    const srcStartRow = parseInt(srcMatch[2])
    const srcEndCol = srcMatch[3] ? columnToNumber(srcMatch[3]) : srcStartCol
    const srcEndRow = srcMatch[4] ? parseInt(srcMatch[4]) : srcStartRow
    
    // 解析目标范围
    const tgtMatch = targetRange.match(/([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?/)
    if (!tgtMatch) {
      return { success: false, error: `无效的目标范围: ${targetRange}` }
    }
    
    const tgtStartCol = columnToNumber(tgtMatch[1])
    const tgtStartRow = parseInt(tgtMatch[2])
    const tgtEndCol = tgtMatch[3] ? columnToNumber(tgtMatch[3]) : tgtStartCol
    const tgtEndRow = tgtMatch[4] ? parseInt(tgtMatch[4]) : tgtStartRow
    
    // 收集源数据
    const sourceData = []
    for (let r = srcStartRow; r <= srcEndRow; r++) {
      const rowData = []
      for (let c = srcStartCol; c <= srcEndCol; c++) {
        const cell = worksheet.getCell(r, c)
        rowData.push({
          value: cell.value,
          style: {
            font: cell.font,
            fill: cell.fill,
            alignment: cell.alignment,
            border: cell.border,
            numFmt: cell.numFmt
          }
        })
      }
      sourceData.push(rowData)
    }
    
    // 填充目标范围
    let filledCount = 0
    const srcRows = sourceData.length
    const srcCols = sourceData[0]?.length || 0
    
    for (let r = tgtStartRow; r <= tgtEndRow; r++) {
      for (let c = tgtStartCol; c <= tgtEndCol; c++) {
        const srcRowIdx = (r - tgtStartRow) % srcRows
        const srcColIdx = (c - tgtStartCol) % srcCols
        const srcCell = sourceData[srcRowIdx]?.[srcColIdx]
        
        if (srcCell) {
          const cell = worksheet.getCell(r, c)
          
          if (fillType === 'series' && typeof srcCell.value === 'number') {
            // 序列填充：数字递增
            const increment = r - tgtStartRow + 1
            cell.value = srcCell.value + increment
          } else if (fillType === 'formula' && srcCell.value?.formula) {
            // 公式填充：调整相对引用（简化处理）
            const rowOffset = r - srcStartRow
            const colOffset = c - srcStartCol
            let formula = srcCell.value.formula
            
            // 简单调整行号（更复杂的需要完整的公式解析器）
            formula = formula.replace(/([A-Z]+)(\d+)/g, (match, col, row) => {
              const newRow = parseInt(row) + rowOffset
              return col + newRow
            })
            
            cell.value = { formula }
          } else {
            // 复制填充
            cell.value = srcCell.value
          }
          
          // 复制样式
          if (srcCell.style.font) cell.font = srcCell.style.font
          if (srcCell.style.fill) cell.fill = srcCell.style.fill
          if (srcCell.style.alignment) cell.alignment = srcCell.style.alignment
          if (srcCell.style.border) cell.border = srcCell.style.border
          if (srcCell.style.numFmt) cell.numFmt = srcCell.style.numFmt
          
          filledCount++
        }
      }
    }
    
    await saveWorkbook(filePath)
    clearWorkbookCache(filePath)
    
    console.log(`[Excel AutoFill] 成功填充 ${filledCount} 个单元格`)
    return { success: true, filledCells: filledCount, fillType }
  } catch (error) {
    console.error('[Excel AutoFill] 失败:', error)
    return { success: false, error: error.message }
  }
}

// 【新增】设置列宽和行高
async function excelSetDimensions(filePath, sheetName, options) {
  try {
    const { columns = [], rows = [] } = options
    
    clearWorkbookCache(filePath)
    const workbook = await getWorkbook(filePath)
    const worksheet = workbook.getWorksheet(sheetName)
    if (!worksheet) {
      return { success: false, error: `工作表 "${sheetName}" 不存在` }
    }
    
    console.log(`[Excel Dimensions] 设置 ${columns.length} 列宽, ${rows.length} 行高`)
    
    // 设置列宽
    for (const col of columns) {
      const colNum = typeof col.column === 'string' ? columnToNumber(col.column) : col.column
      const column = worksheet.getColumn(colNum)
      if (col.width !== undefined) column.width = col.width
      if (col.hidden !== undefined) column.hidden = col.hidden
      if (col.style) column.style = col.style
    }
    
    // 设置行高
    for (const row of rows) {
      const rowObj = worksheet.getRow(row.row)
      if (row.height !== undefined) rowObj.height = row.height
      if (row.hidden !== undefined) rowObj.hidden = row.hidden
    }
    
    await saveWorkbook(filePath)
    clearWorkbookCache(filePath)
    
    return { success: true, columnsSet: columns.length, rowsSet: rows.length }
  } catch (error) {
    console.error('[Excel Dimensions] 失败:', error)
    return { success: false, error: error.message }
  }
}

// 【新增】创建图表（简化版）
async function excelAddChart(filePath, sheetName, options) {
  try {
    const { 
      type = 'column', // column, bar, line, pie, scatter, area
      dataRange,
      title = '',
      position = { col: 1, row: 1 },
      size = { width: 600, height: 400 }
    } = options
    
    clearWorkbookCache(filePath)
    const workbook = await getWorkbook(filePath)
    const worksheet = workbook.getWorksheet(sheetName)
    if (!worksheet) {
      return { success: false, error: `工作表 "${sheetName}" 不存在` }
    }
    
    console.log(`[Excel Chart] 添加 ${type} 图表到 ${sheetName}`)
    
    // ExcelJS 对图表的支持有限，这里我们创建一个基本的图表配置
    // 实际上 ExcelJS 不直接支持图表创建，需要通过其他方式
    // 这里我们记录图表配置，用户可以在 Excel 中手动创建
    
    // 作为替代，我们可以在指定位置添加一个注释说明
    const cell = worksheet.getCell(position.row, position.col)
    cell.note = {
      texts: [
        { text: `图表配置:\n类型: ${type}\n数据范围: ${dataRange}\n标题: ${title || '无'}`, font: { size: 10 } }
      ]
    }
    
    await saveWorkbook(filePath)
    clearWorkbookCache(filePath)
    
    // 返回图表信息（实际图表需要用 Excel 打开后手动创建）
    return { 
      success: true, 
      message: 'ExcelJS 不直接支持图表创建，已在指定位置添加配置说明。请在 Excel 中手动创建图表。',
      chartConfig: { type, dataRange, title, position, size }
    }
  } catch (error) {
    console.error('[Excel Chart] 失败:', error)
    return { success: false, error: error.message }
  }
}

// 【新增】计算公式（获取公式计算结果）
async function excelCalculate(filePath, sheetName, addresses) {
  try {
    clearWorkbookCache(filePath)
    const workbook = await getWorkbook(filePath)
    const worksheet = workbook.getWorksheet(sheetName)
    if (!worksheet) {
      return { success: false, error: `工作表 "${sheetName}" 不存在` }
    }
    
    console.log(`[Excel Calculate] 获取 ${addresses.length} 个单元格的计算结果`)
    
    const results = []
    for (const address of addresses) {
      const addr = parseCellAddress(address)
      if (!addr) continue
      
      const cell = worksheet.getCell(addr.r + 1, addr.c + 1)
      const value = cell.value
      
      let result = {
        address,
        value: null,
        formula: null,
        type: 'unknown'
      }
      
      if (value && typeof value === 'object') {
        if (value.formula) {
          result.formula = value.formula
          result.value = value.result !== undefined ? value.result : '计算中...'
          result.type = 'formula'
        } else if (value.richText) {
          result.value = value.richText.map(t => t.text).join('')
          result.type = 'richText'
        } else if (value.hyperlink) {
          result.value = value.text || value.hyperlink
          result.type = 'hyperlink'
        }
      } else {
        result.value = value
        result.type = typeof value
      }
      
      results.push(result)
    }
    
    return { success: true, results }
  } catch (error) {
    console.error('[Excel Calculate] 失败:', error)
    return { success: false, error: error.message }
  }
}

// 【新增】创建新的 Excel 文件
async function excelCreate(filePath, options = {}) {
  try {
    const { 
      sheets = [{ name: 'Sheet1', data: [] }], 
      openAfterCreate = true,
      defaultStyle = null,  // 全局默认样式
      headerStyle = null    // 表头默认样式
    } = options
    
    console.log(`[Excel Create] 创建新文件: ${filePath}`)
    
    // 检查文件是否已存在
    if (fs.existsSync(filePath)) {
      console.log(`[Excel Create] 文件已存在，将覆盖: ${filePath}`)
    }
    
    // 创建新工作簿
    const workbook = new ExcelJS.Workbook()
    workbook.creator = '智启文档 AI'
    workbook.created = new Date()
    
    // 默认表头样式（如果用户没有指定）
    const defaultHeaderStyle = headerStyle || {
      font: { bold: true, size: 12, color: { argb: 'FFFFFFFF' } },
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } },
      alignment: { horizontal: 'center', vertical: 'middle' },
      border: {
        top: { style: 'thin', color: { argb: 'FF000000' } },
        bottom: { style: 'thin', color: { argb: 'FF000000' } },
        left: { style: 'thin', color: { argb: 'FF000000' } },
        right: { style: 'thin', color: { argb: 'FF000000' } }
      }
    }
    
    // 默认数据单元格样式
    const defaultCellStyle = defaultStyle || {
      font: { size: 11 },
      alignment: { vertical: 'middle' },
      border: {
        top: { style: 'thin', color: { argb: 'FFD0D0D0' } },
        bottom: { style: 'thin', color: { argb: 'FFD0D0D0' } },
        left: { style: 'thin', color: { argb: 'FFD0D0D0' } },
        right: { style: 'thin', color: { argb: 'FFD0D0D0' } }
      }
    }
    
    // 辅助函数：解析简化的样式参数
    const parseSimpleStyle = (styleStr) => {
      if (!styleStr || typeof styleStr !== 'string') return null
      const style = {}
      // 解析类似 "bold,center,#FF0000,14" 的简化格式
      const parts = styleStr.split(',').map(s => s.trim())
      for (const part of parts) {
        if (part === 'bold') {
          style.font = style.font || {}
          style.font.bold = true
        } else if (part === 'italic') {
          style.font = style.font || {}
          style.font.italic = true
        } else if (part === 'underline') {
          style.font = style.font || {}
          style.font.underline = true
        } else if (part === 'center') {
          style.alignment = style.alignment || {}
          style.alignment.horizontal = 'center'
        } else if (part === 'left') {
          style.alignment = style.alignment || {}
          style.alignment.horizontal = 'left'
        } else if (part === 'right') {
          style.alignment = style.alignment || {}
          style.alignment.horizontal = 'right'
        } else if (part.startsWith('#')) {
          // 颜色
          style.font = style.font || {}
          style.font.color = { argb: 'FF' + part.slice(1) }
        } else if (part.startsWith('bg#')) {
          // 背景色
          style.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + part.slice(3) } }
        } else if (/^\d+$/.test(part)) {
          // 字号
          style.font = style.font || {}
          style.font.size = parseInt(part)
        }
      }
      return Object.keys(style).length > 0 ? style : null
    }
    
    // 添加工作表和数据
    for (const sheetConfig of sheets) {
      const worksheet = workbook.addWorksheet(sheetConfig.name || 'Sheet1')
      
      // 是否应用默认样式（默认开启）
      const applyDefaultStyles = sheetConfig.applyDefaultStyles !== false
      // 第一行是否为表头（默认是）
      const firstRowIsHeader = sheetConfig.firstRowIsHeader !== false
      
      // 如果有数据，填充数据
      if (sheetConfig.data && Array.isArray(sheetConfig.data)) {
        sheetConfig.data.forEach((rowData, rowIndex) => {
          if (Array.isArray(rowData)) {
            const row = worksheet.getRow(rowIndex + 1)
            const isHeaderRow = rowIndex === 0 && firstRowIsHeader
            
            // 设置行高
            if (isHeaderRow) {
              row.height = sheetConfig.headerHeight || 25
            } else {
              row.height = sheetConfig.rowHeight || 20
            }
            
            rowData.forEach((cellValue, colIndex) => {
              const cell = row.getCell(colIndex + 1)
              
              // 支持对象格式 { value: ..., style: ... } 或 { v: ..., s: ... }
              if (cellValue && typeof cellValue === 'object' && ('value' in cellValue || 'v' in cellValue)) {
                cell.value = cellValue.value ?? cellValue.v
                
                // 应用样式
                const cellStyle = cellValue.style || cellValue.s
                if (cellStyle) {
                  // 如果是字符串，解析简化格式
                  const parsedStyle = typeof cellStyle === 'string' ? parseSimpleStyle(cellStyle) : cellStyle
                  if (parsedStyle) {
                    if (parsedStyle.font) cell.font = { ...cell.font, ...parsedStyle.font }
                    if (parsedStyle.fill) cell.fill = parsedStyle.fill
                    if (parsedStyle.alignment) cell.alignment = { ...cell.alignment, ...parsedStyle.alignment }
                    if (parsedStyle.border) cell.border = parsedStyle.border
                    if (parsedStyle.numFmt) cell.numFmt = parsedStyle.numFmt
                  }
                }
              } else {
                // 检测公式字符串（以=开头）
                if (typeof cellValue === 'string' && cellValue.startsWith('=')) {
                  cell.value = { formula: cellValue.slice(1) }
                } else {
                  cell.value = cellValue
                }
              }
              
              // 应用默认样式
              if (applyDefaultStyles) {
                if (isHeaderRow) {
                  // 表头样式（如果单元格没有自定义样式）
                  if (!cell.font || !cell.font.bold) {
                    cell.font = { ...defaultHeaderStyle.font, ...cell.font }
                  }
                  if (!cell.fill) {
                    cell.fill = defaultHeaderStyle.fill
                  }
                  if (!cell.alignment) {
                    cell.alignment = defaultHeaderStyle.alignment
                  }
                  if (!cell.border) {
                    cell.border = defaultHeaderStyle.border
                  }
                } else {
                  // 数据行样式
                  if (!cell.font) {
                    cell.font = defaultCellStyle.font
                  }
                  if (!cell.alignment) {
                    cell.alignment = defaultCellStyle.alignment
                  }
                  if (!cell.border) {
                    cell.border = defaultCellStyle.border
                  }
                }
              }
            })
            row.commit()
          }
        })
      }
      
      // 设置列宽（如果提供）
      if (sheetConfig.columnWidths && Array.isArray(sheetConfig.columnWidths)) {
        sheetConfig.columnWidths.forEach((width, index) => {
          if (width) {
            worksheet.getColumn(index + 1).width = width
          }
        })
      } else if (sheetConfig.data && sheetConfig.data.length > 0) {
        // 自动计算列宽
        const firstRow = sheetConfig.data[0]
        if (Array.isArray(firstRow)) {
          firstRow.forEach((_, colIndex) => {
            // 根据内容计算列宽，最小10，最大50
            let maxWidth = 10
            sheetConfig.data.forEach(rowData => {
              if (Array.isArray(rowData) && rowData[colIndex] != null) {
                const val = rowData[colIndex]
                const text = typeof val === 'object' ? String(val.value ?? val.v ?? '') : String(val)
                // 中文字符算2个宽度
                const len = text.split('').reduce((acc, char) => acc + (char.charCodeAt(0) > 127 ? 2 : 1), 0)
                maxWidth = Math.max(maxWidth, Math.min(len + 2, 50))
              }
            })
            worksheet.getColumn(colIndex + 1).width = maxWidth
          })
        }
      }
      
      // 设置合并单元格（如果提供）
      if (sheetConfig.merges && Array.isArray(sheetConfig.merges)) {
        sheetConfig.merges.forEach(range => {
          try {
            worksheet.mergeCells(range)
          } catch (e) {
            console.warn(`[Excel Create] 合并单元格失败: ${range}`, e.message)
          }
        })
      }
      
      // 冻结表头
      if (firstRowIsHeader && sheetConfig.freezeHeader !== false) {
        worksheet.views = [{ state: 'frozen', ySplit: 1 }]
      }
    }
    
    // 确保目录存在
    const dir = path.dirname(filePath)
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true })
    }
    
    // 保存文件
    await workbook.xlsx.writeFile(filePath)
    
    console.log(`[Excel Create] 文件创建成功: ${filePath}`)
    
    return { 
      success: true, 
      filePath,
      sheetsCreated: sheets.map(s => s.name || 'Sheet1'),
      openAfterCreate
    }
  } catch (error) {
    console.error('[Excel Create] 失败:', error)
    return { success: false, error: error.message }
  }
}

// 关闭文件时清除缓存
async function excelClose(filePath) {
  clearWorkbookCache(filePath)
  return { success: true }
}

// 重新加载 Excel 文件（刷新缓存）
async function excelReload(filePath) {
  clearWorkbookCache(filePath)
  // 触发重新打开
  return await excelOpen(filePath)
}

// 【新增】设置自动筛选 (AutoFilter)
async function excelSetFilter(filePath, sheetName, options) {
  try {
    const { range, remove = false } = options || {}
    
    clearWorkbookCache(filePath)
    const workbook = await getWorkbook(filePath)
    const worksheet = workbook.getWorksheet(sheetName)
    if (!worksheet) {
      return { success: false, error: `工作表 "${sheetName}" 不存在` }
    }
    
    if (remove) {
      worksheet.autoFilter = undefined
      console.log(`[Excel Filter] 清除 ${sheetName} 的自动筛选`)
    } else if (range) {
      worksheet.autoFilter = range
      console.log(`[Excel Filter] 设置 ${sheetName} 的自动筛选范围: ${range}`)
    } else {
      // 如果没有指定范围，自动检测数据范围
      const dimensions = worksheet.dimensions
      if (dimensions) {
        const autoRange = `${dimensions.top}:${dimensions.bottom}`.replace(/(\d+):(\d+)/, (m, t, b) => {
          const topAddr = worksheet.getCell(parseInt(t), 1).address
          const bottomAddr = worksheet.getCell(parseInt(t), dimensions.right).address
          return `${topAddr}:${bottomAddr}`
        })
        worksheet.autoFilter = { from: dimensions.tl, to: { row: 1, col: dimensions.right } }
        console.log(`[Excel Filter] 自动设置筛选范围`)
      }
    }
    
    await saveWorkbook(filePath)
    clearWorkbookCache(filePath)
    
    return { 
      success: true, 
      message: remove ? '已清除自动筛选' : `已设置自动筛选范围: ${range || '自动检测'}`
    }
  } catch (error) {
    console.error('[Excel Filter] 失败:', error)
    return { success: false, error: error.message }
  }
}

// 【新增】设置数据验证 (Data Validation)
async function excelSetValidation(filePath, sheetName, options) {
  try {
    const { 
      range, 
      type = 'list', // list, whole, decimal, date, textLength
      values,        // 对于 list 类型
      min,           // 对于数值类型
      max,           // 对于数值类型
      allowBlank = true,
      showError = true,
      errorTitle = '输入错误',
      errorMessage = '请输入有效的值',
      remove = false
    } = options || {}
    
    if (!range) {
      return { success: false, error: '请指定单元格范围 (range)' }
    }
    
    clearWorkbookCache(filePath)
    const workbook = await getWorkbook(filePath)
    const worksheet = workbook.getWorksheet(sheetName)
    if (!worksheet) {
      return { success: false, error: `工作表 "${sheetName}" 不存在` }
    }
    
    // 解析范围并应用到每个单元格
    const rangeMatch = range.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/i)
    if (!rangeMatch && !range.match(/^[A-Z]+\d+$/i)) {
      return { success: false, error: `无效的范围格式: ${range}` }
    }
    
    const applyValidation = (cell) => {
      if (remove) {
        cell.dataValidation = undefined
        return
      }
      
      const validation = {
        type: type,
        allowBlank: allowBlank,
        showErrorMessage: showError,
        errorTitle: errorTitle,
        error: errorMessage
      }
      
      if (type === 'list' && values) {
        // 列表类型
        const listValues = Array.isArray(values) ? values : [values]
        validation.formulae = ['"' + listValues.join(',') + '"']
        validation.showDropDown = true
      } else if (type === 'whole' || type === 'decimal') {
        // 数值类型
        validation.operator = 'between'
        validation.formulae = [min !== undefined ? min : 0, max !== undefined ? max : 999999999]
      } else if (type === 'textLength') {
        // 文本长度
        validation.operator = 'between'
        validation.formulae = [min !== undefined ? min : 0, max !== undefined ? max : 255]
      }
      
      cell.dataValidation = validation
    }
    
    if (rangeMatch) {
      // 范围格式 A1:B10
      const startCol = rangeMatch[1].toUpperCase()
      const startRow = parseInt(rangeMatch[2])
      const endCol = rangeMatch[3].toUpperCase()
      const endRow = parseInt(rangeMatch[4])
      
      for (let row = startRow; row <= endRow; row++) {
        for (let colCode = startCol.charCodeAt(0); colCode <= endCol.charCodeAt(0); colCode++) {
          const col = String.fromCharCode(colCode)
          const cell = worksheet.getCell(`${col}${row}`)
          applyValidation(cell)
        }
      }
    } else {
      // 单个单元格
      const cell = worksheet.getCell(range)
      applyValidation(cell)
    }
    
    await saveWorkbook(filePath)
    clearWorkbookCache(filePath)
    
    console.log(`[Excel Validation] ${remove ? '清除' : '设置'}数据验证: ${range}, 类型: ${type}`)
    
    return { 
      success: true, 
      message: remove ? `已清除 ${range} 的数据验证` : `已设置 ${range} 的${type === 'list' ? '下拉列表' : '数据'}验证`
    }
  } catch (error) {
    console.error('[Excel Validation] 失败:', error)
    return { success: false, error: error.message }
  }
}

// 【新增】设置超链接 (Hyperlink)
async function excelSetHyperlink(filePath, sheetName, options) {
  try {
    const { 
      cell, 
      url, 
      text,
      tooltip,
      remove = false
    } = options || {}
    
    if (!cell) {
      return { success: false, error: '请指定单元格地址 (cell)' }
    }
    
    clearWorkbookCache(filePath)
    const workbook = await getWorkbook(filePath)
    const worksheet = workbook.getWorksheet(sheetName)
    if (!worksheet) {
      return { success: false, error: `工作表 "${sheetName}" 不存在` }
    }
    
    const targetCell = worksheet.getCell(cell)
    
    if (remove) {
      // 清除超链接，保留文本
      const currentText = targetCell.text || targetCell.value
      targetCell.value = currentText
      targetCell.font = { ...targetCell.font, color: undefined, underline: false }
    } else {
      if (!url) {
        return { success: false, error: '请指定链接地址 (url)' }
      }
      
      // 设置超链接
      targetCell.value = {
        text: text || url,
        hyperlink: url,
        tooltip: tooltip || url
      }
      
      // 设置超链接样式
      targetCell.font = {
        ...targetCell.font,
        color: { argb: 'FF0000FF' },
        underline: true
      }
    }
    
    await saveWorkbook(filePath)
    clearWorkbookCache(filePath)
    
    console.log(`[Excel Hyperlink] ${remove ? '清除' : '设置'}超链接: ${cell}`)
    
    return { 
      success: true, 
      message: remove ? `已清除 ${cell} 的超链接` : `已在 ${cell} 设置超链接: ${url}`
    }
  } catch (error) {
    console.error('[Excel Hyperlink] 失败:', error)
    return { success: false, error: error.message }
  }
}

// 【新增】查找替换 (Find and Replace)
async function excelFindReplace(filePath, sheetName, options) {
  try {
    const { 
      find, 
      replace = '',
      matchCase = false,
      matchWholeCell = false,
      allSheets = false
    } = options || {}
    
    if (!find) {
      return { success: false, error: '请指定要查找的内容 (find)' }
    }
    
    clearWorkbookCache(filePath)
    const workbook = await getWorkbook(filePath)
    
    let totalCount = 0
    const results = []
    
    const processSheet = (worksheet) => {
      let sheetCount = 0
      
      worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
          let cellValue = cell.value
          
          // 处理富文本
          if (cellValue && typeof cellValue === 'object' && cellValue.richText) {
            cellValue = cellValue.richText.map(r => r.text).join('')
          }
          
          // 处理超链接
          if (cellValue && typeof cellValue === 'object' && cellValue.text) {
            cellValue = cellValue.text
          }
          
          if (typeof cellValue === 'string') {
            const searchValue = matchCase ? find : find.toLowerCase()
            const compareValue = matchCase ? cellValue : cellValue.toLowerCase()
            
            let shouldReplace = false
            if (matchWholeCell) {
              shouldReplace = compareValue === searchValue
            } else {
              shouldReplace = compareValue.includes(searchValue)
            }
            
            if (shouldReplace) {
              // 执行替换
              const regex = new RegExp(
                find.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'),
                matchCase ? 'g' : 'gi'
              )
              
              if (matchWholeCell) {
                cell.value = replace
              } else {
                cell.value = cellValue.replace(regex, replace)
              }
              
              sheetCount++
              results.push({
                sheet: worksheet.name,
                cell: cell.address,
                oldValue: cellValue,
                newValue: cell.value
              })
            }
          }
        })
      })
      
      return sheetCount
    }
    
    if (allSheets) {
      workbook.eachSheet((worksheet) => {
        totalCount += processSheet(worksheet)
      })
    } else {
      const worksheet = workbook.getWorksheet(sheetName)
      if (!worksheet) {
        return { success: false, error: `工作表 "${sheetName}" 不存在` }
      }
      totalCount = processSheet(worksheet)
    }
    
    if (totalCount > 0) {
      await saveWorkbook(filePath)
      clearWorkbookCache(filePath)
    }
    
    console.log(`[Excel Find/Replace] 替换了 ${totalCount} 处: "${find}" → "${replace}"`)
    
    return { 
      success: true, 
      count: totalCount,
      message: totalCount > 0 
        ? `已将 ${totalCount} 处 "${find}" 替换为 "${replace}"`
        : `未找到 "${find}"`,
      details: results.slice(0, 20) // 最多返回20条详情
    }
  } catch (error) {
    console.error('[Excel Find/Replace] 失败:', error)
    return { success: false, error: error.message }
  }
}

// 【新增】插入图表（生成图片版本 - 使用 QuickChart API）
async function excelInsertChart(filePath, sheetName, options) {
  try {
    const { 
      type = 'column', // column, bar, line, pie, area, scatter, doughnut
      dataRange,
      title = '',
      position = 'E1',
      width = 500,
      height = 300,
      backgroundColor = '#ffffff'
    } = options || {}
    
    if (!dataRange) {
      return { success: false, error: '请指定数据范围 (dataRange)' }
    }
    
    clearWorkbookCache(filePath)
    const workbook = await getWorkbook(filePath)
    const worksheet = workbook.getWorksheet(sheetName)
    if (!worksheet) {
      return { success: false, error: `工作表 "${sheetName}" 不存在` }
    }
    
    console.log(`[Excel Chart] 图表请求: 类型=${type}, 数据=${dataRange}, 位置=${position}`)
    
    // 1. 解析数据范围并读取数据
    const rangeMatch = dataRange.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/i)
    if (!rangeMatch) {
      return { success: false, error: `无效的数据范围格式: ${dataRange}` }
    }
    
    const startCol = rangeMatch[1].toUpperCase()
    const startRow = parseInt(rangeMatch[2])
    const endCol = rangeMatch[3].toUpperCase()
    const endRow = parseInt(rangeMatch[4])
    
    // 读取数据
    const labels = []
    const datasets = []
    const dataColumns = {}
    
    // 假设第一行是标题，第一列是标签
    for (let row = startRow; row <= endRow; row++) {
      const labelCell = worksheet.getCell(`${startCol}${row}`)
      let labelValue = labelCell.value
      if (labelValue && typeof labelValue === 'object') {
        labelValue = labelValue.text || labelValue.result || String(labelValue)
      }
      
      if (row === startRow) {
        // 第一行是系列标题
        for (let colCode = startCol.charCodeAt(0) + 1; colCode <= endCol.charCodeAt(0); colCode++) {
          const col = String.fromCharCode(colCode)
          const headerCell = worksheet.getCell(`${col}${row}`)
          let headerValue = headerCell.value
          if (headerValue && typeof headerValue === 'object') {
            headerValue = headerValue.text || headerValue.result || String(headerValue)
          }
          dataColumns[col] = {
            label: headerValue || `系列${col}`,
            data: []
          }
        }
      } else {
        // 数据行
        labels.push(labelValue || `行${row}`)
        for (let colCode = startCol.charCodeAt(0) + 1; colCode <= endCol.charCodeAt(0); colCode++) {
          const col = String.fromCharCode(colCode)
          const dataCell = worksheet.getCell(`${col}${row}`)
          let cellValue = dataCell.value
          if (cellValue && typeof cellValue === 'object') {
            cellValue = cellValue.result || cellValue.text || 0
          }
          const numValue = typeof cellValue === 'number' ? cellValue : parseFloat(cellValue) || 0
          if (dataColumns[col]) {
            dataColumns[col].data.push(numValue)
          }
        }
      }
    }
    
    // 构建 datasets
    const colors = [
      'rgba(54, 162, 235, 0.8)',
      'rgba(255, 99, 132, 0.8)',
      'rgba(75, 192, 192, 0.8)',
      'rgba(255, 206, 86, 0.8)',
      'rgba(153, 102, 255, 0.8)',
      'rgba(255, 159, 64, 0.8)',
      'rgba(199, 199, 199, 0.8)',
      'rgba(83, 102, 255, 0.8)'
    ]
    
    const borderColors = colors.map(c => c.replace('0.8', '1'))
    
    let colorIndex = 0
    for (const col in dataColumns) {
      datasets.push({
        label: dataColumns[col].label,
        data: dataColumns[col].data,
        backgroundColor: type === 'pie' || type === 'doughnut' 
          ? colors.slice(0, dataColumns[col].data.length)
          : colors[colorIndex % colors.length],
        borderColor: type === 'pie' || type === 'doughnut'
          ? borderColors.slice(0, dataColumns[col].data.length)
          : borderColors[colorIndex % borderColors.length],
        borderWidth: 1
      })
      colorIndex++
    }
    
    // 如果只有一列数据（没有标题行），直接用第一列作为标签
    if (datasets.length === 0 && labels.length > 0) {
      // 单列数据，第一列作为标签，需要重新解析
      labels.length = 0
      const singleData = []
      for (let row = startRow; row <= endRow; row++) {
        const labelCell = worksheet.getCell(`${startCol}${row}`)
        const valueCell = worksheet.getCell(`${endCol}${row}`)
        let labelValue = labelCell.value
        let dataValue = valueCell.value
        
        if (labelValue && typeof labelValue === 'object') {
          labelValue = labelValue.text || labelValue.result || String(labelValue)
        }
        if (dataValue && typeof dataValue === 'object') {
          dataValue = dataValue.result || dataValue.text || 0
        }
        
        labels.push(labelValue || `项${row}`)
        singleData.push(typeof dataValue === 'number' ? dataValue : parseFloat(dataValue) || 0)
      }
      
      datasets.push({
        label: title || '数据',
        data: singleData,
        backgroundColor: type === 'pie' || type === 'doughnut'
          ? colors.slice(0, singleData.length)
          : colors[0],
        borderColor: type === 'pie' || type === 'doughnut'
          ? borderColors.slice(0, singleData.length)
          : borderColors[0],
        borderWidth: 1
      })
    }
    
    console.log(`[Excel Chart] 标签: ${labels.length} 个, 数据系列: ${datasets.length} 个`)
    
    // 2. 构建 QuickChart 配置
    const chartTypeMap = {
      'column': 'bar',
      'bar': 'horizontalBar',
      'line': 'line',
      'pie': 'pie',
      'doughnut': 'doughnut',
      'area': 'line',
      'scatter': 'scatter'
    }
    
    const chartConfig = {
      type: chartTypeMap[type] || 'bar',
      data: {
        labels: labels,
        datasets: datasets
      },
      options: {
        title: {
          display: !!title,
          text: title,
          fontSize: 16
        },
        legend: {
          display: datasets.length > 1 || type === 'pie' || type === 'doughnut'
        },
        plugins: {
          datalabels: {
            display: type === 'pie' || type === 'doughnut',
            color: '#fff',
            font: { weight: 'bold' }
          }
        }
      }
    }
    
    // 面积图特殊处理
    if (type === 'area') {
      chartConfig.data.datasets = chartConfig.data.datasets.map(ds => ({
        ...ds,
        fill: true
      }))
    }
    
    // 3. 调用 QuickChart API 生成图片
    // 使用 GET 方法更稳定
    const chartConfigEncoded = encodeURIComponent(JSON.stringify(chartConfig))
    const quickChartUrl = `https://quickchart.io/chart?c=${chartConfigEncoded}&w=${width}&h=${height}&bkg=${encodeURIComponent(backgroundColor)}&f=png`
    
    console.log('[Excel Chart] 调用 QuickChart API...')
    console.log('[Excel Chart] 图表配置:', JSON.stringify(chartConfig).substring(0, 200))
    
    const response = await fetch(quickChartUrl)
    
    if (!response.ok) {
      const errorText = await response.text()
      console.error('[Excel Chart] API 错误:', errorText)
      throw new Error(`QuickChart API 返回错误: ${response.status} ${response.statusText}`)
    }
    
    const arrayBuffer = await response.arrayBuffer()
    const imageBuffer = Buffer.from(arrayBuffer)
    
    if (imageBuffer.length < 1000) {
      // 图片太小，可能是错误响应
      console.error('[Excel Chart] 图片数据太小，可能生成失败:', imageBuffer.length)
      throw new Error('图表生成失败：返回数据异常')
    }
    
    console.log(`[Excel Chart] 图片生成成功, 大小: ${imageBuffer.length} bytes`)
    
    // 4. 保存图片到临时文件（ExcelJS 对 buffer 支持有时不稳定）
    const tempDir = require('os').tmpdir()
    const tempImagePath = path.join(tempDir, `chart_${Date.now()}.png`)
    fs.writeFileSync(tempImagePath, imageBuffer)
    console.log(`[Excel Chart] 临时图片保存到: ${tempImagePath}`)
    
    // 5. 将图片插入到 Excel（使用文件路径而不是 buffer）
    const imageId = workbook.addImage({
      filename: tempImagePath,
      extension: 'png'
    })
    
    // 解析位置
    const posMatch = position.match(/([A-Z]+)(\d+)/i)
    if (!posMatch) {
      // 清理临时文件
      try { fs.unlinkSync(tempImagePath) } catch {}
      return { success: false, error: `无效的位置格式: ${position}` }
    }
    
    const posCol = posMatch[1].toUpperCase().charCodeAt(0) - 64 // A=1, B=2...
    const posRow = parseInt(posMatch[2])
    
    // 使用 tl + br 方式定位（更稳定）
    // 计算结束位置
    const imgEndCol = posCol - 1 + Math.ceil(width / 72)  // 假设每列约 72 像素
    const imgEndRow = posRow - 1 + Math.ceil(height / 20) // 假设每行约 20 像素
    
    worksheet.addImage(imageId, {
      tl: { col: posCol - 1, row: posRow - 1 },
      br: { col: imgEndCol, row: imgEndRow }
    })
    
    await saveWorkbook(filePath)
    clearWorkbookCache(filePath)
    
    // 清理临时文件
    try { fs.unlinkSync(tempImagePath) } catch {}
    
    console.log(`[Excel Chart] 图表图片已插入到 ${position}`)
    
    return { 
      success: true, 
      message: `已在 ${position} 插入${type === 'column' ? '柱状' : type === 'line' ? '折线' : type === 'pie' ? '饼' : type}图`,
      chartConfig: { type, dataRange, title, position, width, height, labelsCount: labels.length, datasetsCount: datasets.length }
    }
  } catch (error) {
    console.error('[Excel Chart] 失败:', error)
    return { success: false, error: error.message }
  }
}

// HTML 转义函数

  return {
    checkLibreOffice,
    excelConvertXlsToXlsx,
    excelOpen,
    excelReadCells,
    excelSearch,
    excelWriteCells,
    excelInsertRows,
    excelInsertColumns,
    excelAddSheet,
    excelDeleteRows,
    excelDeleteColumns,
    excelDeleteSheet,
    excelListSheets,
    excelMergeCells,
    excelUnmergeCells,
    excelSetFormula,
    excelSort,
    excelConditionalFormat,
    excelAutoFill,
    excelSetDimensions,
    excelAddChart,
    excelCalculate,
    excelCreate,
    excelClose,
    excelReload,
    excelSetFilter,
    excelSetValidation,
    excelSetHyperlink,
    excelFindReplace,
    excelInsertChart,
  }
}

module.exports = {
  createExcelService,
}
