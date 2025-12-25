import { useMemo, useState, useCallback, useRef, useEffect, CSSProperties, MouseEvent, KeyboardEvent, ClipboardEvent } from 'react'
import type { ExcelSheet, ExcelCell, ExcelCellUpdate } from '../types/electron'
import {
  ClipboardPaste,
  Scissors,
  Copy,
  Paintbrush,
  Bold as BoldIcon,
  Italic as ItalicIcon,
  Underline as UnderlineIcon,
  Type as FontIcon,
  ChevronUp,
  ChevronDown,
  AlignLeft,
  AlignCenter,
  AlignRight,
  AlignJustify,
  AlignVerticalDistributeStart,
  AlignVerticalDistributeCenter,
  AlignVerticalDistributeEnd,
  Percent,
  Minus,
  Plus,
  Table,
  Grid2X2,
  Grid3X3,
  Columns,
  Rows,
  Filter,
  Search,
  Sparkles,
  FileSearch,
  FileCode,
  Palette,
  Sigma,
  Undo2,
  Redo2,
  Trash2,
  RowsIcon,
  ColumnsIcon,
  Merge,
  SplitSquareHorizontal,
  Save,
  X,
  Check,
} from 'lucide-react'

type ExcelPreviewProps = {
  sheets: ExcelSheet[]
  initialSheetIndex?: number
  filePath?: string
  onRefresh?: () => Promise<boolean>
}

type MatrixCell = {
  cell?: ExcelCell
  merge?: { rowSpan: number; colSpan: number }
}

type SheetMatrix = {
  rows: MatrixCell[][]
  rowCount: number
  colCount: number
}

// 编辑历史记录类型
type HistoryAction = {
  type: 'cell_edit' | 'delete' | 'paste' | 'insert_row' | 'insert_col' | 'delete_row' | 'delete_col' | 'merge' | 'unmerge' | 'format'
  sheetIndex: number
  data: any
  undoData: any
}

// 剪贴板数据类型
type ClipboardData = {
  cells: { r: number; c: number; value: any; formula?: string }[]
  minR: number
  maxR: number
  minC: number
  maxC: number
  isCut: boolean
}

function convertCellStyle(cell?: ExcelCell): CSSProperties {
  // 默认使用单独的边框属性，避免与简写冲突
  const defaultBorder = '1px solid #d4d4d4'
  const style: CSSProperties = {
    padding: '2px 4px',
    fontSize: '11pt', // Excel 默认字号
    color: '#000',
    backgroundColor: '#fff',
    borderTop: defaultBorder,
    borderBottom: defaultBorder,
    borderLeft: defaultBorder,
    borderRight: defaultBorder,
    boxSizing: 'border-box',
    fontFamily: `'等线', 'Microsoft YaHei', '宋体', Calibri, sans-serif`, // 默认字体
    overflow: 'hidden',
    verticalAlign: 'middle', // 默认垂直居中
  }
  if (!cell?.s) return style

  const s = cell.s as any
  
  // 支持两种格式：ExcelJS 格式（嵌套对象）和 SheetJS 格式（扁平结构）
  // ExcelJS: { font: { name, sz, bold... }, fill: {...}, alignment: {...} }
  // SheetJS: { font: {...}, fgColor: {...}, alignment: {...} } 或直接 { bold: true, sz: 12 }
  
  const font = s.font || s || {}
  const fill = s.fill || {}
  const alignment = s.alignment || {}
  const border = s.border || {}

  // 字体名称
  const fallbacks = `'等线', 'Microsoft YaHei', '宋体', Calibri, sans-serif`
  const fontName = font.name || s.name
  if (fontName) {
    style.fontFamily = `'${fontName}', ${fallbacks}`
  }
  
  // 字号 - 支持多种格式
  // ExcelJS: font.sz (number)
  // SheetJS: font.sz (number) 或直接 s.sz
  const fontSize = font.sz || s.sz
  if (fontSize) {
    style.fontSize = `${Number(fontSize)}pt`
  }
  
  // 字体样式
  if (font.bold || s.bold) style.fontWeight = 'bold'
  if (font.italic || s.italic) style.fontStyle = 'italic'
  
  // 下划线
  const underline = font.underline || s.underline
  if (underline) {
    if (underline === 'double') {
      style.textDecoration = 'underline double'
    } else if (underline === true || underline === 'single') {
      style.textDecoration = 'underline'
    }
  }
  
  // 删除线
  if (font.strike || s.strike) {
    style.textDecoration = style.textDecoration ? `${style.textDecoration} line-through` : 'line-through'
  }
  
  // 字体颜色 - 支持多种格式
  const fontColor = font.color || s.color
  if (fontColor) {
    const argb = fontColor.argb || fontColor.rgb || (typeof fontColor === 'string' ? fontColor : null)
    if (argb) {
      // ARGB 格式，去掉 alpha
      const rgb = argb.length === 8 ? argb.slice(2) : argb
      if (rgb && rgb !== '000000') {
        style.color = `#${rgb}`
      }
    }
  }

  // 背景色 - 支持多种格式
  const fgColor = fill.fgColor || s.fgColor || s.patternFill?.fgColor
  if (fgColor) {
    const argb = fgColor.argb || fgColor.rgb || (typeof fgColor === 'string' ? fgColor : null)
    if (argb) {
      const rgb = argb.length === 8 ? argb.slice(2) : argb
      // 排除默认的白色和透明
      if (rgb && rgb !== 'FFFFFF' && rgb !== '000000' && argb !== '00000000') {
        style.backgroundColor = `#${rgb}`
      }
    }
  }

  // 水平对齐 - 支持多种格式
  const hAlign = alignment.horizontal || s.horizontal
  if (hAlign) {
    if (hAlign === 'center' || hAlign === 'centerContinuous') {
      style.textAlign = 'center'
    } else if (hAlign === 'right') {
      style.textAlign = 'right'
    } else if (hAlign === 'left' || hAlign === 'general') {
      style.textAlign = 'left'
    } else if (hAlign === 'justify' || hAlign === 'distributed') {
      style.textAlign = 'justify'
    }
  } else if (cell?.t === 'n' || cell?.t === '2') { // 数字类型默认右对齐
    style.textAlign = 'right'
  }
  
  // 垂直对齐 - 支持多种格式
  const vAlign = alignment.vertical || s.vertical
  if (vAlign) {
    if (vAlign === 'middle' || vAlign === 'center') {
      style.verticalAlign = 'middle'
    } else if (vAlign === 'top') {
      style.verticalAlign = 'top'
    } else if (vAlign === 'bottom') {
      style.verticalAlign = 'bottom'
    }
  }
  
  // 文本换行 - 支持多种格式
  const wrapText = alignment.wrapText || s.wrapText
  if (wrapText) {
    style.whiteSpace = 'pre-wrap'
    style.wordBreak = 'break-word'
  } else {
    style.whiteSpace = 'nowrap'
  }
  
  // 缩进 - 支持多种格式
  const indent = alignment.indent || s.indent
  if (indent) {
    style.textIndent = `${indent * 10}px`
  }
  
  // 缩小字体填充 - 支持多种格式
  const shrinkToFit = alignment.shrinkToFit || s.shrinkToFit
  if (shrinkToFit) {
    style.whiteSpace = 'nowrap'
    style.textOverflow = 'ellipsis'
  }

  // 边框映射
  const applyBorder = (side: string, target: keyof CSSProperties) => {
    const b = border[side]
    if (!b || !b.style) return
    
    // 边框颜色
    let color = '#000'
    if (b.color) {
      const argb = b.color.argb || b.color.rgb
      if (argb) {
        const rgb = argb.length === 8 ? argb.slice(2) : argb
        color = `#${rgb}`
      }
    }
    
    // 边框宽度
    const borderStyle = b.style
    let width = '1px'
    if (borderStyle === 'medium' || borderStyle === 'mediumDashed') {
      width = '2px'
    } else if (borderStyle === 'thick') {
      width = '3px'
    }
    
    // 边框样式
    let lineStyle = 'solid'
    if (borderStyle === 'dashed' || borderStyle === 'mediumDashed') {
      lineStyle = 'dashed'
    } else if (borderStyle === 'dotted') {
      lineStyle = 'dotted'
    } else if (borderStyle === 'double') {
      lineStyle = 'double'
    }
    
    ;(style as any)[target] = `${width} ${lineStyle} ${color}`
  }
  
  applyBorder('top', 'borderTop')
  applyBorder('bottom', 'borderBottom')
  applyBorder('left', 'borderLeft')
  applyBorder('right', 'borderRight')

  return style
}

function buildMatrix(sheet: ExcelSheet): SheetMatrix {
  // 计算实际数据范围
  const dataRowEnd = Math.max(sheet.range.e.r + 1, 1)
  const dataColEnd = Math.max(sheet.range.e.c + 1, 1)
  // 像 Excel 一样显示更多的行和列
  // 最少显示 100 行，26 列（A-Z），或者数据范围 + 额外空白
  const rowCount = Math.max(dataRowEnd + 50, 100)
  const colCount = Math.max(dataColEnd + 10, 26)
  const rows: MatrixCell[][] = Array.from({ length: rowCount }, () =>
    Array.from({ length: colCount }, () => ({} as MatrixCell))
  )

  const skip = new Set<string>()

  // 合并
  sheet.merges.forEach((m) => {
    rows[m.s.r][m.s.c].merge = {
      rowSpan: m.e.r - m.s.r + 1,
      colSpan: m.e.c - m.s.c + 1,
    }
    for (let r = m.s.r; r <= m.e.r; r++) {
      for (let c = m.s.c; c <= m.e.c; c++) {
        if (r === m.s.r && c === m.s.c) continue
        skip.add(`${r}-${c}`)
      }
    }
  })

  // 填充单元格
  sheet.cells.forEach((cell) => {
    if (cell.r >= rowCount || cell.c >= colCount) return
    rows[cell.r][cell.c].cell = cell
  })

  return { rows, rowCount, colCount }
}

// 选区类型
type Selection = {
  start: { r: number; c: number }
  end: { r: number; c: number }
}

// 获取列标签（A, B, ..., Z, AA, AB, ...）
function getColumnLabel(i: number) {
  let label = ''
  let n = i
  while (n >= 0) {
    label = String.fromCharCode((n % 26) + 65) + label
    n = Math.floor(n / 26) - 1
  }
  return label
}

// 获取单元格地址（如 A1, B2）
function getCellAddress(r: number, c: number) {
  return `${getColumnLabel(c)}${r + 1}`
}

function ExcelPreview({ sheets, initialSheetIndex = 0, filePath, onRefresh }: ExcelPreviewProps) {
  const [activeIndex, setActiveIndex] = useState(initialSheetIndex)
  const [zoom, setZoom] = useState(100)
  
  // 本地数据状态（用于编辑）
  const [localSheets, setLocalSheets] = useState<ExcelSheet[]>(sheets)
  
  // 选区状态
  const [selection, setSelection] = useState<Selection>({
    start: { r: 0, c: 0 },
    end: { r: 0, c: 0 },
  })
  const [isSelecting, setIsSelecting] = useState(false)
  
  // 编辑状态
  const [editingCell, setEditingCell] = useState<{ r: number; c: number } | null>(null)
  const [editValue, setEditValue] = useState('')
  const [isFormulaBarEditing, setIsFormulaBarEditing] = useState(false)
  
  // 历史记录（撤销/重做）
  const [history, setHistory] = useState<HistoryAction[]>([])
  const [historyIndex, setHistoryIndex] = useState(-1)
  
  // 剪贴板
  const [clipboard, setClipboard] = useState<ClipboardData | null>(null)
  
  // 保存状态
  const [isSaving, setIsSaving] = useState(false)
  const [hasChanges, setHasChanges] = useState(false)
  
  // 状态消息
  const [statusMessage, setStatusMessage] = useState('就绪')
  
  const tableRef = useRef<HTMLTableElement>(null)
  const editInputRef = useRef<HTMLInputElement>(null)
  const formulaInputRef = useRef<HTMLInputElement>(null)
  const containerRef = useRef<HTMLDivElement>(null)
  
  const activeSheet = localSheets[activeIndex]
  
  // 同步外部 sheets 变化
  useEffect(() => {
    setLocalSheets(sheets)
    setHasChanges(false)
  }, [sheets])
  
  // 聚焦编辑输入框
  useEffect(() => {
    if (editingCell && editInputRef.current) {
      editInputRef.current.focus()
      editInputRef.current.select()
    }
  }, [editingCell])
  
  const handleZoomChange = (delta: number) => {
    setZoom((prev) => Math.min(200, Math.max(25, prev + delta)))
  }

  // 获取规范化的选区（确保 start <= end）
  const normalizedSelection = useMemo(() => {
    const minR = Math.min(selection.start.r, selection.end.r)
    const maxR = Math.max(selection.start.r, selection.end.r)
    const minC = Math.min(selection.start.c, selection.end.c)
    const maxC = Math.max(selection.start.c, selection.end.c)
    return { minR, maxR, minC, maxC }
  }, [selection])

  // 检查单元格是否在选区内
  const isCellSelected = useCallback(
    (r: number, c: number) => {
      const { minR, maxR, minC, maxC } = normalizedSelection
      return r >= minR && r <= maxR && c >= minC && c <= maxC
    },
    [normalizedSelection]
  )

  // 检查是否是选区的起始单元格（活动单元格）
  const isActiveCell = useCallback(
    (r: number, c: number) => {
      return r === selection.start.r && c === selection.start.c
    },
    [selection.start]
  )
  
  // 添加历史记录
  const pushHistory = useCallback((action: HistoryAction) => {
    setHistory(prev => {
      // 删除当前位置之后的历史
      const newHistory = prev.slice(0, historyIndex + 1)
      newHistory.push(action)
      // 限制历史记录数量
      if (newHistory.length > 100) {
        newHistory.shift()
      }
      return newHistory
    })
    setHistoryIndex(prev => Math.min(prev + 1, 99))
    setHasChanges(true)
  }, [historyIndex])
  
  // 更新单元格值（本地）
  const updateCellValue = useCallback((r: number, c: number, value: any, formula?: string) => {
    setLocalSheets(prev => {
      const newSheets = [...prev]
      const sheet = { ...newSheets[activeIndex] }
      const cells = [...sheet.cells]
      
      const existingIndex = cells.findIndex(cell => cell.r === r && cell.c === c)
      const newCell: ExcelCell = {
        r,
        c,
        v: value,
        w: String(value),
        f: formula,
        display: String(value),
      }
      
      if (existingIndex >= 0) {
        cells[existingIndex] = { ...cells[existingIndex], ...newCell }
      } else {
        cells.push(newCell)
      }
      
      sheet.cells = cells
      newSheets[activeIndex] = sheet
      return newSheets
    })
  }, [activeIndex])
  
  // 删除单元格值（本地）
  const deleteCellValue = useCallback((r: number, c: number) => {
    setLocalSheets(prev => {
      const newSheets = [...prev]
      const sheet = { ...newSheets[activeIndex] }
      const cells = sheet.cells.filter(cell => !(cell.r === r && cell.c === c))
      sheet.cells = cells
      newSheets[activeIndex] = sheet
      return newSheets
    })
  }, [activeIndex])
  
  // 获取单元格值
  const getCellValue = useCallback((row: number, col: number): { value: any; formula?: string } | null => {
    const cell = activeSheet?.cells.find(cell => cell.r === row && cell.c === col)
    if (!cell) return null
    return { value: cell.v, formula: cell.f }
  }, [activeSheet])
  
  // 确认编辑
  const confirmEdit = useCallback(() => {
    if (!editingCell) return
    
    const { r, c } = editingCell
    const oldCell = activeSheet?.cells.find(cell => cell.r === r && cell.c === c)
    const oldValue = oldCell?.v
    const oldFormula = oldCell?.f
    
    // 判断是公式还是普通值
    const isFormula = editValue.startsWith('=')
    const newValue = isFormula ? editValue : editValue
    const newFormula = isFormula ? editValue : undefined
    
    // 记录历史
    pushHistory({
      type: 'cell_edit',
      sheetIndex: activeIndex,
      data: { r, c, value: newValue, formula: newFormula },
      undoData: { r, c, value: oldValue, formula: oldFormula }
    })
    
    // 更新本地数据
    updateCellValue(r, c, isFormula ? editValue : editValue, newFormula)
    
    setEditingCell(null)
    setEditValue('')
    setIsFormulaBarEditing(false)
    setStatusMessage('单元格已修改')
  }, [editingCell, editValue, activeSheet, activeIndex, pushHistory, updateCellValue])
  
  // 取消编辑
  const cancelEdit = useCallback(() => {
    setEditingCell(null)
    setEditValue('')
    setIsFormulaBarEditing(false)
  }, [])
  
  // 开始编辑单元格
  const startEdit = useCallback((r: number, c: number, initialValue?: string) => {
    const cell = activeSheet?.cells.find(cell => cell.r === r && cell.c === c)
    const value = initialValue ?? (cell?.f || cell?.v || '')
    setEditingCell({ r, c })
    setEditValue(String(value))
  }, [activeSheet])
  
  // 删除选中单元格
  const deleteSelectedCells = useCallback(() => {
    const { minR, maxR, minC, maxC } = normalizedSelection
    const deletedCells: { r: number; c: number; value: any; formula?: string }[] = []
    
    for (let r = minR; r <= maxR; r++) {
      for (let c = minC; c <= maxC; c++) {
        const cell = activeSheet?.cells.find(cell => cell.r === r && cell.c === c)
        if (cell) {
          deletedCells.push({ r, c, value: cell.v, formula: cell.f })
        }
      }
    }
    
    if (deletedCells.length === 0) return
    
    // 记录历史
    pushHistory({
      type: 'delete',
      sheetIndex: activeIndex,
      data: { cells: deletedCells.map(c => ({ r: c.r, c: c.c })) },
      undoData: { cells: deletedCells }
    })
    
    // 删除单元格
    setLocalSheets(prev => {
      const newSheets = [...prev]
      const sheet = { ...newSheets[activeIndex] }
      const cells = sheet.cells.filter(cell => {
        return !(cell.r >= minR && cell.r <= maxR && cell.c >= minC && cell.c <= maxC)
      })
      sheet.cells = cells
      newSheets[activeIndex] = sheet
      return newSheets
    })
    
    setHasChanges(true)
    setStatusMessage(`已删除 ${deletedCells.length} 个单元格`)
  }, [normalizedSelection, activeSheet, activeIndex, pushHistory])
  
  // 复制选中单元格
  const copySelectedCells = useCallback((isCut: boolean = false) => {
    const { minR, maxR, minC, maxC } = normalizedSelection
    const cells: ClipboardData['cells'] = []
    
    for (let r = minR; r <= maxR; r++) {
      for (let c = minC; c <= maxC; c++) {
        const cell = activeSheet?.cells.find(cell => cell.r === r && cell.c === c)
        cells.push({
          r: r - minR,
          c: c - minC,
          value: cell?.v ?? '',
          formula: cell?.f
        })
      }
    }
    
    setClipboard({
      cells,
      minR,
      maxR,
      minC,
      maxC,
      isCut
    })
    
    // 同时复制到系统剪贴板（纯文本格式）
    const textData = []
    for (let r = minR; r <= maxR; r++) {
      const rowData = []
      for (let c = minC; c <= maxC; c++) {
        const cell = activeSheet?.cells.find(cell => cell.r === r && cell.c === c)
        rowData.push(cell?.w ?? cell?.v ?? '')
      }
      textData.push(rowData.join('\t'))
    }
    
    navigator.clipboard.writeText(textData.join('\n')).catch(() => {})
    
    setStatusMessage(isCut ? '已剪切' : '已复制')
  }, [normalizedSelection, activeSheet])
  
  // 粘贴
  const pasteFromClipboard = useCallback(async () => {
    const { r: startR, c: startC } = selection.start
    
    // 优先使用内部剪贴板
    if (clipboard) {
      const pastedCells: { r: number; c: number; value: any; formula?: string }[] = []
      const undoCells: { r: number; c: number; value: any; formula?: string }[] = []
      
      clipboard.cells.forEach(cellData => {
        const targetR = startR + cellData.r
        const targetC = startC + cellData.c
        
        // 保存旧值用于撤销
        const oldCell = activeSheet?.cells.find(c => c.r === targetR && c.c === targetC)
        undoCells.push({
          r: targetR,
          c: targetC,
          value: oldCell?.v,
          formula: oldCell?.f
        })
        
        pastedCells.push({
          r: targetR,
          c: targetC,
          value: cellData.value,
          formula: cellData.formula
        })
      })
      
      // 记录历史
      pushHistory({
        type: 'paste',
        sheetIndex: activeIndex,
        data: { cells: pastedCells },
        undoData: { cells: undoCells }
      })
      
      // 更新单元格
      setLocalSheets(prev => {
        const newSheets = [...prev]
        const sheet = { ...newSheets[activeIndex] }
        const cells = [...sheet.cells]
        
        pastedCells.forEach(pc => {
          const existingIndex = cells.findIndex(c => c.r === pc.r && c.c === pc.c)
          const newCell: ExcelCell = {
            r: pc.r,
            c: pc.c,
            v: pc.value,
            w: String(pc.value),
            f: pc.formula,
            display: String(pc.value)
          }
          
          if (existingIndex >= 0) {
            cells[existingIndex] = { ...cells[existingIndex], ...newCell }
          } else {
            cells.push(newCell)
          }
        })
        
        sheet.cells = cells
        newSheets[activeIndex] = sheet
        return newSheets
      })
      
      // 如果是剪切，删除源单元格
      if (clipboard.isCut) {
        setLocalSheets(prev => {
          const newSheets = [...prev]
          const sheet = { ...newSheets[activeIndex] }
          const cells = sheet.cells.filter(cell => {
            return !(cell.r >= clipboard.minR && cell.r <= clipboard.maxR &&
                    cell.c >= clipboard.minC && cell.c <= clipboard.maxC)
          })
          sheet.cells = cells
          newSheets[activeIndex] = sheet
          return newSheets
        })
        setClipboard(null)
      }
      
      setHasChanges(true)
      setStatusMessage('已粘贴')
      return
    }
    
    // 尝试从系统剪贴板读取
    try {
      const text = await navigator.clipboard.readText()
      if (!text) return
      
      const rows = text.split('\n')
      const pastedCells: { r: number; c: number; value: any }[] = []
      const undoCells: { r: number; c: number; value: any; formula?: string }[] = []
      
      rows.forEach((row, rOffset) => {
        const cols = row.split('\t')
        cols.forEach((value, cOffset) => {
          const targetR = startR + rOffset
          const targetC = startC + cOffset
          
          const oldCell = activeSheet?.cells.find(c => c.r === targetR && c.c === targetC)
          undoCells.push({
            r: targetR,
            c: targetC,
            value: oldCell?.v,
            formula: oldCell?.f
          })
          
          pastedCells.push({
            r: targetR,
            c: targetC,
            value: value.trim()
          })
        })
      })
      
      // 记录历史
      pushHistory({
        type: 'paste',
        sheetIndex: activeIndex,
        data: { cells: pastedCells },
        undoData: { cells: undoCells }
      })
      
      // 更新单元格
      setLocalSheets(prev => {
        const newSheets = [...prev]
        const sheet = { ...newSheets[activeIndex] }
        const cells = [...sheet.cells]
        
        pastedCells.forEach(pc => {
          const existingIndex = cells.findIndex(c => c.r === pc.r && c.c === pc.c)
          const newCell: ExcelCell = {
            r: pc.r,
            c: pc.c,
            v: pc.value,
            w: String(pc.value),
            display: String(pc.value)
          }
          
          if (existingIndex >= 0) {
            cells[existingIndex] = { ...cells[existingIndex], ...newCell }
          } else {
            cells.push(newCell)
          }
        })
        
        sheet.cells = cells
        newSheets[activeIndex] = sheet
        return newSheets
      })
      
      setHasChanges(true)
      setStatusMessage('已粘贴')
    } catch (err) {
      console.error('粘贴失败:', err)
    }
  }, [selection.start, clipboard, activeSheet, activeIndex, pushHistory])
  
  // 撤销
  const undo = useCallback(() => {
    if (historyIndex < 0) return
    
    const action = history[historyIndex]
    
    switch (action.type) {
      case 'cell_edit':
      case 'paste': {
        // 恢复旧值
        setLocalSheets(prev => {
          const newSheets = [...prev]
          const sheet = { ...newSheets[action.sheetIndex] }
          const cells = [...sheet.cells]
          
          action.undoData.cells.forEach((uc: any) => {
            const existingIndex = cells.findIndex(c => c.r === uc.r && c.c === uc.c)
            
            if (uc.value === undefined) {
              // 删除单元格
              if (existingIndex >= 0) {
                cells.splice(existingIndex, 1)
              }
            } else {
              const newCell: ExcelCell = {
                r: uc.r,
                c: uc.c,
                v: uc.value,
                w: String(uc.value),
                f: uc.formula,
                display: String(uc.value)
              }
              
              if (existingIndex >= 0) {
                cells[existingIndex] = { ...cells[existingIndex], ...newCell }
              } else {
                cells.push(newCell)
              }
            }
          })
          
          sheet.cells = cells
          newSheets[action.sheetIndex] = sheet
          return newSheets
        })
        break
      }
      case 'delete': {
        // 恢复删除的单元格
        setLocalSheets(prev => {
          const newSheets = [...prev]
          const sheet = { ...newSheets[action.sheetIndex] }
          const cells = [...sheet.cells]
          
          action.undoData.cells.forEach((uc: any) => {
            if (uc.value !== undefined) {
              cells.push({
                r: uc.r,
                c: uc.c,
                v: uc.value,
                w: String(uc.value),
                f: uc.formula,
                display: String(uc.value)
              })
            }
          })
          
          sheet.cells = cells
          newSheets[action.sheetIndex] = sheet
          return newSheets
        })
        break
      }
    }
    
    setHistoryIndex(prev => prev - 1)
    setStatusMessage('已撤销')
  }, [history, historyIndex])
  
  // 重做
  const redo = useCallback(() => {
    if (historyIndex >= history.length - 1) return
    
    const action = history[historyIndex + 1]
    
    switch (action.type) {
      case 'cell_edit': {
        const { r, c, value, formula } = action.data
        updateCellValue(r, c, value, formula)
        break
      }
      case 'paste': {
        setLocalSheets(prev => {
          const newSheets = [...prev]
          const sheet = { ...newSheets[action.sheetIndex] }
          const cells = [...sheet.cells]
          
          action.data.cells.forEach((pc: any) => {
            const existingIndex = cells.findIndex(c => c.r === pc.r && c.c === pc.c)
            const newCell: ExcelCell = {
              r: pc.r,
              c: pc.c,
              v: pc.value,
              w: String(pc.value),
              f: pc.formula,
              display: String(pc.value)
            }
            
            if (existingIndex >= 0) {
              cells[existingIndex] = { ...cells[existingIndex], ...newCell }
            } else {
              cells.push(newCell)
            }
          })
          
          sheet.cells = cells
          newSheets[action.sheetIndex] = sheet
          return newSheets
        })
        break
      }
      case 'delete': {
        setLocalSheets(prev => {
          const newSheets = [...prev]
          const sheet = { ...newSheets[action.sheetIndex] }
          const cells = sheet.cells.filter(cell => {
            return !action.data.cells.some((dc: any) => dc.r === cell.r && dc.c === cell.c)
          })
          sheet.cells = cells
          newSheets[action.sheetIndex] = sheet
          return newSheets
        })
        break
      }
    }
    
    setHistoryIndex(prev => prev + 1)
    setStatusMessage('已重做')
  }, [history, historyIndex, updateCellValue])
  
  // 保存到文件
  const saveToFile = useCallback(async () => {
    if (!filePath || !window.electronAPI) {
      setStatusMessage('无法保存：未指定文件路径')
      return
    }
    
    setIsSaving(true)
    setStatusMessage('正在保存...')
    
    try {
      // 收集所有单元格更新
      const cellUpdates: ExcelCellUpdate[] = []
      
      activeSheet.cells.forEach(cell => {
        cellUpdates.push({
          address: getCellAddress(cell.r, cell.c),
          value: cell.v,
        })
      })
      
      const result = await window.electronAPI.excelWriteCells(
        filePath,
        activeSheet.name,
        cellUpdates
      )
      
      if (result.success) {
        setHasChanges(false)
        setStatusMessage('已保存')
        if (onRefresh) {
          await onRefresh()
        }
      } else {
        setStatusMessage(`保存失败: ${result.error}`)
      }
    } catch (err) {
      setStatusMessage(`保存失败: ${(err as Error).message}`)
    } finally {
      setIsSaving(false)
    }
  }, [filePath, activeSheet, onRefresh])
  
  // 插入行
  const insertRows = useCallback(async (position: 'above' | 'below', count: number = 1) => {
    if (!filePath || !window.electronAPI) {
      setStatusMessage('无法操作：未指定文件路径')
      return
    }
    
    const { minR, maxR } = normalizedSelection
    const startRow = position === 'above' ? minR + 1 : maxR + 2 // 1-based
    
    setStatusMessage('正在插入行...')
    
    try {
      const result = await window.electronAPI.excelInsertRows(
        filePath,
        activeSheet.name,
        startRow,
        count
      )
      
      if (result.success) {
        setStatusMessage(`已插入 ${count} 行`)
        if (onRefresh) {
          await onRefresh()
        }
      } else {
        setStatusMessage(`插入失败: ${result.error}`)
      }
    } catch (err) {
      setStatusMessage(`插入失败: ${(err as Error).message}`)
    }
  }, [filePath, activeSheet, normalizedSelection, onRefresh])
  
  // 插入列
  const insertColumns = useCallback(async (position: 'left' | 'right', count: number = 1) => {
    if (!filePath || !window.electronAPI) {
      setStatusMessage('无法操作：未指定文件路径')
      return
    }
    
    const { minC, maxC } = normalizedSelection
    const startCol = position === 'left' ? minC + 1 : maxC + 2 // 1-based
    
    setStatusMessage('正在插入列...')
    
    try {
      const result = await window.electronAPI.excelInsertColumns(
        filePath,
        activeSheet.name,
        startCol,
        count
      )
      
      if (result.success) {
        setStatusMessage(`已插入 ${count} 列`)
        if (onRefresh) {
          await onRefresh()
        }
      } else {
        setStatusMessage(`插入失败: ${result.error}`)
      }
    } catch (err) {
      setStatusMessage(`插入失败: ${(err as Error).message}`)
    }
  }, [filePath, activeSheet, normalizedSelection, onRefresh])
  
  // 删除行
  const deleteRows = useCallback(async () => {
    if (!filePath || !window.electronAPI) {
      setStatusMessage('无法操作：未指定文件路径')
      return
    }
    
    const { minR, maxR } = normalizedSelection
    const count = maxR - minR + 1
    
    setStatusMessage('正在删除行...')
    
    try {
      const result = await window.electronAPI.excelDeleteRows(
        filePath,
        activeSheet.name,
        minR + 1, // 1-based
        count
      )
      
      if (result.success) {
        setStatusMessage(`已删除 ${count} 行`)
        if (onRefresh) {
          await onRefresh()
        }
      } else {
        setStatusMessage(`删除失败: ${result.error}`)
      }
    } catch (err) {
      setStatusMessage(`删除失败: ${(err as Error).message}`)
    }
  }, [filePath, activeSheet, normalizedSelection, onRefresh])
  
  // 删除列
  const deleteColumns = useCallback(async () => {
    if (!filePath || !window.electronAPI) {
      setStatusMessage('无法操作：未指定文件路径')
      return
    }
    
    const { minC, maxC } = normalizedSelection
    const count = maxC - minC + 1
    
    setStatusMessage('正在删除列...')
    
    try {
      const result = await window.electronAPI.excelDeleteColumns(
        filePath,
        activeSheet.name,
        minC + 1, // 1-based
        count
      )
      
      if (result.success) {
        setStatusMessage(`已删除 ${count} 列`)
        if (onRefresh) {
          await onRefresh()
        }
      } else {
        setStatusMessage(`删除失败: ${result.error}`)
      }
    } catch (err) {
      setStatusMessage(`删除失败: ${(err as Error).message}`)
    }
  }, [filePath, activeSheet, normalizedSelection, onRefresh])
  
  // 合并单元格
  const mergeCells = useCallback(async () => {
    if (!filePath || !window.electronAPI) {
      setStatusMessage('无法操作：未指定文件路径')
      return
    }
    
    const { minR, maxR, minC, maxC } = normalizedSelection
    
    // 至少需要选择两个单元格
    if (minR === maxR && minC === maxC) {
      setStatusMessage('请选择多个单元格进行合并')
      return
    }
    
    const range = `${getColumnLabel(minC)}${minR + 1}:${getColumnLabel(maxC)}${maxR + 1}`
    
    setStatusMessage('正在合并单元格...')
    
    try {
      const result = await window.electronAPI.excelMergeCells(
        filePath,
        activeSheet.name,
        range
      )
      
      if (result.success) {
        setStatusMessage('单元格已合并')
        if (onRefresh) {
          await onRefresh()
        }
      } else {
        setStatusMessage(`合并失败: ${result.error}`)
      }
    } catch (err) {
      setStatusMessage(`合并失败: ${(err as Error).message}`)
    }
  }, [filePath, activeSheet, normalizedSelection, onRefresh])
  
  // 取消合并
  const unmergeCells = useCallback(async () => {
    if (!filePath || !window.electronAPI) {
      setStatusMessage('无法操作：未指定文件路径')
      return
    }
    
    const { minR, maxR, minC, maxC } = normalizedSelection
    const range = `${getColumnLabel(minC)}${minR + 1}:${getColumnLabel(maxC)}${maxR + 1}`
    
    setStatusMessage('正在取消合并...')
    
    try {
      const result = await window.electronAPI.excelUnmergeCells(
        filePath,
        activeSheet.name,
        range
      )
      
      if (result.success) {
        setStatusMessage('已取消合并')
        if (onRefresh) {
          await onRefresh()
        }
      } else {
        setStatusMessage(`取消合并失败: ${result.error}`)
      }
    } catch (err) {
      setStatusMessage(`取消合并失败: ${(err as Error).message}`)
    }
  }, [filePath, activeSheet, normalizedSelection, onRefresh])
  
  // 应用格式到选中单元格
  const applyFormat = useCallback(async (formatType: 'bold' | 'italic' | 'underline' | 'align' | 'bgColor' | 'fontColor', value?: any) => {
    if (!filePath || !window.electronAPI) {
      setStatusMessage('无法操作：未指定文件路径')
      return
    }
    
    const { minR, maxR, minC, maxC } = normalizedSelection
    const cellUpdates: ExcelCellUpdate[] = []
    
    for (let r = minR; r <= maxR; r++) {
      for (let c = minC; c <= maxC; c++) {
        const address = getCellAddress(r, c)
        const existingCell = activeSheet.cells.find(cell => cell.r === r && cell.c === c)
        
        const update: ExcelCellUpdate = { address }
        
        switch (formatType) {
          case 'bold':
            update.style = { font: { bold: true } }
            break
          case 'italic':
            update.style = { font: { italic: true } }
            break
          case 'underline':
            update.style = { font: { underline: true } }
            break
          case 'align':
            update.style = { alignment: { horizontal: value as 'left' | 'center' | 'right' } }
            break
          case 'bgColor':
            update.style = { fill: { fgColor: { argb: value.replace('#', 'FF') } } }
            break
          case 'fontColor':
            update.style = { font: { color: { argb: value.replace('#', 'FF') } } }
            break
        }
        
        // 保留原值
        if (existingCell) {
          update.value = existingCell.v
        }
        
        cellUpdates.push(update)
      }
    }
    
    setStatusMessage('正在应用格式...')
    
    try {
      const result = await window.electronAPI.excelWriteCells(
        filePath,
        activeSheet.name,
        cellUpdates
      )
      
      if (result.success) {
        setStatusMessage('格式已应用')
        if (onRefresh) {
          await onRefresh()
        }
      } else {
        setStatusMessage(`格式应用失败: ${result.error}`)
      }
    } catch (err) {
      setStatusMessage(`格式应用失败: ${(err as Error).message}`)
    }
  }, [filePath, activeSheet, normalizedSelection, onRefresh])
  
  // 排序
  const sortData = useCallback(async (ascending: boolean) => {
    if (!filePath || !window.electronAPI) {
      setStatusMessage('无法操作：未指定文件路径')
      return
    }
    
    const { minR, maxR, minC, maxC } = normalizedSelection
    const range = `${getColumnLabel(minC)}${minR + 1}:${getColumnLabel(maxC)}${maxR + 1}`
    
    setStatusMessage('正在排序...')
    
    try {
      const result = await window.electronAPI.excelSort(
        filePath,
        activeSheet.name,
        {
          range,
          column: minC + 1, // 1-based column
          ascending,
          hasHeader: minR === 0 // 如果从第一行开始，假设有表头
        }
      )
      
      if (result.success) {
        setStatusMessage(`已${ascending ? '升序' : '降序'}排序`)
        if (onRefresh) {
          await onRefresh()
        }
      } else {
        setStatusMessage(`排序失败: ${result.error}`)
      }
    } catch (err) {
      setStatusMessage(`排序失败: ${(err as Error).message}`)
    }
  }, [filePath, activeSheet, normalizedSelection, onRefresh])
  
  // 设置筛选
  const toggleFilter = useCallback(async () => {
    if (!filePath || !window.electronAPI) {
      setStatusMessage('无法操作：未指定文件路径')
      return
    }
    
    const { minR, maxR, minC, maxC } = normalizedSelection
    const range = `${getColumnLabel(minC)}${minR + 1}:${getColumnLabel(maxC)}${maxR + 1}`
    
    setStatusMessage('正在设置筛选...')
    
    try {
      const result = await window.electronAPI.excelSetFilter(
        filePath,
        activeSheet.name,
        { range }
      )
      
      if (result.success) {
        setStatusMessage('筛选已设置')
        if (onRefresh) {
          await onRefresh()
        }
      } else {
        setStatusMessage(`设置筛选失败: ${result.error}`)
      }
    } catch (err) {
      setStatusMessage(`设置筛选失败: ${(err as Error).message}`)
    }
  }, [filePath, activeSheet, normalizedSelection, onRefresh])

  // 鼠标按下开始选择
  const handleCellMouseDown = useCallback(
    (r: number, c: number, e: MouseEvent) => {
      e.preventDefault()
      
      // 如果正在编辑，先确认
      if (editingCell) {
        confirmEdit()
      }
      
      if (e.shiftKey) {
        // Shift+Click：扩展选区
        setSelection((prev) => ({
          start: prev.start,
          end: { r, c },
        }))
      } else {
        // 普通点击：开始新选区
        setSelection({
          start: { r, c },
          end: { r, c },
        })
        setIsSelecting(true)
      }
    },
    [editingCell, confirmEdit]
  )
  
  // 双击进入编辑模式
  const handleCellDoubleClick = useCallback((r: number, c: number) => {
    startEdit(r, c)
  }, [startEdit])

  // 鼠标移动扩展选区
  const handleCellMouseEnter = useCallback(
    (r: number, c: number) => {
      if (isSelecting) {
        setSelection((prev) => ({
          start: prev.start,
          end: { r, c },
        }))
      }
    },
    [isSelecting]
  )

  // 鼠标释放结束选择
  const handleMouseUp = useCallback(() => {
    setIsSelecting(false)
  }, [])

  // 点击列头选择整列
  const handleColHeaderClick = useCallback(
    (c: number, e: MouseEvent, rowCount: number) => {
      e.preventDefault()
      if (editingCell) confirmEdit()
      
      if (e.shiftKey) {
        // Shift+Click：扩展到该列
        setSelection((prev) => ({
          start: { r: 0, c: prev.start.c },
          end: { r: rowCount - 1, c },
        }))
      } else {
        // 选择整列
        setSelection({
          start: { r: 0, c },
          end: { r: rowCount - 1, c },
        })
      }
    },
    [editingCell, confirmEdit]
  )

  // 点击行头选择整行
  const handleRowHeaderClick = useCallback(
    (r: number, e: MouseEvent, colCount: number) => {
      e.preventDefault()
      if (editingCell) confirmEdit()
      
      if (e.shiftKey) {
        // Shift+Click：扩展到该行
        setSelection((prev) => ({
          start: { r: prev.start.r, c: 0 },
          end: { r, c: colCount - 1 },
        }))
      } else {
        // 选择整行
        setSelection({
          start: { r, c: 0 },
          end: { r, c: colCount - 1 },
        })
      }
    },
    [editingCell, confirmEdit]
  )

  // 点击左上角选择全部
  const handleSelectAll = useCallback(
    (rowCount: number, colCount: number) => {
      if (editingCell) confirmEdit()
      setSelection({
        start: { r: 0, c: 0 },
        end: { r: rowCount - 1, c: colCount - 1 },
      })
    },
    [editingCell, confirmEdit]
  )
  
  // 键盘事件处理
  const handleKeyDown = useCallback((e: KeyboardEvent) => {
    // 如果正在编辑，只处理特定键
    if (editingCell) {
      if (e.key === 'Enter') {
        e.preventDefault()
        confirmEdit()
        // 移动到下一行
        setSelection(prev => ({
          start: { r: prev.start.r + 1, c: prev.start.c },
          end: { r: prev.start.r + 1, c: prev.start.c }
        }))
      } else if (e.key === 'Tab') {
        e.preventDefault()
        confirmEdit()
        // 移动到下一列
        setSelection(prev => ({
          start: { r: prev.start.r, c: prev.start.c + 1 },
          end: { r: prev.start.r, c: prev.start.c + 1 }
        }))
      } else if (e.key === 'Escape') {
        e.preventDefault()
        cancelEdit()
      }
      return
    }
    
    const { r, c } = selection.start
    
    // Ctrl 组合键
    if (e.ctrlKey || e.metaKey) {
      switch (e.key.toLowerCase()) {
        case 'c':
          e.preventDefault()
          copySelectedCells(false)
          break
        case 'x':
          e.preventDefault()
          copySelectedCells(true)
          break
        case 'v':
          e.preventDefault()
          pasteFromClipboard()
          break
        case 'z':
          e.preventDefault()
          if (e.shiftKey) {
            redo()
          } else {
            undo()
          }
          break
        case 'y':
          e.preventDefault()
          redo()
          break
        case 's':
          e.preventDefault()
          saveToFile()
          break
        case 'a':
          e.preventDefault()
          if (activeSheet) {
            const maxRow = Math.max(activeSheet.range.e.r + 50, 100)
            const maxCol = Math.max(activeSheet.range.e.c + 10, 26)
            handleSelectAll(maxRow, maxCol)
          }
          break
      }
      return
    }
    
    // 方向键导航
    switch (e.key) {
      case 'ArrowUp':
        e.preventDefault()
        if (r > 0) {
          const newR = r - 1
          setSelection(e.shiftKey 
            ? { start: selection.start, end: { r: newR, c: selection.end.c } }
            : { start: { r: newR, c }, end: { r: newR, c } }
          )
        }
        break
      case 'ArrowDown':
        e.preventDefault()
        {
          // Calculate max rows from activeSheet
          const maxRow = activeSheet ? Math.max(activeSheet.range.e.r + 50, 99) : 99
          if (r < maxRow) {
            const newR = r + 1
            setSelection(e.shiftKey 
              ? { start: selection.start, end: { r: newR, c: selection.end.c } }
              : { start: { r: newR, c }, end: { r: newR, c } }
            )
          }
        }
        break
      case 'ArrowLeft':
        e.preventDefault()
        if (c > 0) {
          const newC = c - 1
          setSelection(e.shiftKey 
            ? { start: selection.start, end: { r: selection.end.r, c: newC } }
            : { start: { r, c: newC }, end: { r, c: newC } }
          )
        }
        break
      case 'ArrowRight':
        e.preventDefault()
        {
          // Calculate max cols from activeSheet
          const maxCol = activeSheet ? Math.max(activeSheet.range.e.c + 9, 25) : 25
          if (c < maxCol) {
            const newC = c + 1
            setSelection(e.shiftKey 
              ? { start: selection.start, end: { r: selection.end.r, c: newC } }
              : { start: { r, c: newC }, end: { r, c: newC } }
            )
          }
        }
        break
      case 'Tab':
        e.preventDefault()
        {
          const maxCol = activeSheet ? Math.max(activeSheet.range.e.c + 9, 25) : 25
          if (c < maxCol) {
            setSelection({ start: { r, c: c + 1 }, end: { r, c: c + 1 } })
          }
        }
        break
      case 'Enter':
        e.preventDefault()
        {
          const maxRow = activeSheet ? Math.max(activeSheet.range.e.r + 50, 99) : 99
          if (r < maxRow) {
            setSelection({ start: { r: r + 1, c }, end: { r: r + 1, c } })
          }
        }
        break
      case 'Delete':
      case 'Backspace':
        e.preventDefault()
        deleteSelectedCells()
        break
      case 'F2':
        e.preventDefault()
        startEdit(r, c)
        break
      default:
        // 开始输入（字母数字键）
        if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
          startEdit(r, c, e.key)
        }
        break
    }
  }, [
    editingCell, confirmEdit, cancelEdit, selection, activeSheet,
    copySelectedCells, pasteFromClipboard, undo, redo, 
    saveToFile, handleSelectAll, deleteSelectedCells, startEdit
  ])
  
  // 聚焦容器以接收键盘事件
  useEffect(() => {
    if (containerRef.current && !editingCell) {
      containerRef.current.focus()
    }
  }, [editingCell, selection])

  const ribbonGroups = useMemo(
    () => [
      {
        title: '剪贴板',
        layout: (
          <div className="ribbon-group-layout">
            <button className="ribbon-btn big" onClick={() => pasteFromClipboard()}>
              <ClipboardPaste size={18} />
              <span>粘贴</span>
            </button>
            <div className="ribbon-btn-col">
              <button className="ribbon-btn" onClick={() => copySelectedCells(true)}>
                <Scissors size={14} />
                <span>剪切</span>
              </button>
              <button className="ribbon-btn" onClick={() => copySelectedCells(false)}>
                <Copy size={14} />
                <span>复制</span>
              </button>
              <button className="ribbon-btn">
                <Paintbrush size={14} />
                <span>刷子</span>
              </button>
            </div>
          </div>
        ),
      },
      {
        title: '撤销',
        layout: (
          <div className="ribbon-group-layout">
            <button 
              className={`ribbon-btn big ${historyIndex < 0 ? 'disabled' : ''}`}
              onClick={undo}
              disabled={historyIndex < 0}
            >
              <Undo2 size={18} />
              <span>撤销</span>
            </button>
            <button 
              className={`ribbon-btn big ${historyIndex >= history.length - 1 ? 'disabled' : ''}`}
              onClick={redo}
              disabled={historyIndex >= history.length - 1}
            >
              <Redo2 size={18} />
              <span>重做</span>
            </button>
          </div>
        ),
      },
      {
        title: '字体',
        layout: (
          <div className="ribbon-group-layout font-group">
            <div className="ribbon-row">
              <button className="ribbon-btn pill">
                <FontIcon size={14} />
                <span>宋体</span>
              </button>
              <button className="ribbon-btn pill">12</button>
              <button className="ribbon-btn icon-only">
                <ChevronUp size={14} />
              </button>
              <button className="ribbon-btn icon-only">
                <ChevronDown size={14} />
              </button>
            </div>
            <div className="ribbon-row">
              <button className="ribbon-btn icon-only" onClick={() => applyFormat('bold')} title="加粗">
                <BoldIcon size={14} />
              </button>
              <button className="ribbon-btn icon-only" onClick={() => applyFormat('italic')} title="斜体">
                <ItalicIcon size={14} />
              </button>
              <button className="ribbon-btn icon-only" onClick={() => applyFormat('underline')} title="下划线">
                <UnderlineIcon size={14} />
              </button>
              <button className="ribbon-btn" title="边框">边框</button>
              <label className="ribbon-btn" title="背景色">
                填充
                <input 
                  type="color" 
                  style={{ width: 0, height: 0, opacity: 0, position: 'absolute' }}
                  onChange={(e) => applyFormat('bgColor', e.target.value)}
                />
              </label>
              <label className="ribbon-btn" title="字体颜色">
                字体色
                <input 
                  type="color" 
                  style={{ width: 0, height: 0, opacity: 0, position: 'absolute' }}
                  onChange={(e) => applyFormat('fontColor', e.target.value)}
                />
              </label>
            </div>
          </div>
        ),
      },
      {
        title: '对齐',
        layout: (
          <div className="ribbon-group-layout">
            <div className="ribbon-row">
              <button className="ribbon-btn icon-only" onClick={() => applyFormat('align', 'left')} title="左对齐">
                <AlignLeft size={14} />
              </button>
              <button className="ribbon-btn icon-only" onClick={() => applyFormat('align', 'center')} title="居中">
                <AlignCenter size={14} />
              </button>
              <button className="ribbon-btn icon-only" onClick={() => applyFormat('align', 'right')} title="右对齐">
                <AlignRight size={14} />
              </button>
              <button className="ribbon-btn icon-only" onClick={() => applyFormat('align', 'justify')} title="两端对齐">
                <AlignJustify size={14} />
              </button>
            </div>
            <div className="ribbon-row">
              <button className="ribbon-btn icon-only" title="顶端对齐">
                <AlignVerticalDistributeStart size={14} />
              </button>
              <button className="ribbon-btn icon-only" title="垂直居中">
                <AlignVerticalDistributeCenter size={14} />
              </button>
              <button className="ribbon-btn icon-only" title="底端对齐">
                <AlignVerticalDistributeEnd size={14} />
              </button>
              <button className="ribbon-btn">换行</button>
            </div>
          </div>
        ),
      },
      {
        title: '数字',
        layout: (
          <div className="ribbon-group-layout">
            <div className="ribbon-row">
              <button className="ribbon-btn pill">常规 ▾</button>
              <button className="ribbon-btn icon-only">
                <Percent size={14} />
              </button>
              <button className="ribbon-btn icon-only">,</button>
            </div>
            <div className="ribbon-row">
              <button className="ribbon-btn icon-only">
                <Plus size={14} />
              </button>
              <button className="ribbon-btn icon-only">
                <Minus size={14} />
              </button>
            </div>
          </div>
        ),
      },
      {
        title: '单元格',
        layout: (
          <div className="ribbon-group-layout">
            <div className="ribbon-btn-col">
              <button className="ribbon-btn" onClick={deleteSelectedCells}>
                <Trash2 size={14} />
                <span>删除内容</span>
            </button>
              <button className="ribbon-btn" onClick={mergeCells}>
                <Merge size={14} />
                <span>合并</span>
            </button>
              <button className="ribbon-btn" onClick={unmergeCells}>
                <SplitSquareHorizontal size={14} />
                <span>拆分</span>
            </button>
            </div>
          </div>
        ),
      },
      {
        title: '行列',
        layout: (
          <div className="ribbon-group-layout">
            <div className="ribbon-btn-col">
              <button className="ribbon-btn" onClick={() => insertRows('above')}>
                <RowsIcon size={14} />
                <span>上方插入行</span>
              </button>
              <button className="ribbon-btn" onClick={() => insertRows('below')}>
                <RowsIcon size={14} />
                <span>下方插入行</span>
              </button>
              <button className="ribbon-btn" onClick={deleteRows}>
                <Minus size={14} />
                <span>删除行</span>
              </button>
            </div>
            <div className="ribbon-btn-col">
              <button className="ribbon-btn" onClick={() => insertColumns('left')}>
                <ColumnsIcon size={14} />
                <span>左侧插入列</span>
            </button>
              <button className="ribbon-btn" onClick={() => insertColumns('right')}>
                <ColumnsIcon size={14} />
                <span>右侧插入列</span>
              </button>
              <button className="ribbon-btn" onClick={deleteColumns}>
                <Minus size={14} />
                <span>删除列</span>
            </button>
            </div>
          </div>
        ),
      },
      {
        title: '数据',
        layout: (
          <div className="ribbon-group-layout">
            <div className="ribbon-btn-col">
              <button className="ribbon-btn" onClick={() => sortData(true)} title="升序排序">
                <ChevronUp size={14} />
                <span>升序</span>
            </button>
              <button className="ribbon-btn" onClick={() => sortData(false)} title="降序排序">
                <ChevronDown size={14} />
                <span>降序</span>
            </button>
              <button className="ribbon-btn" onClick={toggleFilter} title="设置筛选">
                <Filter size={14} />
                <span>筛选</span>
            </button>
            </div>
          </div>
        ),
      },
      {
        title: '保存',
        layout: (
          <div className="ribbon-group-layout">
            <button 
              className={`ribbon-btn big ${!hasChanges ? 'disabled' : ''}`}
              onClick={saveToFile}
              disabled={!hasChanges || isSaving}
            >
              <Save size={18} />
              <span>{isSaving ? '保存中...' : '保存'}</span>
            </button>
          </div>
        ),
      },
    ],
    [
      pasteFromClipboard, copySelectedCells, undo, redo, 
      historyIndex, history.length, deleteSelectedCells,
      saveToFile, hasChanges, isSaving,
      insertRows, insertColumns, deleteRows, deleteColumns,
      mergeCells, unmergeCells, applyFormat, sortData, toggleFilter
    ]
  )

  const matrix = useMemo(() => (activeSheet ? buildMatrix(activeSheet) : null), [activeSheet])

  if (!activeSheet || !matrix) {
    return <div style={{ padding: 16, color: '#666' }}>暂无可预览的工作表</div>
  }

  const colWidths = activeSheet.colWidths || []
  const rowHeights = activeSheet.rowHeights || []

  const colGroup = useMemo(() => {
    const cols = [
      <col key="row-header" style={{ width: '40px', minWidth: '40px', background: '#f3f3f3' }} />,
    ]
    for (let c = 0; c < matrix.colCount; c += 1) {
      const w = colWidths[c] || 72
      cols.push(<col key={c} style={{ width: `${w}px`, minWidth: `${w}px` }} />)
    }
    return cols
  }, [colWidths, matrix.colCount])

  const colHeaders = useMemo(
    () => Array.from({ length: matrix.colCount }).map((_, i) => getColumnLabel(i)),
    [matrix.colCount]
  )

  const selectedCell = useMemo(() => {
    return activeSheet.cells.find(
      cell => cell.r === selection.start.r && cell.c === selection.start.c
    )
  }, [activeSheet.cells, selection.start.r, selection.start.c])

  const selectedValue = useMemo(() => {
    if (!selectedCell) return ''
    // 显示公式或值
    return selectedCell.f || String(selectedCell.w ?? selectedCell.v ?? '')
  }, [selectedCell])

  // 选区范围文本（用于名称框）
  const selectionRangeText = useMemo(() => {
    const { minR, maxR, minC, maxC } = normalizedSelection
    const startLabel = `${getColumnLabel(minC)}${minR + 1}`
    
    // 单个单元格
    if (minR === maxR && minC === maxC) {
      return startLabel
    }
    
    const endLabel = `${getColumnLabel(maxC)}${maxR + 1}`
    return `${startLabel}:${endLabel}`
  }, [normalizedSelection])

  // 生成列标 (A, B, ..., Z, AA, ...)
  return (
    <div 
      className="excel-preview"
      ref={containerRef}
      tabIndex={0}
      onKeyDown={handleKeyDown}
      style={{ outline: 'none' }}
    >
      {/* 顶部 Ribbon 占位 */}
      <div className="excel-ribbon">
        <div className="excel-ribbon-tabs">
          {['开始', '插入', '页面布局', '公式', '数据', '审阅', '视图'].map((tab) => (
            <div key={tab} className="excel-ribbon-tab">
              {tab}
            </div>
          ))}
        </div>
        <div className="excel-ribbon-tools">
          {ribbonGroups.map((group, idx) => (
            <div key={group.title} className="excel-tool-group">
              {group.layout}
              <div className="tool-title">{group.title}</div>
              {idx < ribbonGroups.length - 1 && <div className="ribbon-divider" />}
            </div>
          ))}
        </div>
      </div>

      {/* 顶部公式栏模拟 */}
      <div className="excel-toolbar">
        <div className="excel-name-box">{selectionRangeText}</div>
        <div className="excel-formula-bar">
          <div className="fx-icon">fx</div>
          {isFormulaBarEditing || editingCell ? (
            <div className="formula-input-wrapper">
              <input
                ref={formulaInputRef}
                type="text"
                className="formula-input-edit"
                value={editValue}
                onChange={(e) => setEditValue(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') {
                    confirmEdit()
                  } else if (e.key === 'Escape') {
                    cancelEdit()
                  }
                }}
                onFocus={() => setIsFormulaBarEditing(true)}
              />
              <button className="formula-btn confirm" onClick={confirmEdit}>
                <Check size={14} />
              </button>
              <button className="formula-btn cancel" onClick={cancelEdit}>
                <X size={14} />
              </button>
            </div>
          ) : (
            <div 
              className="formula-input"
              onClick={() => {
                if (selection.start) {
                  startEdit(selection.start.r, selection.start.c)
                  setIsFormulaBarEditing(true)
                }
              }}
            >
              {selectedValue}
            </div>
          )}
        </div>
      </div>

      {/* 主表格区域 */}
      <div 
        className="excel-table-wrapper"
        onMouseUp={handleMouseUp}
        onMouseLeave={handleMouseUp}
      >
        <div className="excel-table-scaler" style={{ transform: `scale(${zoom / 100})`, transformOrigin: 'top left' }}>
        <table className="excel-table" ref={tableRef}>
          <colgroup>{colGroup}</colgroup>
          <thead>
            <tr>
              <th 
                className="excel-corner-header"
                onClick={() => handleSelectAll(matrix.rowCount, matrix.colCount)}
                title="选择全部"
              ></th>
              {colHeaders.map((label, i) => {
                const { minC, maxC } = normalizedSelection
                const isColSelected = i >= minC && i <= maxC
                return (
                  <th
                    key={i}
                    className={`excel-col-header ${isColSelected ? 'active' : ''}`}
                    onMouseDown={(e) => handleColHeaderClick(i, e, matrix.rowCount)}
                  >
                    {label}
                  </th>
                )
              })}
            </tr>
          </thead>
          <tbody>
            {matrix.rows.map((row, rIdx) => {
              const { minR, maxR } = normalizedSelection
              const isRowSelected = rIdx >= minR && rIdx <= maxR
              
              return (
                <tr key={rIdx} style={rowHeights[rIdx] ? { height: `${rowHeights[rIdx]}px` } : { height: '20px' }}>
                  {/* 行号 */}
                  <td 
                    className={`excel-row-header ${isRowSelected ? 'active' : ''}`}
                    onMouseDown={(e) => handleRowHeaderClick(rIdx, e, matrix.colCount)}
                  >
                    {rIdx + 1}
                  </td>
                  {row.map((item, cIdx) => {
                    const isMergedSkipped =
                      localSheets[activeIndex].merges &&
                      localSheets[activeIndex].merges.some(
                        (m) => rIdx >= m.s.r && rIdx <= m.e.r && cIdx >= m.s.c && cIdx <= m.e.c && !(rIdx === m.s.r && cIdx === m.s.c)
                      )
                    if (isMergedSkipped) return null

                    const cell = item.cell
                    const merge = item.merge
                    const display = cell
                      ? String(cell.display ?? cell.w ?? cell.v ?? '')
                      : ''
                    
                    const selected = isCellSelected(rIdx, cIdx)
                    const active = isActiveCell(rIdx, cIdx)
                    const isEditing = editingCell?.r === rIdx && editingCell?.c === cIdx
                    
                    // 判断选区边界
                    const { minR, maxR, minC, maxC } = normalizedSelection
                    const isTopEdge = rIdx === minR && selected
                    const isBottomEdge = rIdx === maxR && selected
                    const isLeftEdge = cIdx === minC && selected
                    const isRightEdge = cIdx === maxC && selected

                    return (
                      <td
                        key={cIdx}
                        rowSpan={merge?.rowSpan}
                        colSpan={merge?.colSpan}
                        style={convertCellStyle(cell)}
                        className={`excel-cell ${selected ? 'excel-cell-in-selection' : ''} ${active ? 'excel-cell-active' : ''} ${isTopEdge ? 'selection-top' : ''} ${isBottomEdge ? 'selection-bottom' : ''} ${isLeftEdge ? 'selection-left' : ''} ${isRightEdge ? 'selection-right' : ''} ${isEditing ? 'excel-cell-editing' : ''}`}
                        onMouseDown={(e) => handleCellMouseDown(rIdx, cIdx, e)}
                        onMouseEnter={() => handleCellMouseEnter(rIdx, cIdx)}
                        onDoubleClick={() => handleCellDoubleClick(rIdx, cIdx)}
                      >
                        {isEditing ? (
                          <input
                            ref={editInputRef}
                            type="text"
                            className="cell-edit-input"
                            value={editValue}
                            onChange={(e) => setEditValue(e.target.value)}
                            onKeyDown={(e) => {
                              if (e.key === 'Enter') {
                                e.preventDefault()
                                confirmEdit()
                              } else if (e.key === 'Escape') {
                                e.preventDefault()
                                cancelEdit()
                              } else if (e.key === 'Tab') {
                                e.preventDefault()
                                confirmEdit()
                                setSelection(prev => ({
                                  start: { r: prev.start.r, c: prev.start.c + 1 },
                                  end: { r: prev.start.r, c: prev.start.c + 1 }
                                }))
                              }
                              e.stopPropagation()
                            }}
                            onBlur={() => {
                              // 延迟确认，避免与其他点击冲突
                              setTimeout(() => {
                                if (editingCell) {
                                  confirmEdit()
                                }
                              }, 100)
                            }}
                            onClick={(e) => e.stopPropagation()}
                          />
                        ) : (
                          display
                        )}
                      </td>
                    )
                  })}
                </tr>
              )
            })}
          </tbody>
        </table>
        </div>
      </div>

      {/* 底部 Sheet 标签栏 */}
      <div className="excel-footer">
        <div className="excel-sheet-area">
          <div className="excel-sheet-scroll-btn">◀</div>
          <div className="excel-sheet-scroll-btn">▶</div>
          <div className="excel-tabs">
            {localSheets.map((sheet, idx) => (
              <button
                key={sheet.name}
                className={`excel-tab ${idx === activeIndex ? 'active' : ''}`}
                onClick={() => setActiveIndex(idx)}
              >
                {sheet.name}
              </button>
            ))}
            <button className="excel-new-sheet">+</button>
          </div>
        </div>
        <div className="excel-status">
          <span className="excel-status-text">
            {statusMessage}
            {hasChanges && ' (未保存)'}
          </span>
          <div className="excel-zoom">
            <button className="zoom-btn" onClick={() => handleZoomChange(-10)}>−</button>
            <div className="zoom-track" title={`${zoom}%`}>
              <div 
                className="zoom-thumb" 
                style={{ left: `${((zoom - 25) / 175) * 100}%` }}
              />
            </div>
            <button className="zoom-btn" onClick={() => handleZoomChange(10)}>+</button>
            <span className="zoom-label">{zoom}%</span>
          </div>
        </div>
      </div>
    </div>
  )
}

export default ExcelPreview
