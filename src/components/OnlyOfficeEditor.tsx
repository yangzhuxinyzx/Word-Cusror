import { useEffect, useRef, useState, useCallback } from 'react'
import { useDocument } from '../context/DocumentContext'
import { Loader2, AlertCircle, RefreshCw, FileText, CheckCircle } from 'lucide-react'

// 文档服务器地址
const DOCUMENT_SERVER_URL = 'http://localhost:8080'

// 生成唯一文档 key
function generateDocKey(): string {
  return `doc_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`
}

// ONLYOFFICE API 全局声明
declare const Api: any  // ONLYOFFICE Document Builder API
declare const Asc: any  // ONLYOFFICE 内部命名空间

declare global {
  interface Window {
    DocsAPI?: {
      DocEditor: new (elementId: string, config: any) => any
    }
    Asc?: any  // ONLYOFFICE 内部命名空间
    // 暴露给 AI 的编辑器操作接口
    onlyOfficeConnector?: {
      // 获取文档文本内容
      getDocumentText: () => Promise<string>
      // 获取文档结构信息（包含格式描述）
      getDocumentStructure: () => Promise<string>
      // 搜索并替换文本
      searchAndReplace: (searchText: string, replaceText: string, replaceAll?: boolean) => Promise<boolean>
      // 在光标位置插入文本
      insertText: (text: string) => Promise<boolean>
      // 获取选中的文本
      getSelectedText: () => Promise<string>
      // 替换选中的文本
      replaceSelectedText: (text: string) => Promise<boolean>
      // 保存文档
      saveDocument: () => Promise<boolean>
      // 下载文档
      downloadDocument: () => void
      // 执行 Document Builder 脚本
      executeScript: (script: string) => Promise<boolean>
      // 添加带格式的段落
      addFormattedParagraph: (text: string, options?: {
        fontSize?: number
        fontFamily?: string
        bold?: boolean
        italic?: boolean
        color?: string
        alignment?: 'left' | 'center' | 'right' | 'justify'
      }) => Promise<boolean>
      // 添加表格
      addTable: (rows: number, cols: number, data?: string[][]) => Promise<boolean>
      // 添加标题
      addHeading: (text: string, level: 1 | 2 | 3 | 4 | 5 | 6) => Promise<boolean>
      // 全选
      selectAll: () => Promise<boolean>
      // 复制
      copy: () => Promise<boolean>
      // 粘贴
      paste: () => Promise<boolean>
      // 清空内容
      clearContent: () => Promise<boolean>
      // 批量替换
      batchReplace: (replacements: Array<{search: string, replace: string}>) => Promise<number>
      // 创建带格式的文档内容
      createFormattedContent: (elements: Array<{
        type: 'heading' | 'paragraph' | 'table'
        content?: string
        level?: number
        bold?: boolean
        fontSize?: number
        fontFamily?: string
        alignment?: 'left' | 'center' | 'right' | 'justify'
        rows?: number
        cols?: number
        data?: string[][]
      }>) => Promise<boolean>
    }
    // ONLYOFFICE Document Builder API
    Api?: any
  }
}

export default function OnlyOfficeEditor() {
  const { currentFile, isElectron, setDocument } = useDocument()
  const editorContainerRef = useRef<HTMLDivElement>(null)
  const editorInstanceRef = useRef<any>(null)
  const connectorRef = useRef<any>(null)
  const [isLoading, setIsLoading] = useState(true)
  const [error, setError] = useState<string | null>(null)
  const [isApiLoaded, setIsApiLoaded] = useState(false)
  const [docKey, setDocKey] = useState(generateDocKey())
  const [fileUrl, setFileUrl] = useState<string | null>(null)
  const [isEditorReady, setIsEditorReady] = useState(false)

  // 当文件改变时，获取文件 URL 并更新 docKey
  useEffect(() => {
    const getUrl = async () => {
      if (currentFile && isElectron && window.electronAPI?.getFileUrl) {
        try {
          const url = await window.electronAPI.getFileUrl(currentFile.path)
          console.log('获取文件 URL:', url)
          setFileUrl(url)
          setDocKey(generateDocKey()) // 新文件需要新的 key
          setIsEditorReady(false)
        } catch (e) {
          console.error('获取文件 URL 失败:', e)
          setFileUrl(null)
        }
      } else {
        setFileUrl(null)
      }
    }
    getUrl()
  }, [currentFile, isElectron])

  // 获取当前有效的 connector（每次调用时动态获取，避免闭包问题）
  const getConnector = useCallback(() => {
    return connectorRef.current
  }, [])

  // 初始化全局连接器接口（供 AI 调用）
  const initConnectorInterface = useCallback(() => {
    if (!connectorRef.current) {
      console.warn('ONLYOFFICE connector 未就绪')
      return
    }

    console.log('🔧 初始化 ONLYOFFICE connector 接口')

    window.onlyOfficeConnector = {
      // 获取文档文本内容
      getDocumentText: async (): Promise<string> => {
        const connector = getConnector()
        if (!connector) {
          console.warn('connector 不可用')
          return ''
        }
        return new Promise((resolve) => {
          connector.executeMethod('GetAllContentControls', [], () => {
            // 尝试获取文档文本
            connector.executeMethod('GetDocumentText', [], (text: string) => {
              resolve(text || '')
            })
          })
        })
      },

      // 获取文档结构信息（包含格式描述和可复用的JSON格式）
      getDocumentStructure: async (): Promise<string> => {
        const connector = getConnector()
        if (!connector) {
          console.warn('connector 不可用')
          return ''
        }
        return new Promise((resolve) => {
          try {
            connector.callCommand(
              function() {
                // @ts-ignore
                const oDocument = Api.GetDocument();
                const elements = [];
                const count = oDocument.GetElementsCount();
                
                for (let i = 0; i < Math.min(count, 30); i++) { // 最多30个元素
                  const elem = oDocument.GetElement(i);
                  if (!elem) continue;
                  
                  const classType = elem.GetClassType();
                  
                  if (classType === 'paragraph') {
                    const text = elem.GetText();
                    if (!text.trim()) continue;
                    
                    const style = elem.GetStyle();
                    const styleName = style ? style.GetName() : 'Normal';
                    
                    // 判断是否是标题
                    const isHeading = styleName.includes('Heading') || styleName.includes('标题');
                    let headingLevel = 0;
                    if (isHeading) {
                      const levelMatch = styleName.match(/(\d)/);
                      headingLevel = levelMatch ? parseInt(levelMatch[1]) : 1;
                    }
                    
                    // 获取对齐方式
                    let alignment = 'left';
                    try {
                      const jc = elem.GetJc();
                      if (jc === 'center') alignment = 'center';
                      else if (jc === 'right') alignment = 'right';
                      else if (jc === 'both') alignment = 'justify';
                    } catch(e) {}
                    
                    if (isHeading) {
                      elements.push({
                        type: 'heading',
                        level: headingLevel,
                        content: text,
                        alignment: alignment
                      });
                    } else {
                      elements.push({
                        type: 'paragraph',
                        content: text.substring(0, 200) + (text.length > 200 ? '...' : ''),
                        alignment: alignment
                      });
                    }
                  } else if (classType === 'table') {
                    const rowCount = elem.GetRowsCount();
                    const colCount = elem.GetRow(0) ? elem.GetRow(0).GetCellsCount() : 0;
                    
                    // 获取表格所有数据
                    const tableData = [];
                    for (let r = 0; r < Math.min(rowCount, 10); r++) { // 最多10行
                      const row = elem.GetRow(r);
                      if (!row) continue;
                      const rowData = [];
                      for (let c = 0; c < colCount; c++) {
                        const cell = row.GetCell(c);
                        if (cell) {
                          const content = cell.GetContent();
                          if (content && content.GetElement) {
                            const para = content.GetElement(0);
                            rowData.push(para ? para.GetText() : '');
                          } else {
                            rowData.push('');
                          }
                        } else {
                          rowData.push('');
                        }
                      }
                      tableData.push(rowData);
                    }
                    
                    elements.push({
                      type: 'table',
                      rows: rowCount,
                      cols: colCount,
                      data: tableData
                    });
                  }
                }
                
                return JSON.stringify(elements);
              },
              (result: any) => {
                if (result) {
                  try {
                    const elements = JSON.parse(result);
                    
                    // 生成人类可读的描述 - 重点标注可替换的内容
                    let description = '【文档结构 - 以下文字可用于 search 参数】\n';
                    description += '⚠️ 替换时 search 必须与下面的文字完全一致！\n\n';
                    
                    for (const elem of elements) {
                      if (elem.type === 'heading') {
                        description += `📌 标题: "${elem.content}"\n`;
                      } else if (elem.type === 'paragraph') {
                        // 显示完整内容（不截断），方便 AI 精确匹配
                        description += `📝 段落: "${elem.content}"\n`;
                      } else if (elem.type === 'table') {
                        description += `📊 表格 (${elem.rows}行×${elem.cols}列):\n`;
                        // 显示每个单元格的精确内容
                        if (elem.data && elem.data.length > 0) {
                          for (let r = 0; r < elem.data.length; r++) {
                            for (let c = 0; c < elem.data[r].length; c++) {
                              const cellContent = elem.data[r][c];
                              if (cellContent && cellContent.trim()) {
                                description += `   [${r+1},${c+1}]: "${cellContent}"\n`;
                              }
                            }
                          }
                        }
                      }
                    }
                    
                    description += '\n【使用说明】\n';
                    description += '- 用 create_from_template 时，search 参数必须从上面的引号内复制\n';
                    description += '- 例如要替换表格中的地点，search 应该是: "精工园3-102"\n';
                    
                    resolve(description);
                  } catch (e) {
                    resolve('');
                  }
                } else {
                  resolve('');
                }
              }
            );
          } catch (e) {
            console.error('获取文档结构失败:', e);
            resolve('');
          }
        });
      },

      // 执行 Document Builder 脚本（创建/修改带格式的内容）
      executeScript: async (script: string): Promise<boolean> => {
        const connector = getConnector()
        if (!connector) return false
        return new Promise((resolve) => {
          try {
            connector.callCommand(() => {
              // 这里执行 Document Builder 脚本
              eval(script)
            }, (result: any) => {
              console.log('执行脚本结果:', result)
              resolve(true)
            })
          } catch (e) {
            console.error('执行脚本失败:', e)
            resolve(false)
          }
        })
      },

      // 在文档末尾添加带格式的段落
      addFormattedParagraph: async (text: string, options?: {
        fontSize?: number
        fontFamily?: string
        bold?: boolean
        italic?: boolean
        color?: string
        alignment?: 'left' | 'center' | 'right' | 'justify'
      }): Promise<boolean> => {
        const connector = getConnector()
        if (!connector) return false
        return new Promise((resolve) => {
          try {
            // 构建脚本字符串，在 callCommand 内部执行
            const script = `
              var oDocument = Api.GetDocument();
              var oParagraph = Api.CreateParagraph();
              var oRun = Api.CreateRun();
              oRun.AddText("${text.replace(/"/g, '\\"').replace(/\n/g, '\\n')}");
              ${options?.fontSize ? `oRun.SetFontSize(${options.fontSize * 2});` : ''}
              ${options?.fontFamily ? `oRun.SetFontFamily("${options.fontFamily}");` : ''}
              ${options?.bold ? 'oRun.SetBold(true);' : ''}
              ${options?.italic ? 'oRun.SetItalic(true);' : ''}
              oParagraph.AddElement(oRun);
              ${options?.alignment ? `oParagraph.SetJc("${options.alignment === 'justify' ? 'both' : options.alignment}");` : ''}
              oDocument.Push(oParagraph);
            `
            
            connector.callCommand(
              function() { eval(script) },
              (result: any) => {
                console.log('添加段落结果:', result)
                resolve(true)
              }
            )
          } catch (e) {
            console.error('添加段落失败:', e)
            resolve(false)
          }
        })
      },

      // 添加表格
      addTable: async (rows: number, cols: number, data?: string[][]): Promise<boolean> => {
        const connector = getConnector()
        if (!connector) return false
        return new Promise((resolve) => {
          try {
            const dataStr = data ? JSON.stringify(data) : 'null'
            const script = `
              var oDocument = Api.GetDocument();
              var oTable = Api.CreateTable(${cols}, ${rows});
              oTable.SetWidth("percent", 100);
              var data = ${dataStr};
              if (data) {
                for (var i = 0; i < Math.min(${rows}, data.length); i++) {
                  var oRow = oTable.GetRow(i);
                  for (var j = 0; j < Math.min(${cols}, (data[i] ? data[i].length : 0)); j++) {
                    var oCell = oRow.GetCell(j);
                    var oCellContent = oCell.GetContent();
                    var oParagraph = oCellContent.GetElement(0);
                    oParagraph.AddText(data[i][j] || "");
                  }
                }
              }
              oDocument.Push(oTable);
            `
            
            connector.callCommand(
              function() { eval(script) },
              (result: any) => {
                console.log('添加表格结果:', result)
                resolve(true)
              }
            )
          } catch (e) {
            console.error('添加表格失败:', e)
            resolve(false)
          }
        })
      },

      // 添加标题
      addHeading: async (text: string, level: 1 | 2 | 3 | 4 | 5 | 6): Promise<boolean> => {
        const connector = getConnector()
        if (!connector) return false
        return new Promise((resolve) => {
          try {
            const script = `
              var oDocument = Api.GetDocument();
              var oParagraph = Api.CreateParagraph();
              oParagraph.AddText("${text.replace(/"/g, '\\"').replace(/\n/g, '\\n')}");
              oParagraph.SetStyle(oDocument.GetStyle("Heading ${level}"));
              oDocument.Push(oParagraph);
            `
            
            connector.callCommand(
              function() { eval(script) },
              (result: any) => {
                console.log('添加标题结果:', result)
                resolve(true)
              }
            )
          } catch (e) {
            console.error('添加标题失败:', e)
            resolve(false)
          }
        })
      },

      // 全选文档内容
      selectAll: async (): Promise<boolean> => {
        const connector = getConnector()
        if (!connector) return false
        return new Promise((resolve) => {
          try {
            const script = `
              var oDocument = Api.GetDocument();
              oDocument.SelectAll();
            `
            connector.callCommand(
              function() { eval(script) },
              () => resolve(true)
            )
          } catch (e) {
            resolve(false)
          }
        })
      },

      // 复制选中内容
      copy: async (): Promise<boolean> => {
        const connector = getConnector()
        if (!connector) return false
        return new Promise((resolve) => {
          try {
            connector.executeMethod('Copy', [], () => resolve(true))
          } catch (e) {
            resolve(false)
          }
        })
      },

      // 粘贴
      paste: async (): Promise<boolean> => {
        const connector = getConnector()
        if (!connector) return false
        return new Promise((resolve) => {
          try {
            connector.executeMethod('Paste', [], () => resolve(true))
          } catch (e) {
            resolve(false)
          }
        })
      },

      // 清空文档（保留格式模板）
      clearContent: async (): Promise<boolean> => {
        const connector = getConnector()
        if (!connector) return false
        return new Promise((resolve) => {
          try {
            const script = `
              var oDocument = Api.GetDocument();
              var nCount = oDocument.GetElementsCount();
              for (var i = nCount - 1; i > 0; i--) {
                oDocument.RemoveElement(i);
              }
              var oFirstPara = oDocument.GetElement(0);
              if (oFirstPara) {
                var oRange = oFirstPara.GetRange(0, oFirstPara.GetText().length);
                oRange.Delete();
              }
            `
            connector.callCommand(
              function() { eval(script) },
              () => resolve(true)
            )
          } catch (e) {
            resolve(false)
          }
        })
      },

      // 批量替换多个内容（用于模板填充）
      batchReplace: async (replacements: Array<{search: string, replace: string}>): Promise<number> => {
        let successCount = 0
        for (const item of replacements) {
          const result = await window.onlyOfficeConnector?.searchAndReplace(item.search, item.replace, true)
          if (result) successCount++
        }
        return successCount
      },

      // 创建带格式的文档内容（逐个添加元素）
      createFormattedContent: async (elements: Array<{
        type: 'heading' | 'paragraph' | 'table'
        content?: string
        level?: number
        bold?: boolean
        fontSize?: number
        fontFamily?: string
        alignment?: 'left' | 'center' | 'right' | 'justify'
        rows?: number
        cols?: number
        data?: string[][]
      }>): Promise<boolean> => {
        console.log('createFormattedContent 开始执行，元素数量:', elements.length)
        
        try {
          // 逐个添加元素
          for (let i = 0; i < elements.length; i++) {
            const elem = elements[i]
            console.log(`添加元素 ${i + 1}/${elements.length}:`, elem.type)
            
            if (elem.type === 'heading' && elem.content) {
              await window.onlyOfficeConnector?.addHeading(elem.content, (elem.level || 1) as 1|2|3|4|5|6)
            } else if (elem.type === 'paragraph' && elem.content) {
              await window.onlyOfficeConnector?.addFormattedParagraph(elem.content, {
                bold: elem.bold,
                fontSize: elem.fontSize,
                fontFamily: elem.fontFamily,
                alignment: elem.alignment
              })
            } else if (elem.type === 'table' && elem.rows && elem.cols) {
              await window.onlyOfficeConnector?.addTable(elem.rows, elem.cols, elem.data)
            }
            
            // 每个元素之间稍等一下
            await new Promise(resolve => setTimeout(resolve, 200))
          }
          
          console.log('createFormattedContent 完成')
          return true
        } catch (e) {
          console.error('创建格式化内容失败:', e)
          return false
        }
      },

      // 搜索并替换文本 - 使用 callCommand 直接操作文档 API
      searchAndReplace: async (searchText: string, replaceText: string, _replaceAll = true): Promise<boolean> => {
        const connector = getConnector()
        if (!connector) {
          console.warn('searchAndReplace: connector 不可用')
          return false
        }
        
        console.log('🔍 开始搜索替换:', searchText, '->', replaceText)
        
        return new Promise((resolve) => {
          try {
            // 使用 Asc.scope 传递参数到 callCommand 内部
            // @ts-ignore
            if (!window.Asc) window.Asc = {}
            // @ts-ignore
            window.Asc.scope = {
              searchText: searchText,
              replaceText: replaceText
            }
            
            // 使用 callCommand 直接操作 Document Builder API
            connector.callCommand(
              function() {
                // @ts-ignore - Api 和 Asc 是 ONLYOFFICE 全局对象
                var searchStr = Asc.scope.searchText;
                var replaceStr = Asc.scope.replaceText;
                var oDocument = Api.GetDocument();
                
                // 方法1: 使用 Search API（最直接的方式）
                var searchResults = oDocument.Search(searchStr, true); // true = 高亮
                
                if (searchResults && searchResults.length > 0) {
                  for (var i = 0; i < searchResults.length; i++) {
                    searchResults[i].SetText(replaceStr);
                  }
                  return true;
                }
                
                // 方法2: 如果 Search 没找到，尝试遍历所有段落和表格
                var found = false;
                var count = oDocument.GetElementsCount();
                
                for (var i = 0; i < count; i++) {
                  var elem = oDocument.GetElement(i);
                  if (!elem) continue;
                  
                  var classType = elem.GetClassType();
                  
                  if (classType === 'paragraph') {
                    var text = elem.GetText();
                    if (text && text.indexOf(searchStr) !== -1) {
                      // 找到了，再次搜索并替换
                      var results2 = oDocument.Search(searchStr, false);
                      if (results2 && results2.length > 0) {
                        for (var j = 0; j < results2.length; j++) {
                          results2[j].SetText(replaceStr);
                        }
                        found = true;
                      }
                      break;
                    }
                  } else if (classType === 'table') {
                    // 表格内搜索
                    var rowCount = elem.GetRowsCount();
                    for (var r = 0; r < rowCount && !found; r++) {
                      var row = elem.GetRow(r);
                      if (!row) continue;
                      var cellCount = row.GetCellsCount();
                      for (var c = 0; c < cellCount && !found; c++) {
                        var cell = row.GetCell(c);
                        if (!cell) continue;
                        var content = cell.GetContent();
                        if (content && content.GetElement) {
                          var para = content.GetElement(0);
                          if (para && para.GetText) {
                            var cellText = para.GetText();
                            if (cellText && cellText.indexOf(searchStr) !== -1) {
                              // 在表格中找到了
                              var results3 = oDocument.Search(searchStr, false);
                              if (results3 && results3.length > 0) {
                                for (var k = 0; k < results3.length; k++) {
                                  results3[k].SetText(replaceStr);
                                }
                                found = true;
                              }
                            }
                          }
                        }
                      }
                    }
                  }
                }
                
                return found;
              },
              (result: any) => {
                console.log('callCommand 搜索替换结果:', result)
                resolve(result === true)
              },
              true // isNoCalc - 不重新计算，提高性能
            )
          } catch (e) {
            console.error('搜索替换失败:', e)
            resolve(false)
          }
        })
      },

      // 在光标位置插入文本
      insertText: async (text: string): Promise<boolean> => {
        const connector = getConnector()
        if (!connector) return false
        return new Promise((resolve) => {
          try {
            connector.executeMethod('InsertText', [text], (result: any) => {
              console.log('插入文本结果:', result)
              resolve(true)
            })
          } catch (e) {
            console.error('插入文本失败:', e)
            resolve(false)
          }
        })
      },

      // 获取选中的文本
      getSelectedText: async (): Promise<string> => {
        const connector = getConnector()
        if (!connector) return ''
        return new Promise((resolve) => {
          connector.executeMethod('GetSelectedText', [], (text: string) => {
            resolve(text || '')
          })
        })
      },

      // 替换选中的文本
      replaceSelectedText: async (text: string): Promise<boolean> => {
        const connector = getConnector()
        if (!connector) return false
        return new Promise((resolve) => {
          try {
            connector.executeMethod('ReplaceSelectedText', [text], (result: any) => {
              console.log('替换选中文本结果:', result)
              resolve(true)
            })
          } catch (e) {
            console.error('替换选中文本失败:', e)
            resolve(false)
          }
        })
      },

      // 保存文档
      saveDocument: async (): Promise<boolean> => {
        const connector = getConnector()
        if (!connector) return false
        return new Promise((resolve) => {
          try {
            connector.executeMethod('Save', [], (result: any) => {
              console.log('保存文档结果:', result)
              resolve(true)
            })
          } catch (e) {
            console.error('保存文档失败:', e)
            resolve(false)
          }
        })
      },

      // 下载文档
      downloadDocument: () => {
        if (editorInstanceRef.current) {
          editorInstanceRef.current.downloadAs()
        }
      }
    }

    console.log('✅ ONLYOFFICE connector 接口已初始化')
  }, [getConnector])

  // 加载 ONLYOFFICE API 脚本
  useEffect(() => {
    // 检查是否已加载
    if (window.DocsAPI) {
      setIsApiLoaded(true)
      setIsLoading(false)
      return
    }

    const existingScript = document.querySelector(
      `script[src="${DOCUMENT_SERVER_URL}/web-apps/apps/api/documents/api.js"]`
    )
    
    if (existingScript) {
      // 等待脚本加载完成
      const checkApi = setInterval(() => {
        if (window.DocsAPI) {
          setIsApiLoaded(true)
          setIsLoading(false)
          clearInterval(checkApi)
        }
      }, 500)
      
      setTimeout(() => {
        clearInterval(checkApi)
        if (!window.DocsAPI) {
          setError('ONLYOFFICE API 加载超时')
          setIsLoading(false)
        }
      }, 30000)
      return
    }

    const script = document.createElement('script')
    script.src = `${DOCUMENT_SERVER_URL}/web-apps/apps/api/documents/api.js`
    script.async = true
    
    script.onload = () => {
      console.log('ONLYOFFICE API 脚本已加载')
      // 等待 DocsAPI 对象可用
      const checkApi = setInterval(() => {
        if (window.DocsAPI) {
          console.log('ONLYOFFICE DocsAPI 已就绪')
          setIsApiLoaded(true)
          setIsLoading(false)
          clearInterval(checkApi)
        }
      }, 200)
      
      setTimeout(() => {
        clearInterval(checkApi)
        if (!window.DocsAPI) {
          setError('ONLYOFFICE API 初始化失败')
          setIsLoading(false)
        }
      }, 10000)
    }
    
    script.onerror = (e) => {
      console.error('加载 ONLYOFFICE API 失败:', e)
      setError('无法连接到 ONLYOFFICE 服务器。请确保 Docker 容器正在运行。')
      setIsLoading(false)
    }

    document.head.appendChild(script)
  }, [])

  // 初始化编辑器
  useEffect(() => {
    if (!isApiLoaded || !window.DocsAPI) {
      return
    }

    // 确保容器存在
    const container = document.getElementById('onlyoffice-editor-container')
    if (!container) {
      console.error('编辑器容器不存在')
      return
    }

    // 清空容器
    container.innerHTML = ''

    // 销毁现有编辑器实例
    if (editorInstanceRef.current) {
      try {
        editorInstanceRef.current.destroyEditor()
      } catch (e) {
        console.log('销毁编辑器:', e)
      }
      editorInstanceRef.current = null
      connectorRef.current = null
    }

    // 如果没有文件或没有文件 URL，显示提示
    if (!currentFile || !fileUrl) {
      container.innerHTML = `
        <div style="display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100%; color: #888;">
          <svg width="64" height="64" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5">
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
            <polyline points="14 2 14 8 20 8"></polyline>
            <line x1="16" y1="13" x2="8" y2="13"></line>
            <line x1="16" y1="17" x2="8" y2="17"></line>
            <polyline points="10 9 9 9 8 9"></polyline>
          </svg>
          <p style="margin-top: 16px; font-size: 16px;">请从左侧选择一个文档</p>
          <p style="margin-top: 8px; font-size: 12px; color: #666;">支持 .docx, .xlsx, .pptx 格式</p>
        </div>
      `
      return
    }

    // 获取文件类型
    const fileExt = currentFile.name.split('.').pop()?.toLowerCase() || 'docx'
    let documentType: 'word' | 'cell' | 'slide' = 'word'
    
    if (['xlsx', 'xls', 'ods', 'csv'].includes(fileExt)) {
      documentType = 'cell'
    } else if (['pptx', 'ppt', 'odp'].includes(fileExt)) {
      documentType = 'slide'
    }

    // 创建编辑器配置
    const config = {
      document: {
        fileType: fileExt,
        key: docKey,
        title: currentFile.name,
        url: fileUrl,
        permissions: {
          edit: true,
          download: true,
          print: true,
          review: true,
          comment: true,
        }
      },
      documentType,
      editorConfig: {
        mode: 'edit',
        lang: 'zh-CN',
        callbackUrl: '', // 本地测试不需要回调
        user: {
          id: 'user-1',
          name: '当前用户'
        },
        customization: {
          autosave: false,
          chat: false,
          comments: true,
          compactHeader: true,
          compactToolbar: false,
          feedback: false,
          forcesave: false,
          help: false,
          hideRightMenu: false,
          hideRulers: false,
          logo: {
            image: '',
            imageDark: '',
            url: ''
          },
          macros: false,
          plugins: true, // 启用插件以支持 connector
          toolbarHideFileName: false,
          toolbarNoTabs: false,
          uiTheme: 'theme-dark',
          unit: 'cm',
          zoom: 100
        }
      },
      height: '100%',
      width: '100%',
      type: 'desktop',
      events: {
        onAppReady: () => {
          console.log('✅ ONLYOFFICE 编辑器已就绪，文件:', currentFile.name)
          setIsEditorReady(true)
          
          // 获取 connector - 需要检查方法是否存在
          console.log('editorInstanceRef.current:', editorInstanceRef.current)
          console.log('editorInstanceRef.current 方法:', editorInstanceRef.current ? Object.keys(editorInstanceRef.current) : 'null')
          
          if (editorInstanceRef.current) {
            try {
              // 检查 createConnector 方法是否存在
              console.log('typeof createConnector:', typeof editorInstanceRef.current.createConnector)
              
              if (typeof editorInstanceRef.current.createConnector === 'function') {
                connectorRef.current = editorInstanceRef.current.createConnector()
                console.log('✅ ONLYOFFICE connector 已创建:', connectorRef.current)
                console.log('connector 方法:', connectorRef.current ? Object.keys(connectorRef.current) : 'null')
                initConnectorInterface()
              } else {
                console.warn('⚠️ createConnector 方法不可用')
                console.log('尝试直接访问 window.Asc.plugin.connector...')
                
                // 备用方式：尝试从全局获取 connector
                setTimeout(() => {
                  // 方式1：再次尝试 createConnector
                  if (editorInstanceRef.current && typeof editorInstanceRef.current.createConnector === 'function') {
                    connectorRef.current = editorInstanceRef.current.createConnector()
                    console.log('✅ ONLYOFFICE connector 已创建（延迟）')
                    initConnectorInterface()
                  } else {
                    // 方式2：尝试从 window.Asc 获取
                    const asc = (window as any).Asc
                    if (asc && asc.plugin && asc.plugin.connector) {
                      connectorRef.current = asc.plugin.connector
                      console.log('✅ 从 window.Asc.plugin.connector 获取到 connector')
                      initConnectorInterface()
                    } else {
                      console.error('❌ 无法获取 ONLYOFFICE connector')
                    }
                  }
                }, 3000)
              }
            } catch (e) {
              console.error('创建 connector 失败:', e)
            }
          }
        },
        onDocumentStateChange: (event: any) => {
          console.log('文档状态:', event.data ? '已修改' : '未修改')
          // 更新 DocumentContext 中的修改状态
          if (setDocument) {
            setDocument(prev => ({
              ...prev,
              isModified: event.data
            }))
          }
        },
        onError: (event: any) => {
          console.error('ONLYOFFICE 错误:', event)
        },
        onWarning: (event: any) => {
          console.warn('ONLYOFFICE 警告:', event)
        },
        onDownloadAs: (event: any) => {
          console.log('文档下载:', event)
        }
      }
    }

    console.log('初始化 ONLYOFFICE 编辑器，配置:', {
      ...config,
      document: { ...config.document, url: fileUrl }
    })

    try {
      editorInstanceRef.current = new window.DocsAPI.DocEditor(
        'onlyoffice-editor-container',
        config
      )
      console.log('编辑器实例已创建')
    } catch (e) {
      console.error('创建编辑器失败:', e)
      setError(`创建编辑器失败: ${e}`)
    }

    return () => {
      if (editorInstanceRef.current) {
        try {
          editorInstanceRef.current.destroyEditor()
        } catch (e) {
          // 忽略
        }
        editorInstanceRef.current = null
        connectorRef.current = null
        window.onlyOfficeConnector = undefined
      }
    }
  }, [isApiLoaded, docKey, currentFile, fileUrl, initConnectorInterface, setDocument])

  // 重试
  const handleRetry = () => {
    setError(null)
    setIsLoading(true)
    setIsApiLoaded(false)
    
    // 移除旧脚本
    const oldScript = document.querySelector(
      `script[src="${DOCUMENT_SERVER_URL}/web-apps/apps/api/documents/api.js"]`
    )
    if (oldScript) {
      oldScript.remove()
    }
    
    // 重新加载
    setTimeout(() => {
      window.location.reload()
    }, 500)
  }

  // 渲染加载状态
  if (isLoading) {
    return (
      <div className="flex-1 flex flex-col items-center justify-center bg-background">
        <Loader2 className="w-12 h-12 text-primary animate-spin mb-4" />
        <p className="text-text-muted">正在加载 ONLYOFFICE 编辑器...</p>
        <p className="text-xs text-text-dim mt-2">首次加载可能需要 30-60 秒</p>
      </div>
    )
  }

  // 渲染错误状态
  if (error) {
    return (
      <div className="flex-1 flex flex-col items-center justify-center bg-background p-8">
        <AlertCircle className="w-12 h-12 text-red-500 mb-4" />
        <p className="text-red-400 mb-4 text-center">{error}</p>
        <button
          onClick={handleRetry}
          className="flex items-center gap-2 px-4 py-2 bg-primary text-white rounded-lg hover:bg-primary-hover transition-colors"
        >
          <RefreshCw className="w-4 h-4" />
          重试
        </button>
        <div className="mt-6 p-4 bg-surface rounded-lg max-w-md">
          <p className="text-sm text-text-muted mb-2">请确保：</p>
          <ul className="text-xs text-text-dim space-y-1 list-disc list-inside">
            <li>Docker Desktop 正在运行</li>
            <li>ONLYOFFICE 容器已启动</li>
            <li>端口 8080 可以访问</li>
          </ul>
          <div className="mt-3 p-2 bg-background rounded text-xs font-mono">
            <p className="text-text-dim mb-1">启动命令：</p>
            <code className="text-green-400">docker start onlyoffice-ds</code>
          </div>
        </div>
      </div>
    )
  }

  return (
    <div className="flex-1 flex flex-col bg-background overflow-hidden" style={{ height: '100%' }}>
      {/* 状态栏 */}
      {currentFile && fileUrl && (
        <div className="px-4 py-2 bg-green-500/10 border-b border-green-500/20 flex items-center gap-2 flex-shrink-0">
          <CheckCircle className="w-4 h-4 text-green-400" />
          <span className="text-xs text-green-400">
            ONLYOFFICE - {currentFile.name}
          </span>
          {isEditorReady && (
            <span className="ml-auto text-xs text-green-400/60">
              AI 已就绪
            </span>
          )}
        </div>
      )}
      
      {!currentFile && (
        <div className="px-4 py-2 bg-blue-500/10 border-b border-blue-500/20 flex items-center gap-2 flex-shrink-0">
          <FileText className="w-4 h-4 text-blue-400" />
          <span className="text-xs text-blue-400">
            ONLYOFFICE 编辑器 - 请选择一个文档
          </span>
        </div>
      )}
      
      {/* 编辑器容器 - 使用 iframe 样式确保正确显示 */}
      <div 
        ref={editorContainerRef}
        id="onlyoffice-editor-container"
        className="flex-1 w-full"
        style={{ 
          height: 'calc(100% - 40px)', 
          minHeight: '500px',
          position: 'relative'
        }}
      />
    </div>
  )
}