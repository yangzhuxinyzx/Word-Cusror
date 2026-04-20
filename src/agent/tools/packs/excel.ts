import type { FileItem } from '../../../types'
import type { ExcelDomainAdapter } from '../../adapters/excel/ExcelDomainAdapter'
import { defineAgentTool } from '../contracts'
import type { ExecutableAgentTool } from '../executor'

export interface ExcelToolPackDeps {
  adapter: ExcelDomainAdapter
  registerToolActivity: (tool: string, label: string) => string
  completeToolActivity: (
    activityId: string,
    status: 'success' | 'error' | 'skipped',
    detail?: string,
  ) => void
  updateAgentAction: (action: string) => void
  truncateLabel: (text: string, limit?: number) => string
}

function getExcelContext(deps: ExcelToolPackDeps, tool: string) {
  const context = deps.adapter.ensureWorkbook()
  if (!context.ok) {
    return {
      ok: false as const,
      result: { tool, success: false, message: context.message },
    }
  }

  return {
    ok: true as const,
    excelFilePath: context.excelFilePath,
  }
}

function ok(tool: string, message: string, data?: Record<string, unknown>) {
  return { tool, success: true, message, data }
}

function fail(tool: string, message: string) {
  return { tool, success: false, message }
}

export function createExcelToolPack(
  deps: ExcelToolPackDeps,
): ExecutableAgentTool[] {
  return [
    defineAgentTool({
      id: 'excel_read',
      displayName: 'Excel Read',
      description: 'Read cell ranges from the active Excel workbook',
      domain: 'excel',
      mutation: 'read',
      concurrency: 'parallel_safe',
      async handler(args) {
        const context = getExcelContext(deps, 'excel_read')
        if (!context.ok) return context.result
        const sheet = args.sheet || 'Sheet1'
        const range = args.range || 'A1'
        const activityId = deps.registerToolActivity('excel_read', `Read: ${sheet}!${range}`)
        try {
          const result = await deps.adapter.readCells(context.excelFilePath, sheet, range)
          if (result.success && result.cells) {
            deps.completeToolActivity(activityId, 'success', `${result.cells.length} cells`)
            return ok('excel_read', `Read ${sheet}!${range}.`, {
              cells: result.cells,
              count: result.cells.length,
            })
          }
          deps.completeToolActivity(activityId, 'error', result.error)
          return fail('excel_read', result.error || 'Read failed')
        } catch (error) {
          deps.completeToolActivity(activityId, 'error', 'Read failed')
          return fail('excel_read', `Read failed: ${error}`)
        }
      },
    }),
    defineAgentTool({
      id: 'excel_search',
      displayName: 'Excel Search',
      description: 'Search within the active Excel workbook',
      domain: 'excel',
      mutation: 'read',
      concurrency: 'parallel_safe',
      async handler(args) {
        const context = getExcelContext(deps, 'excel_search')
        if (!context.ok) return context.result
        const sheet = args.sheet || 'Sheet1'
        const text = args.text || args.searchText || ''
        if (!text) return fail('excel_search', 'Missing search text.')
        const activityId = deps.registerToolActivity('excel_search', `Search: ${deps.truncateLabel(text, 20)}`)
        try {
          const result = await deps.adapter.search(context.excelFilePath, sheet, text)
          if (result.success) {
            deps.completeToolActivity(activityId, 'success', `${result.count || 0} hits`)
            return ok('excel_search', `Found ${result.count || 0} matches in ${sheet}.`, {
              results: result.results,
              count: result.count,
            })
          }
          deps.completeToolActivity(activityId, 'error', result.error)
          return fail('excel_search', result.error || 'Search failed')
        } catch (error) {
          deps.completeToolActivity(activityId, 'error', 'Search failed')
          return fail('excel_search', `Search failed: ${error}`)
        }
      },
    }),
    defineAgentTool({
      id: 'excel_write',
      displayName: 'Excel Write',
      description: 'Write cell values into the active Excel workbook',
      domain: 'excel',
      mutation: 'write',
      concurrency: 'serial',
      async handler(args) {
        const context = getExcelContext(deps, 'excel_write')
        if (!context.ok) return context.result
        const sheet = args.sheet || 'Sheet1'
        let updates: Array<{ address: string; value?: unknown; style?: unknown }> = []
        if (args.updates) {
          try {
            updates = JSON.parse(args.updates)
          } catch {
            return fail('excel_write', 'Invalid updates JSON.')
          }
        }
        if (updates.length === 0) return fail('excel_write', 'Missing updates.')
        const activityId = deps.registerToolActivity('excel_write', `Write: ${sheet}`)
        deps.updateAgentAction(`Writing ${updates.length} cells...`)
        try {
          const result = await deps.adapter.writeCells(context.excelFilePath, sheet, updates)
          if (result.success) {
            await deps.adapter.refreshWorkbookPreview()
            deps.completeToolActivity(activityId, 'success', `${result.count || 0} cells`)
            return ok('excel_write', `Updated ${result.count || 0} cells.`, {
              updatedCells: result.updatedCells,
              count: result.count,
            })
          }
          deps.completeToolActivity(activityId, 'error', result.error)
          return fail('excel_write', result.error || 'Write failed')
        } catch (error) {
          deps.completeToolActivity(activityId, 'error', 'Write failed')
          return fail('excel_write', `Write failed: ${error}`)
        }
      },
    }),
    defineAgentTool({
      id: 'excel_insert_rows',
      displayName: 'Excel Insert Rows',
      description: 'Insert rows into the active Excel workbook',
      domain: 'excel',
      mutation: 'write',
      concurrency: 'serial',
      async handler(args) {
        const context = getExcelContext(deps, 'excel_insert_rows')
        if (!context.ok) return context.result
        const sheet = args.sheet || 'Sheet1'
        const startRow = parseInt(args.startRow, 10) || 1
        const count = parseInt(args.count, 10) || 1
        let data: unknown[][] | undefined
        if (args.data) {
          try {
            data = JSON.parse(args.data)
          } catch {
            data = undefined
          }
        }
        const activityId = deps.registerToolActivity('excel_insert_rows', `Insert rows: ${startRow}`)
        try {
          const result = await deps.adapter.insertRows(
            context.excelFilePath,
            sheet,
            startRow,
            count,
            data,
          )
          if (result.success) {
            await deps.adapter.refreshWorkbookPreview()
            deps.completeToolActivity(activityId, 'success', `${count} rows`)
            return ok('excel_insert_rows', `Inserted ${count} row(s).`, {
              insertedAt: result.insertedAt,
              count: result.count,
            })
          }
          deps.completeToolActivity(activityId, 'error', result.error)
          return fail('excel_insert_rows', result.error || 'Insert rows failed')
        } catch (error) {
          deps.completeToolActivity(activityId, 'error', 'Insert rows failed')
          return fail('excel_insert_rows', `Insert rows failed: ${error}`)
        }
      },
    }),
    defineAgentTool({
      id: 'excel_insert_columns',
      displayName: 'Excel Insert Columns',
      description: 'Insert columns into the active Excel workbook',
      domain: 'excel',
      mutation: 'write',
      concurrency: 'serial',
      async handler(args) {
        const context = getExcelContext(deps, 'excel_insert_columns')
        if (!context.ok) return context.result
        const sheet = args.sheet || 'Sheet1'
        const startCol = parseInt(args.startCol, 10) || 1
        const count = parseInt(args.count, 10) || 1
        const activityId = deps.registerToolActivity('excel_insert_columns', `Insert cols: ${startCol}`)
        try {
          const result = await deps.adapter.insertColumns(
            context.excelFilePath,
            sheet,
            startCol,
            count,
          )
          if (result.success) {
            await deps.adapter.refreshWorkbookPreview()
            deps.completeToolActivity(activityId, 'success', `${count} cols`)
            return ok('excel_insert_columns', `Inserted ${count} column(s).`, {
              insertedAt: result.insertedAt,
              count: result.count,
            })
          }
          deps.completeToolActivity(activityId, 'error', result.error)
          return fail('excel_insert_columns', result.error || 'Insert columns failed')
        } catch (error) {
          deps.completeToolActivity(activityId, 'error', 'Insert columns failed')
          return fail('excel_insert_columns', `Insert columns failed: ${error}`)
        }
      },
    }),
    defineAgentTool({
      id: 'excel_delete_rows',
      displayName: 'Excel Delete Rows',
      description: 'Delete rows from the active Excel workbook',
      domain: 'excel',
      mutation: 'delete',
      concurrency: 'serial',
      async handler(args) {
        const context = getExcelContext(deps, 'excel_delete_rows')
        if (!context.ok) return context.result
        const sheet = args.sheet || 'Sheet1'
        const startRow = parseInt(args.startRow, 10) || 1
        const count = parseInt(args.count, 10) || 1
        const activityId = deps.registerToolActivity('excel_delete_rows', `Delete rows: ${startRow}`)
        try {
          const result = await deps.adapter.deleteRows(
            context.excelFilePath,
            sheet,
            startRow,
            count,
          )
          if (result.success) {
            await deps.adapter.refreshWorkbookPreview()
            deps.completeToolActivity(activityId, 'success', `${count} rows`)
            return ok('excel_delete_rows', `Deleted ${count} row(s).`, {
              deletedFrom: result.deletedFrom,
              count: result.count,
            })
          }
          deps.completeToolActivity(activityId, 'error', result.error)
          return fail('excel_delete_rows', result.error || 'Delete rows failed')
        } catch (error) {
          deps.completeToolActivity(activityId, 'error', 'Delete rows failed')
          return fail('excel_delete_rows', `Delete rows failed: ${error}`)
        }
      },
    }),
    defineAgentTool({
      id: 'excel_delete_columns',
      displayName: 'Excel Delete Columns',
      description: 'Delete columns from the active Excel workbook',
      domain: 'excel',
      mutation: 'delete',
      concurrency: 'serial',
      async handler(args) {
        const context = getExcelContext(deps, 'excel_delete_columns')
        if (!context.ok) return context.result
        const sheet = args.sheet || 'Sheet1'
        const startCol = parseInt(args.startCol, 10) || 1
        const count = parseInt(args.count, 10) || 1
        const activityId = deps.registerToolActivity('excel_delete_columns', `Delete cols: ${startCol}`)
        try {
          const result = await deps.adapter.deleteColumns(
            context.excelFilePath,
            sheet,
            startCol,
            count,
          )
          if (result.success) {
            await deps.adapter.refreshWorkbookPreview()
            deps.completeToolActivity(activityId, 'success', `${count} cols`)
            return ok('excel_delete_columns', `Deleted ${count} column(s).`, {
              deletedFrom: result.deletedFrom,
              count: result.count,
            })
          }
          deps.completeToolActivity(activityId, 'error', result.error)
          return fail('excel_delete_columns', result.error || 'Delete columns failed')
        } catch (error) {
          deps.completeToolActivity(activityId, 'error', 'Delete columns failed')
          return fail('excel_delete_columns', `Delete columns failed: ${error}`)
        }
      },
    }),
    defineAgentTool({
      id: 'excel_add_sheet',
      displayName: 'Excel Add Sheet',
      description: 'Add a worksheet to the active Excel workbook',
      domain: 'excel',
      mutation: 'create',
      concurrency: 'serial',
      async handler(args) {
        const context = getExcelContext(deps, 'excel_add_sheet')
        if (!context.ok) return context.result
        const name = args.name || 'Sheet2'
        const activityId = deps.registerToolActivity('excel_add_sheet', `Add sheet: ${name}`)
        try {
          const result = await deps.adapter.addSheet(context.excelFilePath, name)
          if (result.success) {
            await deps.adapter.refreshWorkbookPreview()
            deps.completeToolActivity(activityId, 'success')
            return ok('excel_add_sheet', `Added sheet ${result.sheetName || name}.`, {
              sheetName: result.sheetName || name,
            })
          }
          deps.completeToolActivity(activityId, 'error', result.error)
          return fail('excel_add_sheet', result.error || 'Add sheet failed')
        } catch (error) {
          deps.completeToolActivity(activityId, 'error', 'Add sheet failed')
          return fail('excel_add_sheet', `Add sheet failed: ${error}`)
        }
      },
    }),
    defineAgentTool({
      id: 'excel_delete_sheet',
      displayName: 'Excel Delete Sheet',
      description: 'Delete a worksheet from the active Excel workbook',
      domain: 'excel',
      mutation: 'delete',
      concurrency: 'serial',
      async handler(args) {
        const context = getExcelContext(deps, 'excel_delete_sheet')
        if (!context.ok) return context.result
        const name = args.name || ''
        if (!name) return fail('excel_delete_sheet', 'Missing sheet name.')
        const activityId = deps.registerToolActivity('excel_delete_sheet', `Delete sheet: ${name}`)
        try {
          const result = await deps.adapter.deleteSheet(context.excelFilePath, name)
          if (result.success) {
            await deps.adapter.refreshWorkbookPreview()
            deps.completeToolActivity(activityId, 'success')
            return ok('excel_delete_sheet', `Deleted sheet ${result.deletedSheet || name}.`, {
              deletedSheet: result.deletedSheet || name,
            })
          }
          deps.completeToolActivity(activityId, 'error', result.error)
          return fail('excel_delete_sheet', result.error || 'Delete sheet failed')
        } catch (error) {
          deps.completeToolActivity(activityId, 'error', 'Delete sheet failed')
          return fail('excel_delete_sheet', `Delete sheet failed: ${error}`)
        }
      },
    }),
    defineAgentTool({
      id: 'excel_merge',
      displayName: 'Excel Merge',
      description: 'Merge cells in the active Excel workbook',
      domain: 'excel',
      mutation: 'write',
      concurrency: 'serial',
      async handler(args) {
        const context = getExcelContext(deps, 'excel_merge')
        if (!context.ok) return context.result
        const sheet = args.sheet || 'Sheet1'
        const range = args.range || ''
        if (!range) return fail('excel_merge', 'Missing merge range.')
        const activityId = deps.registerToolActivity('excel_merge', `Merge: ${range}`)
        try {
          const result = await deps.adapter.mergeCells(context.excelFilePath, sheet, range)
          if (result.success) {
            await deps.adapter.refreshWorkbookPreview()
            deps.completeToolActivity(activityId, 'success')
            return ok('excel_merge', `Merged ${range}.`, {
              mergedRange: result.mergedRange,
            })
          }
          deps.completeToolActivity(activityId, 'error', result.error)
          return fail('excel_merge', result.error || 'Merge failed')
        } catch (error) {
          deps.completeToolActivity(activityId, 'error', 'Merge failed')
          return fail('excel_merge', `Merge failed: ${error}`)
        }
      },
    }),
    defineAgentTool({
      id: 'excel_unmerge',
      displayName: 'Excel Unmerge',
      description: 'Unmerge cells in the active Excel workbook',
      domain: 'excel',
      mutation: 'write',
      concurrency: 'serial',
      async handler(args) {
        const context = getExcelContext(deps, 'excel_unmerge')
        if (!context.ok) return context.result
        const sheet = args.sheet || 'Sheet1'
        const range = args.range || ''
        if (!range) return fail('excel_unmerge', 'Missing unmerge range.')
        const activityId = deps.registerToolActivity('excel_unmerge', `Unmerge: ${range}`)
        try {
          const result = await deps.adapter.unmergeCells(context.excelFilePath, sheet, range)
          if (result.success) {
            await deps.adapter.refreshWorkbookPreview()
            deps.completeToolActivity(activityId, 'success')
            return ok('excel_unmerge', `Unmerged ${range}.`, {
              unmergedRange: result.unmergedRange,
            })
          }
          deps.completeToolActivity(activityId, 'error', result.error)
          return fail('excel_unmerge', result.error || 'Unmerge failed')
        } catch (error) {
          deps.completeToolActivity(activityId, 'error', 'Unmerge failed')
          return fail('excel_unmerge', `Unmerge failed: ${error}`)
        }
      },
    }),
    defineAgentTool({
      id: 'excel_create',
      displayName: 'Excel Create',
      description: 'Create a new Excel workbook',
      domain: 'excel',
      mutation: 'create',
      concurrency: 'serial',
      async handler(args) {
        const workspacePath = deps.adapter.getWorkspacePath()
        if (!workspacePath) {
          return fail('excel_create', 'Please open a workspace folder first.')
        }
        const filename = args.filename || args.name || 'workbook.xlsx'
        let sheets: Array<{ name?: string; data?: unknown[][] }> = []
        if (args.sheets) {
          try {
            sheets = JSON.parse(args.sheets)
          } catch {
            sheets = []
          }
        }
        if (sheets.length === 0 && args.data) {
          try {
            const data = JSON.parse(args.data)
            sheets = [{ name: args.sheetName || 'Sheet1', data }]
          } catch {
            return fail('excel_create', 'Invalid workbook data JSON.')
          }
        }
        if (sheets.length === 0) sheets = [{ name: 'Sheet1', data: [] }]
        const finalFilename = filename.toLowerCase().endsWith('.xlsx') ? filename : `${filename}.xlsx`
        const filePath = `${workspacePath}/${finalFilename}`
        const activityId = deps.registerToolActivity('excel_create', `Create: ${finalFilename}`)
        try {
          const result = await deps.adapter.createWorkbook(filePath, {
            sheets,
            openAfterCreate: true,
          })
          if (result.success) {
            await deps.adapter.refreshWorkspaceFiles()
            if (result.openAfterCreate && result.filePath) {
              await deps.adapter.openWorkbook({
                name: finalFilename,
                path: result.filePath,
                type: 'file',
              })
            }
            deps.completeToolActivity(activityId, 'success')
            return ok('excel_create', `Created workbook ${finalFilename}.`, {
              filePath: result.filePath,
              fileName: finalFilename,
              sheetsCreated: result.sheetsCreated,
            })
          }
          deps.completeToolActivity(activityId, 'error', result.error)
          return fail('excel_create', result.error || 'Create workbook failed')
        } catch (error) {
          deps.completeToolActivity(activityId, 'error', 'Create workbook failed')
          return fail('excel_create', `Create workbook failed: ${error}`)
        }
      },
    }),
    defineAgentTool({
      id: 'excel_formula',
      displayName: 'Excel Formula',
      description: 'Set formulas in the active Excel workbook',
      domain: 'excel',
      mutation: 'write',
      concurrency: 'serial',
      async handler(args) {
        const context = getExcelContext(deps, 'excel_formula')
        if (!context.ok) return context.result
        const sheet = args.sheet || 'Sheet1'
        let formulas: Array<{ address: string; formula: string; numberFormat?: string }> = []
        try {
          if (args.formulas) {
            formulas = JSON.parse(args.formulas)
          } else if (args.address && args.formula) {
            formulas = [{ address: args.address, formula: args.formula, numberFormat: args.numberFormat }]
          }
        } catch {
          return fail('excel_formula', 'Invalid formulas JSON.')
        }
        if (formulas.length === 0) return fail('excel_formula', 'Missing formulas.')
        const activityId = deps.registerToolActivity('excel_formula', `Formula: ${formulas.length}`)
        try {
          const result = await deps.adapter.setFormula(context.excelFilePath, sheet, formulas)
          if (result.success) {
            await deps.adapter.refreshWorkbookPreview()
            deps.completeToolActivity(activityId, 'success')
            return ok('excel_formula', `Applied ${result.count || 0} formula(s).`, result)
          }
          deps.completeToolActivity(activityId, 'error', result.error)
          return fail('excel_formula', result.error || 'Set formula failed')
        } catch (error) {
          deps.completeToolActivity(activityId, 'error', 'Set formula failed')
          return fail('excel_formula', `Set formula failed: ${error}`)
        }
      },
    }),
    defineAgentTool({
      id: 'excel_sort',
      displayName: 'Excel Sort',
      description: 'Sort a range in the active Excel workbook',
      domain: 'excel',
      mutation: 'write',
      concurrency: 'serial',
      async handler(args) {
        const context = getExcelContext(deps, 'excel_sort')
        if (!context.ok) return context.result
        const sheet = args.sheet || 'Sheet1'
        const range = args.range || ''
        if (!range) return fail('excel_sort', 'Missing sort range.')
        const activityId = deps.registerToolActivity('excel_sort', `Sort: ${range}`)
        try {
          const result = await deps.adapter.sort(context.excelFilePath, sheet, {
            range,
            column: args.column || 'A',
            ascending: args.ascending !== 'false',
            hasHeader: args.hasHeader !== 'false',
          })
          if (result.success) {
            await deps.adapter.refreshWorkbookPreview()
            deps.completeToolActivity(activityId, 'success')
            return ok('excel_sort', `Sorted range ${range}.`, result)
          }
          deps.completeToolActivity(activityId, 'error', result.error)
          return fail('excel_sort', result.error || 'Sort failed')
        } catch (error) {
          deps.completeToolActivity(activityId, 'error', 'Sort failed')
          return fail('excel_sort', `Sort failed: ${error}`)
        }
      },
    }),
    defineAgentTool({
      id: 'excel_autofill',
      displayName: 'Excel Autofill',
      description: 'Autofill a range in the active Excel workbook',
      domain: 'excel',
      mutation: 'write',
      concurrency: 'serial',
      async handler(args) {
        const context = getExcelContext(deps, 'excel_autofill')
        if (!context.ok) return context.result
        const sourceRange = args.sourceRange || args.source || ''
        const targetRange = args.targetRange || args.target || ''
        if (!sourceRange || !targetRange) {
          return fail('excel_autofill', 'Missing source or target range.')
        }
        const activityId = deps.registerToolActivity('excel_autofill', `Autofill: ${sourceRange}`)
        try {
          const result = await deps.adapter.autoFill(
            context.excelFilePath,
            args.sheet || 'Sheet1',
            {
              sourceRange,
              targetRange,
              fillType: (args.fillType || args.type || 'copy') as
                | 'copy'
                | 'series'
                | 'formula',
            },
          )
          if (result.success) {
            await deps.adapter.refreshWorkbookPreview()
            deps.completeToolActivity(activityId, 'success')
            return ok('excel_autofill', `Autofilled ${targetRange}.`, result)
          }
          deps.completeToolActivity(activityId, 'error', result.error)
          return fail('excel_autofill', result.error || 'Autofill failed')
        } catch (error) {
          deps.completeToolActivity(activityId, 'error', 'Autofill failed')
          return fail('excel_autofill', `Autofill failed: ${error}`)
        }
      },
    }),
    defineAgentTool({
      id: 'excel_dimensions',
      displayName: 'Excel Dimensions',
      description: 'Set row and column dimensions in the active Excel workbook',
      domain: 'excel',
      mutation: 'write',
      concurrency: 'serial',
      async handler(args) {
        const context = getExcelContext(deps, 'excel_dimensions')
        if (!context.ok) return context.result
        let columns: Array<{ column: string | number; width?: number; hidden?: boolean }> = []
        let rows: Array<{ row: number; height?: number; hidden?: boolean }> = []
        try {
          if (args.columns) columns = JSON.parse(args.columns)
          if (args.rows) rows = JSON.parse(args.rows)
        } catch {
          return fail('excel_dimensions', 'Invalid dimensions JSON.')
        }
        const activityId = deps.registerToolActivity('excel_dimensions', `Dimensions: ${columns.length}/${rows.length}`)
        try {
          const result = await deps.adapter.setDimensions(
            context.excelFilePath,
            args.sheet || 'Sheet1',
            { columns, rows },
          )
          if (result.success) {
            await deps.adapter.refreshWorkbookPreview()
            deps.completeToolActivity(activityId, 'success')
            return ok('excel_dimensions', 'Updated dimensions.', result)
          }
          deps.completeToolActivity(activityId, 'error', result.error)
          return fail('excel_dimensions', result.error || 'Set dimensions failed')
        } catch (error) {
          deps.completeToolActivity(activityId, 'error', 'Set dimensions failed')
          return fail('excel_dimensions', `Set dimensions failed: ${error}`)
        }
      },
    }),
    defineAgentTool({
      id: 'excel_conditional_format',
      displayName: 'Excel Conditional Format',
      description: 'Apply conditional formatting rules in the active Excel workbook',
      domain: 'excel',
      mutation: 'write',
      concurrency: 'serial',
      async handler(args) {
        const context = getExcelContext(deps, 'excel_conditional_format')
        if (!context.ok) return context.result
        const range = args.range || ''
        if (!range) return fail('excel_conditional_format', 'Missing range.')
        let rules: unknown[] = []
        try {
          if (args.rules) {
            rules = JSON.parse(args.rules)
          } else if (args.type) {
            rules = [
              {
                type: args.type,
                operator: args.operator,
                value: args.value,
                fill: args.fill ? { bgColor: args.fill } : undefined,
              },
            ]
          }
        } catch {
          return fail('excel_conditional_format', 'Invalid rules JSON.')
        }
        const activityId = deps.registerToolActivity('excel_conditional_format', `Conditional: ${range}`)
        try {
          const result = await deps.adapter.conditionalFormat(
            context.excelFilePath,
            args.sheet || 'Sheet1',
            { range, rules },
          )
          if (result.success) {
            await deps.adapter.refreshWorkbookPreview()
            deps.completeToolActivity(activityId, 'success')
            return ok('excel_conditional_format', 'Conditional formatting updated.', result)
          }
          deps.completeToolActivity(activityId, 'error', result.error)
          return fail('excel_conditional_format', result.error || 'Conditional formatting failed')
        } catch (error) {
          deps.completeToolActivity(activityId, 'error', 'Conditional formatting failed')
          return fail('excel_conditional_format', `Conditional formatting failed: ${error}`)
        }
      },
    }),
    defineAgentTool({
      id: 'excel_calculate',
      displayName: 'Excel Calculate',
      description: 'Calculate formula results in the active Excel workbook',
      domain: 'excel',
      mutation: 'read',
      concurrency: 'parallel_safe',
      async handler(args) {
        const context = getExcelContext(deps, 'excel_calculate')
        if (!context.ok) return context.result
        let addresses: string[] = []
        try {
          if (args.addresses) {
            addresses = JSON.parse(args.addresses)
          } else if (args.address) {
            addresses = [args.address]
          }
        } catch {
          return fail('excel_calculate', 'Invalid address list.')
        }
        if (addresses.length === 0) return fail('excel_calculate', 'Missing addresses.')
        try {
          const result = await deps.adapter.calculate(
            context.excelFilePath,
            args.sheet || 'Sheet1',
            addresses,
          )
          if (result.success) {
            return ok('excel_calculate', `Calculated ${result.results?.length || 0} cell(s).`, {
              results: result.results,
            })
          }
          return fail('excel_calculate', result.error || 'Calculation failed')
        } catch (error) {
          return fail('excel_calculate', `Calculation failed: ${error}`)
        }
      },
    }),
    defineAgentTool({
      id: 'excel_filter',
      displayName: 'Excel Filter',
      description: 'Set or remove filters in the active Excel workbook',
      domain: 'excel',
      mutation: 'write',
      concurrency: 'serial',
      async handler(args) {
        const context = getExcelContext(deps, 'excel_filter')
        if (!context.ok) return context.result
        try {
          const result = await deps.adapter.setFilter(
            context.excelFilePath,
            args.sheet || 'Sheet1',
            {
              range: args.range || '',
              remove: (args.action || 'set').toLowerCase() === 'remove',
            },
          )
          if (result.success) {
            await deps.adapter.refreshWorkbookPreview()
            return ok('excel_filter', result.message || 'Filter updated.')
          }
          return fail('excel_filter', result.error || 'Filter failed')
        } catch (error) {
          return fail('excel_filter', `Filter failed: ${error}`)
        }
      },
    }),
    defineAgentTool({
      id: 'excel_validation',
      displayName: 'Excel Validation',
      description: 'Set data validation in the active Excel workbook',
      domain: 'excel',
      mutation: 'write',
      concurrency: 'serial',
      async handler(args) {
        const context = getExcelContext(deps, 'excel_validation')
        if (!context.ok) return context.result
        const range = args.range || ''
        if (!range) return fail('excel_validation', 'Missing validation range.')
        let values: string[] = []
        if (args.values) {
          try {
            values = JSON.parse(args.values)
          } catch {
            values = args.values.split(',').map((value) => value.trim())
          }
        }
        try {
          const result = await deps.adapter.setValidation(
            context.excelFilePath,
            args.sheet || 'Sheet1',
            {
              range,
              type: (args.type || 'list') as 'list' | 'whole' | 'decimal',
              values,
              min: args.min ? parseFloat(args.min) : undefined,
              max: args.max ? parseFloat(args.max) : undefined,
              remove: (args.action || 'set').toLowerCase() === 'remove',
            },
          )
          if (result.success) {
            await deps.adapter.refreshWorkbookPreview()
            return ok('excel_validation', result.message || 'Validation updated.')
          }
          return fail('excel_validation', result.error || 'Validation failed')
        } catch (error) {
          return fail('excel_validation', `Validation failed: ${error}`)
        }
      },
    }),
    defineAgentTool({
      id: 'excel_hyperlink',
      displayName: 'Excel Hyperlink',
      description: 'Set hyperlinks in the active Excel workbook',
      domain: 'excel',
      mutation: 'write',
      concurrency: 'serial',
      async handler(args) {
        const context = getExcelContext(deps, 'excel_hyperlink')
        if (!context.ok) return context.result
        const cell = args.cell || ''
        if (!cell) return fail('excel_hyperlink', 'Missing cell address.')
        try {
          const result = await deps.adapter.setHyperlink(
            context.excelFilePath,
            args.sheet || 'Sheet1',
            {
              cell,
              url: args.url,
              text: args.text || args.url,
              tooltip: args.tooltip,
              remove: (args.action || 'set').toLowerCase() === 'remove',
            },
          )
          if (result.success) {
            await deps.adapter.refreshWorkbookPreview()
            return ok('excel_hyperlink', result.message || 'Hyperlink updated.')
          }
          return fail('excel_hyperlink', result.error || 'Hyperlink failed')
        } catch (error) {
          return fail('excel_hyperlink', `Hyperlink failed: ${error}`)
        }
      },
    }),
    defineAgentTool({
      id: 'excel_find_replace',
      displayName: 'Excel Find Replace',
      description: 'Find and replace text in the active Excel workbook',
      domain: 'excel',
      mutation: 'write',
      concurrency: 'serial',
      async handler(args) {
        const context = getExcelContext(deps, 'excel_find_replace')
        if (!context.ok) return context.result
        const find = args.find || ''
        if (!find) return fail('excel_find_replace', 'Missing find text.')
        try {
          const result = await deps.adapter.findReplace(
            context.excelFilePath,
            args.sheet || 'Sheet1',
            {
              find,
              replace: args.replace || '',
              matchCase: args.matchCase === 'true',
              matchWholeCell: args.matchWholeCell === 'true',
              allSheets: args.allSheets === 'true',
            },
          )
          if (result.success) {
            await deps.adapter.refreshWorkbookPreview()
            return ok('excel_find_replace', result.message || `Replaced ${result.count || 0} match(es).`, {
              count: result.count,
              details: result.details,
            })
          }
          return fail('excel_find_replace', result.error || 'Find/replace failed')
        } catch (error) {
          return fail('excel_find_replace', `Find/replace failed: ${error}`)
        }
      },
    }),
    defineAgentTool({
      id: 'excel_chart',
      displayName: 'Excel Chart',
      description: 'Insert a chart into the active Excel workbook',
      domain: 'excel',
      mutation: 'create',
      concurrency: 'serial',
      async handler(args) {
        const context = getExcelContext(deps, 'excel_chart')
        if (!context.ok) return context.result
        const dataRange = args.dataRange || ''
        if (!dataRange) return fail('excel_chart', 'Missing chart data range.')
        try {
          const result = await deps.adapter.insertChart(
            context.excelFilePath,
            args.sheet || 'Sheet1',
            {
              type: (args.type || 'column') as 'column' | 'bar' | 'line' | 'pie',
              dataRange,
              title: args.title || '',
              position: args.position || 'E1',
              width: args.width ? parseInt(args.width, 10) : 500,
              height: args.height ? parseInt(args.height, 10) : 300,
            },
          )
          if (result.success) {
            await deps.adapter.refreshWorkbookPreview()
            return ok('excel_chart', result.message || 'Chart inserted.', {
              chartConfig: result.chartConfig,
            })
          }
          return fail('excel_chart', result.error || 'Chart insertion failed')
        } catch (error) {
          return fail('excel_chart', `Chart insertion failed: ${error}`)
        }
      },
    }),
  ]
}
