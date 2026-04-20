import type { AISettings } from '../../../types'
import type { PptDomainAdapter, PptEditContextSnapshot } from '../../adapters/ppt/PptDomainAdapter'
import { defineAgentTool } from '../contracts'
import type { ExecutableAgentTool } from '../executor'

export interface PptToolPackDeps {
  adapter: PptDomainAdapter
  settings: AISettings
  pendingImages: string[]
  pptEditContext: PptEditContextSnapshot | null
  registerToolActivity: (tool: string, label: string) => string
  completeToolActivity: (
    activityId: string,
    status: 'success' | 'error' | 'skipped',
    detail?: string,
  ) => void
  updateAgentAction: (action: string) => void
  completeAgentStep: () => void
  updateAgentFile: (
    updates: Partial<{
      name: string
      additions: number
      deletions: number
      status: 'pending' | 'writing' | 'done'
    }>,
  ) => void
  addAgentFileOperation: (operation: string) => void
  finishAgentProgress: () => void
  truncateLabel: (text: string, limit?: number) => string
}

export function createPptToolPack(
  deps: PptToolPackDeps,
): ExecutableAgentTool[] {
  return [
    defineAgentTool({
      id: 'ppt_create',
      displayName: 'PPT Create',
      description: 'Generate a new PPTX deck',
      domain: 'ppt',
      mutation: 'create',
      concurrency: 'serial',
      async handler(args) {
        const title = args.title || '新建演示文稿'
        const theme = args.theme || ''
        const style = args.style || ''
        const outline = args.outline || ''
        const activityId = deps.registerToolActivity(
          'ppt_create',
          `PPT：${deps.truncateLabel(title, 24)}`,
        )

        if (!deps.adapter.canUsePptTools()) {
          deps.completeToolActivity(activityId, 'error', '仅桌面版支持')
          return {
            tool: 'ppt_create',
            success: false,
            message: 'PPT 生成仅支持桌面版（Electron）',
          }
        }

        if (!outline || outline.trim().length < 10) {
          deps.completeToolActivity(activityId, 'error', '缺少大纲')
          return {
            tool: 'ppt_create',
            success: false,
            message: '缺少 outline 参数，需要 PPT 大纲内容',
          }
        }
        const output = deps.adapter.buildOutputPath(title)
        if (!output) {
          deps.completeToolActivity(activityId, 'error', 'missing workspace')
          return {
            tool: 'ppt_create',
            success: false,
            message: 'Missing workspace path. Please open a file or workspace first.',
          }
        }

        const { fileName: pptxName, outputPath } = output
        const slideCountMatch = outline.match(/第\s*(\d+)\s*页/g)
        const estimatedSlideCount = slideCountMatch
          ? slideCountMatch.length
          : 3

        try {
          deps.updateAgentAction('正在生成 PPT 视觉设计...')
          deps.completeAgentStep()
          deps.updateAgentFile({ status: 'writing', name: pptxName })
          deps.addAgentFileOperation(
            `PPT: preparing ${estimatedSlideCount} slide prompts`,
          )
          const geminiResult = await deps.adapter.generatePromptSlides({
            title,
            theme,
            style,
            outline,
            outputPath,
            pendingImages: deps.pendingImages,
            settings: deps.settings,
          })

          if (!geminiResult.success || !geminiResult.slides) {
            deps.completeToolActivity(activityId, 'error', '提示词生成失败')
            return {
              tool: 'ppt_create',
              success: false,
              message: `设计提示词生成失败: ${geminiResult.error || '未知错误'}`,
            }
          }

          const deckDesignConcept = geminiResult.designConcept || ''
          const deckColorPalette = geminiResult.colorPalette || ''
          const slides = geminiResult.slides.map((slide) => ({
            prompt: slide.prompt,
            negativePrompt: slide.negativePrompt,
          }))

          deps.updateAgentAction(`正在生成 ${slides.length} 页 PPT 图像...`)
          deps.addAgentFileOperation(`PPT: generating ${slides.length} slide images`)

          const negativeDefault =
            'watermark, logo, brand name text, badge, QR code, UI, screenshot, HUD, sci-fi interface, holographic UI, futuristic dashboard, neon cyberpunk, neon cyan, bright cyan, fluorescent cyan, neon teal, cheap turquoise, generic isometric city, isometric cityscape, circuit-board city, lowres, blurry, garbled Chinese, wrong characters, text distortion, misspelling, random letters, gibberish, extra text, english text, ugly typography, amateur layout, noisy background, oversaturated, cheap plastic, toy-like, glossy, harsh specular, overbloom, stock 3d icons, generic template, ai artifacts, uncanny'
          const result = await deps.adapter.generateDeck({
            outputPath,
            slides: slides.map((slide) => ({
              prompt: slide.prompt,
              negativePrompt: slide.negativePrompt || negativeDefault,
            })),
            designConcept: deckDesignConcept,
            colorPalette: deckColorPalette,
            settings: deps.settings,
          })

          if (!result.success || !result.path) {
            deps.completeToolActivity(activityId, 'error', result.error || '导出失败')
            return {
              tool: 'ppt_create',
              success: false,
              message: `PPT 生成失败: ${result.error || '未知错误'}`,
            }
          }

          await deps.adapter.refreshFiles()
          await deps.adapter.openGeneratedDeck(pptxName, result.path)

          deps.updateAgentFile({
            additions: slides.length,
            status: 'done',
            name: pptxName,
          })
          deps.finishAgentProgress()
          deps.completeToolActivity(activityId, 'success', `${slides.length} 页`)

          return {
            tool: 'ppt_create',
            success: true,
            message: `已生成 PPT：${pptxName}，共 ${slides.length} 页，并已打开`,
            data: {
              fileName: pptxName,
              path: result.path,
              slideCount: slides.length,
            },
          }
        } catch (e) {
          deps.completeToolActivity(activityId, 'error', '异常')
          return {
            tool: 'ppt_create',
            success: false,
            message: `PPT 生成失败: ${e}`,
          }
        }
      },
    }),
    defineAgentTool({
      id: 'ppt_edit',
      displayName: 'PPT Edit',
      description: 'Edit an existing PPT deck or selected slide region',
      domain: 'ppt',
      mutation: 'write',
      concurrency: 'serial',
      async handler(args) {
        const pageNumber = Number(args.pageNumber) || 1
        const mode =
          args.mode === 'partial_edit' ? 'partial_edit' : 'regenerate'
        const feedback = args.feedback || ''
        const pptxPath = args.pptxPath || deps.pptEditContext?.pptxPath || ''

        let regionRect:
          | { x: number; y: number; w: number; h: number }
          | undefined = deps.pptEditContext?.regionRect
        if (typeof args.regionRect === 'string' && args.regionRect.trim()) {
          try {
            regionRect = JSON.parse(args.regionRect)
          } catch {
            // ignore malformed region rect
          }
        }
        const regionScreenshot =
          typeof args.regionScreenshot === 'string' &&
          args.regionScreenshot.trim()
            ? args.regionScreenshot
            : deps.pptEditContext?.imageBase64

        const modeLabel =
          mode === 'regenerate' ? '整页重做' : '局部编辑'
        const activityId = deps.registerToolActivity(
          'ppt_edit',
          `PPT ${modeLabel}：第 ${pageNumber} 页`,
        )

        if (!deps.adapter.canUsePptTools()) {
          deps.completeToolActivity(activityId, 'error', '仅桌面版支持')
          return {
            tool: 'ppt_edit',
            success: false,
            message: 'PPT 编辑仅支持桌面版（Electron）',
          }
        }

        if (!pptxPath) {
          deps.completeToolActivity(activityId, 'error', '缺少路径')
          return {
            tool: 'ppt_edit',
            success: false,
            message: '缺少 PPTX 文件路径',
          }
        }

        try {
          deps.updateAgentAction(`正在${modeLabel}第 ${pageNumber} 页...`)
          const result = await deps.adapter.editSlides({
            pptxPath,
            pageNumber,
            mode,
            feedback,
            regionRect,
            regionScreenshot,
            settings: deps.settings,
          })

          if (!result.success) {
            deps.completeToolActivity(activityId, 'error', result.error || '失败')
            return {
              tool: 'ppt_edit',
              success: false,
              message: `PPT 编辑失败: ${result.error || '未知错误'}`,
            }
          }
          await deps.adapter.refreshFiles()
          await deps.adapter.reopenCurrentDeckIfNeeded(pptxPath)

          deps.completeToolActivity(activityId, 'success', modeLabel)
          return {
            tool: 'ppt_edit',
            success: true,
            message: `已完成第 ${pageNumber} 页的${modeLabel}`,
            data: {
              pageNumber,
              mode,
              fileName: pptxPath.split(/[\\/]/).pop() || '',
              pptxPath,
            },
          }
        } catch (e) {
          deps.completeToolActivity(activityId, 'error', '异常')
          return {
            tool: 'ppt_edit',
            success: false,
            message: `PPT 编辑失败: ${e}`,
          }
        }
      },
    }),
    defineAgentTool({
      id: 'ppt_text_edit',
      displayName: 'PPT Text Edit',
      description: 'Detect text boxes on an image-only PPT slide and apply text replacements in place',
      domain: 'ppt',
      mutation: 'write',
      concurrency: 'serial',
      async handler(args) {
        const pageNumber = Number(args.pageNumber) || 1
        const pptxPath = args.pptxPath || deps.pptEditContext?.pptxPath || ''
        const rawEdits = args.edits
        const activityId = deps.registerToolActivity(
          'ppt_text_edit',
          `PPT 改字：第 ${pageNumber} 页`,
        )

        if (!deps.adapter.canUsePptTools()) {
          deps.completeToolActivity(activityId, 'error', '仅桌面版支持')
          return {
            tool: 'ppt_text_edit',
            success: false,
            message: 'PPT 图片页改字仅支持桌面版（Electron）',
          }
        }

        if (!pptxPath) {
          deps.completeToolActivity(activityId, 'error', '缺少路径')
          return {
            tool: 'ppt_text_edit',
            success: false,
            message: '缺少 PPTX 文件路径',
          }
        }

        let edits = []
        if (Array.isArray(rawEdits)) {
          edits = rawEdits
        } else if (typeof rawEdits === 'string' && rawEdits.trim()) {
          try {
            edits = JSON.parse(rawEdits)
          } catch {
            edits = []
          }
        }

        if (!Array.isArray(edits) || edits.length === 0) {
          deps.completeToolActivity(activityId, 'error', '缺少 edits')
          return {
            tool: 'ppt_text_edit',
            success: false,
            message: '缺少 edits 参数，至少需要一条 boxId/fromText/toText 变更',
          }
        }

        try {
          deps.updateAgentAction(`正在识别并改写第 ${pageNumber} 页文字...`)
          const result = await deps.adapter.applyTextEdits({
            pptxPath,
            pageNumber,
            edits,
          })
          if (!result.success) {
            deps.completeToolActivity(activityId, 'error', result.error || '失败')
            return {
              tool: 'ppt_text_edit',
              success: false,
              message: result.error || 'PPT 改字失败',
              data: {
                fallbackSuggested: result.fallbackSuggested === true,
              },
            }
          }

          await deps.adapter.refreshFiles()
          await deps.adapter.reopenCurrentDeckIfNeeded(pptxPath)
          deps.completeToolActivity(activityId, 'success', `${edits.length} 处`)

          return {
            tool: 'ppt_text_edit',
            success: true,
            message: `已完成第 ${pageNumber} 页 ${edits.length} 处文字替换`,
            data: {
              pageNumber,
              editsApplied: edits.length,
              imageDataUrl: result.imageDataUrl,
              pptxPath,
            },
          }
        } catch (e) {
          deps.completeToolActivity(activityId, 'error', '异常')
          return {
            tool: 'ppt_text_edit',
            success: false,
            message: `PPT 改字失败: ${e}`,
          }
        }
      },
    }),
  ]
}
