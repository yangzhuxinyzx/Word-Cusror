import mammoth from 'mammoth'
import { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType } from 'docx'

export interface ParsedDocument {
  html: string
  markdown: string
  rawText: string
}

// 将 .docx 文件解析为 HTML 和 Markdown
export async function parseDocxFile(file: File): Promise<ParsedDocument> {
  const arrayBuffer = await file.arrayBuffer()
  
  // 使用 mammoth 解析为 HTML
  const htmlResult = await mammoth.convertToHtml({ arrayBuffer })
  const html = htmlResult.value
  
  // 转换为 Markdown
  const markdown = htmlToMarkdown(html)
  
  // 提取纯文本
  const textResult = await mammoth.extractRawText({ arrayBuffer })
  const rawText = textResult.value
  
  return { html, markdown, rawText }
}

// 简单的 HTML 转 Markdown
function htmlToMarkdown(html: string): string {
  let md = html
  
  // 标题
  md = md.replace(/<h1[^>]*>(.*?)<\/h1>/gi, '# $1\n\n')
  md = md.replace(/<h2[^>]*>(.*?)<\/h2>/gi, '## $1\n\n')
  md = md.replace(/<h3[^>]*>(.*?)<\/h3>/gi, '### $1\n\n')
  md = md.replace(/<h4[^>]*>(.*?)<\/h4>/gi, '#### $1\n\n')
  
  // 段落
  md = md.replace(/<p[^>]*>(.*?)<\/p>/gi, '$1\n\n')
  
  // 粗体和斜体
  md = md.replace(/<strong[^>]*>(.*?)<\/strong>/gi, '**$1**')
  md = md.replace(/<b[^>]*>(.*?)<\/b>/gi, '**$1**')
  md = md.replace(/<em[^>]*>(.*?)<\/em>/gi, '*$1*')
  md = md.replace(/<i[^>]*>(.*?)<\/i>/gi, '*$1*')
  
  // 列表
  md = md.replace(/<ul[^>]*>(.*?)<\/ul>/gis, (_, content) => {
    return content.replace(/<li[^>]*>(.*?)<\/li>/gi, '- $1\n')
  })
  md = md.replace(/<ol[^>]*>(.*?)<\/ol>/gis, (_, content) => {
    let index = 0
    return content.replace(/<li[^>]*>(.*?)<\/li>/gi, () => {
      index++
      return `${index}. $1\n`
    })
  })
  
  // 链接
  md = md.replace(/<a[^>]*href="([^"]*)"[^>]*>(.*?)<\/a>/gi, '[$2]($1)')
  
  // 换行
  md = md.replace(/<br\s*\/?>/gi, '\n')
  
  // 移除其他 HTML 标签
  md = md.replace(/<[^>]+>/g, '')
  
  // 清理多余空行
  md = md.replace(/\n{3,}/g, '\n\n')
  
  // 解码 HTML 实体
  md = md.replace(/&nbsp;/g, ' ')
  md = md.replace(/&amp;/g, '&')
  md = md.replace(/&lt;/g, '<')
  md = md.replace(/&gt;/g, '>')
  md = md.replace(/&quot;/g, '"')
  
  return md.trim()
}

// 将 Markdown 转换为 docx 文档
export async function markdownToDocx(markdown: string, title: string): Promise<Blob> {
  const lines = markdown.split('\n')
  const children: Paragraph[] = []
  
  let inList = false
  let listItems: string[] = []

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i]
    
    // 标题
    if (line.startsWith('# ')) {
      children.push(new Paragraph({
        text: line.slice(2),
        heading: HeadingLevel.HEADING_1,
        spacing: { before: 400, after: 200 },
      }))
    } else if (line.startsWith('## ')) {
      children.push(new Paragraph({
        text: line.slice(3),
        heading: HeadingLevel.HEADING_2,
        spacing: { before: 300, after: 150 },
      }))
    } else if (line.startsWith('### ')) {
      children.push(new Paragraph({
        text: line.slice(4),
        heading: HeadingLevel.HEADING_3,
        spacing: { before: 240, after: 120 },
      }))
    }
    // 无序列表
    else if (line.startsWith('- ') || line.startsWith('* ')) {
      children.push(new Paragraph({
        children: [new TextRun(line.slice(2))],
        bullet: { level: 0 },
        spacing: { before: 60, after: 60 },
      }))
    }
    // 有序列表
    else if (/^\d+\.\s/.test(line)) {
      const text = line.replace(/^\d+\.\s/, '')
      children.push(new Paragraph({
        children: [new TextRun(text)],
        numbering: { reference: 'default-numbering', level: 0 },
        spacing: { before: 60, after: 60 },
      }))
    }
    // 普通段落
    else if (line.trim()) {
      // 解析粗体和斜体
      const textRuns = parseInlineFormatting(line)
      children.push(new Paragraph({
        children: textRuns,
        spacing: { before: 120, after: 120 },
      }))
    }
    // 空行
    else {
      children.push(new Paragraph({}))
    }
  }

  const doc = new Document({
    sections: [{
      properties: {
        page: {
          margin: {
            top: 1440,    // 1 inch = 1440 twips
            right: 1440,
            bottom: 1440,
            left: 1440,
          },
        },
      },
      children,
    }],
    numbering: {
      config: [{
        reference: 'default-numbering',
        levels: [{
          level: 0,
          format: 'decimal',
          text: '%1.',
          alignment: AlignmentType.START,
        }],
      }],
    },
  })

  return await Packer.toBlob(doc)
}

// 解析行内格式（粗体、斜体）
function parseInlineFormatting(text: string): TextRun[] {
  const runs: TextRun[] = []
  let remaining = text
  
  // 简化处理：先处理粗体，再处理斜体
  const regex = /(\*\*(.+?)\*\*|\*(.+?)\*|([^*]+))/g
  let match
  
  while ((match = regex.exec(text)) !== null) {
    if (match[2]) {
      // 粗体 **text**
      runs.push(new TextRun({ text: match[2], bold: true }))
    } else if (match[3]) {
      // 斜体 *text*
      runs.push(new TextRun({ text: match[3], italics: true }))
    } else if (match[4]) {
      // 普通文本
      runs.push(new TextRun({ text: match[4] }))
    }
  }
  
  if (runs.length === 0) {
    runs.push(new TextRun({ text }))
  }
  
  return runs
}

// 读取拖放的文件
export function handleFileDrop(e: DragEvent): File | null {
  e.preventDefault()
  const files = e.dataTransfer?.files
  if (files && files.length > 0) {
    const file = files[0]
    if (file.name.endsWith('.docx') || file.name.endsWith('.doc')) {
      return file
    }
  }
  return null
}

