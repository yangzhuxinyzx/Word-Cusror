import JSZip from 'jszip'

export async function summarizePptxBase64(params: {
  base64: string
  maxSlides: number
  maxChars: number
  truncateWithNote: (text: string, maxLen: number, note: string) => string
}): Promise<string> {
  const zip = await JSZip.loadAsync(params.base64, { base64: true })
  const slidePaths = Object.keys(zip.files)
    .filter((name) => name.startsWith('ppt/slides/slide') && name.endsWith('.xml'))
    .sort((left, right) => {
      const getNum = (value: string) =>
        parseInt(value.match(/slide(\d+)\.xml/)?.[1] || '0', 10)
      return getNum(left) - getNum(right)
    })

  const decodeXml = (input: string) =>
    input
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&#39;/g, "'")

  const lines: string[] = [`页数: ${slidePaths.length}`]

  for (let i = 0; i < Math.min(params.maxSlides, slidePaths.length); i += 1) {
    const slideXml = await zip.file(slidePaths[i])!.async('string')
    const texts = Array.from(slideXml.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)).map(
      (match) => decodeXml(match[1]),
    )
    const combined = texts.join(' ').replace(/\s+/g, ' ').trim()
    if (combined) {
      lines.push(`- 第 ${i + 1} 页：${combined.slice(0, 200)}`)
    }
  }

  return params.truncateWithNote(lines.join('\n'), params.maxChars, 'PPT 摘要')
}
