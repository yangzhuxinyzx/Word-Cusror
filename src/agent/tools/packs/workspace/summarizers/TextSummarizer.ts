export function summarizeTextContent(
  content: string,
  maxChars: number,
  truncateWithNote: (text: string, maxLen: number, note: string) => string,
  fileLabel: string,
): string {
  return truncateWithNote(content || '', maxChars, `${fileLabel} 摘要`)
}
