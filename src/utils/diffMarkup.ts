export const INLINE_DIFF_PAIR_CLASS = 'diff-change-pair'

export function buildInlineDiffTokenHtml(
  kind: 'old' | 'new',
  diffId: string,
  html: string,
  extraAttributes = '',
): string {
  return `<span class="diff-${kind}" data-diff-id="${diffId}"${extraAttributes}>${html}</span>`
}

export function buildInlineDiffPairHtml(
  diffId: string,
  oldHtml: string,
  newHtml: string,
): string {
  return `<span class="${INLINE_DIFF_PAIR_CLASS}" data-diff-id="${diffId}">${buildInlineDiffTokenHtml('old', diffId, oldHtml)}${buildInlineDiffTokenHtml('new', diffId, newHtml)}</span>`
}
