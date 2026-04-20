export interface CollapseContextItem {
  id: string
  priority: number
  content: string
}

export function collapseContextToBudget(
  items: readonly CollapseContextItem[],
  budgetChars: number,
): string[] {
  const sorted = [...items].sort((left, right) => right.priority - left.priority)
  const kept: string[] = []
  let used = 0

  for (const item of sorted) {
    if (!item.content.trim()) continue
    if (used + item.content.length > budgetChars) continue
    kept.push(item.content)
    used += item.content.length
  }

  return kept
}
