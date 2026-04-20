export type ThemeMode = 'system' | 'light' | 'dark'

const STORAGE_KEY = 'word-cursor-theme-mode'

function getSystemTheme(): 'light' | 'dark' {
  if (typeof window === 'undefined') return 'dark'
  return window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light'
}

export function getThemeMode(): ThemeMode {
  try {
    const raw = localStorage.getItem(STORAGE_KEY)
    if (raw === 'light' || raw === 'dark' || raw === 'system') return raw
  } catch {
    // ignore
  }
  return 'dark'
}

export function applyTheme(mode: ThemeMode) {
  const resolved = mode === 'system' ? getSystemTheme() : mode
  document.documentElement.dataset.theme = resolved
}

export function setThemeMode(mode: ThemeMode) {
  try {
    localStorage.setItem(STORAGE_KEY, mode)
  } catch {
    // ignore
  }
  applyTheme(mode)
}

export function initTheme() {
  const mode = getThemeMode()
  applyTheme(mode)

  // 跟随系统：监听系统主题变化
  if (mode === 'system' && window.matchMedia) {
    const mql = window.matchMedia('(prefers-color-scheme: dark)')
    const onChange = () => applyTheme('system')
    // Safari/old Chromium 兼容
    if (typeof mql.addEventListener === 'function') mql.addEventListener('change', onChange)
    else (mql as any).addListener?.(onChange)
  }
}


