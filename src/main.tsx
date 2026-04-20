import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App'
import './index.css'
import './fonts/onlyoffice-fonts.css'
import { initTheme } from './utils/theme'
import { loadOnlyOfficeBundledFonts } from './fonts/loadBundledFonts'
import { initWorkspaceFonts } from './fonts/loadWorkspaceFonts'

type WordCursorWindow = Window & typeof globalThis & {
  __wordCursorReactRoot?: ReactDOM.Root
}

initTheme()
// 非阻塞：尝试注入打包字体（若 manifest 为空则无副作用）
void loadOnlyOfficeBundledFonts()
// 非阻塞：Electron 环境下从 Fonts/ 目录加载字体
void initWorkspaceFonts()

const appWindow = window as WordCursorWindow

function renderApp() {
  const container = document.getElementById('root')
  if (!container) return

  if (!appWindow.__wordCursorReactRoot) {
    appWindow.__wordCursorReactRoot = ReactDOM.createRoot(container)
  }

  appWindow.__wordCursorReactRoot.render(
    <React.StrictMode>
      <App />
    </React.StrictMode>,
  )
}

function rootLooksBlank(): boolean {
  const container = document.getElementById('root')
  if (!container) return false
  return !container.firstElementChild && (container.textContent || '').trim().length === 0
}

function scheduleMountHealthCheck(delayMs: number) {
  window.setTimeout(() => {
    if (!rootLooksBlank()) return
    console.warn('[bootstrap] root rendered blank, attempting remount')
    renderApp()
  }, delayMs)
}

renderApp()
scheduleMountHealthCheck(300)
scheduleMountHealthCheck(1200)
scheduleMountHealthCheck(2500)

