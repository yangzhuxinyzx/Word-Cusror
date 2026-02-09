import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App'
import './index.css'
import './fonts/onlyoffice-fonts.css'
import { initTheme } from './utils/theme'
import { loadOnlyOfficeBundledFonts } from './fonts/loadBundledFonts'
import { initWorkspaceFonts } from './fonts/loadWorkspaceFonts'

initTheme()
// 非阻塞：尝试注入打包字体（若 manifest 为空则无副作用）
void loadOnlyOfficeBundledFonts()
// 非阻塞：Electron 环境下从 Fonts/ 目录加载字体
void initWorkspaceFonts()

ReactDOM.createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
)

