# ONLYOFFICE 字体目录（前端打包用）

把你要打包进前端（Electron/网页）使用的字体文件放到这里，例如：

- `*.woff2`（推荐，体积小、加载快）
- `*.woff`
- `*.ttf` / `*.otf`（可用但体积更大）

同时维护 `manifest.json`（或在 `src/fonts/fontManifest.ts` 里维护清单）。应用启动时会自动注入对应的 `@font-face`，并在 `WordEditor` 字体选择器中展示这些字体。













