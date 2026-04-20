# Word Render Lab

## 目标

为 `Word-Cursor` 提供一条稳定的“打开 DOCX -> 等待渲染完成 -> 自动截图 -> 可选参考图 diff”闭环，方便持续缩小和原生 Word / WPS 的渲染差距。

这套能力优先服务前端排版调优，因此先落地为仓库内置的 Playwright 自动化脚本；页面本身同时做成了 MCP 友好的浏览器实验页，后续可直接接官方 Playwright MCP。

## 浏览器实验页

启动开发服务器后，打开：

```text
http://127.0.0.1:3000/?render-lab=1&fixture=render-lab-sample-cn.docx
```

默认会从 [public/render-lab/fixtures](/Users/mac/Desktop/Word-Cursor/public/render-lab/fixtures) 读取 DOCX fixture，并进入最小化的 Word 渲染实验页。

页面会暴露这些稳定入口，方便 Playwright / MCP 使用：

- `window.__wordCursorRenderLab`
- `window.__wordCursorLayoutDebug`
- `window.__wordCursorWordEditorDebug`

以及这些稳定选择器：

- `[data-testid="render-lab-root"]`
- `[data-testid="word-render-canvas-preview"]`
- `[data-testid="word-render-page-1"] canvas`

## 安装

首次使用需要安装 Playwright 浏览器：

```bash
npm run render:lab:install
```

## 自动截图

直接截图第一页：

```bash
npm run render:lab:capture -- --fixture render-lab-sample-cn.docx --page 1
```

脚本会：

1. 尝试复用现有 `127.0.0.1:3000` 开发服务器
2. 若未启动，则自动拉起本地 Vite
3. 打开 Render Lab
4. 等待文档解析和分页布局完成
5. 截图并把结果写入 [outputs/render-lab](/Users/mac/Desktop/Word-Cursor/outputs/render-lab)

可用参数：

- `--fixture xxx.docx`
- `--page 1`
- `--page all`
- `--output /absolute/path/to/output.png`
- `--reference /absolute/path/to/reference.png`
- `--url http://127.0.0.1:3000`
- `--no-server`

## 参考图 Diff

如果你已经有一张来自 Word / WPS 的参考截图，可以直接生成 diff：

```bash
npm run render:lab:capture -- \
  --fixture render-lab-sample-cn.docx \
  --page 1 \
  --reference /absolute/path/to/word-page-1.png
```

输出会包含：

- 当前截图 `.png`
- 元数据 `.json`
- 差异图 `.diff.png`

差异图中红色区域表示与参考图不一致的位置。

## 与 MCP 的衔接

当前项目内还没有通用 MCP client，所以第一阶段先把页面、状态和脚本做稳定，让我们能马上迭代渲染。

如果后续要让外部智能体直接操作浏览器，推荐接官方 Playwright MCP：

```json
{
  "mcpServers": {
    "playwright": {
      "command": "npx",
      "args": ["@playwright/mcp@latest"]
    }
  }
}
```

接入后，MCP 侧可以直接打开上面的 Render Lab URL，并依赖已经暴露好的状态与选择器做截图、观察和回归。
