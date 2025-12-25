<p align="center">
  <img src="public/favicon.svg" width="84" height="84" alt="Word-Cursor Logo">
</p>

<h1 align="center">Word-Cursor</h1>

<p align="center">
  <strong>AI 驱动的智能办公文档编辑器</strong><br>
  把 Cursor 级别的"对话式编辑 + 工具调用 + 可审阅变更"带进 Word / Excel / PowerPoint
</p>

<p align="center">
  📮 <strong>如遇部署问题，欢迎咨询 QQ：2935076541</strong>
</p>

<p align="center">
  <img src="https://img.shields.io/badge/版本-v1.0.0-blue" alt="Version">
  <img src="https://img.shields.io/badge/React-18-61dafb?logo=react" alt="React">
  <img src="https://img.shields.io/badge/TypeScript-5-3178c6?logo=typescript" alt="TypeScript">
  <img src="https://img.shields.io/badge/Electron-33-47848f?logo=electron" alt="Electron">
  <img src="https://img.shields.io/badge/Vite-5-646cff?logo=vite" alt="Vite">
  <img src="https://img.shields.io/badge/License-MIT-green" alt="License">
</p>

<p align="center">
  <a href="#-为什么-word-cursor-是革新性的">革新性</a> •
  <a href="#-你可以用它做什么">应用场景</a> •
  <a href="#-功能全景">功能全景</a> •
  <a href="#-部署运行">部署运行</a> •
  <a href="#-配置说明">配置说明</a> •
  <a href="#-使用指南">使用指南</a> •
  <a href="#-仓库结构">仓库结构</a> •
  <a href="#-faq">FAQ</a>
</p>

---

## 🌟 为什么 Word-Cursor 是革新性的

传统办公软件的核心范式是“**人手工操作 UI**”；而 **Word-Cursor** 把范式升级为“**人描述意图 → AI 通过工具执行 → 结果可审阅、可回滚**”：

- **从“编辑器”变成“工作流执行器”**：你不再只是改字、调格式，而是用自然语言驱动一整段业务动作（生成报告、统一公文格式、批量改表、做 PPT、联网调研并引用）。
- **把 AI 变成可控的“协作者”而不是黑箱**：AI 的修改以可视化差异（diff）呈现，支持接受/拒绝，降低“AI 乱改”的风险。
- **把“上下文”从聊天窗口拉回到真实文件系统**：以工作区为中心，AI 可以理解你打开的文件夹、文件树、当前文档内容与结构。
- **把搜索与生成能力内置到桌面端**：联网搜索、PPT 生图/编辑等能力在 Electron 桌面应用中直接可用，形成端到端生产力闭环。

一句话：**Word-Cursor = Cursor 的"可审阅 AI 工作流" + Office 的"文档生产场景"**。

---

## 📸 项目截图

### 主界面
![主界面](./assets/主界面1.png)

### Word 文档编辑
![Word界面](./assets/word界面.png)

### Excel 表格预览
![Excel界面](./assets/excel界面.png)

### PPT 演示文稿
![PPT界面](./assets/ppt界面.png)

### PPT 作品展示

<p align="center">
  <img src="./assets/ppt作品展示1.png" width="80%" alt="PPT作品展示1">
</p>

<p align="center">
  <img src="./assets/ppt作品展示2.png" width="80%" alt="PPT作品展示2">
</p>

<p align="center">
  <img src="./assets/ppt作品展示3.png" width="80%" alt="PPT作品展示3">
</p>

---

## 🎯 你可以用它做什么

- **写作/公文/报告**：生成初稿 → 统一格式 → 分章节润色 → 输出可交付的 `.docx`
- **会议纪要/周报**：从要点快速生成结构化文档，并按“可审阅修改”逐条确认
- **表格理解与基础处理**：打开 `.xlsx/.xls` 预览工作表，做结构化分析与基础操作
- **PPT 端到端生成与再编辑**：从大纲生成整套 PPTX；不满意就“整页重做”或“局部编辑”
- **联网调研**：在对话中直接搜索并整理要点，把证据链写回文档

---

## ✨ 功能全景

### 系统架构（从“聊天”到“改文件”）

```text
┌───────────────────────────┐
│        React 渲染进程       │
│  - Word/Excel/PPT 预览 UI   │
│  - ChatPanel 对话与指令编排  │
│  - diff 审阅（接受/拒绝）    │
└───────────────┬───────────┘
                │ Electron IPC（window.electronAPI.*）
┌───────────────▼───────────┐
│        Electron 主进程      │
│  - 本地文件系统/工作区       │
│  - Brave Search MCP（内置）  │
│  - PPT：Gemini 提示词→生图→PPTX │
│  - 图像后处理（sharp）       │
└───────────────┬───────────┘
                │ HTTP
┌───────────────▼───────────┐
│ 外部能力（按需）            │
│ - 主模型（OpenAI 兼容 API）  │
│ - OpenRouter（Gemini）       │
│ - DashScope（生图/编辑）     │
│ - ONLYOFFICE Document Server │
└───────────────────────────┘
```

### 能力矩阵（你能得到什么）

| 模块 | 核心能力 | 是否需要桌面端（Electron） |
|---|---|---|
| **Word 文档** | 富文本编辑、打印布局、AI 选区快改、可审阅 diff（接受/拒绝） | 否（但打开本地工作区更完整） |
| **Excel 表格** | 工作表预览、基础读取与分析、与 AI 联动处理 | 是（本地文件读写） |
| **PPT 演示** | 端到端生成（提示词→生图→PPTX）、预览、整页重做、局部编辑 | 是（PPT 生成/编辑在主进程执行） |
| **联网搜索** | 内置 Brave Search MCP，结构化返回 Web/News/FAQ/Videos/Discussions | 是（默认）；Web 模式需后备接口 |
| **本地补全** | `Tab` 低延迟续写（本地模型服务） | 否（但桌面端体验更好） |
| **ONLYOFFICE** | 真编辑器模式（高兼容性/复杂排版） | 是（且需本地启动 Document Server） |

### 1) 文档编辑（Word 体验）

- **富文本编辑**：
  - 标题（1/2/3 级）、段落、列表、引用
  - 字体/字号/颜色/高亮、粗体/斜体/下划线/删除线、上下标
  - 表格插入与编辑、图片插入、超链接
- **打印布局（A4）预览**：
  - 多页白纸背景、页边距留白、页间距
  - 缩放：`Ctrl + 滚轮` / `Ctrl + 0/+/-`
- **AI 选区快改**：
  - 选中一段文字，右键或 `Ctrl + K` 输入指令
  - 支持“尽量保留原格式”的输出（对 HTML 选区做格式保留）
- **工作区与文件**（桌面端）：
  - 打开本地文件夹作为工作区，浏览文件树、打开/保存文档
  - 支持拖拽文件到对话区用于补充上下文（用于“文件夹理解/资料整合”）

### 2) “可审阅”的 AI 修改（核心体验）

- **diff 标记**：AI 的“删改增”会高亮显示
- **接受 / 拒绝**：一键应用或撤销 AI 建议（降低不可控风险）
- **修改面板**：集中查看与定位修改（适合长文审阅）

### 3) AI 智能补全（可选）

- `Tab` 触发 **本地模型**低延迟续写（可在设置中开关、配置本地服务地址与模型名）

### 4) PPT：生成、预览、编辑（重点能力）

#### 生成流水线（端到端）

- **大纲 → 视觉提示词 → 生图 → 打包 PPTX** 一次完成
- 视觉提示词由 **Gemini（OpenRouter）**生成
- 图片由 **DashScope（通义万相）**生成，并做尺寸/黑边等后处理，最终导出 `.pptx`

支持多种风格方向（可在提示词中指定）：
- Ethereal Glass & Gradient（通透玻璃/渐变）
- Swiss International（经典现代主义网格）
- Midnight Pro（暗夜高级质感）
- Neo-Chinese Zen（新中式留白）

#### 编辑能力（更像“图像级 PPT 编辑”）

- **整页重做**：对某页不满意，按该页内容与风格重新生成并替换
- **局部编辑**：框选区域，生成编辑提示词，仅修改局部元素并回填
- **预览器**：缩略图导航、页码、全屏、缩放

### 5) Web 搜索（桌面端内置）

- **内置 Brave Search MCP（In-Memory 直连）**：无需单独起外部 MCP 服务，只要配置 `Brave Search API Key`
- 支持返回：Web / News / FAQ / Videos / Discussions 等结构化结果

### 6) Excel 预览（桌面端）

- `.xlsx/.xls` 工作表预览与基础数据读取
- 常见工作流：让 AI 解释表结构、提取关键指标、生成分析结论、生成图表建议与口径说明
-（说明：目前更偏“预览 + AI 辅助处理”，不是完整 Excel 替代品）

### 7) ONLYOFFICE 真编辑模式（可选）

如果你需要更接近“真实 Office”级别的编辑体验（特别是兼容性、复杂排版、原生协作生态），项目内提供 **ONLYOFFICE Document Server** 的集成编辑器模式（需要你本地启动 Document Server）。

---

## 🚀 部署运行

### 环境要求

- **Node.js** >= 18
- **npm** >= 9
- **OS**：Windows 10+ / macOS 10.15+

### 运行模式说明（非常重要）

- **桌面端（Electron）**：功能最完整
  - ✅ 工作区（打开本地文件夹/读写文件）
  - ✅ Brave 搜索（内置 MCP）
  - ✅ PPT 生成/编辑（主进程执行：生图/打包/写文件）
  - ✅ ONLYOFFICE（可选，需你本地起 Document Server）
- **浏览器（Web）**：适合做 UI/交互开发
  - ✅ Word 编辑器 UI（Tiptap）
  - ✅ 纯前端可运行部分
  - ⚠️ 缺少本地文件系统/桌面端搜索/PPT 生成（需要额外后备接口或自行改造）

### 安装依赖

```bash
npm install
```

### 开发模式（推荐：桌面端）

```bash
npm run dev:electron
```

说明：
- 桌面端模式才能使用：**文件工作区/本地文件读写、Brave 搜索、PPT 生成与编辑** 等能力。

### 仅 Web 模式（UI 开发用）

```bash
npm run dev
```

说明：
- 浏览器模式缺少 Electron 主进程能力，某些功能会受限（例如：工作区文件系统、桌面端搜索、PPT 生成）。
- 如果你确实要在 Web 模式启用“联网搜索”，可以配置后备接口（见下方“配置说明 → Web 搜索后备接口”）。

### 构建与打包（桌面应用）

```bash
# 构建 Web 资源（dist/）
npm run build

# 打包桌面应用（Windows NSIS / macOS DMG）
npm run build:electron
```

构建产物输出到 `release/`。

### 打包产物说明（你会得到什么）

- **Windows**：`release/` 下会生成 NSIS 安装包（`.exe`）与相关文件
- **macOS**：`release/` 下会生成 DMG（`.dmg`）

> 如果你只想验证功能，优先用 `npm run dev:electron`；打包用于分发给其他机器安装。

---

## ⚙️ 配置说明

项目支持两种配置方式：
- **应用内设置（推荐）**：右上角设置入口，自动本地保存
- **`.env`（可选）**：Electron 主进程会读取根目录 `.env`

> 安全提示：**不要把真实 API Key 提交到公开仓库**。仓库已提供 `.env.example` 作为模板。

### 配置优先级（避免“我配了但没生效”）

- **桌面端**：
  - 优先使用 **应用内设置**（UI 中填写并自动保存）
  - 其次读取根目录 `.env`
- **Web 模式**：
  - 只能读取 `VITE_*` 前缀的环境变量（Vite 规则）
  - 其他 Key 建议仍通过 UI 配置或自行实现后端代理

### Key 获取地址（官方入口）

- **主模型 API Key（推荐）**：[https://linapi.net/register?aff=FGJ7](https://linapi.net/register?aff=FGJ7)
- **Brave Search API**：`https://brave.com/search/api/`
- **OpenRouter**：`https://openrouter.ai/`
- **DashScope（阿里云百炼）**：`https://dashscope.console.aliyun.com/`

### 配置项一览（对应 `AISettings`）

| 配置项 | 用途 | 典型值/说明 |
|---|---|---|
| `apiKey` | 主模型 Key（必需） | OpenAI/兼容网关/自建服务 |
| `baseUrl` | 主模型网关地址（必需） | 形如 `https://.../v1` 或兼容 OpenAI 的 base URL |
| `model` | 主模型名称（必需） | 如 `gemini-2.5-flash-preview-05-20` / 其他兼容模型 |
| `temperature` | 生成随机性 | 建议 0.3~0.9（按场景调整） |
| `maxTokens` | 输出长度上限 | 视模型与成本而定 |
| `localModel.enabled` | 是否启用 Tab 补全 | `true/false` |
| `localModel.baseUrl` | 本地补全服务地址 | 如 `http://127.0.0.1:8080/v1` |
| `localModel.model` | 本地补全模型名 | 如 `gpt-oss-20b` |
| `openRouterApiKey` | PPT 视觉提示词（Gemini） | OpenRouter Key（建议配置） |
| `dashscopeApiKey` | PPT 生图/局部编辑 | DashScope Key（建议配置） |
| `pptImageModel` | PPT 生图模型 | `z-image-turbo`（快）/ `qwen-image-plus`（质感） |
| `braveApiKey` | 联网搜索 | Brave Search API Key（强烈推荐） |

### 1) 主模型（必需）

用于 AI 对话、文档编辑等能力（支持 OpenAI 兼容接口）。

- **API Key**：你的模型服务 Key
- **Base URL**：兼容 OpenAI 的 `baseUrl`（例如某些 Gemini/第三方网关）
- **Model**：模型名（如 `gemini-2.5-flash-preview-05-20` 等）
- **Temperature / Max tokens**：生成参数

### 2) PPT 生成（建议配置）

#### OpenRouter（Gemini 生成视觉提示词）

- **OpenRouter API Key**：用于调用 Gemini 生成每页视觉设计提示词

#### DashScope（通义万相，生图/图像编辑）

- **DashScope API Key**：用于 `qwen-image-plus`/`z-image-turbo` 等生图，以及 `qwen-image-edit-plus` 局部编辑

也可通过 `.env` 配置（可选）：

```env
DASHSCOPE_API_KEY=你的_key
```

### 3) Brave Web 搜索（可选，但强烈推荐）

用于联网搜索与资料调研。可在应用设置中填写，也可用 `.env`：

```env
BRAVE_API_KEY=你的_key
```

### 4) 本地补全模型（可选）

用于 `Tab` 低延迟补全：
- **enabled**：是否启用
- **baseUrl**：例如 `http://127.0.0.1:8080/v1`
- **model**：例如 `gpt-oss-20b`

### 5) Web 搜索后备接口（仅 Web 模式可选）

在浏览器模式下，如果你有自建的“搜索代理服务”，可配置：

```env
VITE_WEB_SEARCH_ENDPOINT=https://your-web-search-endpoint
VITE_WEB_SEARCH_KEY=your-token
```

---

## 🧭 使用指南

### 三分钟上手（从 0 到出结果）

1. **启动桌面端**
   - 运行：`npm run dev:electron`
2. **打开工作区**
   - 在左侧打开一个本地文件夹（建议新建一个空文件夹专门测试）
3. **配置 Key**
   - 主模型：必须（否则无法对话与编辑）
   - Brave：建议（联网搜索）
   - OpenRouter + DashScope：做 PPT 必配（视觉提示词 + 生图/编辑）
4. **做一次完整闭环验证**
   - 新建/打开一个 `.docx`
   - 选中一段文字按 `Ctrl + K`，输入“更正式、更精炼，保留原意”
   - 在底部“待确认修改”里 **接受/拒绝**，确认 diff 审阅链路正常

### 典型工作流（复制即可用）

#### 工作流 A：写一份“可交付”的报告（推荐）

1. 打开/新建文档
2. 对话中输入（示例）：

```
我要写一份《XX 项目阶段总结》，请先给我一个 8~10 个小节的目录大纲（包含每节要点），确认后再逐节扩写。
要求：语气正式、数据口径清晰、每节 3~5 条要点。
```

3. 让 AI 按章节逐步扩写，并在每次修改后通过 **待确认修改**审阅
4. 最后输入：

```
请把全文统一成公文写法：标题层级清晰、用词正式、避免口语；统一“公司/企业”等术语口径，并生成 200 字摘要置于开头。
```

#### 工作流 B：联网调研 → 引用到文档

```
请联网搜索“XX 领域 2025 最新政策/数据”，给出 5 条最有引用价值的结论，每条附来源链接与一句引用摘要。
然后把这些内容整理成文档中的“参考资料与引用”小节。
```

#### 工作流 C：从文档生成 PPT（端到端）

```
请基于当前文档生成 12 页 PPT 大纲（严格输出 JSON），风格：Swiss International，信息密度中等偏高，每页 3~5 个要点。
注意：第一页封面、最后一页总结与行动项。
```

确认后开始生成；生成完成后可继续：

```
第 6 页太花了：整体更克制、留白更多、文字更清晰；保持主题一致，整页重做。
```

#### 工作流 D：局部编辑 PPT（只改一块）

1. 在 PPT 预览中按住 `Ctrl` 拖拽框选区域
2. 对话中输入：

```
只修改我框选的区域：把图标改成线性风格，颜色从亮蓝改为低饱和灰蓝；其他区域保持不变。
```

### 第一次使用（建议流程）

1. 启动桌面端：`npm run dev:electron`
2. 打开一个本地文件夹作为 **工作区**
3. 右上角打开 **设置**：
   - 填写主模型 API（必需）
   - 填写 Brave Key（用于联网搜索）
   - 需要做 PPT 再补充：OpenRouter Key + DashScope Key
4. 在工作区中打开/新建 `.docx/.xlsx/.pptx` 开始使用

### 文档（Word）使用要点

- **选区 AI 快改**：选中文字 → `Ctrl + K` 或右键菜单 → 输入指令（例如“更正式一点、保留原意、不要加段落”）
- **审阅式应用**：AI 改完后会出现“待确认修改”提示，可 **接受/拒绝**；长文建议打开“查看全部修改”
- **打印布局缩放**：`Ctrl + 滚轮`，或 `Ctrl + 0/+/-`

### 快捷命令（在对话框中输入）

| 命令 | 作用 |
|---|---|
| `/润色` | 优化表达，更专业更通顺 |
| `/精简` | 删除冗余，保留关键信息 |
| `/翻译` | 中英互译（自动识别语言） |
| `/格式化` | 统一格式（字体/字号/行距等） |
| `/编号` | 为标题生成编号（适用于报告/公文） |
| `/公文` | 按公文习惯改写与排版（适用于正式材料） |
| `/会议纪要` | 将要点整理为规范会议纪要 |
| `/总结` | 生成摘要/要点总结 |

### 常用快捷键

| 快捷键 | 功能 |
|---|---|
| `Ctrl + S` | 保存文档 |
| `Ctrl + Z` | 撤销 |
| `Ctrl + Y` | 重做 |
| `Ctrl + K` | 选区 AI 快改（需要先选中文字） |
| `Tab` | 触发/接受智能补全（需启用本地模型） |
| `Esc` | 取消补全 / 关闭弹窗 /（有待确认修改时）拒绝全部 |
| `Enter` |（有待确认修改时）接受全部 |
| `Ctrl + 滚轮` / `Ctrl + 0/+/-` | 打印布局缩放 |

### 对话式编辑示例（可直接复制）

```
帮我把这份内容改成正式公文语气，并统一标题层级（一级黑体居中，二级黑体左对齐）。

把全文中的“公司”统一替换为“企业”，但不要改“公司法”“公司治理”这些固定搭配。

对第一章做一个 200 字摘要，然后把摘要插入到正文开头，作为“摘要”小节。

把表格里“精工园3-102”替换成“精工园3-105”，并保持原单元格格式不变。
```

### PPT 生成与编辑（桌面端）

- **生成**：在对话中让 AI 先给“结构化大纲”，确认后开始生成（Gemini 设计 → DashScope 生图 → 导出 PPTX）
- **整页重做**：对某页说“这页重做，风格更高级更克制，文字更清晰”
- **局部编辑**：在预览里 `Ctrl + 拖拽` 框选区域 → 在对话里说“把这里的图标换成更简洁的线性风格，颜色更低饱和”

### PPT 大纲 JSON 规范（强烈建议按这个格式给 AI）

你可以直接对 AI 说：“请严格输出 JSON，不要输出解释文字”，并使用如下结构（字段名支持一定的中英文变体，但建议统一用下面这版）：

```json
{
  "title": "演示文稿标题",
  "theme": "主题（可选）",
  "styleHint": "视觉风格提示（可选）",
  "slides": [
    {
      "pageNumber": 1,
      "pageType": "cover",
      "headline": "主标题",
      "subheadline": "副标题（可选）",
      "bullets": ["要点1", "要点2", "要点3"],
      "footerNote": "页脚/备注（可选）",
      "layoutIntent": "布局意图（可选，如：左文右图/对比/时间线）"
    }
  ]
}
```

**推荐做法**：
- 先让 AI 产出大纲 → 你确认页数与结构 → 再开始生成
- 生成后如果某页不满意：
  - “整页重做”：更适合整体风格/版式/文案密度都不对
  - “局部编辑”：更适合只改某个元素（图标、配色、局部文字清晰度）

### ONLYOFFICE 真编辑模式（可选）

启动 ONLYOFFICE Document Server（示例命令）：

```bash
docker run -i -t -d --name onlyoffice-ds -p 8080:80 onlyoffice/documentserver
```

然后在 Word-Cursor 中切换到 ONLYOFFICE 编辑器模式（`onlyoffice`），即可在 `http://localhost:8080` 上加载编辑器 API。

#### ONLYOFFICE（更详细的建议参数）

如果你需要更稳定的本地持久化与重启恢复，可加上数据卷（可选）：

```bash
docker run -i -t -d ^
  --name onlyoffice-ds ^
  -p 8080:80 ^
  -v onlyoffice_data:/var/www/onlyoffice/Data ^
  -v onlyoffice_logs:/var/log/onlyoffice ^
  onlyoffice/documentserver
```

常见问题：
- 浏览器打不开 `http://localhost:8080`：检查端口占用、防火墙、Docker Desktop 是否启动
- 应用里提示“首次加载可能需要 30-60 秒”：首次拉取脚本与初始化需要时间，属于正常现象

---

## 📁 仓库结构

```
word-cusror/
├── electron/                  # Electron 主进程（文件系统、搜索、PPT 生图/打包等）
├── src/                       # React 渲染进程（编辑器/对话面板/工作区 UI）
├── brave-search-mcp-server/   # Brave Search MCP Server 源码（用于参考/二次开发）
├── scrapeless-mcp-server/     # Scrapeless MCP Server（网页自动化/抓取能力的独立服务）
├── DocumentServer/            # ONLYOFFICE Document Server 相关资料（含 roadmap 等）
├── public/                    # 静态资源
├── dist/                      # 前端构建产物
└── package.json
```

### 子项目说明

- **`electron/`**
  - 桌面端能力入口：文件读写、Brave MCP 内置搜索、PPT 生成/编辑、图像后处理等
- **`src/`**
  - UI 与交互：工作区、Word/Excel/PPT 预览、对话面板、审阅式 diff 等
- **`scrapeless-mcp-server/`**
  - 一个独立的 MCP Server：提供浏览器自动化、抓取动态网页、导出 Markdown/截图等能力
  - 可与 Cursor/Claude 等 MCP 客户端集成，用于更强的“真实世界上下文”获取
- **`DocumentServer/`**
  - ONLYOFFICE 相关资料；项目中提供 ONLYOFFICE 编辑器集成（需要你本地启动 Document Server）

---

## 🧩 技术栈（概览）

- **桌面端**：Electron
- **前端**：React + TypeScript + TailwindCSS + Vite
- **文档/格式**：Tiptap（ProseMirror）、docx / mammoth、exceljs、pptxgenjs、jszip
- **AI 与工具调用**：
  - OpenAI 兼容 API（主模型）
  - OpenRouter（Gemini：PPT 视觉提示词）
  - DashScope（通义万相：生图与图像编辑）
  - Brave Search MCP（联网搜索，桌面端内置）

---

## 🧰 可选扩展：Scrapeless MCP Server（网页自动化/抓取）

仓库中包含 `scrapeless-mcp-server/`，它是一个遵循 MCP 标准的独立服务，适合用于：
- 抓取 JS-heavy 的动态网页（导出 HTML / Markdown / 截图）
- 浏览器自动化（打开页面、点击、滚动、输入）
- 绕过部分反爬/Cloudflare 场景（取决于服务能力与策略）

> 说明：它是“可选增强能力”，不影响 Word-Cursor 主程序的基本使用。

### 在 MCP 客户端中接入（示例：本地 stdio）

在支持 MCP 的客户端（例如 Cursor / Claude Desktop）中添加类似配置（示例格式）：

```json
{
  "mcpServers": {
    "Scrapeless MCP Server": {
      "command": "npx",
      "args": ["-y", "scrapeless-mcp-server"],
      "env": {
        "SCRAPELESS_KEY": "YOUR_SCRAPELESS_KEY"
      }
    }
  }
}
```

如果你想用“托管 API 模式（HTTP）”，请参考 `scrapeless-mcp-server/README.md`。

## ❓ FAQ

### 1) 为什么我在浏览器模式下用不了某些功能？

很多能力依赖 Electron 主进程（文件系统、Brave MCP、PPT 生图/打包）。请用 `npm run dev:electron` 运行桌面端。

### 2) PPT 生成需要哪些 Key？

建议都配：
- **OpenRouter API Key**：Gemini 生成视觉提示词
- **DashScope API Key**：生图与局部编辑

### 3) ONLYOFFICE 编辑器报“无法连接到 ONLYOFFICE 服务器”

请确认：
- Docker 正在运行
- `onlyoffice/documentserver` 容器已启动
- 端口映射为 `-p 8080:80`

### 4) 安全建议有哪些？

- 不要提交真实 API Key
- 生产环境建议使用安全的配置管理/密钥服务
- 对外部网页内容（抓取/搜索结果）在送入模型前做必要的过滤与防注入策略

---

## 🧯 故障排查（更详细）

### 1) 启动相关

- **`npm run dev:electron` 失败**
  - 确认 Node.js >= 18
  - 删除 `node_modules` 后重装：`npm install`
  - Windows 下建议使用 PowerShell/Windows Terminal，并避免路径包含过长/特殊字符

- **白屏/界面加载但功能按钮无响应**
  - 优先看 Electron 主进程控制台输出（是否有端口占用/权限错误/Key 缺失）
  - 先用一个“空工作区”验证（排除超大文件夹/特殊文件名导致的读取问题）

### 2) 联网搜索相关

- **提示“请在设置中配置 Brave Search API Key”**
  - 在设置里填写 `Brave Search API Key`
  - 或在 `.env` 配置 `BRAVE_API_KEY`
  - 注意：桌面端会优先使用设置里填写的 Key；`.env` 仅作为兜底

- **返回结果为空**
  - 尝试把 query 改得更具体（加入年份、地区、机构名）
  - 调大 `num`（如果 UI 暴露了搜索数量）
  - 检查网络代理/公司内网策略是否拦截外部请求

### 3) PPT 生成相关

- **提示缺少 OpenRouter / DashScope Key**
  - OpenRouter：用于 Gemini 生成“每页视觉提示词”
  - DashScope：用于生图与局部编辑（没有它无法生成 PPT）
- **生成很慢或失败**
  - 生图接口可能触发并发/限流/安全审核；可降低页数先验证链路
  - 如果某页反复失败：改写该页内容（避免敏感词）、降低“写进图片的文字密度”，再重试

- **PPT 看起来“AI 味很重/排版不稳定”**
  - 先固定风格（如 Swiss International / Midnight Pro），再调整信息密度
  - 让 AI 在大纲里写清楚每页的“布局意图”（`layoutIntent`）
  - 通过“整页重做”统一审美，不要只用局部改动堆补丁

### 4) ONLYOFFICE 相关

- **提示无法连接 ONLYOFFICE**
  - 确认 Docker Desktop 正在运行
  - 确认容器已启动：`docker ps`
  - 确认端口：`http://localhost:8080`
  - 若端口 8080 被占用：把命令改为 `-p 18080:80`，并同步修改应用中 Document Server 地址（如有配置项）

### 5) 本地补全（Tab）相关

- **Tab 没反应**
  - 确认设置里启用了本地补全（`localModel.enabled=true`）
  - 确认本地模型服务可访问（`localModel.baseUrl`）
  - 补全通常需要一定上下文长度；建议先输入一段再按 Tab

---

## 🔐 安全与合规（建议阅读）

- **密钥安全**
  - 不要把 `.env` 里的真实 Key 上传到公开仓库
  - 建议仅在本机保存，或使用专门的密钥管理方案
- **提示注入（Prompt Injection）**
  - 来自网页搜索/抓取的内容默认“不可信”
  - 不要把原始网页内容不加过滤地直接塞进 system prompt
- **内容与版权**
  - PPT 生图结果可能受模型与平台条款约束；用于商业分发前请确认合规性

### 安全实践清单（落地版）

- **最小权限**：只配置你要用的 Key；不用就不填
- **最小暴露**：不要把 Key 写进截图/录屏/演示文档
- **最小信任**：网页内容先“提取结构化要点”再送入模型，避免把整页原文直接喂给模型

---

## 🤝 贡献

欢迎提交 Issue / PR（功能建议、Bug 修复、文档完善都非常欢迎）。

---

## 📄 License

MIT

