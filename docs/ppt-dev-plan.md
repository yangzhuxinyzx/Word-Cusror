# PPTX 预览与真编辑（纯前端渲染）开发文档

## 目标
- **目标体验**：像 PowerPoint 一样的 PPT 体验：左侧缩略图、主画布、顶部 Ribbon 工具栏、缩放/全屏放映、复制粘贴、撤销重做。
- **兼容目标**：能打开常见 `.pptx` 并正确预览；在 MVP 范围内能编辑并保存回 `.pptx`，用 PowerPoint 打开仍可继续编辑。
- **工程约束**：优先“可落地 + 可扩展”，避免一次性实现全量 Office 功能导致失控。

## 现状与关联能力
- 已有：Electron 文件系统 IPC、工作区文件树、以及 image-only 的 `.pptx` 导出能力（`pptxgenjs` 打包）。
- 新增目标：让 `.pptx` 变成“一等公民”：**打开即可预览**，并在编辑器内完成**真编辑**。

## 总体策略（必须先锁定范围）
“真编辑 + 纯前端解析/渲染 pptx”本质是在做迷你 PowerPoint。为避免爆炸式增长，必须分阶段：

- **Phase 0（只读预览）**：解析 `.pptx` → 渲染缩略图/主画布 → 翻页/缩放/全屏。
- **Phase 1（MVP 真编辑）**：支持最核心对象的编辑与保存回 `.pptx`：
  - 文本框（编辑文字、字号/颜色/对齐、粗斜下划线）
  - 图片（移动/缩放/替换图片）
  - 基础形状（矩形/圆/线/箭头，填充/描边）
  - 背景（纯色/背景图）
- **Phase 2（兼容增强）**：母版/主题（部分）、更多形状、段落高级排版、组合/对齐分布等。
- **Phase 3（可选）**：表格/图表/SmartArt/动画/切换（风险最高，建议最后再评估）。

> 注意：Phase 1 是“真编辑”的合理最小闭环。若要求 SmartArt/图表/动画等全支持，周期会极长。

## MVP 降级策略（建议默认接受）
- **母版/主题**：不完全还原，允许近似样式与字体回退。
- **动画/切换**：先不支持（只做静态页）。
- **SmartArt/复杂图表**：先静态化（渲染为位图）或跳过提示。
- **字体**：不内嵌字体，按系统字体回退。

## PPTX（OpenXML）结构速览
`.pptx` 是 zip，关键路径：
- `ppt/presentation.xml`：演示文稿结构、slide 列表顺序
- `ppt/slides/slide1.xml`...：每页内容（shape、text、pic 等）
- `ppt/slides/_rels/slide1.xml.rels`：该页资源引用（图片、外链等）
- `ppt/media/*`：图片/媒体资源
- `ppt/theme/*`、`ppt/slideMasters/*`、`ppt/slideLayouts/*`：主题/母版/布局（兼容复杂）

## 核心数据模型（编辑器内存态）
### 统一坐标与单位
- PPTX 内部使用 EMU（English Metric Unit）
- 渲染层使用 px（或逻辑像素）
- 必须有稳定换算：
  - `EMU_PER_INCH = 914400`
  - `px = (emu / EMU_PER_INCH) * DPI`
  - DPI：建议用 96 作为逻辑 DPI（与浏览器常见一致）

### PresentationModel（建议结构）
- `PresentationModel`
  - `meta`: size、themeRefs、defaultTextStyle
  - `slides: SlideModel[]`
- `SlideModel`
  - `id`, `index`
  - `background: SlideBackground`
  - `elements: ShapeModel[]`
  - `notes?: ...`（可选）
- `ShapeModel`（联合类型）
  - `TextBoxShape`：`x,y,w,h,rotation,zIndex,textRuns[],paragraphStyle`
  - `ImageShape`：`x,y,w,h,rotation,mediaRef,crop`
  - `BasicShape`：`shapeType, fill, stroke`

## 渲染架构（建议 SVG/DOM 为主）
### 为什么不推荐纯 Canvas
Canvas 处理中文文本编辑、选区定位、字体回退很难；SVG/DOM 更适合文本清晰与可编辑。

### 推荐方案
- **主画布**：SVG（形状/图片可用 `<image>` + `<path>`，文本用 `<foreignObject>` 或 DOM overlay）
- **文字编辑**：选中后切换为 DOM overlay（contenteditable/富文本引擎），提交后回写到模型。
- **缩略图**：复用同一渲染器，降低分辨率/禁用编辑层。

```mermaid
flowchart TD
  PptxZip[PPTX_zip(JSZip)] --> Parse[ParseXML_to_Model]
  Parse --> Model[PresentationModel]
  Model --> RenderThumb[Render_Thumbnails]
  Model --> RenderStage[Render_Stage]
  UIEvents[UI_Events] --> EditOps[Edit_Operations]
  EditOps --> Model
  Model --> Serialize[Model_to_XML]
  Serialize --> ZipWrite[WriteZip_PPTX]
```

## 编辑交互（Phase 1 必须具备）
- **选择/多选**：点击选中、Shift 多选、框选
- **变换**：拖拽移动、8向缩放、旋转（可后置）
- **对齐辅助**：参考线/吸附（Phase 2）
- **快捷键**：Ctrl/Cmd + Z/Y、Ctrl/Cmd + C/V、Delete
- **Undo/Redo**：基于 operation log（建议：命令模式 + 可逆操作）

## 写回 PPTX（保存）策略
### 两种策略
- **A. 全量重写（MVP 推荐）**
  - 重写 `ppt/slides/slideX.xml`（以及对应 rels）
  - `ppt/presentation.xml` 仅更新 slide 顺序/新增删除页
  - 不动主题/母版（除非必要）
  - 优点：实现简单、可控；缺点：可能丢失部分高级属性
- **B. 增量 patch（Phase 2+）**
  - 精确修改 XML 节点与 rels
  - 优点：保留更多原始信息；缺点：实现复杂、易出边缘 bug

### media 管理（Phase 1）
- 替换图片 = 写入 `ppt/media/xxx.png` + 更新该 slide 的 `.rels`
- 新增图片 = 分配新 rid + 新文件名 + 写入 media

## 依赖与实现选型（建议）
- zip：`jszip`（仓库已用）
- XML：建议引入 `fast-xml-parser`（更适合双向：parse + build），或 DOMParser+XMLSerializer（实现更细但更难维护）
- 渲染：自研 SVG/DOM
- 交互：可用 `konva`/`fabric`（若接受 Canvas），但文本编辑会更难；SVG/DOM 仍建议自研轻量控制器

## 与现有项目的集成点（建议路径）
### 文件打开与预览
- 在 `DocumentContext.openFile()` 中识别 `.pptx`：
  - Phase 0：读取 zip（Electron main IPC `read-file` 返回 base64）→ 前端解析 → 显示 `PptPreview` 组件

### UI 结构
- 新增 `src/components/PptEditor.tsx`
  - 左：缩略图列表
  - 中：舞台渲染
  - 上：Ribbon 样式操作区

### 保存链路
- 前端生成新的 zip buffer（或把变更集发给主进程）
- 通过 `electronAPI.writeBinaryFile()` 保存为 `.pptx`

## 里程碑与任务拆解（建议）
### Phase 0：只读预览（2–4 周）
- 解析：presentation.xml + slideX.xml（只处理常见 shape：text、pic、basic shapes）
- 渲染：缩略图 + 主舞台 + 翻页/缩放
- 兼容：字体回退、主题/母版先忽略

### Phase 1：MVP 真编辑（4–8 周）
- 文本框编辑（含基础样式）
- 图片替换与基本变换
- 基础形状样式调整
- 保存回 pptx（全量重写 slides XML + rels + media）

### Phase 2：增强兼容（持续迭代）
- 母版/主题部分支持
- 对齐分布、组合、更多形状
- 表格/图表/SmartArt 的策略（静态化 vs 真编辑）

## 风险清单（必须提前暴露）
- **字体与中文排版一致性**：PowerPoint 使用的字体可能用户机器没有 → 回退会导致换行变化
- **主题/母版解析复杂**：不做母版会导致部分 pptx 看起来“差一点”，但能大幅加快交付
- **写回兼容性**：保存后的 pptx 必须能被 PowerPoint 打开并继续编辑，否则体验崩溃
- **性能**：大 deck（100+ 页）与复杂页渲染需要虚拟列表与分层渲染

## 测试计划（必须建立样例集）
### 样例集建议（按优先级收集）
1. **基础**：纯文本框、纯图片、形状组合、背景图
2. **常见模板**：学校汇报/论文答辩/商业汇报（多中文、多图）
3. **复杂**：母版、主题字体、SmartArt、图表、表格
4. **边界**：超长文本、多字体混排、旋转文本、透明度、裁剪图片

### 建议建立的样例目录结构
建议在仓库里建立（或在本地工作区建立）一套固定命名的样例集，方便回归：
- `samples/pptx/basic/`
- `samples/pptx/templates/`
- `samples/pptx/complex/`
- `samples/pptx/edge/`

每个样例同时保存：
- 原始文件：`xxx.original.pptx`
- 编辑器保存后的文件：`xxx.saved.pptx`
- 预期截图（可选）：`xxx.expected.png`（用于快速对比）

### 验收标准（每个样例必测）
- **打开预览一致性**：
  - 缩略图页序正确
  - 主要元素位置误差 < 2px（以 1920×1080/1200 的逻辑尺寸计）
  - 中文不乱码、不缺字、不明显错行
- **编辑能力**（Phase 1 样例）：
  - 文本框可改字并保存
  - 图片可替换并保存
  - 形状样式可改并保存
- **保存回 PowerPoint**：
  - PowerPoint 能正常打开
  - 再次编辑不报错、不丢页、不丢图

### 自动化/半自动化验收建议
- **打开/保存一致性**：对每个样例执行“打开→不改动→保存”为 `*.saved.pptx`，再用 PowerPoint 打开确认无修复提示。\n
- **像素对比（可选）**：将 `original` 与 `saved` 的每页都渲染成 PNG（可先用 PowerPoint/LibreOffice 手动导出），做快速视觉 diff。\n
- **压力样例**：准备一个 80–150 页的 deck，验证缩略图虚拟列表与主画布渲染性能。\n
- **中文字体回退**：准备“微软雅黑/宋体/黑体/苹方/思源”等混用样例，验证回退策略与换行变化可接受。\n
- **媒体引用正确性**：包含重复引用同一张图片、多格式图片（png/jpg）与裁剪的样例，验证 rels 不错乱。\n

## 备注：与“DashScope 海报式生成”能力的关系
- 当前 image-only 生成导出可以作为 **快速生产**路径（适合一键做汇报）。
- 真编辑路线用于“打开现有 pptx 做二次编辑”的场景。
- 两条路径可共存：生成的 image-only pptx 也应能被预览器打开（即使不可编辑细粒度对象）。


