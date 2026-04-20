# Word-Cursor 智能体 Runtime 重构开发计划

## 1. 背景

Word-Cursor 现在已经具备较强的业务能力：

- Word 文档打开、编辑、导出
- Excel 预览与操作
- PPT 生成、编辑、预览
- 工作区文件访问
- Web 搜索
- 本地 memory

但当前智能体实现更接近“把很多业务能力接到一个大聊天组件里”。

现状的核心问题是：

- [ChatPanel.tsx](/C:/Users/yangz/Desktop/Github/Word-Cursor/src/components/ChatPanel.tsx) 同时承担 UI、tool dispatch、tool execution、结果回写、业务编排
- [AIContext.tsx](/C:/Users/yangz/Desktop/Github/Word-Cursor/src/context/AIContext.tsx) 同时承担 prompt、loop、消息管理、tool 协议解析、memory 注入
- 文档领域能力很强，但 agent runtime 抽象不够稳定
- 动态上下文、工具描述、业务约束大多直接堆进主 prompt
- 新功能一旦增加，容易继续放大 `ChatPanel`、`AIContext`、`electron/main.cjs` 的复杂度

Claude Code 值得借鉴的不是某一段 prompt，而是它的分层方式：

- 会话引擎
- prompt/context 组装
- tool 抽象与注册
- tool 执行与并发调度
- skills
- subagents
- attachments / memory / delta context

本计划的目标，是把 Word-Cursor 从“业务能力集中在聊天层”重构为“智能体 runtime 驱动业务域”。

## 2. 本次重构范围

本计划覆盖：

- 智能体 runtime 分层
- tool 抽象与迁移
- prompt / context / attachment 体系
- Word / Excel / PPT / workspace / search / memory 的 agent 接入方式
- skill 机制
- subagent 机制
- Electron 服务拆层
- UI 与 runtime 解耦

本计划暂不把“权限与安全系统”作为里程碑阻塞项。

这不等于删除现有保护逻辑，而是：

- 不以 Claude Code 那套 permission system 为本轮主要目标
- 相关接口预留，但不优先投入重构精力
- 先把 runtime 结构搭稳，再考虑更精细的权限分层

## 3. 重构总目标

重构完成后，目标形态应为：

1. UI 不再直接承担业务 tool 编排。
2. agent runtime 成为一个独立层，负责会话、prompt、上下文、tools、skills、subagents。
3. 各业务域通过统一 tool contract 接入 runtime。
4. 动态上下文不再全部塞进主 system prompt，而是分成静态 prompt、动态 section、attachments。
5. Word 编辑作为第一核心域，具备最强的 agent 接入能力。
6. Excel、PPT、workspace、search、memory 成为可插拔的 domain tool packs。
7. Electron 主进程从“全能后端”逐步拆成更清晰的服务层。

## 4. 目标架构

建议新增一个明确的 runtime 目录，例如：

```text
src/agent/
  core/
    AgentSessionEngine.ts
    AgentLoop.ts
    MessageStore.ts
    ConversationState.ts
    ModelGateway.ts
    runtimeTypes.ts
  prompt/
    SystemPromptComposer.ts
    ContextAssembler.ts
    PromptSections.ts
  attachments/
    AttachmentManager.ts
    AttachmentTypes.ts
    builders/
  tools/
    contracts.ts
    ir.ts
    registry.ts
    executor.ts
    scheduler.ts
    results.ts
    packs/
      word/
      excel/
      ppt/
      workspace/
      web/
      memory/
    search/
  skills/
    SkillRegistry.ts
    SkillExecutor.ts
    builtins/
  subagents/
    SubagentManager.ts
    AgentProfiles.ts
    builtins/
  adapters/
    document/
    electron/
    editor/
    providers/
  tasks/
    TaskRegistry.ts
    TaskRunner.ts
    TaskNotifications.ts
  storage/
    SessionTranscriptStore.ts
    ToolResultStore.ts
    ReplayStore.ts
  hooks/
    HookRegistry.ts
    HookDispatcher.ts
  compaction/
    AutoCompact.ts
    ReactiveCompact.ts
    ContextCollapse.ts
```

同时逐步把 Electron 侧拆成：

```text
electron/
  services/
    files/
    search/
    memory/
    excel/
    ppt/
    ai/
    fonts/
  ipc/
    registerFileIpc.cjs
    registerSearchIpc.cjs
    registerMemoryIpc.cjs
    registerExcelIpc.cjs
    registerPptIpc.cjs
    registerAiIpc.cjs
```

除了目录拆分，目标架构还必须明确包含这些运行时子系统：

- `ModelGateway` provider 适配层
- `ToolCallIR / ToolResultIR / ToolProgressIR / ToolErrorIR`
- background task / task framework
- transcript / replay / resume 存储层
- compaction / collapse / result budget
- hooks / extension points
- deferred tools / tool search

这些不是“高级增强项”，而是接近 Claude Code 架构时必须前置考虑的基础设施。

## 5. 核心设计原则

### 5.1 Runtime First

先稳定 runtime，再迁业务能力。

不要继续在现有 [ChatPanel.tsx](/C:/Users/yangz/Desktop/Github/Word-Cursor/src/components/ChatPanel.tsx) 上叠逻辑。后续新增能力默认进入 `src/agent/*`。

### 5.2 Tool First-Class

所有 agent 能力都必须先成为标准化 tool，再进入编排。

tool 至少要有：

- name
- description
- input schema
- output schema
- handler
- domain tag
- read/write classification
- concurrency policy
- result rendering metadata

### 5.2.1 Tool Call IR 必须前置定义

如果目标是让工具调用方式全面向 Claude Code 靠齐，那么不能只定义“tool metadata”，还必须在最前面定义统一内部协议。

建议新增内部 IR：

```text
ToolCallIR
ToolResultIR
ToolProgressIR
ToolErrorIR
```

至少应包含：

- `toolCallId`
- `toolName`
- `input`
- `status`
- `startedAt / finishedAt`
- `progress events`
- `result payload`
- `error payload`
- `source`（native / legacy / synthetic）

后续所有调用链都应该围绕 IR 运转：

- 模型原生工具调用 → IR
- 旧 `[TOOL_CALL]` 文本协议 → IR
- UI 展示 → 读取 IR
- attachments / telemetry / replay / logs → 读取 IR

不要把“统一内部 tool call IR”放到最后才做。

如果 IR 太晚定义，前面迁移出来的 tool packs 仍会隐式依赖旧文本协议，后续返工成本会明显上升。

### 5.2.2 Tool 对象不只是 handler

如果要追求稳定性和可扩展性，tool contract 不应只停留在：

- name
- schema
- handler

建议进一步明确每个 tool 至少具备：

- `prompt()` 或 tool-facing description
- `input schema`
- `output schema`
- `validateInput()`
- `checkPermissions()`
- `isReadOnly()`
- `isDestructive()`
- `interruptBehavior()`
- `concurrencyPolicy`
- `renderToolUse()`
- `renderToolResult()`
- `toTelemetryPayload()`

这样 runtime 才能真正承接：

- 参数校验
- 权限控制
- 并发调度
- 中断恢复
- UI 呈现
- 工具日志与回放

### 5.2.3 Tool Execution Pipeline 必须被文档化

如果目标是“功能更强、调用更稳”，那么工具执行链本身必须是架构的一部分，而不是散落在聊天组件里。

建议把统一执行流水线定义为：

1. `parse`：把模型输出转成 `ToolCallIR`
2. `schema validate`：结构校验
3. `semantic validate`：语义校验
4. `permission check`
5. `schedule`
6. `execute`
7. `collect progress`
8. `collect result`
9. `post-process`
10. `write back to conversation state`

并明确：

- 失败必须产生结构化 `ToolErrorIR`
- 中断必须产生可解释的 synthetic result / cancelled result
- 并发工具需要有稳定的 pairing 机制
- 任何工具结果回写都不应依赖 UI 组件本身的局部状态

### 5.2.4 主链路目标应当是 provider-native tool calling

如果目标是全面向 Claude Code 靠齐，那么新 runtime 的长期目标不能只是“把文本协议搬到新目录里”，而应当明确：

- **主链路优先支持 provider-native tool calling**
- 旧的 `[TOOL_CALL] ... [/TOOL_CALL]`、XML 标签协议只作为兼容 adapter

也就是说：

- 新功能默认走 native tool calling 或统一 IR
- 旧协议只负责把历史逻辑翻译成 IR
- 不再让新增 domain tool pack 直接依赖文本协议解析

这一点非常关键。

如果继续长期把文本协议当作主链路，那么以下问题仍会持续：

- DSL / JSON 参数脆弱
- 模型工具参数污染
- tool/result pairing 不稳定
- 最终反馈与中间状态混在 UI 逻辑里
- 截断恢复与并发调用难以做稳

因此，本计划虽然保留分阶段兼容策略，但应明确：

- **兼容旧协议不等于继续以旧协议为核心**
- **新 runtime 的主心骨应该是 IR + native tool calling**

### 5.2.5 Provider Adapter 必须单独成层

如果目标是“真正向 Claude Code 靠”，那么 provider 差异不能继续散落在 `AIContext` 或聊天 UI 中。

建议至少明确三类 adapter：

- `AnthropicToolUseAdapter`
- `OpenAICompatibleToolCallAdapter`
- `LegacyTextToolAdapter`

它们的唯一职责是：

- 读取 provider / legacy 协议输出
- 映射到统一 `ToolCallIR`
- 将统一 `ToolResultIR` / `ToolErrorIR` 映射回 provider 需要的回填结构

必须明确：

- provider-specific block 结构不允许直接泄漏到 `ToolExecutor`
- `AIContext` 不允许继续直接拼接或解析 provider 工具协议
- 所有 provider 差异都应止步于 adapter 层

### 5.2.6 IR 必须定义 block 映射规则

不仅要有 `ToolCallIR` 这个名字，还要明确不同来源如何映射：

- Anthropic `tool_use` -> `ToolCallIR`
- OpenAI-compatible `tool_calls` -> `ToolCallIR`
- legacy `[TOOL_CALL] ... [/TOOL_CALL]` -> `ToolCallIR`
- legacy XML / `<tool_use>` -> `ToolCallIR`
- synthetic cancelled / timeout / fallback -> `ToolErrorIR` 或 synthetic `ToolResultIR`

否则后续实现会变成：

- 每个模块自己定义一套“差不多的 IR”
- pairing、progress、retry、telemetry 再次分裂

### 5.2.7 ToolExecutor 必须以 IR 为唯一主输入

必须增加一条硬要求：

- 新 runtime 中，`ToolExecutor` 的主入口只接受 `ToolCallIR`
- 新 runtime 中，`ToolExecutor` 的主输出只返回 `ToolResultIR / ToolErrorIR / ToolProgressIR`

兼容层可以保留：

- `execute(toolName, args)` 这样的 legacy wrapper

但这只能是 adapter，不能再成为真正主入口。

### 5.2.8 工具生命周期状态必须写入 Session Runtime

为了接近 Claude Code 的稳定性，工具执行状态不能只停留在本地 UI 状态里。

建议统一状态流：

- `parsed`
- `validated`
- `scheduled`
- `executing`
- `completed`
- `failed`
- `cancelled`

这些状态必须进入 session runtime / snapshot，供以下模块统一读取：

- UI 展示
- replay
- telemetry
- bg task / subagent 通知
- retry / truncation recovery

### 5.2.9 ChatPanel 必须退出主工具编排

虽然本计划允许分阶段迁移，但必须加一条更强的工程约束：

- `ChatPanel` 不再承载主工具编排逻辑
- `ChatPanel` 只允许：
  - 发送用户输入
  - 订阅 runtime snapshot / event stream
  - 展示 tool progress / result / attachments

如果某项工具执行逻辑仍写在 `ChatPanel` 中，它只能作为临时兼容过渡，不是长期允许状态。

### 5.2.10 必须明确 native tool calling 的切换时机

文档应明确：

- 从哪个 phase 开始，新功能禁止继续走 legacy 文本协议
- 从哪个 phase 开始，provider-native tool calling 开始进入主链路
- 从哪个 phase 开始，legacy `[TOOL_CALL]` / XML 只作为 fallback adapter

建议最低要求：

- Phase 0：定义 IR 与 adapter
- Phase 1：新迁移 tool 默认走 IR
- Phase 2：`AgentLoop / ModelGateway / ToolExecutor` 内部统一 IR
- Phase 3：开始正式为 provider-native tool calling 接入 schema 注入与结果回填
- Phase 9：彻底下线旧协议主链路

### 5.2.11 ModelGateway 必须有能力矩阵

`ModelGateway` 不能只是“统一 fetch 一下模型”。

它必须明确描述每种 provider 的能力边界，否则 runtime 最后仍会默认按一种 provider 写死。

建议至少维护如下矩阵：

| Provider 类型 | 典型接口 | native tool-use | reasoning/thinking | 多模态 | prompt cache | deferred tools | 备注 |
|---|---|---|---|---|---|---|---|
| OpenAI-compatible | `/chat/completions` | 部分支持，取决于 provider 是否实现 `tools/tool_calls` | 部分支持，取决于具体模型与供应商 | 部分支持 | 通常无统一标准 | 通常无 provider 原生能力 | 这是 Word-Cursor 当前实际主链路来源 |
| Anthropic Messages API | `messages` | 强支持，`tool_use/tool_result` 是一等公民 | 强支持，thinking / redacted thinking 有明确协议 | 支持 | 支持 | 支持 `defer_loading` 等机制 | Claude Code 的核心原生路径 |
| Legacy Text Adapter | `[TOOL_CALL]` / XML / `<tool_use>` 文本协议 | 不支持 | 弱，靠文本约定 | 仅透传 | 弱 | 不支持 | 仅作兼容层，不应再是主链路 |

对我们项目的要求是：

- 当前 OpenAI-compatible 路径继续保留，但要通过 `OpenAICompatibleToolCallAdapter`
- 未来必须预留 `AnthropicToolUseAdapter`
- `ModelGateway` 必须返回能力声明，而不是让上层猜

建议能力声明至少包含：

- `supportsNativeToolUse`
- `supportsReasoning`
- `supportsMultimodal`
- `supportsPromptCache`
- `supportsDeferredTools`
- `supportsStructuredToolSchema`

### 5.2.12 Tool lifecycle 必须定义成状态机

仅有“执行流水线”还不够，必须进一步定义工具生命周期状态机。

建议统一为：

```text
parsed
validated
scheduled
executing
progress
completed
failed
cancelled
```

并明确：

- `tool_use -> progress -> tool_result` 是标准正向路径
- `tool_use -> cancelled -> synthetic result` 是中断路径
- `tool_use -> failed -> ToolErrorIR` 是失败路径

对我们项目尤其重要，因为：

- Word 长文档修改可能被中断
- PPT 生成和资料检索很适合后台跑
- Excel 批量操作可能运行时间较长

### 5.2.13 Tool pairing、synthetic result 与并发结果顺序必须写死

Claude Code 在这块非常强，我们也必须写成规范。

必须明确：

- 每个 `ToolCallIR` 必须有稳定的 `toolCallId`
- 每个 `ToolResultIR / ToolErrorIR / ToolProgressIR` 必须显式携带 `toolCallId`
- tool/result pairing 不能靠文本顺序推断

对于 aborted / cancelled / synthetic result，需要规定：

- 当工具被用户中断时，必须产生 `cancelled` 事件
- 当工具因为 sibling error / fallback / provider 中断而失效时，必须产生 synthetic result 或 synthetic error
- synthetic result 也必须进入 transcript / runtime snapshot，而不是只在 UI 层丢一条提示

并发时还要明确：

- **执行顺序** 可以并发
- **展示顺序** 必须稳定
- **结果回填顺序** 应按 `toolCallId` 或原始 tool call 序列决定，而不是按 Promise 返回时机乱序

### 5.2.14 Deferred tools / tool search 需要前置规划

不是所有工具都应该一开始全量暴露给模型。

Claude Code 值得学的是：

- 常用工具常驻
- 长尾工具延迟暴露
- 通过 tool search / deferred loading 减少 prompt 噪声

对 Word-Cursor 而言，建议区分：

- 常驻工具：Word 基础编辑、workspace 读取、web search、memory search
- 延迟工具：PPT 高级编辑、Excel 高级数据处理、模板填充、图表生成、批量文档生成

这不仅能降低 prompt 体积，也能减少模型误选工具的概率。

### 5.2.15 Hooks / extension points 必须预留

即便本轮不把权限/安全系统做重，也必须把 hook 位点预留清楚。

建议至少明确：

- `PreToolUse`
- `PostToolUse`
- `PostToolUseFailure`
- `PermissionDenied`
- `InstructionsLoaded`

它们的用途包括：

- 审计
- 自动验证
- 插件扩展
- 文档加载后的自动处理
- 任务完成后的自动通知

如果这些位点不提前定义，后面要做插件或自动验证时会大面积返工。

### 5.2.16 Background task / task framework 必须独立成层

Claude Code 值得学的不是“能起 subagent”，而是它把长任务当作一等对象来管理。

这类任务应具备：

- 任务 ID
- 前台 / 后台切换能力
- output file
- transcript
- 完成通知
- 终止 / 恢复 / 跟踪

对 Word-Cursor 尤其重要的任务包括：

- 长文档生成
- PPT 生成 / 重做
- 资料检索
- 批量格式化
- 深度 workspace exploration

建议在 runtime 内独立抽象：

- `TaskRegistry`
- `TaskRunner`
- `TaskNotificationCenter`
- `TaskOutputStore`
- `TaskResumeManager`

### 5.2.17 Session transcript / replay / resume 必须提前定义

`MessageStore` 只能解决会话内存态，不等于 transcript 系统。

必须前置明确：

- 会话如何落盘
- 工具调用如何记录到 transcript
- resume 如何恢复
- replay 如何回放
- subagent transcript 如何隔离

建议拆分：

- `SessionTranscriptStore`
- `AgentTranscriptStore`
- `ReplayStore`
- `ResumeLoader`

如果这一层不提前设计，后续将很难稳定支持：

- resume
- bg task 恢复
- subagent 隔离
- replay / telemetry 对齐

### 5.3 Prompt 分段化

主 prompt 只保留稳定、长期有效的规则。

变化频繁的内容必须下沉到：

- dynamic sections
- per-turn attachments
- tool descriptions
- skill descriptions

### 5.3.1 Context compaction / context collapse 必须前置设计

Claude Code 真正强的不只是 prompt，而是上下文控制。

对 Word-Cursor 这种长文档场景来说，这一层比普通 CLI 更关键。

建议在架构上提前定义：

- `auto compact`
- `reactive compact`
- `tool result budget`
- `context collapse`

职责划分建议：

- `auto compact`：在上下文接近阈值时自动触发摘要压缩
- `reactive compact`：provider 报上下文过长后触发救援
- `tool result budget`：控制单轮工具结果注入上限
- `context collapse`：把旧轮次的重型上下文投影成轻量视图，而不是简单截断

### 5.3.2 Tool result storage 必须独立成层

超长工具结果不能继续直接塞回对话文本。

Claude Code 的做法值得学：

- 超大结果持久化到文件
- 对模型只给摘要 + 引用
- 需要时再读取详细结果

对我们项目特别重要的场景包括：

- workspace 大目录总结
- 长文档结构上下文
- PPT 生成中间材料
- Excel 批量结果

所以必须单独设计：

- `ToolResultStore`
- result preview
- persisted result reference
- result budget 与清理策略

### 5.3.3 Prompt cache 稳定性工程必须作为主设计项

这是最容易被忽视、但会直接影响稳定性的部分。

Claude Code 非常重视：

- 静态 prompt 与动态 prompt 分离
- 动态 agent list / mcp instructions 走 attachment delta
- tool schema cache
- system prompt boundary

Word-Cursor 也必须明确：

- 静态 system prompt 不应频繁变化
- 动态工具列表、workspace 状态、memory 命中结果尽量走 attachment / delta
- tool schema 序列化应缓存
- system prompt 需要清晰的 static/dynamic boundary

### 5.3.4 `/init` 本质上是 Workspace Profile 构建能力

Claude Code 的 `/init` 值得学的不是命令名，而是背后的能力：

- 在正式执行前先建立工作区认知
- 对高价值文件做初步理解
- 形成可复用的 workspace profile

对 Word-Cursor 来说，这个能力比纯代码项目更重要，因为工作区里常常同时存在：

- `docx`
- `pdf`
- `pptx`
- `xlsx`
- `txt / md`
- 模板文件
- 参考材料 / 制式文件 / 会议纪要

因此 `/init` 不应只是“列文件”，而应被定义为一套只读工作流：

1. `workspace.list`
2. `workspace.read / workspace.summarize`
3. 按文件类型走专门 summarizer
4. 生成 `workspace_profile`
5. 可选写入 memory

理想输出应是双轨：

- `workspace_profile.json`
- 聊天里的简要总结

并加两条硬约束：

- `/init` 默认只读，不修改文件
- `/init` 结果必须结构化，不能只给散文式说明

### 5.4 UI Thin Layer

UI 负责展示，不负责智能体业务决策。

具体来说：

- `ChatPanel` 负责消息展示、输入、结果展示
- `WordEditor` 负责编辑器交互
- runtime 负责 tool loop 和 domain orchestration

### 5.5 Domain Adapter 化

Word、Excel、PPT、workspace、web、memory 都不应该直接暴露给聊天 UI，而应该通过 adapter 接入 tool packs。

### 5.6 Built-in Specialist Agents 必须定义职责边界

Claude Code 值得学的不是“有 subagent”，而是 agent 类型职责非常明确。

对 Word-Cursor 也应该写成硬规范：

- `doc-explore`
  - 只读
  - 负责大纲、结构、样式、工作区资料探索
- `doc-plan`
  - 只规划
  - 不直接改文档
- `doc-review`
  - 负责审校、对比、批注建议
  - 默认不落地 destructive 修改
- `verification`
  - 对抗式验证
  - 负责证明结果真实可用，不参与实现
- `ppt-builder`
  - 负责 PPT 生成/重构
  - 适合后台运行
- `excel-operator`
  - 负责表格批量处理和计算

必须明确哪些 agent：

- 只读
- 只规划
- 只验证
- 可写
- 可后台运行

## 6. 现有代码到目标结构的映射

### 6.1 当前核心模块

- [AIContext.tsx](/C:/Users/yangz/Desktop/Github/Word-Cursor/src/context/AIContext.tsx)
- [ChatPanel.tsx](/C:/Users/yangz/Desktop/Github/Word-Cursor/src/components/ChatPanel.tsx)
- [DocumentContext.tsx](/C:/Users/yangz/Desktop/Github/Word-Cursor/src/context/DocumentContext.tsx)
- [WordEditor.tsx](/C:/Users/yangz/Desktop/Github/Word-Cursor/src/components/WordEditor.tsx)
- [main.cjs](/C:/Users/yangz/Desktop/Github/Word-Cursor/electron/main.cjs)
- [preload.cjs](/C:/Users/yangz/Desktop/Github/Word-Cursor/electron/preload.cjs)

### 6.2 重构后的职责迁移

`AIContext` 预计拆分为：

- `AgentSessionEngine`
- `ModelGateway`
- `SystemPromptComposer`
- `ContextAssembler`
- `MessageStore`

`ChatPanel` 预计拆分为：

- `ChatPanel` 只保留 UI
- `ToolRegistry`
- `ToolExecutor`
- `ResultRenderer`
- `AttachmentManager`

`DocumentContext` 预计保留为文档域核心，但通过 adapter 暴露给 runtime：

- `DocumentAgentAdapter`
- `WordToolPack`

`electron/main.cjs` 预计拆分为多个 service + IPC register 文件。

## 7. 分阶段开发计划

## Phase 0：重构基线与兼容层

### 目标

建立新 runtime 的骨架，但不立即打断现有能力。

### 主要工作

- 新建 `src/agent/` 目录和基础模块骨架
- 定义统一 tool contract
- 定义统一 tool call / result / progress IR
- 定义 provider adapter 接口与 ModelGateway 能力矩阵
- 定义统一 skill contract
- 定义统一 attachment contract
- 定义 runtime 内部消息格式
- 定义 transcript / task / hook / result storage 骨架
- 建立“兼容层”，让旧的 `ChatPanel -> AIContext` 调用路径还能工作

### 输出物

- `src/agent/tools/contracts.ts`
- `src/agent/tools/results.ts`
- `src/agent/tools/ir.ts`
- `src/agent/core/MessageStore.ts`
- `src/agent/storage/SessionTranscriptStore.ts`
- `src/agent/storage/ToolResultStore.ts`
- `src/agent/tasks/TaskRegistry.ts`
- `src/agent/hooks/HookRegistry.ts`
- `src/agent/attachments/AttachmentTypes.ts`
- `src/agent/skills/SkillRegistry.ts`
- 旧逻辑与新逻辑之间的 adapter shim

### 验收标准

- 不改功能行为的前提下，项目仍可构建
- 新 runtime 目录结构成型
- tool / skill / attachment 的类型边界明确
- 新旧工具调用都能映射到统一内部 IR
- provider 能力矩阵与 adapter 边界明确
- transcript / task / hook / result storage 至少完成接口级骨架

### 影响文件

- [AIContext.tsx](/C:/Users/yangz/Desktop/Github/Word-Cursor/src/context/AIContext.tsx)
- [ChatPanel.tsx](/C:/Users/yangz/Desktop/Github/Word-Cursor/src/components/ChatPanel.tsx)

## Phase 1：Tool 系统抽象化

### 目标

把目前散落在 `ChatPanel` 内的大量工具逻辑抽离成标准工具包。

### 主要工作

把当前这些能力迁移为标准 tools：

- Word 文档编辑工具
- Workspace 文件工具
- PPT 工具
- Excel 工具
- Web 搜索工具
- Memory 工具

建议目录：

```text
src/agent/tools/packs/
  word/
  workspace/
  ppt/
  excel/
  web/
  memory/
```

### 重点改造

- 从 [ChatPanel.tsx](/C:/Users/yangz/Desktop/Github/Word-Cursor/src/components/ChatPanel.tsx) 中抽出各 tool handler
- 每个 tool 拥有单独定义文件
- 引入 registry 统一注册
- 引入 scheduler，为后续并发与顺序执行做准备
- 引入 executor，统一承接 parse / validate / schedule / execute / result write-back
- 要求所有新迁移 tool 默认接入统一 IR，而不是继续返回松散字符串参数
- 引入 deferred tools / tool search 规划
- 明确 tool lifecycle 状态机、pairing、synthetic result 与并发输出顺序

### 建议首批工具

- `word.replace_text`
- `word.insert_blocks`
- `word.delete_blocks`
- `word.apply_ops`
- `workspace.list`
- `workspace.read`
- `workspace.open`
- `ppt.create`
- `ppt.edit`
- `excel.read_sheet`
- `excel.update_range`
- `web.search`
- `memory.search`

### 验收标准

- `ChatPanel` 不再直接包含大段业务执行逻辑
- 所有现有主要工具都可从 registry 调起
- 新老路径在过渡期可以并存
- 至少首批迁移工具不再直接依赖 UI 层局部状态完成执行
- 工具执行链具备可观测的 parse / validate / execute / result 边界
- 常驻工具与长尾工具暴露策略明确
- `ToolExecutor` 的主入口开始围绕 IR 运转

### 影响文件

- [ChatPanel.tsx](/C:/Users/yangz/Desktop/Github/Word-Cursor/src/components/ChatPanel.tsx)
- [DocumentContext.tsx](/C:/Users/yangz/Desktop/Github/Word-Cursor/src/context/DocumentContext.tsx)
- [webSearch.ts](/C:/Users/yangz/Desktop/Github/Word-Cursor/src/utils/webSearch.ts)

## Phase 2：Agent Session Engine 抽离

### 目标

把当前 `AIContext` 里的会话循环、消息管理、tool loop、状态收敛成独立 engine。

### 主要工作

- 抽出 `AgentSessionEngine`
- 抽出 `AgentLoop`
- 抽出 `MessageStore`
- 抽出 `ModelGateway`
- 抽出 transcript / replay / resume
- 抽出 background task / task runner
- 把模型返回解析和 tool 调度从 React context 中移走

### 关键原则

- React context 只存 session state 与 UI binding
- 真正的 runtime 执行不再依赖 React 组件生命周期
- transcript / task / subagent 状态不再依赖 UI 生命周期

### 兼容策略

Phase 2 不强行切模型协议。

先兼容当前自定义 `[TOOL_CALL] ... [/TOOL_CALL]` 协议，再在后续 phase 里升级。

但需要补一条硬约束：

- Phase 2 虽然继续兼容旧协议，但 **AgentLoop / ModelGateway / ToolExecutor 内部必须统一以 IR 运转**
- 旧协议只能作为“输入适配层”，不能继续渗透到新 runtime 深处
- 从 Phase 2 开始，新实现不应再新增对旧文本协议的直接依赖

### 验收标准

- `AIContext` 显著瘦身
- 会话循环可在非 UI 层独立运行
- tool 解析、消息推进、结果回写都有清晰边界
- transcript / replay / resume 可以在 runtime 层独立工作
- background task 至少具备 ID、状态、输出文件和完成通知

### 影响文件

- [AIContext.tsx](/C:/Users/yangz/Desktop/Github/Word-Cursor/src/context/AIContext.tsx)
- [toolCallLogger.ts](/C:/Users/yangz/Desktop/Github/Word-Cursor/src/utils/toolCallLogger.ts)

## Phase 3：Prompt / Context / Attachment 体系重建

### 目标

把现在“prompt 巨石化”的方式，改成 Claude Code 风格的分段上下文工程。

### 主要工作

新增：

- `SystemPromptComposer`
- `ContextAssembler`
- `AttachmentManager`
- `PromptSections`
- `AutoCompact`
- `ReactiveCompact`
- `ContextCollapse`
- `ToolResultStore`

补充目标：

- 为 provider-native tool calling 预留 schema 注入和结果回填位点
- 把“工具描述”与“动态上下文”分开，避免工具 schema 频繁被上下文污染
- 为后续 prompt cache 稳定性做准备
- 为 `/init` / workspace profile 预留 attachment 与 memory 复用位点

把上下文拆成：

- 静态 system prompt
- 动态 system sections
- per-turn attachments
- tool descriptions
- skill descriptions

### 对 Word-Cursor 的 attachment 建议

- `current_document_summary`
- `document_structure_delta`
- `available_tools_delta`
- `workspace_context_delta`
- `workspace_profile`
- `relevant_memories`
- `ppt_edit_context`
- `excel_sheet_context`

### 重点改造

- 当前文档摘要不再每轮整块塞进 prompt
- DSL 结构、页面信息、当前选择区、工作区相关文件改为 attachment 化
- memory 从“统一预注入”改为“相关项注入”
- 明确 static/dynamic prompt boundary
- tool schema cache / prompt cache 稳定性成为正式工程目标
- tool result budget 与 result persistence 有正式策略
- `/init` 产出的 workspace profile 以 attachment / delta 形式复用，而不是反复重扫工作区

### 验收标准

- system prompt 结构化输出
- 动态信息可以按 attachment 单独生成
- prompt 体积和重复注入明显下降
- `workspace_profile` 可以独立生成、独立更新、独立注入

### 影响文件

- [AIContext.tsx](/C:/Users/yangz/Desktop/Github/Word-Cursor/src/context/AIContext.tsx)
- [docxAgentContext.ts](/C:/Users/yangz/Desktop/Github/Word-Cursor/src/utils/docxAgentContext.ts)
- [docxAgentContextWorker.ts](/C:/Users/yangz/Desktop/Github/Word-Cursor/src/workers/docxAgentContextWorker.ts)
- `src/memory/*`

## Phase 4：Word 域优先重构

### 目标

把 Word 域从“被聊天层直接调用”改成“通过 domain adapter + tool pack 接入 runtime”。

### 主要工作

- 新增 `DocumentAgentAdapter`
- 新增 `WordToolPack`
- 把 `DocumentContext` 暴露为稳定的 domain API，而不是让聊天层直接读写其内部状态

### 重点内容

围绕 Word 三态重构：

- HTML 编辑态
- DSL 结构态
- DOCX 导出态

建议 runtime 中明确区分：

- 文本编辑工具
- 结构编辑工具
- 格式/样式工具
- 审阅与 diff 工具

### 代表性工具

- `word.read_selection`
- `word.read_outline`
- `word.replace_via_dsl`
- `word.insert_via_dsl`
- `word.delete_via_dsl`
- `word.preview_ops`
- `word.apply_ops`
- `word.accept_change`
- `word.reject_change`

### 验收标准

- Word 工具不再直接散在 `ChatPanel`
- `DocumentContext` 对 runtime 只暴露稳定接口
- 文档编辑链路可被其他 UI 或 agent 重用

### 影响文件

- [DocumentContext.tsx](/C:/Users/yangz/Desktop/Github/Word-Cursor/src/context/DocumentContext.tsx)
- [WordEditor.tsx](/C:/Users/yangz/Desktop/Github/Word-Cursor/src/components/WordEditor.tsx)
- [docDsl.ts](/C:/Users/yangz/Desktop/Github/Word-Cursor/src/utils/docDsl.ts)
- [htmlToDsl.ts](/C:/Users/yangz/Desktop/Github/Word-Cursor/src/utils/htmlToDsl.ts)
- [docDslToDocx.ts](/C:/Users/yangz/Desktop/Github/Word-Cursor/src/utils/docDslToDocx.ts)

## Phase 5：Excel / PPT / Workspace / Search 工具包化

### 目标

把 Word 以外的业务域标准化为可插拔 tool packs。

### 主要工作

- `ExcelToolPack`
- `PptToolPack`
- `WorkspaceToolPack`
- `WebToolPack`
- `MemoryToolPack`
- 补充按文件类型 summarizer：
  - `DocxSummarizer`
  - `PdfSummarizer`
  - `PptxSummarizer`
  - `XlsxSummarizer`
  - `TextSummarizer`

### 设计要求

- 每个 pack 可独立注册
- 每个 pack 的 context builder 独立
- 每个 pack 的 UI 输出可独立渲染
- pack 之间通过 runtime 协调，不互相直接依赖 UI 组件
- `WorkspaceToolPack` 不只是读文件，还要能构建 `workspace_profile`

### 验收标准

- Excel / PPT / workspace / search 不再把业务执行塞进聊天组件
- 多域协作通过 runtime orchestration 完成
- 工作区初始化可以先建模再执行，不需要用户反复手动指出材料在哪

### 影响文件

- [ExcelPreview.tsx](/C:/Users/yangz/Desktop/Github/Word-Cursor/src/components/ExcelPreview.tsx)
- [PptPreviewHtml.tsx](/C:/Users/yangz/Desktop/Github/Word-Cursor/src/components/PptPreviewHtml.tsx)
- [ChatPanel.tsx](/C:/Users/yangz/Desktop/Github/Word-Cursor/src/components/ChatPanel.tsx)
- [main.cjs](/C:/Users/yangz/Desktop/Github/Word-Cursor/electron/main.cjs)

## Phase 6：Skills 工作流层

### 目标

把高频任务从“长 prompt 指令”升级为“可发现、可调用、可维护”的 skill。

### 适合 Word-Cursor 的首批 skills

- `rewrite-formal`
- `rewrite-report`
- `summarize-to-slides`
- `generate-meeting-minutes`
- `document-proofread`
- `format-normalization`
- `template-based-doc`
- `ppt-from-outline`
- `excel-cleanup`
- `init`
- `初始化`
- `项目理解`

### 主要工作

- 新增 `SkillRegistry`
- 新增 `SkillExecutor`
- 支持内置技能和工作区技能
- 支持 skill 对 tool 的约束与组合
- 将 `/init` 作为内置只读 skill：
  - 先扫描工作区
  - 再抽取高价值文件
  - 再生成 `workspace_profile`
  - 最后输出摘要和建议动作

### 验收标准

- 高频任务不再依赖模型每次临时推导工具链
- skill 可独立维护与测试
- `/init` skill 能稳定产出 `workspace_profile`，而不是每次临时理解目录

## Phase 7：Subagent 体系

### 目标

引入多 agent 编排能力，但聚焦文档场景，不照搬 CLI 型复杂度。

### 建议内建 subagents

- `doc-explore`
- `workspace-explore`
- `doc-editor`
- `ppt-builder`
- `excel-operator`
- `verification`

### 主要工作

- 新增 `SubagentManager`
- 定义 agent profiles
- 支持同步 subagent
- 支持后台 subagent
- 支持对子任务结果做摘要回传
- 支持 specialist agents 的只读 / 只规划 / 只验证规范
- 支持 subagent transcript 隔离
- 支持后台子任务通知、恢复、终止

### 对本项目的收益

- 主 agent 不必自己消化大量中间搜索结果
- 文档分析、资料收集、结构化编辑可拆并行
- 大文档和大工作区场景更稳

### 验收标准

- 主会话可启动子 agent
- 子 agent 可带不同工具集运行
- 子 agent 结果回流主会话
- specialist agents 职责边界清晰，不依赖 prompt 临场发挥

## Phase 8：Electron Backend 拆层

### 目标

降低 [main.cjs](/C:/Users/yangz/Desktop/Github/Word-Cursor/electron/main.cjs) 的单文件复杂度。

### 主要工作

- 按领域拆分 service
- IPC 注册独立化
- 前端只通过 preload 暴露稳定 API

### 建议拆分顺序

1. files
2. memory
3. web/search
4. excel
5. ppt
6. ai/model proxy
7. fonts

### 验收标准

- 主进程文件大幅缩小
- 新功能开发不再默认改 `main.cjs`

### 当前收尾状态

- 已按领域拆出 `files`
- 已按领域拆出 `memory`
- 已按领域拆出 `web/search`
- 已按领域拆出 `excel`
- 已按领域拆出 `fonts`
- 已按领域拆出 `ai/model proxy`
- 已按领域拆出 `ppt`
- [main.cjs](/C:/Users/yangz/Desktop/Github/Word-Cursor/electron/main.cjs) 现在主要负责依赖装配、窗口生命周期和 IPC 注册
- `preload` API 形状在本阶段保持稳定
- 当前剩余问题不再属于 Phase 8 计划内的切片拆分，而是后续独立清理项，例如 Document Builder 历史逻辑和现有构建 warning

## Phase 9：协议升级与旧系统下线

### 目标

在 runtime 稳定后，逐步淘汰旧路径。

### 主要工作

- 将旧的 `[TOOL_CALL]` 解析层包成兼容 adapter
- 逐步迁移到统一内部 tool call IR
- 新功能只允许走新 runtime
- 删除旧的 `ChatPanel` 内部业务逻辑残留
- 删除 `AIContext` 中已无必要的旧逻辑

补充说明：

这里的“协议升级”不应理解为“从这一阶段才开始设计稳定工具调用方式”。

更准确的定位应当是：

- Phase 0-2 已经完成 IR 与执行链收口
- Phase 3 以后新功能优先走更稳定的 schema / native tool calling 路线
- Phase 9 负责的是**下线旧文本协议主链路**，不是从零开始思考工具协议

### 验收标准

- 新旧链路切换完成
- 核心能力全部走新架构
- 旧逻辑仅保留必要兼容层，随后清理

### 当前开发起点

- 先不追求一次性删除全部 legacy 代码
- 第一批收口目标是把旧协议从“主链路核心”降级为“兼容 adapter”
- 主循环内部优先统一到 tool call IR
- provider-aware tool call parser 负责按 provider 能力解析响应
- legacy `[TOOL_CALL]` / XML / `<tool_use>` 继续保留，但只作为兼容输入层
- 请求链路已开始支持 provider-native `tools schema` 注入
- Electron ai proxy 已开始承接 native tool call 的结构化响应
- 主循环已开始为 native tool call 生成结构化 assistant/tool result 对话消息
- system prompt 组装已开始压制 legacy `[TOOL_CALL]` 指令，native tool calling 优先
- browser / 非 Electron 路径仍显式保留在 legacy fallback，避免半成品 native 链路误入
- native 模式下主 prompt 已开始裁掉 legacy `<available_tools>` 大段示例，减少协议冲突与上下文噪音

### 当前收尾状态

- provider-aware tool call parser 已进入主循环
- native `tools schema` 已进入 Electron 请求链路
- native tool call 的 assistant/tool result 对话消息已开始回填
- system prompt 已明确 native tool calling 优先
- legacy `[TOOL_CALL]` / XML / `<tool_use>` 仍保留，但定位已降级为兼容层
- 当前剩余工作已不再是主架构改造，而是运行态 smoke test 与少量 prompt/compat 清理

## 8. 每个 Phase 的交付要求

每个 phase 至少应包含：

- 架构变更说明
- 影响模块说明
- 迁移清单
- 手工验证步骤
- 回退策略

每个 phase 的最小验证：

- `npm run build`
- `npm run dev:electron`
- 对应域的手工 smoke test

## 9. 推荐实施顺序

严格建议按下面顺序推进：

1. Phase 0
2. Phase 1
3. Phase 2
4. Phase 3
5. Phase 4
6. Phase 5
7. Phase 6
8. Phase 7
9. Phase 8
10. Phase 9

不要一开始就做 subagent，不要一开始就做 Electron 全拆。

最先要做的是：

- runtime contract
- tool system
- session engine
- prompt/context/attachment

这是后面所有重构的地基。

## 10. 第一阶段建议直接开工的文件

如果按最优顺序开始，建议第一批实际动手的文件是：

- [AIContext.tsx](/C:/Users/yangz/Desktop/Github/Word-Cursor/src/context/AIContext.tsx)
- [ChatPanel.tsx](/C:/Users/yangz/Desktop/Github/Word-Cursor/src/components/ChatPanel.tsx)
- [DocumentContext.tsx](/C:/Users/yangz/Desktop/Github/Word-Cursor/src/context/DocumentContext.tsx)
- [main.cjs](/C:/Users/yangz/Desktop/Github/Word-Cursor/electron/main.cjs)
- [preload.cjs](/C:/Users/yangz/Desktop/Github/Word-Cursor/electron/preload.cjs)

但不是直接大改这几个文件本体，而是：

- 先新增 `src/agent/*`
- 再把旧逻辑一块一块迁出去

## 11. 最终目标状态

最终我们希望 Word-Cursor 的智能体层达到下面这个结构：

- `ChatPanel` 是 UI
- `AIContext` 是轻状态壳
- `AgentSessionEngine` 是会话核心
- `ToolRegistry + ToolExecutor` 是能力入口
- `SystemPromptComposer + AttachmentManager` 是上下文工程层
- `SkillRegistry` 是工作流层
- `SubagentManager` 是多 agent 层
- `DocumentAgentAdapter / ExcelToolPack / PptToolPack / WorkspaceToolPack` 是业务域接入层
- Electron 主进程退化为清晰的本地服务总线

这才是适合 Word-Cursor 长期演进的智能体架构。

## 12. 结论

本次重构不应被理解为“优化一下 prompt”或“把工具整理整理”。

这应该是一次明确的架构升级：

- 从聊天驱动业务
- 升级为 runtime 驱动业务

Claude Code 给我们的核心启发是：

- prompt 不是核心，runtime 才是核心
- tool 不是附属能力，tool system 本身就是智能体骨架
- memory、skills、attachments、subagents 都应该是 runtime 一等公民

Word-Cursor 后续所有 AI 文档编辑、生成、协作能力，都应建立在这次重构之后的新 runtime 之上。
## 13. Word Tool Convergence

## 13. Word Tool Convergence

- Model-visible Word tools should converge to:
- `word.read`
- `word.create`
- `word.edit`
- `word.format`
- `word.resolve_change`
- `word.chart`
- Legacy Word tool ids remain compatibility aliases only:
- `replace`
- `review`
- `insert`
- `delete`
- `create`
- `create_from_template`
- `copy_template`
- `word_edit_ops`
- `word.preview_ops`
- `word.apply_ops`
- `word_chart`
- `word.edit` must be DSL-first, with block-scoped edits preferred over raw text mutation.
- Native tool schemas, skills, subagent boundaries, and future prompt cleanup should all reference canonical ids first.
