# Word-Cursor 官网开发文档

## 一、项目概述

**Word-Cursor** 是一款 AI 驱动的智能办公文档编辑器，把 Cursor 级别的"对话式编辑 + 工具调用 + 可审阅变更"带进 Word / Excel / PowerPoint。

- 版本：v1.0.1
- GitHub：https://github.com/Yangyangxxx/Word-Cursor
- 域名：http://yangyzx.com/
- 标语：对话即编辑，自然语言驱动办公

---

## 二、服务器信息

- **IP**: 8.141.124.194
- **系统**: Ubuntu 24.04 (阿里云 ECS, 2核2GB)
- **SSH**: `ssh -i ~/.ssh/id_ed25519 root@8.141.124.194`
- **面板**: 宝塔面板
- **Web 服务**: nginx
- **网站根目录**: `/www/wwwroot/yangyzx.com/`
- **nginx vhost 配置**: `/www/server/panel/vhost/nginx/html_yangyzx.com.conf`
- **无 Node.js**，纯静态 HTML/CSS/JS 部署
- **部署方式**: SCP 上传文件到网站根目录即可

### 当前文件结构
```
/www/wwwroot/yangyzx.com/
├── index.html          # 主页面（当前版本有 bug，需重做）
├── favicon.svg         # 网站图标
├── assets/             # 项目截图
│   ├── 主界面1.png      (323KB)
│   ├── word界面.png     (670KB)
│   ├── excel界面.png    (712KB)
│   ├── ppt界面.png      (2.1MB)
│   ├── ppt作品展示1.png  (1.3MB)
│   ├── ppt作品展示2.png  (4.3MB)
│   └── ppt作品展示3.png  (1.6MB)
├── 404.html
├── .htaccess
└── .user.ini
```

### 部署命令
```bash
# 上传文件
scp -i ~/.ssh/id_ed25519 index.html root@8.141.124.194:/www/wwwroot/yangyzx.com/
scp -i ~/.ssh/id_ed25519 -r assets/ root@8.141.124.194:/www/wwwroot/yangyzx.com/

# 修复权限 + 重载 nginx
ssh -i ~/.ssh/id_ed25519 root@8.141.124.194 "chown -R www-data:www-data /www/wwwroot/yangyzx.com/ && nginx -s reload"
```

---

## 三、项目核心特色（官网需要展示的内容）

### 3.1 核心卖点
1. **对话式编辑** — 用自然语言描述修改意图，AI 自动执行精准的文档编辑
2. **可审阅变更** — 每次 AI 修改以 Diff 形式呈现，逐条接受/拒绝
3. **多格式支持** — Word 深度编辑 + Excel 智能分析 + PPT 一键生成

### 3.2 Word 编辑功能
- 基于 Tiptap (ProseMirror) 的富文本编辑器
- A4 打印预览布局
- Ctrl+K 选中即编辑
- 文档审查（/审查）：语法、逻辑、措辞、错别字、格式分类标记
- 格式刷、样式管理、页眉页脚、水印、目录生成
- 模板填充与公文格式化
- 支持 .docx 导入导出

### 3.3 Excel 功能
- .xlsx/.xls 工作簿预览
- AI 辅助数据分析
- 单元格读写、公式批量设置
- 条件格式、数据验证、排序筛选
- 柱状图/饼图/折线图一键生成
- 跨工作表公式

### 3.4 PPT 生成
- 端到端生成：大纲 → 视觉提示词(Gemini) → 图像生成(DashScope) → PPTX 打包
- 6 种视觉风格：玻璃质感渐变、瑞士国际排版、午夜专业、新中式禅意等
- 整页重做或局部微调
- 缩略图预览、全屏放映、导出 .pptx

### 3.5 其他功能
- **Web 搜索**: 内置 Brave Search，对话中直接调研引用
- **AI 智能补全**: Tab 键触发低延迟续写
- **快捷命令**: /审查 /润色 /精简 /翻译 /格式化 /编号 /公文 /会议纪要 /总结 /校对
- **记忆系统**: 跨会话记忆，自动积累用户偏好

### 3.6 技术栈
- React 18 + TypeScript 5 + Vite 5 + TailwindCSS 3
- Electron 33 (桌面端)
- Tiptap (ProseMirror) 富文本编辑
- docx / mammoth / exceljs / pptxgenjs / jszip (文档 I/O)
- OpenAI 兼容 API + 流式 + 工具调用
- DashScope (阿里云) 图像生成
- Brave Search MCP

### 3.7 对比传统办公
| 功能 | Word-Cursor | 传统 Office |
|------|------------|-------------|
| AI 编辑 | ✅ 原生内置 | ❌ 需插件 |
| 对话交互 | ✅ 自然语言 | ❌ 菜单操作 |
| 智能补全 | ✅ Tab 续写 | ❌ 无 |
| 联网搜索 | ✅ 内置搜索 | ❌ 需切换 |
| PPT 生成 | ✅ AI 设计 | ❌ 仅模板 |
| 变更审阅 | ✅ Diff 对比 | ⚠️ 修订模式 |

---

## 四、可用的视觉素材

所有截图已上传到服务器 `/www/wwwroot/yangyzx.com/assets/`：

1. **主界面1.png** — Word-Cursor 主界面全貌
2. **word界面.png** — Word 文档编辑界面（左侧编辑器 + 右侧 AI 对话）
3. **excel界面.png** — Excel 表格预览界面
4. **ppt界面.png** — PPT 演示文稿界面（缩略图 + 预览）
5. **ppt作品展示1.png** — AI 生成的 PPT 作品（玻璃质感风格）
6. **ppt作品展示2.png** — AI 生成的 PPT 作品（多页展示）
7. **ppt作品展示3.png** — AI 生成的 PPT 作品（另一风格）
8. **favicon.svg** — 项目 Logo/图标

本地源文件在 `C:\Users\29350\Desktop\Word-Cursor\assets\` 和 `C:\Users\29350\Desktop\Word-Cursor\public\favicon.svg`

---

## 五、当前版本问题 & 改进方向

### 已知 Bug
1. **PPT 作品展示区**: gallery 自动滚动用了 setInterval 每 30ms 滚动 1px，导致画面一直闪烁跳动，需要改为平滑的轮播或静态网格展示

### 需要改进
1. **动画效果太单一** — 当前只有基础的 fadeIn + translateY，需要更丰富的动画：
   - Hero 区可以做更炫的粒子/光线效果
   - 功能展示区可以做交互式的 hover 动画
   - 滚动视差效果
   - SVG 路径描边动画
   - 数字滚动计数器
   - 鼠标跟随光效

2. **设计缺乏创意** — 当前是标准的深色科技风模板，可以考虑：
   - 更有品牌感的配色方案
   - 3D 效果或等距插画
   - 交互式 Demo 演示区（模拟对话编辑过程）
   - 更精致的卡片设计（毛玻璃 + 光边效果）
   - 动态渐变背景
   - 打字机效果展示 AI 对话过程

3. **图片展示方式** — 截图直接放太单调，可以：
   - 放在模拟的电脑/笔记本屏幕框内
   - 加阴影和倾斜透视效果
   - 做成可交互的 Tab 切换

### 技术建议
- 继续使用 anime.js（CDN: `https://cdn.jsdelivr.net/npm/animejs@3.2.2/lib/anime.min.js`）
- 可以额外引入 Three.js 做 3D 背景效果
- 保持纯静态，不需要构建工具
- 注意中文文件名在 URL 中需要编码（nginx 已支持）
- 大图片(ppt作品展示2.png 4.3MB)建议压缩或用 loading="lazy"

---

## 六、页面结构参考（可重新设计）

建议保留的区块：
1. **导航栏** — 固定顶部，Logo + 锚点链接 + 下载按钮
2. **Hero 首屏** — 项目名称、标语、CTA 按钮、炫酷背景动画
3. **核心亮点** — 3 个核心卖点卡片
4. **功能展示** — Word / Excel / PPT 三大功能详细介绍 + 截图
5. **PPT 作品画廊** — 展示 AI 生成的 PPT 作品
6. **快捷命令** — 展示 / 命令列表
7. **技术栈** — 使用的技术图标
8. **对比表格** — vs 传统办公
9. **Footer** — 下载链接、GitHub、联系方式、版权

---

## 七、联系信息
- 邮箱: yangyzx@qq.com
- GitHub: https://github.com/Yangyangxxx/Word-Cursor
- Release 下载: https://github.com/Yangyangxxx/Word-Cursor/releases
