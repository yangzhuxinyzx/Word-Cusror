import { useMemo } from 'react'
import type { Editor } from '@tiptap/react'
import type { PageSetup } from '../types'
import {
  Bold,
  Italic,
  Underline as UnderlineIcon,
  Strikethrough,
  AlignLeft,
  AlignCenter,
  AlignRight,
  AlignJustify,
  List,
  ListOrdered,
  Undo,
  Redo,
  Copy,
  Clipboard,
  Table as TableIcon,
  Image as ImageIcon,
  Link as LinkIcon,
  Maximize2,
  Check,
  X,
  Eye,
  ZoomIn,
  ZoomOut,
  MessageSquarePlus,
  ChevronUp,
  ChevronDown,
  SpellCheck,
  FileSearch,
  Type,
} from 'lucide-react'

export type WordRibbonTab = 'home' | 'insert' | 'layout' | 'references' | 'review' | 'view'

type WordRibbonProps = {
  tab: WordRibbonTab
  setTab: (t: WordRibbonTab) => void

  editor: Editor

  // font
  currentFontFamily: string
  setCurrentFontFamily: (v: string) => void
  fontFamilyOptions: string[]
  currentFontSize: string
  setCurrentFontSize: (v: string) => void
  applyFontFamily: (v: string) => void
  applyFontSize: (v: string) => void

  // view
  viewMode: 'print' | 'web'
  setViewMode: (v: 'print' | 'web') => void
  printInteractionMode?: 'stable' | 'table-edit'
  canTogglePrintTableEdit?: boolean
  onTogglePrintTableEdit?: () => void

  // zoom
  zoomLevel: number
  setZoomLevel: (n: number) => void

  // page setup
  pageSetup: PageSetup
  setPageSetup: (setup: Partial<PageSetup>) => void

  // review
  pendingChangesTotal: number
  onOpenRevisionPanel: () => void
  acceptAllChanges: () => void
  rejectAllChanges: () => void
  onAddComment?: () => void
  onOpenCommentPanel?: () => void
  onNavigateChange?: (direction: 'prev' | 'next') => void
  onAIReview?: () => void
  docStats?: { words: number; chars: number }

  // editor mode (optional, for App integrate)
  editorMode?: 'tiptap' | 'onlyoffice'
  setEditorMode?: (m: 'tiptap' | 'onlyoffice') => void
}

const tabs: Array<{ id: WordRibbonTab; label: string }> = [
  { id: 'home', label: '开始' },
  { id: 'insert', label: '插入' },
  { id: 'layout', label: '布局' },
  { id: 'references', label: '引用' },
  { id: 'review', label: '审阅' },
  { id: 'view', label: '视图' },
]

function RibbonButton({
  title,
  onClick,
  disabled,
  active,
  children,
  className = '',
}: {
  title: string
  onClick: () => void
  disabled?: boolean
  active?: boolean
  children: React.ReactNode
  className?: string
}) {
  return (
    <button
      type="button"
      title={title}
      onClick={onClick}
      disabled={disabled}
      className={[
        'relative flex items-center justify-center p-1.5 rounded-md transition-all duration-150',
        disabled ? 'opacity-40 cursor-not-allowed' : 'cursor-pointer active:scale-95',
        active
          ? 'bg-accent text-white shadow-sm shadow-accent/20'
          : 'text-text-secondary hover:text-text hover:bg-black/5 dark:hover:bg-white/10',
        className,
      ].join(' ')}
    >
      {children}
    </button>
  )
}

function RibbonGroup({
  label,
  children,
}: {
  label: string
  children: React.ReactNode
}) {
  return (
    <div className="flex flex-col items-center px-2 border-r border-black/10 dark:border-white/10 last:border-0 h-full justify-center">
      <div className="flex items-center gap-1 flex-wrap justify-center">{children}</div>
      <div className="text-[10px] text-text-dim mt-1 select-none tracking-tight">{label}</div>
    </div>
  )
}

function DisabledStub({ label }: { label: string }) {
  return (
    <div className="px-3 py-2 text-xs text-text-dim">
      {label}：暂未实现（先保留 UI 位置）
    </div>
  )
}

export default function WordRibbon(props: WordRibbonProps) {
  const {
    tab,
    setTab,
    editor,
    currentFontFamily,
    setCurrentFontFamily,
    fontFamilyOptions,
    currentFontSize,
    setCurrentFontSize,
    applyFontFamily,
    applyFontSize,
    viewMode,
    setViewMode,
    printInteractionMode,
    canTogglePrintTableEdit,
    onTogglePrintTableEdit,
    zoomLevel,
    setZoomLevel,
    pageSetup,
    setPageSetup,
    pendingChangesTotal,
    onOpenRevisionPanel,
    acceptAllChanges,
    rejectAllChanges,
    onAddComment,
    onOpenCommentPanel,
    onNavigateChange,
    onAIReview,
    docStats,
    editorMode,
    setEditorMode,
  } = props

  const isActive = useMemo(() => {
    return {
      bold: editor.isActive('bold'),
      italic: editor.isActive('italic'),
      underline: editor.isActive('underline'),
      strike: editor.isActive('strike'),
      bulletList: editor.isActive('bulletList'),
      orderedList: editor.isActive('orderedList'),
      blockquote: editor.isActive('blockquote'),
      alignLeft: editor.isActive({ textAlign: 'left' }),
      alignCenter: editor.isActive({ textAlign: 'center' }),
      alignRight: editor.isActive({ textAlign: 'right' }),
      alignJustify: editor.isActive({ textAlign: 'justify' }),
      link: editor.isActive('link'),
    }
  }, [editor, editor.state])

  return (
    <div className="w-full select-none">
      {/* Tabs row */}
      <div className="flex items-center gap-1 px-3 py-1.5 glass border-b border-border">
        <div className="flex items-center gap-1">
          {tabs.map((t) => (
            <button
              key={t.id}
              type="button"
              onClick={() => setTab(t.id)}
              className={[
                'px-3 py-1.5 text-xs font-semibold rounded-lg transition-colors',
                tab === t.id
                  ? 'bg-black/10 dark:bg-white/10 text-text border border-black/10 dark:border-white/10'
                  : 'text-text-muted hover:text-text hover:bg-black/5 dark:hover:bg-white/5 border border-transparent',
              ].join(' ')}
            >
              {t.label}
            </button>
          ))}
        </div>

        {/* Right side (review + editor mode quicks) */}
        <div className="ml-auto flex items-center gap-2">
          {pendingChangesTotal > 0 && (
            <button
              type="button"
              onClick={onOpenRevisionPanel}
              className="flex items-center gap-1.5 px-2.5 py-1.5 text-xs rounded-lg bg-accent/12 border border-accent/25 text-accent hover:bg-accent/18 transition-colors"
              title="打开修订面板"
            >
              <Eye className="w-3.5 h-3.5" />
              待审阅 ×{pendingChangesTotal}
            </button>
          )}

          {setEditorMode && (
            <div className="flex items-center gap-1 bg-black/5 dark:bg-white/5 border border-black/10 dark:border-white/10 rounded-lg p-1">
              <button
                type="button"
                onClick={() => setEditorMode('tiptap')}
                className={[
                  'px-2.5 py-1 text-[11px] rounded-md transition-colors',
                  editorMode === 'tiptap'
                    ? 'bg-accent/14 text-accent border border-accent/25'
                    : 'text-text-muted hover:text-text hover:bg-white/5 border border-transparent',
                ].join(' ')}
                title="切换到内置编辑器（Word Ribbon）"
              >
                内置
              </button>
              <button
                type="button"
                onClick={() => setEditorMode('onlyoffice')}
                className={[
                  'px-2.5 py-1 text-[11px] rounded-md transition-colors',
                  editorMode === 'onlyoffice'
                    ? 'bg-accent/14 text-accent border border-accent/25'
                    : 'text-text-muted hover:text-text hover:bg-white/5 border border-transparent',
                ].join(' ')}
                title="切换到 ONLYOFFICE"
              >
                ONLYOFFICE
              </button>
            </div>
          )}
        </div>
      </div>

      {/* Ribbon content */}
      <div className="w-full glass border-b border-border shadow-sm overflow-x-auto custom-scrollbar">
        <div className="flex items-stretch px-2 py-2 min-w-max h-[108px]">
          {tab === 'home' && (
            <>
              <RibbonGroup label="剪贴板">
                <RibbonButton
                  title="复制"
                  onClick={() => {
                    const { from, to } = editor.state.selection
                    const text = editor.state.doc.textBetween(from, to, '\n')
                    if (text) navigator.clipboard.writeText(text).catch(() => {})
                  }}
                  active={false}
                >
                  <Copy className="w-4 h-4" />
                </RibbonButton>
                <RibbonButton
                  title="粘贴"
                  onClick={() =>
                    navigator.clipboard
                      .readText()
                      .then((t) => t && editor.commands.insertContent(t))
                      .catch(() => {})
                  }
                >
                  <Clipboard className="w-4 h-4" />
                </RibbonButton>
              </RibbonGroup>

              <RibbonGroup label="字体">
                <div className="flex flex-col gap-2">
                  <div className="flex items-center gap-1">
                    <input
                      list="word-ribbon-font-family-list"
                      value={currentFontFamily}
                      onChange={(e) => {
                        const v = e.target.value
                        setCurrentFontFamily(v)
                        applyFontFamily(v)
                      }}
                      className="w-32 h-7 px-2 text-xs bg-black/5 dark:bg-white/5 border border-black/10 dark:border-white/10 rounded-md text-text focus:border-accent focus:ring-1 focus:ring-accent outline-none"
                      placeholder="字体"
                    />
                    <datalist id="word-ribbon-font-family-list">
                      {fontFamilyOptions.map((f) => (
                        <option key={f} value={f} />
                      ))}
                    </datalist>

                    <select
                      value={currentFontSize}
                      onChange={(e) => {
                        const v = e.target.value
                        setCurrentFontSize(v)
                        applyFontSize(v)
                      }}
                      className="w-20 h-7 px-1 text-xs bg-black/5 dark:bg-white/5 border border-black/10 dark:border-white/10 rounded-md text-text focus:border-accent focus:ring-1 focus:ring-accent outline-none"
                    >
                      <option value="9pt">小五</option>
                      <option value="10.5pt">五号</option>
                      <option value="12pt">小四</option>
                      <option value="14pt">四号</option>
                      <option value="15pt">小三</option>
                      <option value="16pt">三号</option>
                      <option value="18pt">小二</option>
                      <option value="22pt">二号</option>
                      <option value="24pt">小一</option>
                      <option value="26pt">一号</option>
                      <option value="36pt">小初</option>
                      <option value="42pt">初号</option>
                    </select>

                    <RibbonButton
                      title="增大字号"
                      onClick={() => {
                        const sizes = ['9pt','10.5pt','12pt','14pt','15pt','16pt','18pt','22pt','24pt','26pt','36pt','42pt']
                        const idx = sizes.indexOf(currentFontSize)
                        const next = idx >= 0 ? sizes[Math.min(sizes.length - 1, idx + 1)] : '14pt'
                        setCurrentFontSize(next)
                        applyFontSize(next)
                      }}
                    >
                      <ZoomIn className="w-3.5 h-3.5" />
                    </RibbonButton>
                    <RibbonButton
                      title="减小字号"
                      onClick={() => {
                        const sizes = ['9pt','10.5pt','12pt','14pt','15pt','16pt','18pt','22pt','24pt','26pt','36pt','42pt']
                        const idx = sizes.indexOf(currentFontSize)
                        const prev = idx >= 0 ? sizes[Math.max(0, idx - 1)] : '12pt'
                        setCurrentFontSize(prev)
                        applyFontSize(prev)
                      }}
                    >
                      <ZoomOut className="w-3.5 h-3.5" />
                    </RibbonButton>
                  </div>

                  <div className="flex items-center gap-0.5">
                    <RibbonButton title="粗体" onClick={() => editor.chain().focus().toggleBold().run()} active={isActive.bold}>
                      <Bold className="w-4 h-4" />
                    </RibbonButton>
                    <RibbonButton title="斜体" onClick={() => editor.chain().focus().toggleItalic().run()} active={isActive.italic}>
                      <Italic className="w-4 h-4" />
                    </RibbonButton>
                    <RibbonButton title="下划线" onClick={() => editor.chain().focus().toggleUnderline().run()} active={isActive.underline}>
                      <UnderlineIcon className="w-4 h-4" />
                    </RibbonButton>
                    <RibbonButton title="删除线" onClick={() => editor.chain().focus().toggleStrike().run()} active={isActive.strike}>
                      <Strikethrough className="w-4 h-4" />
                    </RibbonButton>
                  </div>
                </div>
              </RibbonGroup>

              <RibbonGroup label="段落">
                <div className="flex flex-col gap-2">
                  <div className="flex items-center gap-0.5">
                    <RibbonButton title="无序列表" onClick={() => editor.chain().focus().toggleBulletList().run()} active={isActive.bulletList}>
                      <List className="w-4 h-4" />
                    </RibbonButton>
                    <RibbonButton title="有序列表" onClick={() => editor.chain().focus().toggleOrderedList().run()} active={isActive.orderedList}>
                      <ListOrdered className="w-4 h-4" />
                    </RibbonButton>
                    <div className="w-px h-4 bg-black/10 dark:bg-white/10 mx-1" />
                    <RibbonButton title="左对齐" onClick={() => editor.chain().focus().setTextAlign('left').run()} active={isActive.alignLeft}>
                      <AlignLeft className="w-4 h-4" />
                    </RibbonButton>
                    <RibbonButton title="居中" onClick={() => editor.chain().focus().setTextAlign('center').run()} active={isActive.alignCenter}>
                      <AlignCenter className="w-4 h-4" />
                    </RibbonButton>
                    <RibbonButton title="右对齐" onClick={() => editor.chain().focus().setTextAlign('right').run()} active={isActive.alignRight}>
                      <AlignRight className="w-4 h-4" />
                    </RibbonButton>
                    <RibbonButton title="两端对齐" onClick={() => editor.chain().focus().setTextAlign('justify').run()} active={isActive.alignJustify}>
                      <AlignJustify className="w-4 h-4" />
                    </RibbonButton>
                  </div>
                </div>
              </RibbonGroup>

              <RibbonGroup label="样式">
                <div className="flex items-center gap-1 bg-black/5 dark:bg-white/5 p-1 rounded-lg border border-black/10 dark:border-white/10">
                  <button
                    type="button"
                    onClick={() => editor.chain().focus().setParagraph().run()}
                    className={[
                      'px-3 py-1.5 rounded-md text-xs transition-all',
                      editor.isActive('paragraph')
                        ? 'bg-white dark:bg-white/20 shadow-sm text-accent font-semibold'
                        : 'text-text-muted hover:text-text hover:bg-white/5',
                    ].join(' ')}
                  >
                    正文
                  </button>
                  <button
                    type="button"
                    onClick={() => editor.chain().focus().toggleHeading({ level: 1 }).run()}
                    className={[
                      'px-3 py-1.5 rounded-md text-xs transition-all',
                      editor.isActive('heading', { level: 1 })
                        ? 'bg-white dark:bg-white/20 shadow-sm text-accent font-bold'
                        : 'text-text-muted hover:text-text hover:bg-white/5',
                    ].join(' ')}
                  >
                    标题 1
                  </button>
                  <button
                    type="button"
                    onClick={() => editor.chain().focus().toggleHeading({ level: 2 }).run()}
                    className={[
                      'px-3 py-1.5 rounded-md text-xs transition-all',
                      editor.isActive('heading', { level: 2 })
                        ? 'bg-white dark:bg-white/20 shadow-sm text-accent font-semibold'
                        : 'text-text-muted hover:text-text hover:bg-white/5',
                    ].join(' ')}
                  >
                    标题 2
                  </button>
                  <button
                    type="button"
                    onClick={() => editor.chain().focus().toggleHeading({ level: 3 }).run()}
                    className={[
                      'px-3 py-1.5 rounded-md text-xs transition-all',
                      editor.isActive('heading', { level: 3 })
                        ? 'bg-white dark:bg-white/20 shadow-sm text-accent font-medium'
                        : 'text-text-muted hover:text-text hover:bg-white/5',
                    ].join(' ')}
                  >
                    标题 3
                  </button>
                </div>
              </RibbonGroup>

              <RibbonGroup label="编辑">
                <div className="flex flex-col gap-1">
                  <RibbonButton title="撤销" onClick={() => editor.chain().focus().undo().run()} disabled={!editor.can().undo()}>
                    <Undo className="w-4 h-4" />
                  </RibbonButton>
                  <RibbonButton title="重做" onClick={() => editor.chain().focus().redo().run()} disabled={!editor.can().redo()}>
                    <Redo className="w-4 h-4" />
                  </RibbonButton>
                </div>
              </RibbonGroup>
            </>
          )}

          {tab === 'insert' && (
            <>
              <RibbonGroup label="插入">
                <RibbonButton
                  title="表格"
                  onClick={() => editor.chain().focus().insertTable({ rows: 3, cols: 3, withHeaderRow: true }).run()}
                >
                  <TableIcon className="w-4 h-4" />
                </RibbonButton>
                <RibbonButton
                  title="图片（URL）"
                  onClick={() => {
                    const url = window.prompt('输入图片链接:')
                    if (url) editor.chain().focus().setImage({ src: url }).run()
                  }}
                >
                  <ImageIcon className="w-4 h-4" />
                </RibbonButton>
                <RibbonButton
                  title="链接"
                  onClick={() => {
                    const url = window.prompt('输入链接地址:')
                    if (url) editor.chain().focus().setLink({ href: url }).run()
                  }}
                  active={isActive.link}
                >
                  <LinkIcon className="w-4 h-4" />
                </RibbonButton>
              </RibbonGroup>
              <DisabledStub label="页眉页脚/公式/文本框" />
            </>
          )}

          {tab === 'layout' && (
            <>
              <RibbonGroup label="页面设置">
                <div className="flex flex-col gap-2">
                  <div className="flex items-center gap-1">
                    <button
                      type="button"
                      onClick={() => setPageSetup({ orientation: 'portrait' })}
                      className={[
                        'px-2.5 py-1 text-xs rounded-md border transition-colors',
                        pageSetup.orientation === 'portrait'
                          ? 'bg-accent/12 border-accent/25 text-accent'
                          : 'bg-black/5 dark:bg-white/5 border-black/10 dark:border-white/10 text-text-muted hover:text-text',
                      ].join(' ')}
                      title="纵向"
                    >
                      纵向
                    </button>
                    <button
                      type="button"
                      onClick={() => setPageSetup({ orientation: 'landscape' })}
                      className={[
                        'px-2.5 py-1 text-xs rounded-md border transition-colors',
                        pageSetup.orientation === 'landscape'
                          ? 'bg-accent/12 border-accent/25 text-accent'
                          : 'bg-black/5 dark:bg-white/5 border-black/10 dark:border-white/10 text-text-muted hover:text-text',
                      ].join(' ')}
                      title="横向"
                    >
                      横向
                    </button>
                  </div>
                  <div className="flex items-center gap-1">
                    <button
                      type="button"
                      onClick={() =>
                        setPageSetup({
                          margins: { top: '2.54cm', bottom: '2.54cm', left: '3.17cm', right: '3.17cm' },
                        })
                      }
                      className="px-2.5 py-1 text-xs rounded-md bg-black/5 dark:bg-white/5 border border-black/10 dark:border-white/10 text-text-muted hover:text-text"
                      title="普通页边距"
                    >
                      普通
                    </button>
                    <button
                      type="button"
                      onClick={() =>
                        setPageSetup({
                          margins: { top: '1.27cm', bottom: '1.27cm', left: '1.27cm', right: '1.27cm' },
                        })
                      }
                      className="px-2.5 py-1 text-xs rounded-md bg-black/5 dark:bg-white/5 border border-black/10 dark:border-white/10 text-text-muted hover:text-text"
                      title="窄页边距"
                    >
                      窄
                    </button>
                  </div>
                </div>
              </RibbonGroup>
              <DisabledStub label="分栏/水印/缩进" />
            </>
          )}

          {tab === 'references' && <DisabledStub label="引用（目录/脚注/引用）" />}

          {tab === 'review' && (
            <>
              <RibbonGroup label="校对">
                <RibbonButton title="AI 审查文档" onClick={() => onAIReview?.()}>
                  <SpellCheck className="w-4 h-4" />
                </RibbonButton>
                <RibbonButton title="字数统计" onClick={() => {}} disabled>
                  <Type className="w-4 h-4" />
                </RibbonButton>
                {docStats && (
                  <div className="text-[10px] text-text-dim leading-tight px-1">
                    <div>字 {docStats.chars}</div>
                    <div>词 {docStats.words}</div>
                  </div>
                )}
              </RibbonGroup>
              <RibbonGroup label="批注">
                <RibbonButton title="添加批注" onClick={() => onAddComment?.()}>
                  <MessageSquarePlus className="w-4 h-4" />
                </RibbonButton>
                <RibbonButton title="批注面板" onClick={() => onOpenCommentPanel?.()}>
                  <FileSearch className="w-4 h-4" />
                </RibbonButton>
              </RibbonGroup>
              <RibbonGroup label="修订">
                <RibbonButton title="打开修订面板" onClick={onOpenRevisionPanel}>
                  <Eye className="w-4 h-4" />
                </RibbonButton>
                <RibbonButton title="全部接受" onClick={acceptAllChanges} disabled={pendingChangesTotal === 0}>
                  <Check className="w-4 h-4" />
                </RibbonButton>
                <RibbonButton title="全部拒绝" onClick={rejectAllChanges} disabled={pendingChangesTotal === 0}>
                  <X className="w-4 h-4" />
                </RibbonButton>
              </RibbonGroup>
              <RibbonGroup label="更改">
                <RibbonButton title="上一处修订" onClick={() => onNavigateChange?.('prev')} disabled={pendingChangesTotal === 0}>
                  <ChevronUp className="w-4 h-4" />
                </RibbonButton>
                <RibbonButton title="下一处修订" onClick={() => onNavigateChange?.('next')} disabled={pendingChangesTotal === 0}>
                  <ChevronDown className="w-4 h-4" />
                </RibbonButton>
              </RibbonGroup>
            </>
          )}

          {tab === 'view' && (
            <>
              <RibbonGroup label="视图">
                <button
                  type="button"
                  onClick={() => setViewMode('print')}
                  className={[
                    'px-2.5 py-1 text-xs rounded-md border transition-colors',
                    viewMode === 'print'
                      ? 'bg-accent/12 border-accent/25 text-accent'
                      : 'bg-black/5 dark:bg-white/5 border-black/10 dark:border-white/10 text-text-muted hover:text-text',
                  ].join(' ')}
                  title="打印布局"
                >
                  打印布局
                </button>
                <button
                  type="button"
                  onClick={() => setViewMode('web')}
                  className={[
                    'px-2.5 py-1 text-xs rounded-md border transition-colors',
                    viewMode === 'web'
                      ? 'bg-accent/12 border-accent/25 text-accent'
                      : 'bg-black/5 dark:bg-white/5 border-black/10 dark:border-white/10 text-text-muted hover:text-text',
                  ].join(' ')}
                  title="网页布局"
                >
                  网页布局
                </button>
                {viewMode === 'print' && onTogglePrintTableEdit && (
                  <button
                    type="button"
                    onClick={onTogglePrintTableEdit}
                    disabled={!canTogglePrintTableEdit}
                    className={[
                      'px-2.5 py-1 text-xs rounded-md border transition-colors disabled:opacity-40 disabled:cursor-not-allowed',
                      printInteractionMode === 'table-edit'
                        ? 'bg-accent/12 border-accent/25 text-accent'
                        : 'bg-black/5 dark:bg-white/5 border-black/10 dark:border-white/10 text-text-muted hover:text-text',
                    ].join(' ')}
                    title={printInteractionMode === 'table-edit' ? '退出表格编辑' : '进入表格编辑'}
                  >
                    {printInteractionMode === 'table-edit' ? '退出表格编辑' : '进入表格编辑'}
                  </button>
                )}
              </RibbonGroup>
              <RibbonGroup label="缩放">
                <div className="flex flex-col gap-2">
                  <div className="flex items-center gap-1">
                    <button
                      type="button"
                      onClick={() => setZoomLevel(Math.max(25, zoomLevel - 10))}
                      className="px-2 py-1 text-xs rounded-md bg-black/5 dark:bg-white/5 border border-black/10 dark:border-white/10 text-text-muted hover:text-text"
                    >
                      −
                    </button>
                    <input
                      type="range"
                      min={25}
                      max={500}
                      step={5}
                      value={zoomLevel}
                      onChange={(e) => setZoomLevel(Number(e.target.value))}
                      className="w-28 accent-[var(--accent)]"
                      title="缩放"
                    />
                    <button
                      type="button"
                      onClick={() => setZoomLevel(Math.min(500, zoomLevel + 10))}
                      className="px-2 py-1 text-xs rounded-md bg-black/5 dark:bg-white/5 border border-black/10 dark:border-white/10 text-text-muted hover:text-text"
                    >
                      +
                    </button>
                  </div>
                  <div className="flex items-center gap-2 text-[11px] text-text-dim justify-center">
                    <Maximize2 className="w-3.5 h-3.5" />
                    {zoomLevel}%
                  </div>
                </div>
              </RibbonGroup>
              <DisabledStub label="阅读模式/导航窗格" />
            </>
          )}
        </div>
      </div>
    </div>
  )
}


