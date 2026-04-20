import { useEffect, useState } from 'react'
import { X, Thermometer, Check, AlertCircle, Image, Palette, Database, Trash2, Cpu, FolderOpen } from 'lucide-react'
import { useAI } from '../context/AIContext'
import { useDocument } from '../context/useDocument'
import { getThemeMode, setThemeMode, ThemeMode } from '../utils/theme'
import { memoryStatus, memoryStatusDetail, memoryClear, memoryRebuildIndex } from '../memory/manager'
import {
  knowledgeListPendingProfile,
  knowledgeListProfileFacts,
  knowledgeRebuild,
  knowledgeResolvePendingProfile,
  knowledgeStatus,
} from '../knowledge/manager'
import { PRESET_MODELS, LOCKED_MODEL, type PresetModelId } from '../config/models'

interface SettingsModalProps {
  onClose: () => void
}

export default function SettingsModal({ onClose }: SettingsModalProps) {
  const { settings, updateSettings } = useAI()
  const { oracleStatus, oracleError } = useDocument()
  const [localSettings, setLocalSettings] = useState(settings)
  const [testStatus, setTestStatus] = useState<'idle' | 'testing' | 'success' | 'error'>('idle')
  const [testError, setTestError] = useState('')
  const [themeMode, setThemeModeState] = useState<ThemeMode>('system')
  const [memoryStatusText, setMemoryStatusText] = useState('')
  const [memoryStatusError, setMemoryStatusError] = useState('')
  const [memoryDetailText, setMemoryDetailText] = useState('')
  const [memoryBusy, setMemoryBusy] = useState(false)
  const [pendingProfiles, setPendingProfiles] = useState<Array<{
    id: string
    category: string
    statement: string
    evidenceText: string
    sourceScope?: string
    sourcePath?: string
  }>>([])
  const [profileFacts, setProfileFacts] = useState<Array<{
    id: string
    category: string
    statement: string
    evidenceText: string
    sourceScope?: string
    sourcePath?: string
  }>>([])

  // 根据当前 settings 匹配预设模型
  const [selectedModelId, setSelectedModelId] = useState<PresetModelId>(() => {
    const match = PRESET_MODELS.find(
      m => m.model === settings.model && m.baseUrl === settings.baseUrl
    )
    return match?.id || PRESET_MODELS[0].id
  })

  useEffect(() => {
    setThemeModeState(getThemeMode())
  }, [])

  const handleSave = () => {
    const preset = PRESET_MODELS.find(m => m.id === selectedModelId) || PRESET_MODELS[0]
    updateSettings({
      ...localSettings,
      apiKey: preset.apiKey,
      baseUrl: preset.baseUrl,
      model: preset.model,
      localModel: {
        enabled: true,
        baseUrl: preset.baseUrl,
        model: preset.model,
        apiKey: preset.apiKey,
      },
    })
    setThemeMode(themeMode)
    onClose()
  }

  const testConnection = async () => {
    const preset = PRESET_MODELS.find(m => m.id === selectedModelId) || PRESET_MODELS[0]

    setTestStatus('testing')
    setTestError('')

    try {
      const response = await fetch(`${preset.baseUrl}/models`, {
        headers: {
          'Authorization': `Bearer ${preset.apiKey}`,
        },
      })

      if (response.ok) {
        setTestStatus('success')
      } else {
        const error = await response.json()
        setTestStatus('error')
        setTestError(error.error?.message || '连接失败')
      }
    } catch (error) {
      setTestStatus('error')
      setTestError(error instanceof Error ? error.message : '网络错误')
    }
  }

  const refreshMemoryStatus = async () => {
    setMemoryBusy(true)
    setMemoryStatusError('')
    try {
      const [knowledge, memory, detail, pending, facts] = await Promise.all([
        knowledgeStatus(),
        memoryStatus(),
        memoryStatusDetail(),
        knowledgeListPendingProfile(),
        knowledgeListProfileFacts(),
      ])

      if (knowledge.success) {
        const parts = [
          knowledge.workspace?.rootPath ? `工作区：${knowledge.workspace.rootPath}` : '',
          knowledge.workspace ? `工作区文件：${knowledge.workspace.indexedFileCount}/${knowledge.workspace.fileCount}` : '',
          knowledge.global?.rootPath ? `长期库：${knowledge.global.rootPath}` : '',
          knowledge.global ? `长期库文件：${knowledge.global.indexedFileCount}/${knowledge.global.fileCount}` : '',
          knowledge.profile ? `待确认：${knowledge.profile.pendingCount}` : '',
          knowledge.profile ? `已确认：${knowledge.profile.factCount}` : '',
        ].filter(Boolean)
        setMemoryStatusText(parts.join(' | ') || '知识与记忆系统已就绪')
      } else {
        setMemoryStatusText('')
        setMemoryStatusError(knowledge.error || '知识系统不可用')
      }

      const memoryParts = []
      if (memory.success) {
        if (typeof memory.fileCount === 'number') memoryParts.push(`记忆文件：${memory.fileCount}`)
        if (typeof memory.chunkCount === 'number') memoryParts.push(`记忆片段：${memory.chunkCount}`)
        if (memory.lastIndexedAt) memoryParts.push(`记忆索引时间：${memory.lastIndexedAt}`)
      }
      if (detail.success) {
        const chunkInfo = (detail.chunkSources || [])
          .map(item => `${item.source}:${item.count}`)
          .join(', ')
        if (chunkInfo) memoryParts.push(`记忆来源(${chunkInfo})`)
      }
      setMemoryDetailText(memoryParts.join(' | '))
      setPendingProfiles(pending.success ? pending.items : [])
      setProfileFacts(facts.success ? facts.items : [])
    } catch (error) {
      setMemoryStatusText('')
      setMemoryStatusError(error instanceof Error ? error.message : '记忆状态读取失败')
      setMemoryDetailText('')
      setPendingProfiles([])
      setProfileFacts([])
    } finally {
      setMemoryBusy(false)
    }
  }

  const handleClearMemory = async () => {
    const ok = window.confirm('确定清空所有记忆数据吗？此操作不可恢复。')
    if (!ok) return
    setMemoryBusy(true)
    setMemoryStatusError('')
    try {
      const result = await memoryClear('all')
      if (result.success) {
        setMemoryStatusText('记忆已清空')
      } else {
        setMemoryStatusText('')
        setMemoryStatusError(result.error || '清空失败')
      }
    } catch (error) {
      setMemoryStatusText('')
      setMemoryStatusError(error instanceof Error ? error.message : '清空失败')
    } finally {
      setMemoryBusy(false)
    }
  }

  const handleRebuildMemory = async () => {
    const ok = window.confirm('将重建全部记忆索引，可能需要一些时间。是否继续？')
    if (!ok) return
    setMemoryBusy(true)
    setMemoryStatusError('')
    try {
      const [knowledgeResult, memoryResult] = await Promise.all([
        knowledgeRebuild('all'),
        memoryRebuildIndex(),
      ])
      if (knowledgeResult.success && memoryResult.success) {
        setMemoryStatusText('知识与记忆索引重建完成')
        await refreshMemoryStatus()
      } else {
        setMemoryStatusText('')
        setMemoryStatusError(knowledgeResult.error || memoryResult.error || '索引重建失败')
      }
    } catch (error) {
      setMemoryStatusText('')
      setMemoryStatusError(error instanceof Error ? error.message : '索引重建失败')
    } finally {
      setMemoryBusy(false)
    }
  }

  const handlePickKnowledgeFolder = async () => {
    if (!window.electronAPI?.selectFolder) return
    const folderPath = await window.electronAPI.selectFolder()
    if (!folderPath) return
    setLocalSettings(prev => ({ ...prev, globalKnowledgePath: folderPath }))
  }

  const handleResolvePendingProfiles = async (action: 'accept' | 'reject', ids: string[]) => {
    if (ids.length === 0) return
    setMemoryBusy(true)
    setMemoryStatusError('')
    try {
      const result = await knowledgeResolvePendingProfile({ ids, action })
      if (!result.success) {
        setMemoryStatusError(result.error || '处理待确认画像失败')
      }
      await refreshMemoryStatus()
    } catch (error) {
      setMemoryStatusError(error instanceof Error ? error.message : '处理待确认画像失败')
    } finally {
      setMemoryBusy(false)
    }
  }

  useEffect(() => {
    void refreshMemoryStatus()
  }, [])

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center animate-fade-in">
      {/* 背景遮罩 */}
      <div 
        className="absolute inset-0 bg-black/60 backdrop-blur-sm transition-opacity"
        onClick={onClose}
      />
      
      {/* 弹窗内容 */}
      <div className="relative w-full max-w-md max-h-[85vh] flex flex-col bg-surface border border-border rounded-xl shadow-2xl overflow-hidden transform transition-all scale-100">
        {/* 头部 */}
        <div className="flex-shrink-0 flex items-center justify-between px-6 py-4 border-b border-border bg-black/10 dark:bg-white/10 backdrop-blur-md z-10">
          <h2 className="text-base font-semibold text-text">Settings</h2>
          <button
            onClick={onClose}
            className="p-1.5 rounded-md text-text-muted hover:text-text hover:bg-surface-hover transition-all"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        {/* 内容区 */}
        <div className="flex-1 overflow-y-auto p-6 space-y-5 custom-scrollbar">
          {/* 主题 */}
          <div className="space-y-2">
            <label className="flex items-center gap-2 text-xs font-medium text-text-muted uppercase tracking-wide">
              <Palette className="w-3.5 h-3.5 text-primary" />
              主题
            </label>
            <select
              value={themeMode}
              onChange={(e) => setThemeModeState(e.target.value as ThemeMode)}
              className="w-full bg-background border border-border rounded-lg px-3 py-2 text-sm text-text focus:outline-none focus:border-primary/50 focus:bg-surface-hover transition-all [&>option]:bg-[#1e1e2e] [&>option]:text-white"
            >
              <option value="system" className="bg-[#1e1e2e] text-white">跟随系统</option>
              <option value="light" className="bg-[#1e1e2e] text-white">浅色</option>
              <option value="dark" className="bg-[#1e1e2e] text-white">深色</option>
            </select>
            <p className="text-[10px] text-text-dim">
              macOS 风格雾面玻璃主题；默认跟随系统，支持手动切换
            </p>
          </div>

          {/* 模型选择 */}
          <div className="space-y-2">
            <label className="flex items-center gap-2 text-xs font-medium text-text-muted uppercase tracking-wide">
              <Cpu className="w-3.5 h-3.5 text-primary" />
              AI 模型
            </label>
            {LOCKED_MODEL || PRESET_MODELS.length <= 1 ? (
              <div className="w-full bg-background border border-border rounded-lg px-3 py-2 text-sm text-text">
                {PRESET_MODELS[0]?.label || '未配置模型'}
              </div>
            ) : (
              <select
                value={selectedModelId}
                onChange={(e) => setSelectedModelId(e.target.value as PresetModelId)}
                className="w-full bg-background border border-border rounded-lg px-3 py-2 text-sm text-text focus:outline-none focus:border-primary/50 focus:bg-surface-hover transition-all [&>option]:bg-[#1e1e2e] [&>option]:text-white"
              >
                {PRESET_MODELS.map(m => (
                  <option key={m.id} value={m.id} className="bg-[#1e1e2e] text-white">{m.label}</option>
                ))}
              </select>
            )}
            <p className="text-[10px] text-text-dim">
              {PRESET_MODELS.find(m => m.id === selectedModelId)?.description || ''}
            </p>
          </div>

          {/* PPT 图像生成模型选择 */}
          <div className="space-y-2">
            <label className="flex items-center gap-2 text-xs font-medium text-text-muted uppercase tracking-wide">
              <Image className="w-3.5 h-3.5 text-orange-400" />
              PPT 图像模型
            </label>
            <div className="w-full bg-background border border-border rounded-lg px-3 py-2 text-sm text-text">
              Gemini 3.1 Flash Image Preview（4K）
            </div>
            <p className="text-[10px] text-text-dim">
              Google Gemini 3.1 Flash 图像生成，性能/成本/延迟最佳平衡，支持 4K 输出
            </p>
          </div>

          {/* 高配无痕改字 API */}
          <div className="space-y-3">
            <label className="flex items-center gap-2 text-xs font-medium text-text-muted uppercase tracking-wide">
              <Image className="w-3.5 h-3.5 text-emerald-400" />
              高配改字 API
            </label>
            <input
              type="text"
              value={localSettings.adobeFireflyClientId || ''}
              onChange={(e) => setLocalSettings(prev => ({ ...prev, adobeFireflyClientId: e.target.value }))}
              className="w-full bg-background border border-border rounded-lg px-3 py-2 text-sm text-text focus:outline-none focus:border-primary/50"
              placeholder="Adobe Firefly Client ID"
            />
            <input
              type="password"
              value={localSettings.adobeFireflyClientSecret || ''}
              onChange={(e) => setLocalSettings(prev => ({ ...prev, adobeFireflyClientSecret: e.target.value }))}
              className="w-full bg-background border border-border rounded-lg px-3 py-2 text-sm text-text focus:outline-none focus:border-primary/50"
              placeholder="Adobe Firefly Client Secret"
            />
            <input
              type="password"
              value={localSettings.bflApiKey || ''}
              onChange={(e) => setLocalSettings(prev => ({ ...prev, bflApiKey: e.target.value }))}
              className="w-full bg-background border border-border rounded-lg px-3 py-2 text-sm text-text focus:outline-none focus:border-primary/50"
              placeholder="BFL API Key（可选）"
            />
            <p className="text-[10px] text-text-dim">
              用于后续接入 Adobe Firefly Fill / Composite 与 FLUX 高配融合候选。未配置时仍走本地主链。
            </p>
          </div>

          {/* 知识与记忆系统 */}
          <div className="space-y-3">
            <label className="flex items-center gap-2 text-xs font-medium text-text-muted uppercase tracking-wide">
              <Database className="w-3.5 h-3.5 text-primary" />
              知识与记忆
            </label>
            <div className="flex items-center gap-2">
              <input
                type="checkbox"
                checked={localSettings.knowledgeEnabled !== false}
                onChange={(e) => setLocalSettings(prev => ({ ...prev, knowledgeEnabled: e.target.checked }))}
                className="h-4 w-4"
              />
              <span className="text-sm text-text">启用本地知识库</span>
            </div>
            <div className="flex items-center gap-2">
              <input
                type="checkbox"
                checked={localSettings.workspaceKnowledgeEnabled !== false}
                onChange={(e) => setLocalSettings(prev => ({ ...prev, workspaceKnowledgeEnabled: e.target.checked }))}
                className="h-4 w-4"
              />
              <span className="text-sm text-text">启用当前工作区知识库</span>
            </div>
            <div className="flex items-center gap-2">
              <input
                type="checkbox"
                checked={localSettings.profileMemoryEnabled !== false}
                onChange={(e) => setLocalSettings(prev => ({ ...prev, profileMemoryEnabled: e.target.checked }))}
                className="h-4 w-4"
              />
              <span className="text-sm text-text">启用用户画像记忆</span>
            </div>
            <div className="space-y-2">
              <div className="flex items-center gap-2">
                <input
                  type="text"
                  value={localSettings.globalKnowledgePath || ''}
                  onChange={(e) => setLocalSettings(prev => ({ ...prev, globalKnowledgePath: e.target.value }))}
                  className="flex-1 bg-background border border-border rounded-lg px-3 py-2 text-sm text-text focus:outline-none focus:border-primary/50"
                  placeholder="长期知识库目录"
                />
                <button
                  onClick={handlePickKnowledgeFolder}
                  className="px-3 py-2 text-xs rounded-md border border-border hover:bg-surface-hover flex items-center gap-1"
                >
                  <FolderOpen className="w-3.5 h-3.5" />
                  选择
                </button>
              </div>
              <p className="text-[10px] text-text-dim">
                只支持单个全局长期知识库目录；目录中的 doc/docx/xls/xlsx/pdf/md/txt/json/xml 会被持续索引。
              </p>
            </div>
            <div className="grid grid-cols-2 gap-2">
              <input
                type="text"
                value={localSettings.embeddingBaseUrl || ''}
                onChange={(e) => setLocalSettings(prev => ({ ...prev, embeddingBaseUrl: e.target.value }))}
                className="w-full bg-background border border-border rounded-lg px-3 py-2 text-sm text-text focus:outline-none focus:border-primary/50"
                placeholder="Embedding Base URL"
              />
              <input
                type="text"
                value={localSettings.embeddingModel || ''}
                onChange={(e) => setLocalSettings(prev => ({ ...prev, embeddingModel: e.target.value }))}
                className="w-full bg-background border border-border rounded-lg px-3 py-2 text-sm text-text focus:outline-none focus:border-primary/50"
                placeholder="Embedding Model"
              />
              <input
                type="password"
                value={localSettings.embeddingApiKey || ''}
                onChange={(e) => setLocalSettings(prev => ({ ...prev, embeddingApiKey: e.target.value }))}
                className="col-span-2 w-full bg-background border border-border rounded-lg px-3 py-2 text-sm text-text focus:outline-none focus:border-primary/50"
                placeholder="Embedding API Key"
              />
            </div>
            <div className="grid grid-cols-2 gap-2">
              <input
                type="number"
                min={1}
                max={20}
                value={localSettings.knowledgeTopK ?? 8}
                onChange={(e) => setLocalSettings(prev => ({ ...prev, knowledgeTopK: Number(e.target.value) || 8 }))}
                className="w-full bg-background border border-border rounded-lg px-3 py-2 text-sm text-text focus:outline-none focus:border-primary/50"
                placeholder="知识 TopK"
                title="知识库检索条数"
              />
              <input
                type="number"
                min={1}
                max={20}
                value={localSettings.memoryTopK ?? 5}
                onChange={(e) => setLocalSettings(prev => ({ ...prev, memoryTopK: Number(e.target.value) || 5 }))}
                className="w-full bg-background border border-border rounded-lg px-3 py-2 text-sm text-text focus:outline-none focus:border-primary/50"
                placeholder="记忆 TopK"
                title="记忆检索条数"
              />
            </div>
            <div className="grid grid-cols-2 gap-2">
              <input
                type="number"
                min={500}
                max={8000}
                value={localSettings.memoryMaxChars ?? 2000}
                onChange={(e) => setLocalSettings(prev => ({ ...prev, memoryMaxChars: Number(e.target.value) || 2000 }))}
                className="w-full bg-background border border-border rounded-lg px-3 py-2 text-sm text-text focus:outline-none focus:border-primary/50"
                placeholder="记忆字符数"
                title="记忆注入最大字符数"
              />
              <input
                type="number"
                min={2000}
                max={50000}
                value={localSettings.memoryFlushThresholdChars ?? 12000}
                onChange={(e) => setLocalSettings(prev => ({ ...prev, memoryFlushThresholdChars: Number(e.target.value) || 12000 }))}
                className="w-full bg-background border border-border rounded-lg px-3 py-2 text-sm text-text focus:outline-none focus:border-primary/50"
                placeholder="刷新阈值"
                title="触发自动记忆刷新的上下文字符阈值"
              />
            </div>
            <div className="flex items-center gap-2">
              <button
                onClick={refreshMemoryStatus}
                disabled={memoryBusy}
                className="px-3 py-1.5 text-xs rounded-md border border-border hover:bg-surface-hover disabled:opacity-50"
              >
                查看知识状态
              </button>
              <button
                onClick={handleRebuildMemory}
                disabled={memoryBusy}
                className="px-3 py-1.5 text-xs rounded-md border border-border hover:bg-surface-hover disabled:opacity-50"
              >
                重建索引
              </button>
              <button
                onClick={handleClearMemory}
                disabled={memoryBusy}
                className="px-3 py-1.5 text-xs rounded-md border border-border text-error hover:bg-error/10 disabled:opacity-50 flex items-center gap-1"
              >
                <Trash2 className="w-3 h-3" />
                清空记忆
              </button>
            </div>
            {memoryStatusText && (
              <p className="text-[11px] text-text-dim">{memoryStatusText}</p>
            )}
            {memoryDetailText && (
              <p className="text-[11px] text-text-muted">{memoryDetailText}</p>
            )}
            {memoryStatusError && (
              <p className="text-[11px] text-error">{memoryStatusError}</p>
            )}
            {pendingProfiles.length > 0 && (
              <div className="space-y-2 rounded-lg border border-border p-3">
                <div className="flex items-center justify-between gap-2">
                  <p className="text-xs font-medium text-text">待确认画像</p>
                  <div className="flex items-center gap-2">
                    <button
                      onClick={() => handleResolvePendingProfiles('accept', pendingProfiles.map(item => item.id))}
                      disabled={memoryBusy}
                      className="px-2 py-1 text-[10px] rounded border border-border hover:bg-surface-hover disabled:opacity-50"
                    >
                      全部接受
                    </button>
                    <button
                      onClick={() => handleResolvePendingProfiles('reject', pendingProfiles.map(item => item.id))}
                      disabled={memoryBusy}
                      className="px-2 py-1 text-[10px] rounded border border-border hover:bg-surface-hover disabled:opacity-50"
                    >
                      全部拒绝
                    </button>
                  </div>
                </div>
                {pendingProfiles.map((item) => (
                  <div key={item.id} className="rounded-md border border-border bg-background px-3 py-2 space-y-1">
                    <div className="text-[11px] text-text-muted uppercase tracking-wide">{item.category}</div>
                    <div className="text-sm text-text">{item.statement}</div>
                    <div className="text-[11px] text-text-secondary">{item.evidenceText}</div>
                    {(item.sourceScope || item.sourcePath) && (
                      <div className="text-[10px] text-text-dim">
                        {(item.sourceScope || '').trim()} {item.sourcePath || ''}
                      </div>
                    )}
                    <div className="flex items-center gap-2 pt-1">
                      <button
                        onClick={() => handleResolvePendingProfiles('accept', [item.id])}
                        disabled={memoryBusy}
                        className="px-2 py-1 text-[10px] rounded border border-border hover:bg-surface-hover disabled:opacity-50"
                      >
                        接受
                      </button>
                      <button
                        onClick={() => handleResolvePendingProfiles('reject', [item.id])}
                        disabled={memoryBusy}
                        className="px-2 py-1 text-[10px] rounded border border-border hover:bg-surface-hover disabled:opacity-50"
                      >
                        拒绝
                      </button>
                    </div>
                  </div>
                ))}
              </div>
            )}
            {profileFacts.length > 0 && (
              <div className="space-y-2 rounded-lg border border-border p-3">
                <p className="text-xs font-medium text-text">已确认画像</p>
                {profileFacts.slice(0, 8).map((item) => (
                  <div key={item.id} className="rounded-md border border-border bg-background px-3 py-2 space-y-1">
                    <div className="text-[11px] text-text-muted uppercase tracking-wide">{item.category}</div>
                    <div className="text-sm text-text">{item.statement}</div>
                    <div className="text-[11px] text-text-secondary">{item.evidenceText}</div>
                  </div>
                ))}
              </div>
            )}
            <p className="text-[10px] text-text-dim">
              知识库只在桌面端生效；工作区知识库与长期知识库走本地索引，用户画像先进入待确认再写入。
            </p>
          </div>

          {(oracleStatus === 'unavailable' || oracleStatus === 'error') && oracleError && (
            <div className="space-y-2">
              <label className="flex items-center gap-2 text-xs font-medium text-text-muted uppercase tracking-wide">
                <AlertCircle className="w-3.5 h-3.5 text-amber-400" />
                Word Oracle
              </label>
              <div className="rounded-lg border border-amber-500/20 bg-amber-500/10 px-3 py-2 text-xs text-text-secondary">
                {oracleStatus === 'unavailable' ? '不可用' : '异常'}：{oracleError}
              </div>
              <p className="text-[10px] text-text-dim">
                仅影响 Word Oracle 对齐校验/导出，不影响常规文档编辑。
              </p>
            </div>
          )}

          {/* Temperature */}
          <div className="space-y-3">
            <div className="flex justify-between items-center">
                <label className="flex items-center gap-2 text-xs font-medium text-text-muted uppercase tracking-wide">
                <Thermometer className="w-3.5 h-3.5 text-primary" />
                Creativity
                </label>
                <span className="text-xs font-mono text-text bg-surface-hover px-2 py-0.5 rounded border border-border">{localSettings.temperature}</span>
            </div>
            
            <input
              type="range"
              min="0"
              max="1"
              step="0.1"
              value={localSettings.temperature}
              onChange={(e) => setLocalSettings(prev => ({ ...prev, temperature: parseFloat(e.target.value) }))}
              className="w-full h-1.5 bg-surface-hover rounded-lg appearance-none cursor-pointer accent-primary hover:accent-primary-hover"
            />
            <div className="flex justify-between text-[10px] text-text-dim">
              <span>Precise</span>
              <span>Creative</span>
            </div>
          </div>

          {/* 测试连接状态 */}
          {testStatus !== 'idle' && (
            <div className={`flex items-center gap-2 p-3 rounded-lg border ${
              testStatus === 'success' ? 'bg-success/10 border-success/20 text-success' :
              testStatus === 'error' ? 'bg-red-500/5 border-red-500/20 text-red-400' :
              'bg-primary/5 border-primary/20 text-primary'
            }`}>
              {testStatus === 'testing' && (
                <>
                  <div className="w-3.5 h-3.5 border-2 border-current border-t-transparent rounded-full animate-spin" />
                  <span className="text-xs">Testing connection...</span>
                </>
              )}
              {testStatus === 'success' && (
                <>
                  <Check className="w-3.5 h-3.5" />
                  <span className="text-xs">Connected successfully!</span>
                </>
              )}
              {testStatus === 'error' && (
                <>
                  <AlertCircle className="w-3.5 h-3.5" />
                  <span className="text-xs">{testError}</span>
                </>
              )}
            </div>
          )}
        </div>

        {/* 底部按钮 */}
        <div className="flex-shrink-0 flex items-center justify-between px-6 py-4 border-t border-border bg-black/10 dark:bg-white/10 backdrop-blur-md z-10">
          <button
            onClick={testConnection}
            className="px-3 py-1.5 text-xs font-medium text-text-muted hover:text-text border border-border rounded-md hover:bg-surface-hover transition-all"
          >
            Test Connection
          </button>
          <div className="flex items-center gap-2">
            <button
              onClick={onClose}
              className="px-3 py-1.5 text-xs font-medium text-text-muted hover:text-text transition-colors"
            >
              Cancel
            </button>
            <button
              onClick={handleSave}
              className="px-4 py-1.5 bg-primary text-white text-xs font-medium rounded-md hover:bg-primary-hover transition-all shadow-glow"
            >
              Save Changes
            </button>
          </div>
        </div>
      </div>
    </div>
  )
}
