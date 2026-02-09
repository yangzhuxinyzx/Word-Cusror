import { useEffect, useState } from 'react'
import { X, Key, Server, Thermometer, Hash, Check, AlertCircle, Monitor, Image, Palette, Database, Trash2 } from 'lucide-react'
import { useAI } from '../context/AIContext'
import { getThemeMode, setThemeMode, ThemeMode } from '../utils/theme'
import { memoryStatus, memoryStatusDetail, memoryClear, memoryRebuildIndex } from '../memory/manager'

interface SettingsModalProps {
  onClose: () => void
}

export default function SettingsModal({ onClose }: SettingsModalProps) {
  const { settings, updateSettings } = useAI()
  const [localSettings, setLocalSettings] = useState(settings)
  const [testStatus, setTestStatus] = useState<'idle' | 'testing' | 'success' | 'error'>('idle')
  const [testError, setTestError] = useState('')
  const [themeMode, setThemeModeState] = useState<ThemeMode>('system')
  const [memoryStatusText, setMemoryStatusText] = useState('')
  const [memoryStatusError, setMemoryStatusError] = useState('')
  const [memoryDetailText, setMemoryDetailText] = useState('')
  const [memoryBusy, setMemoryBusy] = useState(false)

  useEffect(() => {
    setThemeModeState(getThemeMode())
  }, [])

  const handleSave = () => {
    updateSettings(localSettings)
    setThemeMode(themeMode)
    onClose()
  }

  const testConnection = async () => {
    if (!localSettings.apiKey) {
      setTestStatus('error')
      setTestError('请先输入 API Key')
      return
    }

    setTestStatus('testing')
    setTestError('')

    try {
      const response = await fetch(`${localSettings.baseUrl}/models`, {
        headers: {
          'Authorization': `Bearer ${localSettings.apiKey}`,
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
      const result = await memoryStatus()
      if (result.success) {
        const parts = [
          result.memoryDir ? `目录：${result.memoryDir}` : '',
          typeof result.fileCount === 'number' ? `文件数：${result.fileCount}` : '',
          typeof result.chunkCount === 'number' ? `片段数：${result.chunkCount}` : '',
          result.lastIndexedAt ? `索引时间：${result.lastIndexedAt}` : '',
        ].filter(Boolean)
        setMemoryStatusText(parts.join(' | ') || '记忆系统已就绪')
      } else {
        setMemoryStatusText('')
        setMemoryStatusError(result.error || result.message || '记忆系统不可用')
      }
      const detail = await memoryStatusDetail()
      if (detail.success) {
        const chunkInfo = (detail.chunkSources || [])
          .map(item => `${item.source}:${item.count}`)
          .join(', ')
        const fileInfo = (detail.fileSources || [])
          .map(item => `${item.source}:${item.count}`)
          .join(', ')
        const detailLine = [
          chunkInfo ? `片段(${chunkInfo})` : '',
          fileInfo ? `文件(${fileInfo})` : ''
        ].filter(Boolean).join(' | ')
        setMemoryDetailText(detailLine)
      } else {
        setMemoryDetailText('')
      }
    } catch (error) {
      setMemoryStatusText('')
      setMemoryStatusError(error instanceof Error ? error.message : '记忆状态读取失败')
      setMemoryDetailText('')
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
      const result = await memoryRebuildIndex()
      if (result.success) {
        setMemoryStatusText('索引重建完成')
        const chunkInfo = (result.chunkSources || [])
          .map(item => `${item.source}:${item.count}`)
          .join(', ')
        const fileInfo = (result.fileSources || [])
          .map(item => `${item.source}:${item.count}`)
          .join(', ')
        setMemoryDetailText(
          [chunkInfo ? `片段(${chunkInfo})` : '', fileInfo ? `文件(${fileInfo})` : '']
            .filter(Boolean)
            .join(' | ')
        )
      } else {
        setMemoryStatusText('')
        setMemoryStatusError(result.error || '索引重建失败')
      }
    } catch (error) {
      setMemoryStatusText('')
      setMemoryStatusError(error instanceof Error ? error.message : '索引重建失败')
    } finally {
      setMemoryBusy(false)
    }
  }

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
              className="w-full bg-background border border-border rounded-lg px-3 py-2 text-sm text-text focus:outline-none focus:border-primary/50 focus:bg-surface-hover transition-all"
            >
              <option value="system">跟随系统</option>
              <option value="light">浅色</option>
              <option value="dark">深色</option>
            </select>
            <p className="text-[10px] text-text-dim">
              macOS 风格雾面玻璃主题；默认跟随系统，支持手动切换
            </p>
          </div>

          {/* API Key */}
          <div className="space-y-2">
            <label className="flex items-center gap-2 text-xs font-medium text-text-muted uppercase tracking-wide">
              <Key className="w-3.5 h-3.5 text-primary" />
              API Key
            </label>
            <input
              type="password"
              value={localSettings.apiKey}
              onChange={(e) => setLocalSettings(prev => ({ ...prev, apiKey: e.target.value }))}
              placeholder="sk-..."
              className="w-full bg-background border border-border rounded-lg px-3 py-2 text-sm text-text placeholder-text-dim focus:outline-none focus:border-primary/50 focus:bg-surface-hover transition-all"
            />
          </div>

          {/* DashScope API Key - 阿里云百炼，用于 PPT 图像生成 */}
          <div className="space-y-2">
            <label className="flex items-center gap-2 text-xs font-medium text-text-muted uppercase tracking-wide">
              <Key className="w-3.5 h-3.5 text-orange-400" />
              DashScope API Key
            </label>
            <input
              type="password"
              value={localSettings.dashscopeApiKey || ''}
              onChange={(e) => setLocalSettings(prev => ({ ...prev, dashscopeApiKey: e.target.value }))}
              placeholder="sk-..."
              className="w-full bg-background border border-border rounded-lg px-3 py-2 text-sm text-text placeholder-text-dim focus:outline-none focus:border-primary/50 focus:bg-surface-hover transition-all"
            />
            <p className="text-[10px] text-text-dim">
              阿里云百炼 API Key，用于 PPT 图像生成
            </p>
          </div>

          {/* PPT 图像生成模型选择 */}
          <div className="space-y-2">
            <label className="flex items-center gap-2 text-xs font-medium text-text-muted uppercase tracking-wide">
              <Image className="w-3.5 h-3.5 text-orange-400" />
              PPT 图像模型
            </label>
            <select
              value={localSettings.pptImageModel || 'gemini-image'}
              onChange={(e) => setLocalSettings(prev => ({ ...prev, pptImageModel: e.target.value as 'z-image-turbo' | 'qwen-image-plus' | 'gemini-image' }))}
              className="w-full bg-background border border-border rounded-lg px-3 py-2 text-sm text-text focus:outline-none focus:border-primary/50 focus:bg-surface-hover transition-all"
            >
              <option value="gemini-image">Gemini-3-Pro-Image-Preview（推荐，高质量）</option>
              <option value="z-image-turbo">Z-Image Turbo（快速，16:9）</option>
              <option value="qwen-image-plus">Qwen-Image-Plus（高质量，异步）</option>
            </select>
            <p className="text-[10px] text-text-dim">
              {(localSettings.pptImageModel || 'gemini-image') === 'gemini-image' 
                ? 'Gemini-3-Pro-Image-Preview: 使用 LinAPI 调用 Gemini 生图，支持文生图'
                : localSettings.pptImageModel === 'qwen-image-plus' 
                  ? 'Qwen-Image-Plus: 高质量图像，异步生成，需等待轮询'
                  : 'Z-Image Turbo: 快速生成，同步返回，2048×1152 分辨率'}
            </p>
          </div>

          {/* OpenRouter API Key - 用于 Gemini PPT 设计 */}
          <div className="space-y-2">
            <label className="flex items-center gap-2 text-xs font-medium text-text-muted uppercase tracking-wide">
              <Key className="w-3.5 h-3.5 text-system-purple" />
              OpenRouter API Key
            </label>
            <input
              type="password"
              value={localSettings.openRouterApiKey || ''}
              onChange={(e) => setLocalSettings(prev => ({ ...prev, openRouterApiKey: e.target.value }))}
              placeholder="sk-or-..."
              className="w-full bg-background border border-border rounded-lg px-3 py-2 text-sm text-text placeholder-text-dim focus:outline-none focus:border-primary/50 focus:bg-surface-hover transition-all"
            />
            <p className="text-[10px] text-text-dim">
              用于 PPT 生成时调用 Gemini 设计视觉风格（可选，无则使用主模型）
            </p>
          </div>

          {/* Brave Search API Key */}
          <div className="space-y-2">
            <label className="flex items-center gap-2 text-xs font-medium text-text-muted uppercase tracking-wide">
              <Key className="w-3.5 h-3.5 text-accent" />
              Brave Search API Key
            </label>
            <input
              type="password"
              value={localSettings.braveApiKey || ''}
              onChange={(e) => setLocalSettings(prev => ({ ...prev, braveApiKey: e.target.value }))}
              placeholder="BSA..."
              className="w-full bg-background border border-border rounded-lg px-3 py-2 text-sm text-text placeholder-text-dim focus:outline-none focus:border-primary/50 focus:bg-surface-hover transition-all"
            />
            <p className="text-[10px] text-text-dim">
              用于联网搜索和资料调研（<a href="https://brave.com/search/api/" target="_blank" rel="noopener noreferrer" className="text-primary hover:underline">获取 API Key</a>）
            </p>
          </div>

          {/* 记忆系统 */}
          <div className="space-y-3">
            <label className="flex items-center gap-2 text-xs font-medium text-text-muted uppercase tracking-wide">
              <Database className="w-3.5 h-3.5 text-primary" />
              记忆系统
            </label>
            <div className="flex items-center gap-2">
              <input
                type="checkbox"
                checked={localSettings.memoryEnabled !== false}
                onChange={(e) => setLocalSettings(prev => ({ ...prev, memoryEnabled: e.target.checked }))}
                className="h-4 w-4"
              />
              <span className="text-sm text-text">启用本地记忆（用户级）</span>
            </div>
            <div className="grid grid-cols-2 gap-2">
              <input
                type="number"
                min={1}
                max={20}
                value={localSettings.memoryTopK ?? 5}
                onChange={(e) => setLocalSettings(prev => ({ ...prev, memoryTopK: Number(e.target.value) || 5 }))}
                className="w-full bg-background border border-border rounded-lg px-3 py-2 text-sm text-text focus:outline-none focus:border-primary/50"
                placeholder="TopK"
                title="记忆检索条数"
              />
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
                查看记忆状态
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
            <p className="text-[10px] text-text-dim">
              记忆只在桌面端生效；会自动写入每日记忆文件并进行本地检索。
            </p>
          </div>

          {/* Base URL */}
          <div className="space-y-2">
            <label className="flex items-center gap-2 text-xs font-medium text-text-muted uppercase tracking-wide">
              <Server className="w-3.5 h-3.5 text-primary" />
              API Endpoint
            </label>
            <input
              type="text"
              value={localSettings.baseUrl}
              onChange={(e) => setLocalSettings(prev => ({ ...prev, baseUrl: e.target.value }))}
              placeholder="https://api.openai.com/v1"
              className="w-full bg-background border border-border rounded-lg px-3 py-2 text-sm text-text placeholder-text-dim focus:outline-none focus:border-primary/50 focus:bg-surface-hover transition-all"
            />
          </div>

          {/* Model */}
          <div className="space-y-2">
            <label className="flex items-center gap-2 text-xs font-medium text-text-muted uppercase tracking-wide">
              <Hash className="w-3.5 h-3.5 text-primary" />
              Model Name
            </label>
            <input
              type="text"
              value={localSettings.model}
              onChange={(e) => setLocalSettings(prev => ({ ...prev, model: e.target.value }))}
              placeholder="kimi-k2.5"
              className="w-full bg-background border border-border rounded-lg px-3 py-2 text-sm text-text placeholder-text-dim focus:outline-none focus:border-primary/50 focus:bg-surface-hover transition-all"
            />
            <p className="text-[10px] text-text-dim">
              e.g., kimi-k2.5, GLM-4.7, gpt-4, claude-3-opus（推荐 kimi-k2.5，支持深度思考）
            </p>
          </div>

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
