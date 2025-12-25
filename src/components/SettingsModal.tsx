import { useState } from 'react'
import { X, Key, Server, Thermometer, Hash, Check, AlertCircle, Monitor, Image } from 'lucide-react'
import { useAI } from '../context/AIContext'

interface SettingsModalProps {
  onClose: () => void
}

export default function SettingsModal({ onClose }: SettingsModalProps) {
  const { settings, updateSettings } = useAI()
  const [localSettings, setLocalSettings] = useState(settings)
  const [testStatus, setTestStatus] = useState<'idle' | 'testing' | 'success' | 'error'>('idle')
  const [testError, setTestError] = useState('')

  const handleSave = () => {
    updateSettings(localSettings)
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

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center animate-fade-in">
      {/* 背景遮罩 */}
      <div 
        className="absolute inset-0 bg-black/60 backdrop-blur-sm transition-opacity"
        onClick={onClose}
      />
      
      {/* 弹窗内容 */}
      <div className="relative w-full max-w-md bg-surface border border-border rounded-xl shadow-2xl overflow-hidden transform transition-all scale-100">
        {/* 头部 */}
        <div className="flex items-center justify-between px-6 py-4 border-b border-border bg-background/50">
          <h2 className="text-base font-semibold text-text">Settings</h2>
          <button
            onClick={onClose}
            className="p-1.5 rounded-md text-text-muted hover:text-text hover:bg-surface-hover transition-all"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        {/* 内容区 */}
        <div className="p-6 space-y-5">
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
              <Key className="w-3.5 h-3.5 text-purple-400" />
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
              <Key className="w-3.5 h-3.5 text-blue-400" />
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
              placeholder="gpt-oss-20b"
              className="w-full bg-background border border-border rounded-lg px-3 py-2 text-sm text-text placeholder-text-dim focus:outline-none focus:border-primary/50 focus:bg-surface-hover transition-all"
            />
            <p className="text-[10px] text-text-dim">
              e.g., gpt-4, claude-3-opus, or local model name
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
              testStatus === 'success' ? 'bg-green-500/5 border-green-500/20 text-green-400' :
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
        <div className="flex items-center justify-between px-6 py-4 border-t border-border bg-background/50">
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
