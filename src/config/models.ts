/**
 * 预设模型配置
 * 当前文本模型统一走同一套 API 配置，避免前后端出现多套文本模型入口。
 */

export interface PresetModel {
  id: string
  label: string
  description: string
  baseUrl: string
  apiKey: string
  model: string
}

export type PresetModelId = string

const AIR_OUTER_BASE = 'https://new.12ai.org/v1'
const AIR_OUTER_KEY = ''

const LINAPI_BASE = 'https://api.linapi.net/v1'
const LINAPI_KEY = ''

export const PRESET_MODELS: PresetModel[] = [
  {
    id: 'gemini-3.1-pro-preview',
    label: 'Gemini 3.1 Pro Preview（需配置 Key）',
    description: '统一文本 API：Gemini 3.1 Pro Preview，用于对话、Agent 推理与 PPT 提示词生成；请先在设置或 .env 中填写 Key',
    baseUrl: AIR_OUTER_BASE,
    apiKey: AIR_OUTER_KEY,
    model: 'gemini-3.1-pro-preview',
  },
]

/**
 * 默认外部服务配置
 * 公开仓库中不提交真实密钥；请在设置面板或 `.env` 中填写
 */
export const BUILTIN_KEYS = {
  /** Brave Search API Key */
  braveApiKey: '',
  /** LinAPI Key — 用于 Gemini PPT 设计提示词生成 (gemini-3-pro-preview) */
  linapiKey: LINAPI_KEY,
  linapiBaseUrl: LINAPI_BASE,
  /** 阿里云百炼 DashScope — 用于 PPT 图像生成（已弃用，保留兼容） */
  dashscopeApiKey: '',
  /** Gemini Image — PPT 图像生成 */
  geminiImageApiKey: '',
  geminiImageBaseUrl: 'https://cdn.12ai.org',
} as const

/**
 * 锁定模型标志 — 构建时通过 VITE_LOCKED_MODEL=true 启用
 * 启用后主模型固定为当前统一文本模型，用户无法切换
 */
export const LOCKED_MODEL = import.meta.env.VITE_LOCKED_MODEL === 'true'
