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

const AIR_OUTER_BASE = 'https://airouter.service.itstudio.club/v1'
const AIR_OUTER_KEY = 'sk-zsMIboIHKoqE0TVzph0cLikPuwNyrvxNa9Z32wBm4RWXWZ5G'

const LINAPI_BASE = 'https://api.linapi.net/v1'
const LINAPI_KEY = 'sk-GlZipyNFDvia80KAeB6QBCnVU56CYtQuDdhpJsBC9WIPsWKB'

export const PRESET_MODELS: PresetModel[] = [
  {
    id: 'gpt-5.4',
    label: 'GPT-5.4（统一文本模型）',
    description: '统一文本 API：GPT-5.4，用于对话、Agent 推理与主文档编辑',
    baseUrl: AIR_OUTER_BASE,
    apiKey: AIR_OUTER_KEY,
    model: 'gpt-5.4',
  },
]

/**
 * 全局内嵌 API Keys — 用户无需配置，打包时直接内置
 * 仅主 LLM 模型通过设置面板切换，其余服务统一使用以下配置
 */
export const BUILTIN_KEYS = {
  /** Brave Search API Key */
  braveApiKey: 'BSAQYP67rWAAQLmPDo8Ja8QwpbNBtek',
  /** LinAPI Key — 用于 Gemini PPT 设计提示词生成 (gemini-3-pro-preview) */
  linapiKey: LINAPI_KEY,
  linapiBaseUrl: LINAPI_BASE,
  /** 阿里云百炼 DashScope — 用于 PPT 图像生成（已弃用，保留兼容） */
  dashscopeApiKey: 'sk-e5a1e4b639dc4bb38e3d72ec98c08e2a',
  /** Gemini Image — PPT 图像生成 */
  geminiImageApiKey: 'sk-WSy9lOSlEurqEudP517S5PjkXZvimrnhRmFATIYexG8n9Hap',
  geminiImageBaseUrl: 'https://cdn.12ai.org',
} as const

/**
 * 锁定模型标志 — 构建时通过 VITE_LOCKED_MODEL=true 启用
 * 启用后主模型固定为当前统一文本模型，用户无法切换
 */
export const LOCKED_MODEL = import.meta.env.VITE_LOCKED_MODEL === 'true'
