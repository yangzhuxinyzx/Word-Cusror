/// <reference types="vite/client" />

interface ImportMetaEnv {
  readonly VITE_WEB_SEARCH_ENDPOINT?: string
  readonly VITE_WEB_SEARCH_KEY?: string
  readonly [key: string]: string | undefined
}

interface ImportMeta {
  readonly env: ImportMetaEnv
}

