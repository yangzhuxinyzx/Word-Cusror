/** @type {import('tailwindcss').Config} */
export default {
  content: [
    "./index.html",
    "./src/**/*.{js,ts,jsx,tsx}",
  ],
  theme: {
    extend: {
      colors: {
        // 主背景色 - 使用 CSS 变量适配
        background: 'var(--bg-primary)',
        'background-dark': 'var(--bg-secondary)',
        'background-light': '#ffffff', // 纯白
        
        // 表面色（由 CSS 变量适配浅/深色；修复深色模式“过亮”问题）
        surface: 'var(--surface)',
        'surface-solid': 'var(--bg-secondary)',
        'surface-hover': 'var(--surface-hover)',
        'surface-active': 'var(--surface-active)',
        
        // 边框色
        border: 'var(--glass-border)',
        'border-light': 'var(--border-light)',
        // 统一为系统蓝 glow（避免残留的绿色 glow）
        'border-glow': 'rgb(var(--accent-rgb) / 0.2)',
        
        // 主题色 - Sage
        sage: {
          50: '#f0fdf4',
          100: '#dcfce7',
          200: '#bbf7d0',
          300: '#86efac',
          400: '#4ade80',
          500: '#22c55e', // 主色
          600: '#16a34a',
          700: '#15803d',
          800: '#166534',
          900: '#14532d',
          950: '#052e16',
        },
        
        // 主题色 - Lavender
        lavender: {
          50: '#f5f3ff',
          100: '#ede9fe',
          200: '#ddd6fe',
          300: '#c4b5fd',
          400: '#a78bfa',
          500: '#8b5cf6', // 主色
          600: '#7c3aed',
          700: '#6d28d9',
          800: '#5b21b6',
          900: '#4c1d95',
          950: '#2e1065',
        },
        
        // 主题色 - Peach
        peach: {
          50: '#fff7ed',
          100: '#ffedd5',
          200: '#fed7aa',
          300: '#fdba74',
          400: '#fb923c',
          500: '#f97316', // 主色
          600: '#ea580c',
          700: '#c2410c',
          800: '#9a3412',
          900: '#7c2d12',
          950: '#431407',
        },
        
        // 功能色（macOS system colors）
        // 说明：用 rgb(var(--*-rgb) / <alpha-value>) 以支持 Tailwind 的 /xx 透明度写法
        primary: 'rgb(var(--accent-rgb) / <alpha-value>)',
        'primary-hover': 'rgb(var(--accent-hover-rgb) / <alpha-value>)',
        accent: 'rgb(var(--accent-rgb) / <alpha-value>)',
        'accent-hover': 'rgb(var(--accent-hover-rgb) / <alpha-value>)',
        highlight: 'rgb(var(--warning-rgb) / <alpha-value>)',

        // macOS system accent colors（用于图标与小面积点缀）
        'system-purple': 'rgb(var(--lavender-rgb) / <alpha-value>)',
        'system-teal': 'rgb(var(--teal-rgb) / <alpha-value>)',
        'system-gray': 'rgb(var(--system-gray-rgb) / <alpha-value>)',
        
        // 文本色 - 适配浅色模式
        text: 'var(--text-primary)',
        'text-secondary': 'var(--text-secondary)',
        'text-muted': 'var(--text-muted)',
        'text-dim': 'var(--text-dim)',
        
        // 状态色（macOS system）
        success: 'rgb(var(--success-rgb) / <alpha-value>)',
        warning: 'rgb(var(--warning-rgb) / <alpha-value>)',
        error: 'rgb(var(--danger-rgb) / <alpha-value>)',
        info: 'rgb(var(--accent-rgb) / <alpha-value>)',
      },
      fontFamily: {
        // 优先系统字体（macOS: SF / Windows: Segoe UI）
        sans: [
          '-apple-system',
          'BlinkMacSystemFont',
          '"SF Pro Text"',
          '"SF Pro Display"',
          '"Segoe UI"',
          'system-ui',
          '"Noto Sans SC"',
          '"Plus Jakarta Sans"',
          'sans-serif',
        ],
        display: [
          '-apple-system',
          'BlinkMacSystemFont',
          '"SF Pro Display"',
          '"Segoe UI"',
          'system-ui',
          '"Noto Sans SC"',
          '"Sora"',
          'sans-serif',
        ],
        mono: ['"JetBrains Mono"', '"Fira Code"', 'monospace'],
        serif: ['"Newsreader"', '"Noto Serif SC"', 'Georgia', 'serif'],
      },
      boxShadow: {
        'glow': '0 0 20px rgb(var(--accent-rgb) / 0.18)',
        'glass': '0 4px 20px rgba(0, 0, 0, 0.05)',
        'glass-lg': '0 10px 40px rgba(0, 0, 0, 0.08)',
        'paper': '0 2px 10px rgba(0, 0, 0, 0.03)',
      },
      backgroundImage: {
        'gradient-radial': 'radial-gradient(var(--tw-gradient-stops))',
        // 保留名字以兼容旧类名，但实际使用系统蓝
        'sage-gradient': 'linear-gradient(135deg, rgb(var(--accent-rgb) / 0.10) 0%, rgb(var(--accent-rgb) / 0.02) 100%)',
      },
    },
  },
  plugins: [],
}