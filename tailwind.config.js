/** @type {import('tailwindcss').Config} */
export default {
  content: [
    "./index.html",
    "./src/**/*.{js,ts,jsx,tsx}",
  ],
  theme: {
    extend: {
      colors: {
        background: '#09090b', // Zinc 950
        surface: '#18181b',    // Zinc 900
        'surface-hover': '#27272a', // Zinc 800
        border: '#27272a',     // Zinc 800
        primary: '#6366f1',    // Indigo 500
        'primary-hover': '#4f46e5', // Indigo 600
        text: '#e4e4e7',       // Zinc 200
        'text-muted': '#a1a1aa', // Zinc 400
        'text-dim': '#71717a',   // Zinc 500
      },
      fontFamily: {
        sans: ['Inter', 'system-ui', 'sans-serif'],
        mono: ['JetBrains Mono', 'Fira Code', 'monospace'],
        serif: ['Newsreader', 'Merriweather', 'Times New Roman', 'serif'],
      },
      animation: {
        'fade-in': 'fadeIn 0.3s ease-out',
        'slide-in': 'slideIn 0.3s ease-out',
        'pulse-slow': 'pulse 4s cubic-bezier(0.4, 0, 0.6, 1) infinite',
      },
      keyframes: {
        fadeIn: {
          '0%': { opacity: '0' },
          '100%': { opacity: '1' },
        },
        slideIn: {
          '0%': { transform: 'translateX(10px)', opacity: '0' },
          '100%': { transform: 'translateX(0)', opacity: '1' },
        },
      },
      boxShadow: {
        'glow': '0 0 20px rgba(99, 102, 241, 0.15)',
        'paper': '0 4px 20px rgba(0, 0, 0, 0.2)',
      }
    },
  },
  plugins: [],
}
