/** @type {import('tailwindcss').Config} */
export default {
  content: ['./index.html', './src/**/*.{js,jsx}'],
  theme: {
    extend: {
      colors: {
        bg:            '#F7F8FA',
        surface:       '#FFFFFF',
        border:        '#E2E5EA',
        text:          '#1A1D23',
        'text-muted':  '#6B7280',
        primary:       '#1F5FAE',
        'primary-hover':'#184C8C',
        success:       '#1E8E5A',
        warning:       '#C77D17',
        danger:        '#C0392B',
      },
      fontFamily: {
        sans: ['Inter', 'ui-sans-serif', 'system-ui', 'sans-serif'],
        mono: ['JetBrains Mono', 'ui-monospace', 'monospace'],
      },
    },
  },
  plugins: [],
}
