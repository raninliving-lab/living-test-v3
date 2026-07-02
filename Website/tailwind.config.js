/** @type {import('tailwindcss').Config} */
module.exports = {
  content: ['./index.html'],
  theme: {
    extend: {
      colors: {
        ink: '#1A2620',
        'ink-2': '#2C3A33',
        muted: '#6F7A72',
        line: '#E5E2DC',
        'bg-soft': '#F4F3EF',
        sand: '#D9CDBE',
        sage: '#7C8474',
        accent: '#CBAE7A',
      },
      fontFamily: {
        sans: ['"TT Norms"', '"DB Adman X"', 'ui-sans-serif', 'system-ui'],
        display: ['"TT Norms"', '"DB Adman X"', 'ui-sans-serif'],
        serif: ['"Didot"', '"TT Norms"', 'serif'],
        thai: ['"DB Adman X"', '"TT Norms"', 'sans-serif'],
      },
      letterSpacing: {
        tightest: '0em',
        'tighter-2': '0em',
      },
      boxShadow: {
        card: '0 1px 2px rgba(0,0,0,.04), 0 8px 32px rgba(0,0,0,.06)',
        soft: '0 1px 2px rgba(0,0,0,.03), 0 12px 48px rgba(0,0,0,.05)',
      },
      maxWidth: {
        page: '1400px',
      },
    },
  },
  plugins: [],
};
