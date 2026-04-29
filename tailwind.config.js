/** @type {import('tailwindcss').Config} */
export default {
  content: ['./index.html', './src/**/*.{ts,tsx}'],
  theme: {
    extend: {
      colors: {
        'giw-charcoal':    '#2C2C2C',
        'giw-olive':       '#00602B',
        'giw-olive-dark':  '#004D22',
        'giw-olive-light': '#C8E6D4',
        'giw-warm-white':  '#FFFFFF',
        'giw-mid-grey':    '#C0C0C0',
        'giw-border':      '#C0C0C0',
        'giw-cream':       '#FFFFFF',
        'giw-achieved':    '#C6EFCE',
        'giw-scoped':      '#C0C0C0',
        'giw-not-achieved':'#FCE4D6',
        'giw-mandatory':   '#2C2C2C',
      },
      fontFamily: {
        'giw-heading': ['Montserrat', 'sans-serif'],
        'giw-body':    ['Open Sans', 'sans-serif'],
      },
      borderRadius: {
        sm: '2px',
      },
    },
  },
  plugins: [],
}
