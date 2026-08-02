import { defineConfig } from 'vitest/config'
import react from '@vitejs/plugin-react'
import path from 'path'

export default defineConfig({
  plugins: [react()],
  test: {
    environment: 'happy-dom',
    globals: true,
    setupFiles: ['./src/test/setup.ts'],
    include: ['src/**/*.test.{ts,tsx}', 'server/**/*.test.ts', 'tests/**/*.test.ts'],
    coverage: {
      provider: 'v8',
      reporter: ['text', 'json', 'html', 'json-summary'],
      include: ['src/**/*.{ts,tsx}', 'server/**/*.ts'],
      exclude: [
        'src/**/*.test.{ts,tsx}',
        'server/**/*.test.ts',
        'tests/**/*.test.ts',
        'src/test/**',
        'src/main.tsx',
        'server/index.ts',
      ],
    },

    alias: {
      '@': path.resolve(__dirname, './src')
    }
  }
})

