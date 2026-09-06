import { defineConfig } from 'tsup'

export default defineConfig({
  entry: ['src/index.ts'],
  format: ['esm', 'cjs', 'iife'],
  globalName: 'Quire',
  dts: true,
  sourcemap: true,
  clean: true,
  minify: false,
})
