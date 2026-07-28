#!/usr/bin/env node

import path from 'node:path'
import { fileURLToPath } from 'node:url'

const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)

// 引导启动编译后的 CarrotMRO 服务
const serverModulePath = path.join(__dirname, '..', 'dist', 'server', 'index.js')

import(serverModulePath).catch((err) => {
  console.error('CarrotMRO 启动服务异常:', err)
  process.exit(1)
})
