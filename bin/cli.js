#!/usr/bin/env node

import fs from 'node:fs'
import path from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'

const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)

// 解析命令行参数中的 --port / -p (支持具体数字、0 或 random 随机可用端口)
const args = process.argv.slice(2)
for (let i = 0; i < args.length; i++) {
  if ((args[i] === '--port' || args[i] === '-p') && args[i + 1]) {
    const val = args[i + 1]
    process.env.PORT = (val === 'random' || val === '0') ? '0' : String(val)
    break
  }
}

// 引导启动编译后的 CarrotMRO 服务 (优先 dist/server/index.js，回退 server/index.js)
const distPath = path.join(__dirname, '..', 'dist', 'server', 'index.js')
const devPath = path.join(__dirname, '..', 'server', 'index.js')
const serverModulePath = fs.existsSync(distPath) ? distPath : devPath

import(pathToFileURL(serverModulePath).href).catch((err) => {
  console.error('CarrotMRO 启动服务异常:', err)
  process.exit(1)
})
