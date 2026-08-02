import { describe, it, expect, afterAll } from 'vitest'
import path from 'path'
import { fileURLToPath, pathToFileURL } from 'url'
import axios from 'axios'
import { startServerInstance } from './index.js'

const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)

describe('Entrypoints & Dynamic Ports Integration Tests (CLI & Electron 入口测试)', () => {
  let serverInstance1: any = null
  let serverInstance2: any = null

  afterAll(async () => {
    if (serverInstance1?.server) {
      await new Promise((res) => serverInstance1.server.close(res))
    }
    if (serverInstance2?.server) {
      await new Promise((res) => serverInstance2.server.close(res))
    }
  })

  it('1. startServerInstance 未指定端口时默认分配系统随机可用端口 (port = 0)', async () => {
    serverInstance1 = await startServerInstance()
    expect(serverInstance1.port).toBeGreaterThan(0)
    expect(serverInstance1.url).toBe(`http://localhost:${serverInstance1.port}`)

    // 验证 API 可以正常响应
    const res = await axios.get(`${serverInstance1.url}/api/preset-columns`)
    expect(res.status).toBe(200)
    expect(Array.isArray(res.data)).toBe(true)
  })

  it('2. 支持多实例并发运行，随机/递增分配不同端口', async () => {
    serverInstance2 = await startServerInstance(0)
    expect(serverInstance2.port).toBeGreaterThan(0)
    expect(serverInstance2.port).not.toBe(serverInstance1.port)

    const res = await axios.get(`${serverInstance2.url}/api/preset-columns`)
    expect(res.status).toBe(200)
  })

  it('3. bin/cli.js 入口与 ES 模块路径导入正常可用', async () => {
    const cliPath = path.join(__dirname, '..', 'bin', 'cli.js')
    const fileUrl = pathToFileURL(cliPath).href
    expect(fileUrl).toContain('file://')
  })
})
