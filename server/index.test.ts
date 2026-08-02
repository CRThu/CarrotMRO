import { describe, it, expect, beforeAll, afterAll } from 'vitest'
import fs from 'fs/promises'
import fsSync from 'fs'
import path from 'path'
import express from 'express'
import * as XLSX from 'xlsx'

const TEST_PROJECT_NAME = 'integration-test-project'
const TEST_RATECARD_NAME = 'integration-test-ratecard'
const PROJECTS_DIR = path.join(process.cwd(), 'data', 'projects')
const RATECARD_DIR = path.join(process.cwd(), 'data', 'ratecard')

import { app } from './index.js'

async function request(app: express.Express, method: string, urlPath: string, body?: any) {
  return new Promise<{ status: number; body: any }>((resolve, reject) => {
    const server = app.listen(0, '127.0.0.1', async () => {
      const address = server.address() as any
      const port = address.port
      try {
        const options: RequestInit = {
          method,
          headers: { 'Content-Type': 'application/json' },
        }
        if (body) options.body = JSON.stringify(body)
        const res = await fetch(`http://127.0.0.1:${port}${urlPath}`, options)
        const json = await res.json().catch(() => ({}))
        server.close(() => resolve({ status: res.status, body: json }))
      } catch (err) {
        server.close(() => reject(err))
      }
    })
  })
}

async function requestBuffer(app: express.Express, method: string, urlPath: string, body?: any) {
  return new Promise<{ status: number; buffer: Buffer; contentType: string }>((resolve, reject) => {
    const server = app.listen(0, '127.0.0.1', async () => {
      const address = server.address() as any
      const port = address.port
      try {
        const options: RequestInit = {
          method,
          headers: { 'Content-Type': 'application/json' },
        }
        if (body) options.body = JSON.stringify(body)
        const res = await fetch(`http://127.0.0.1:${port}${urlPath}`, options)
        const arrayBuf = await res.arrayBuffer()
        const buffer = Buffer.from(arrayBuf)
        const contentType = res.headers.get('content-type') || ''
        server.close(() => resolve({ status: res.status, buffer, contentType }))
      } catch (err) {
        server.close(() => reject(err))
      }
    })
  })
}

describe('Server Express REST API Integration Tests', () => {
  beforeAll(async () => {
    const projPath = path.join(PROJECTS_DIR, TEST_PROJECT_NAME)
    const ratecardPath = path.join(RATECARD_DIR, `${TEST_RATECARD_NAME}.json`)
    if (fsSync.existsSync(projPath)) {
      await fs.rm(projPath, { recursive: true, force: true }).catch(() => {})
    }
    if (fsSync.existsSync(ratecardPath)) {
      await fs.unlink(ratecardPath).catch(() => {})
    }
  })

  afterAll(async () => {
    const projPath = path.join(PROJECTS_DIR, TEST_PROJECT_NAME)
    const ratecardPath = path.join(RATECARD_DIR, `${TEST_RATECARD_NAME}.json`)
    if (fsSync.existsSync(projPath)) {
      await fs.rm(projPath, { recursive: true, force: true }).catch(() => {})
    }
    if (fsSync.existsSync(ratecardPath)) {
      await fs.unlink(ratecardPath).catch(() => {})
    }
  })

  describe('Preset Columns & Global Settings API', () => {
    it('GET /api/preset-columns 应返回 10 项预制列', async () => {
      const res = await request(app, 'GET', '/api/preset-columns')
      expect(res.status).toBe(200)
      expect(Array.isArray(res.body)).toBe(true)
      expect(res.body).toContain('项目名称')
      expect(res.body.length).toBe(10)
    })

    it('GET & PUT /api/settings 应能正确读取与更新配置', async () => {
      const getRes = await request(app, 'GET', '/api/settings')
      expect(getRes.status).toBe(200)
      expect(getRes.body.llm).toBeDefined()

      const putRes = await request(app, 'PUT', '/api/settings', {
        llm: { activeProvider: 'google' }
      })
      expect(putRes.status).toBe(200)
      expect(putRes.body.success).toBe(true)
    })
  })

  describe('Project Management API', () => {
    it('POST /api/projects/:name 应成功初始化新项目', async () => {
      const res = await request(app, 'POST', `/api/projects/${TEST_PROJECT_NAME}`)
      expect(res.status).toBe(200)
      expect(res.body.success).toBe(true)
    })

    it('GET /api/projects 应包含新建的项目名称', async () => {
      const res = await request(app, 'GET', '/api/projects')
      expect(res.status).toBe(200)
      expect(Array.isArray(res.body)).toBe(true)
      expect(res.body).toContain(TEST_PROJECT_NAME)
    })

    it('GET & PATCH /api/projects/:name/settings 应能读取并更新项目设置', async () => {
      const patchRes = await request(app, 'PATCH', `/api/projects/${TEST_PROJECT_NAME}/settings`, {
        ocr_columns: ['项目名称', '单位', '不含税单价']
      })
      expect(patchRes.status).toBe(200)
      expect(patchRes.body.success).toBe(true)
    })
  })

  describe('Quotations CRUD API', () => {
    let quotationFile = ''

    it('POST /api/projects/:name/quotations 应新建报价单', async () => {
      const res = await request(app, 'POST', `/api/projects/${TEST_PROJECT_NAME}/quotations`, {
        name: '维保采购报价单'
      })
      expect(res.status).toBe(200)
      expect(res.body.file).toBeDefined()
      quotationFile = res.body.file
    })

    it('GET /api/projects/:name/quotations 应列表返回当前项目的报价单', async () => {
      const res = await request(app, 'GET', `/api/projects/${TEST_PROJECT_NAME}/quotations`)
      expect(res.status).toBe(200)
      expect(Array.isArray(res.body.files)).toBe(true)
      expect(res.body.files).toContain(quotationFile)
    })

    it('PUT & GET /api/projects/:name/quotations/:file 应修改并读取报价单数据', async () => {
      const mockItems = [
        { 项目名称: '高压阀门', 单位: '个', 数量: 2, 不含税单价: 150, 不含税总价: 300 }
      ]
      const putRes = await request(app, 'PUT', `/api/projects/${TEST_PROJECT_NAME}/quotations/${quotationFile}`, {
        name: '更新后的报价单',
        items: mockItems
      })
      expect(putRes.status).toBe(200)
      expect(putRes.body.success).toBe(true)

      const getRes = await request(app, 'GET', `/api/projects/${TEST_PROJECT_NAME}/quotations/${quotationFile}`)
      expect(getRes.status).toBe(200)
      expect(getRes.body.name).toBe('更新后的报价单')
      expect(getRes.body.items.length).toBe(1)
    })
  })

  describe('RateCard Management & Match API', () => {
    it('POST /api/ratecards/:name 应初始化协议定价表', async () => {
      const res = await request(app, 'POST', `/api/ratecards/${TEST_RATECARD_NAME}`)
      expect(res.status).toBe(200)
      expect(res.body.success).toBe(true)
    })

    it('POST /api/ratecards/:name/import 应清洗并写入物料条目', async () => {
      const payload = {
        headers: ['项目名称', '单位', '不含税单价', '说明'],
        items: [
          { 项目名称: 'PVC穿线管', 单位: '根', 不含税单价: '12.50', 说明: '规格DN20' },
        ],
        mapping: {
          项目名称: '项目名称',
          单位: '单位',
          不含税单价: '不含税单价',
          说明: '说明'
        }
      }
      const res = await request(app, 'POST', `/api/ratecards/${TEST_RATECARD_NAME}/import`, payload)
      expect(res.status).toBe(200)
      expect(res.body.success).toBe(true)
    })

    it('POST /api/match 应在协议定价表中精确比对物料信息', async () => {
      const res = await request(app, 'POST', '/api/match', {
        queries: ['PVC穿线管'],
        ratecard_name: `${TEST_RATECARD_NAME}.json`
      })
      expect(res.status).toBe(200)
      expect(res.body['PVC穿线管']).toBeDefined()
      expect(Array.isArray(res.body['PVC穿线管'])).toBe(true)
      expect(res.body['PVC穿线管'].length).toBeGreaterThanOrEqual(1)
      expect(res.body['PVC穿线管'][0][2]['项目名称']).toBe('PVC穿线管')
    })

    it('POST /api/projects/:name/quotations/:file/export 应正确导出标准 Excel 并支持反向解析验证', async () => {
      // 准备报价单模拟数据
      const quotationFile = 'quotation-1.json'
      const quotationPath = path.join(PROJECTS_DIR, TEST_PROJECT_NAME, quotationFile)
      const mockQuotationData = {
        columns: ['项目组', '项目名称', '单位', '数量', '不含税单价', '不含税总价', '含税单价', '含税总价', '说明'],
        items: [
          { 项目组: '电工组', 项目名称: '铜芯电缆', 单位: '米', 数量: '100', 不含税单价: '45.00', 不含税总价: '4500.00', 税率: '0.13', 含税单价: '50.85', 含税总价: '5085.00', 说明: '国标' },
          { 项目组: '水暖组', 项目名称: 'PPR水管', 单位: '根', 数量: '50', 不含税单价: '20.00', 不含税总价: '1000.00', 税率: '0.13', 含税单价: '22.60', 含税总价: '1130.00', 说明: '20mm' }
        ]
      }
      await fs.writeFile(quotationPath, JSON.stringify(mockQuotationData, null, 2), 'utf-8')

      // 调用导出 API
      const exportRes = await requestBuffer(app, 'POST', `/api/projects/${TEST_PROJECT_NAME}/quotations/${quotationFile}/export`)
      expect(exportRes.status).toBe(200)
      expect(exportRes.contentType).toContain('spreadsheetml.sheet')

      // 反向解析导出的 Excel 字节流
      const workbook = XLSX.read(exportRes.buffer, { type: 'buffer' })
      expect(workbook.SheetNames).toContain('报价单')

      const sheet = workbook.Sheets['报价单']
      const sheetData: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1 })

      // 1. 验证大标题
      expect(sheetData[0][0]).toBe('报价单')
      expect(sheetData[1][0]).toContain(`项目名称: ${TEST_PROJECT_NAME}`)

      // 2. 验证表头列 (第4行, index=3)
      const headerRow = sheetData[3]
      expect(headerRow[0]).toBe('序号')
      expect(headerRow).toContain('项目名称')
      expect(headerRow).toContain('不含税总价')
      expect(headerRow).toContain('含税总价')

      // 3. 验证数据明细行
      expect(sheetData[4][0]).toBe(1)
      expect(sheetData[4][headerRow.indexOf('项目名称')]).toBe('铜芯电缆')
      expect(sheetData[4][headerRow.indexOf('不含税总价')]).toBe(4500)

      expect(sheetData[5][0]).toBe(2)
      expect(sheetData[5][headerRow.indexOf('项目名称')]).toBe('PPR水管')
      expect(sheetData[5][headerRow.indexOf('不含税总价')]).toBe(1000)

      // 4. 验证合计汇总行
      const totalRow = sheetData[6]
      expect(totalRow[0]).toBe('合计')
      expect(totalRow[headerRow.indexOf('不含税总价')]).toBe(5500)
      expect(totalRow[headerRow.indexOf('含税总价')]).toBe(6215)

      // 5. 验证原生 Excel 公式 (原生乘法公式 * 与原生求和公式 SUM)
      const workbookFormulas = XLSX.read(exportRes.buffer, { type: 'buffer', cellFormula: true })
      const sheetFormulas = workbookFormulas.Sheets['报价单']
      const formulaKeys = Object.keys(sheetFormulas).filter(k => !k.startsWith('!') && sheetFormulas[k]?.f)

      expect(formulaKeys.length).toBeGreaterThan(0)
      const hasMultiplyFormula = formulaKeys.some(k => String(sheetFormulas[k].f).includes('*'))
      const hasSumFormula = formulaKeys.some(k => String(sheetFormulas[k].f).includes('SUM'))

      expect(hasMultiplyFormula).toBe(true)
      expect(hasSumFormula).toBe(true)
    })
  })
})
