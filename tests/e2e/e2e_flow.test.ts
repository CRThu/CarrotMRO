import { describe, it, expect, beforeAll, afterAll } from 'vitest'
import fs from 'fs/promises'
import fsSync from 'fs'
import path from 'path'
import express from 'express'
import { app } from '../../server/index.js'

const TEST_PROJECT_NAME = 'E2E_Full_Lifecycle_Project'
const TEST_RATECARD_NAME = 'E2E_RateCard'
const PROJECTS_DIR = path.join(process.cwd(), 'data', 'projects')
const RATECARD_DIR = path.join(process.cwd(), 'data', 'ratecard')

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

describe('End-to-End (E2E) Business Flow Integration Test', () => {
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

  it('完整业务闭环：创建项目 -> 上传协议定价表 -> 新建报价单 -> 智能匹配带入 -> 公式计算与保存', async () => {
    // 步骤 1: 创建新项目
    const createProjRes = await request(app, 'POST', `/api/projects/${TEST_PROJECT_NAME}`)
    expect(createProjRes.status).toBe(200)
    expect(createProjRes.body.success).toBe(true)

    // 步骤 2: 建立并导入协议定价表
    const initRcRes = await request(app, 'POST', `/api/ratecards/${TEST_RATECARD_NAME}`)
    expect(initRcRes.status).toBe(200)

    const importRcRes = await request(app, 'POST', `/api/ratecards/${TEST_RATECARD_NAME}/import`, {
      headers: ['项目名称', '单位', '不含税单价', '说明'],
      items: [
        { 项目名称: 'HRB400螺纹钢', 单位: '吨', 不含税单价: '3800.00', 说明: '国标12mm' },
        { 项目名称: '304不锈钢无缝管', 单位: '米', 不含税单价: '150.00', 说明: '壁厚3mm' }
      ],
      mapping: { 项目名称: '项目名称', 单位: '单位', 不含税单价: '不含税单价', 说明: '说明' }
    })
    expect(importRcRes.status).toBe(200)
    expect(importRcRes.body.count).toBe(2)

    // 步骤 3: 为项目绑定此协议定价表
    const patchSettingsRes = await request(app, 'PATCH', `/api/projects/${TEST_PROJECT_NAME}/settings`, {
      ratecard_name: `${TEST_RATECARD_NAME}.json`
    })
    expect(patchSettingsRes.status).toBe(200)
    expect(patchSettingsRes.body.settings.ratecard_name).toBe(`${TEST_RATECARD_NAME}.json`)

    // 步骤 4: 新建报价单
    const createQuotationRes = await request(app, 'POST', `/api/projects/${TEST_PROJECT_NAME}/quotations`, {
      name: '项目一期采购报价单'
    })
    expect(createQuotationRes.status).toBe(200)
    const quotationFile = createQuotationRes.body.file
    expect(quotationFile).toBeDefined()

    // 步骤 5: 在定价表中检索精准匹配物料
    const matchRes = await request(app, 'POST', '/api/match', {
      queries: ['螺纹钢'],
      ratecard_name: `${TEST_RATECARD_NAME}.json`
    })
    expect(matchRes.status).toBe(200)
    expect(matchRes.body['螺纹钢']).toBeDefined()
    expect(matchRes.body['螺纹钢'].length).toBeGreaterThan(0)
    const matchedCandidate = matchRes.body['螺纹钢'][0]
    const itemData = matchedCandidate[2]
    expect(itemData['项目名称']).toBe('HRB400螺纹钢')
    expect(itemData['不含税单价']).toBe('3800.00')

    // 步骤 6: 组装报价单行条目并进行联动计算
    const quantity = 10
    const unitPrice = parseFloat(itemData['不含税单价']) // 3800
    const taxRate = 0.13
    const totalPrice = quantity * unitPrice // 38000
    const taxUnitPrice = unitPrice * (1 + taxRate) // 4294
    const taxTotalPrice = quantity * taxUnitPrice // 42940

    const quotationItems = [
      {
        项目组: '建筑钢材',
        项目名称: itemData['项目名称'],
        单位: itemData['单位'],
        数量: quantity,
        不含税单价: unitPrice,
        不含税总价: totalPrice,
        税率: taxRate,
        含税单价: taxUnitPrice,
        含税总价: taxTotalPrice,
        说明: itemData['说明'],
        _matchStatus: 'matched',
        _matchedName: itemData['项目名称']
      }
    ]

    // 步骤 7: 保存并写盘报价单数据
    const saveQuotationRes = await request(app, 'PUT', `/api/projects/${TEST_PROJECT_NAME}/quotations/${quotationFile}`, {
      name: '项目一期采购报价单',
      items: quotationItems,
      remarks: ['已根据 2026 E2E 定价表完成全自动匹配与公式联动核算']
    })
    expect(saveQuotationRes.status).toBe(200)
    expect(saveQuotationRes.body.success).toBe(true)

    // 步骤 8: 重新读取持久化 JSON 文件，断言核对完整性
    const verifyRes = await request(app, 'GET', `/api/projects/${TEST_PROJECT_NAME}/quotations/${quotationFile}`)
    expect(verifyRes.status).toBe(200)
    expect(verifyRes.body.items.length).toBe(1)
    expect(verifyRes.body.items[0].项目名称).toBe('HRB400螺纹钢')
    expect(verifyRes.body.items[0].不含税总价).toBe(38000)
    expect(verifyRes.body.items[0].含税总价).toBe(42940)
    expect(verifyRes.body.remarks[0]).toContain('全自动匹配')
  })
})
