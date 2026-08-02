import express from 'express'
import cors from 'cors'
import path from 'path'
import { fileURLToPath } from 'url'
import fs from 'fs/promises'
import fsSync from 'fs'
import open from 'open'
import multer from 'multer'
import * as XLSX from 'xlsx'
import { getDataDir, loadSettings, getSettings, saveSettings } from './services/settings.js'
import { runOcrWithLlm, testLlmConnection, ImageInput } from './services/llm.js'

const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)

const app = express()
const PORT = process.env.PORT || 3000
const upload = multer({ storage: multer.memoryStorage() })

app.use(cors())
app.use(express.json({ limit: '50mb' }))

// 自动幂等保证基础数据目录(data/projects, data/ratecard, data/template)已建立
app.use(async (req, res, next) => {
  await initDirs()
  next()
})

// 内存中的任务状态缓存
const tasks: Record<
  string,
  {
    status: 'processing' | 'done' | 'error'
    progress?: string
    logs?: string[]
    file?: string
    columns?: string[]
    result?: any
    message?: string
    streamText?: string
  }
> = {}

export function getProjectsDir() { return path.join(getDataDir(), 'projects') }
export function getRatecardDir() { return path.join(getDataDir(), 'ratecard') }
export function getTemplateDir() { return path.join(getDataDir(), 'template') }

// 初始化数据目录与自动加载持久化配置
export async function initDirs() {
  const projectsDir = getProjectsDir()
  const ratecardDir = getRatecardDir()
  const templateDir = getTemplateDir()
  try {
    await fs.mkdir(projectsDir, { recursive: true })
    await fs.mkdir(ratecardDir, { recursive: true })
    await fs.mkdir(templateDir, { recursive: true })
    await loadSettings()
    console.log(`数据目录 [${getDataDir()}] 及 settings.json 成功自动加载！`)
  } catch (err) {
    console.error('初始化数据目录/设置失败:', err)
  }
}

// 全局 10 项标准预制列
const PRESET_COLUMNS = [
  '项目组',
  '项目名称',
  '单位',
  '数量',
  '不含税单价',
  '不含税总价',
  '税率',
  '含税单价',
  '含税总价',
  '说明'
]

// ===== 设置 API =====
app.get('/api/settings', (req, res) => {
  res.json(getSettings())
})

app.put('/api/settings', async (req, res) => {
  try {
    const updated = await saveSettings(req.body)
    res.json({ success: true, settings: updated })
  } catch (err: any) {
    res.status(500).json({ success: false, detail: err.message || String(err) })
  }
})

app.post('/api/settings/test-llm', async (req, res) => {
  try {
    const result = await testLlmConnection(req.body)
    res.json(result)
  } catch (err: any) {
    res.status(500).json({ success: false, message: err.message || String(err) })
  }
})

// ===== 全局预制列 =====
app.get('/api/preset-columns', (req, res) => {
  res.json(PRESET_COLUMNS)
})

// ===== 项目及 Settings API =====
app.get('/api/projects', async (req, res) => {
  try {
    const entries = await fs.readdir(getProjectsDir(), { withFileTypes: true })
    const projectNames = entries.filter(e => e.isDirectory()).map(e => e.name)
    res.json(projectNames)
  } catch (err) {
    res.json([])
  }
})

app.post('/api/projects/:name', async (req, res) => {
  const { name } = req.params
  const projPath = path.join(getProjectsDir(), name)
  try {
    await fs.mkdir(projPath, { recursive: true })
    const settingsPath = path.join(projPath, 'settings.json')
    const settingsData = {
      name,
      created_at: new Date().toISOString(),
      ratecard_name: null,
      template_name: null,
      ocr_columns: ['项目名称', '单位', '数量', '不含税单价', '说明'],
      quotation_columns: PRESET_COLUMNS,
      match_validation_rules: {
        strict_name_match: true,
        check_columns: ['项目名称', '单位'],
        fill_columns: ['单位', '不含税单价', '含税单价', '税率', '说明'],
      },
    }
    await fs.writeFile(settingsPath, JSON.stringify(settingsData, null, 2), 'utf-8')
    res.json({ success: true, name })
  } catch (err: any) {
    res.status(500).json({ detail: err.message })
  }
})

app.get('/api/projects/:name', async (req, res) => {
  const { name } = req.params
  const settingsPath = path.join(getProjectsDir(), name, 'settings.json')
  const legacyConfigPath = path.join(getProjectsDir(), name, 'project.json')

  let settings = {
    name,
    created_at: new Date().toISOString(),
    ratecard_name: null as string | null,
    template_name: null as string | null,
    ocr_columns: ['项目名称', '单位', '数量', '不含税单价', '说明'],
    quotation_columns: PRESET_COLUMNS,
    match_validation_rules: {
      strict_name_match: true,
      check_columns: ['项目名称', '单位'],
      fill_columns: ['单位', '不含税单价', '含税单价', '税率', '说明'],
    },
  }

  try {
    if (fsSync.existsSync(settingsPath)) {
      const content = await fs.readFile(settingsPath, 'utf-8')
      settings = { ...settings, ...JSON.parse(content) }
    } else if (fsSync.existsSync(legacyConfigPath)) {
      const content = await fs.readFile(legacyConfigPath, 'utf-8')
      const legacy = JSON.parse(content)
      settings = {
        ...settings,
        ratecard_name: legacy.ratecard_name || null,
        template_name: legacy.template_name || null,
      }
      await fs.writeFile(settingsPath, JSON.stringify(settings, null, 2), 'utf-8')
    } else {
      await fs.mkdir(path.join(getProjectsDir(), name), { recursive: true })
      await fs.writeFile(settingsPath, JSON.stringify(settings, null, 2), 'utf-8')
    }
  } catch {}

  res.json(settings)
})

app.patch('/api/projects/:name/settings', async (req, res) => {
  const { name } = req.params
  const projPath = path.join(getProjectsDir(), name)
  const settingsPath = path.join(projPath, 'settings.json')
  try {
    await fs.mkdir(projPath, { recursive: true })
    let settings: any = {
      name,
      created_at: new Date().toISOString(),
      ratecard_name: null,
      template_name: null,
      ocr_columns: ['项目名称', '单位', '数量', '不含税单价', '说明'],
      quotation_columns: PRESET_COLUMNS,
    }
    if (fsSync.existsSync(settingsPath)) {
      settings = JSON.parse(await fs.readFile(settingsPath, 'utf-8'))
    }
    settings = { ...settings, ...req.body }
    await fs.writeFile(settingsPath, JSON.stringify(settings, null, 2), 'utf-8')
    res.json({ success: true, settings })
  } catch (err: any) {
    res.status(500).json({ detail: err.message })
  }
})

// 增加路由兼容 patch /ratecard 和 /template
app.patch('/api/projects/:name/ratecard', async (req, res) => {
  const { name } = req.params
  const { ratecard_name } = req.body
  const settingsPath = path.join(getProjectsDir(), name, 'settings.json')
  try {
    let settings: any = {}
    if (fsSync.existsSync(settingsPath)) {
      settings = JSON.parse(await fs.readFile(settingsPath, 'utf-8'))
    }
    settings.ratecard_name = ratecard_name
    await fs.writeFile(settingsPath, JSON.stringify(settings, null, 2), 'utf-8')
    res.json({ success: true, settings })
  } catch (err: any) {
    res.status(500).json({ detail: err.message })
  }
})

app.patch('/api/projects/:name/template', async (req, res) => {
  const { name } = req.params
  const { template_name } = req.body
  const settingsPath = path.join(getProjectsDir(), name, 'settings.json')
  try {
    let settings: any = {}
    if (fsSync.existsSync(settingsPath)) {
      settings = JSON.parse(await fs.readFile(settingsPath, 'utf-8'))
    }
    settings.template_name = template_name
    await fs.writeFile(settingsPath, JSON.stringify(settings, null, 2), 'utf-8')
    res.json({ success: true, settings })
  } catch (err: any) {
    res.status(500).json({ detail: err.message })
  }
})

// ===== 报价单 API (quotation-*.json) =====
app.get('/api/projects/:name/quotations', async (req, res) => {
  const { name } = req.params
  const projPath = path.join(getProjectsDir(), name)
  try {
    await fs.mkdir(projPath, { recursive: true })
    const files = await fs.readdir(projPath)
    const qFiles = files.filter(f => f.startsWith('quotation-') && f.endsWith('.json'))
    res.json({ files: qFiles })
  } catch {
    res.json({ files: [] })
  }
})

app.post('/api/projects/:name/quotations', async (req, res) => {
  const { name } = req.params
  const projPath = path.join(getProjectsDir(), name)
  try {
    await fs.mkdir(projPath, { recursive: true })
    const files = await fs.readdir(projPath)
    const qFiles = files.filter(f => f.startsWith('quotation-') && f.endsWith('.json'))
    const fileIndex = qFiles.length + 1
    const filename = `quotation-${fileIndex}.json`

    const payload = {
      created_at: new Date().toISOString(),
      last_edit_time: new Date().toISOString(),
      items: [],
    }

    await fs.writeFile(path.join(projPath, filename), JSON.stringify(payload, null, 2), 'utf-8')
    res.json({ success: true, file: filename })
  } catch (err: any) {
    res.status(500).json({ detail: err.message })
  }
})

app.get('/api/projects/:name/quotations/:file', async (req, res) => {
  const { name, file } = req.params
  const filePath = path.join(getProjectsDir(), name, file)
  try {
    const content = await fs.readFile(filePath, 'utf-8')
    res.json(JSON.parse(content))
  } catch (err: any) {
    res.status(500).json({ detail: err.message })
  }
})

app.put('/api/projects/:name/quotations/:file', async (req, res) => {
  const { name, file } = req.params
  const filePath = path.join(getProjectsDir(), name, file)
  try {
    const payload = {
      ...req.body,
      last_edit_time: new Date().toISOString(),
    }
    await fs.writeFile(filePath, JSON.stringify(payload, null, 2), 'utf-8')
    res.json({ success: true })
  } catch (err: any) {
    res.status(500).json({ detail: err.message })
  }
})

app.delete('/api/projects/:name/quotations/:file', async (req, res) => {
  const { name, file } = req.params
  const filePath = path.join(getProjectsDir(), name, file)
  try {
    await fs.unlink(filePath)
    res.json({ success: true })
  } catch (err: any) {
    res.status(500).json({ detail: err.message })
  }
})

app.patch('/api/projects/:name/quotations/:file/rename', async (req, res) => {
  const { name, file } = req.params
  let { new_filename } = req.body
  if (!new_filename || typeof new_filename !== 'string') {
    return res.status(400).json({ detail: '新报价单文件名不能为空' })
  }
  new_filename = new_filename.trim()
  if (!new_filename.endsWith('.json')) {
    new_filename = `${new_filename}.json`
  }
  const safeNewFile = path.basename(new_filename)
  const oldPath = path.join(getProjectsDir(), name, file)
  const newPath = path.join(getProjectsDir(), name, safeNewFile)

  try {
    if (!fsSync.existsSync(oldPath)) {
      return res.status(404).json({ detail: '原报价单文件不存在' })
    }
    if (oldPath !== newPath && fsSync.existsSync(newPath)) {
      return res.status(400).json({ detail: '已存在同名的报价单文件' })
    }
    await fs.rename(oldPath, newPath)
    res.json({ success: true, file: safeNewFile })
  } catch (err: any) {
    res.status(500).json({ detail: err.message })
  }
})

// ===== 报价单标准 Excel 导出 API =====
app.post('/api/projects/:name/quotations/:file/export', async (req, res) => {
  const { name, file } = req.params
  const projectDir = path.join(getProjectsDir(), name)
  const quotationPath = path.join(projectDir, file)

  if (!fsSync.existsSync(quotationPath)) {
    return res.status(404).json({ detail: '报价单文件不存在' })
  }

  try {
    // 1. 读取项目配置（获取 quotation_columns）
    const settingsPath = path.join(projectDir, 'settings.json')
    let quotationColumns: string[] = []
    if (fsSync.existsSync(settingsPath)) {
      const settingsContent = await fs.readFile(settingsPath, 'utf-8')
      const settings = JSON.parse(settingsContent)
      if (Array.isArray(settings.quotation_columns) && settings.quotation_columns.length > 0) {
        quotationColumns = settings.quotation_columns
      }
    }

    // 若项目未配置 quotation_columns，默认使用全局 10 项标准预制列
    if (quotationColumns.length === 0) {
      quotationColumns = [...PRESET_COLUMNS]
    }

    // 2. 读取报价单数据
    const quotationContent = await fs.readFile(quotationPath, 'utf-8')
    const quotationData = JSON.parse(quotationContent)
    const items: Record<string, any>[] = Array.isArray(quotationData.items) ? quotationData.items : []

    // 3. 构建二维表格数据数组 (AOA: Array of Arrays)
    const aoa: any[][] = []

    // 大标题行
    aoa.push(['报价单'])

    // 基本信息行
    const exportDate = new Date().toISOString().split('T')[0]
    aoa.push([`项目名称: ${name}`, '', '', `报价单: ${file.replace('.json', '')}`, '', `导出日期: ${exportDate}`])

    // 空行隔开
    aoa.push([])

    // 表头行 (序号 + quotation_columns)
    const tableHeader = ['序号', ...quotationColumns]
    aoa.push(tableHeader)

    // 列索引映射到 Excel 列字母 (0 -> A, 1 -> B, ...)
    const getColLetter = (colIdx: number): string => {
      let letter = ''
      let curr = colIdx
      while (curr >= 0) {
        const remainder = curr % 26
        letter = String.fromCharCode(remainder + 65) + letter
        curr = Math.floor(curr / 26) - 1
      }
      return letter
    }

    const colToLetterMap: Record<string, string> = {}
    tableHeader.forEach((colName, idx) => {
      colToLetterMap[colName] = getColLetter(idx)
    })

    const startRow = 5 // 数据首行 Excel 行号
    const endRow = startRow + items.length - 1 // 数据末行 Excel 行号

    items.forEach((item, index) => {
      const excelRow = startRow + index
      const row: any[] = [index + 1]

      quotationColumns.forEach(col => {
        const val = item[col] ?? ''

        // 1. 不含税总价：使用 Excel 原生乘法公式 =数量*不含税单价
        if (col === '不含税总价' && colToLetterMap['数量'] && colToLetterMap['不含税单价']) {
          const qtyCell = `${colToLetterMap['数量']}${excelRow}`
          const priceCell = `${colToLetterMap['不含税单价']}${excelRow}`
          const qtyVal = Number(item['数量']) || 0
          const priceVal = Number(item['不含税单价']) || 0
          row.push({ t: 'n', f: `${qtyCell}*${priceCell}`, v: Number((qtyVal * priceVal).toFixed(2)) })
        }
        // 2. 含税总价：使用 Excel 原生乘法公式 =数量*含税单价
        else if (col === '含税总价' && colToLetterMap['数量'] && colToLetterMap['含税单价']) {
          const qtyCell = `${colToLetterMap['数量']}${excelRow}`
          const incPriceCell = `${colToLetterMap['含税单价']}${excelRow}`
          const qtyVal = Number(item['数量']) || 0
          const incPriceVal = Number(item['含税单价']) || 0
          row.push({ t: 'n', f: `${qtyCell}*${incPriceCell}`, v: Number((qtyVal * incPriceVal).toFixed(2)) })
        }
        // 3. 普通列判断数值类型写入
        else if (typeof val === 'number') {
          row.push(val)
        } else if (typeof val === 'string' && val.trim() !== '' && !isNaN(Number(val))) {
          if (col === '税率' && val.includes('%')) {
            row.push(val)
          } else {
            row.push(Number(val))
          }
        } else {
          row.push(val)
        }
      })
      aoa.push(row)
    })

    // 4. 底部合计行：使用 Excel 原生 SUM 公式 =SUM(F5:F10)
    if (items.length > 0) {
      const totalRow: any[] = ['合计']
      quotationColumns.forEach(col => {
        const colLetter = colToLetterMap[col]
        if ((col === '不含税总价' || col === '含税总价') && colLetter) {
          let initialSum = 0
          items.forEach(it => { initialSum += Number(it[col]) || 0 })
          totalRow.push({ t: 'n', f: `SUM(${colLetter}${startRow}:${colLetter}${endRow})`, v: Number(initialSum.toFixed(2)) })
        } else {
          totalRow.push('')
        }
      })
      aoa.push(totalRow)
    }

    // 4. 使用 XLSX 生成 Workbook & Worksheet
    const worksheet = XLSX.utils.aoa_to_sheet(aoa)

    // 设置基本列宽自适应
    const colWidths = tableHeader.map(colName => {
      const minWidth = Math.max(String(colName).length * 2 + 4, 12)
      return { wch: minWidth }
    })
    worksheet['!cols'] = colWidths

    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, worksheet, '报价单')

    // 5. 导出 Buffer 并返回 HTTP 响应
    const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' })

    const exportFilename = `${file.replace('.json', '')}.xlsx`
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(exportFilename)}"`)
    res.send(buffer)
  } catch (err: any) {
    res.status(500).json({ detail: `导出失败: ${err.message}` })
  }
})

// ===== OCR 多图图片异步识别与流式日志通知 API =====
app.post('/api/projects/:name/ocr', upload.array('files'), async (req, res) => {
  const { name } = req.params
  const files = (req.files as Express.Multer.File[]) || []

  if (files.length === 0) {
    return res.status(400).json({ detail: '请上传至少一张图片' })
  }

  // 获取项目配置中 ocr_columns
  const settingsPath = path.join(getProjectsDir(), name, 'settings.json')
  let ocrColumns = ['项目名称', '单位', '数量', '不含税单价', '说明']
  if (fsSync.existsSync(settingsPath)) {
    try {
      const settings = JSON.parse(await fs.readFile(settingsPath, 'utf-8'))
      if (Array.isArray(settings.ocr_columns) && settings.ocr_columns.length > 0) {
        ocrColumns = settings.ocr_columns
      }
    } catch {}
  }

  const taskId = `task_${Date.now()}_${Math.random().toString(36).substr(2, 5)}`
  const initStep = `已接收 ${files.length} 张图片，正在进行图像识别预处理...`
  const initTime = new Date().toLocaleTimeString('zh-CN', { hour12: false })

  tasks[taskId] = {
    status: 'processing',
    progress: initStep,
    logs: [`[${initTime}] ${initStep}`],
    streamText: '',
  }
  res.json({ task_id: taskId });

  (async () => {
    try {
      const imageInputs: ImageInput[] = files.map(f => ({
        buffer: f.buffer,
        mimeType: f.mimetype,
      }))

      const ocrRes = await runOcrWithLlm(
        imageInputs,
        ocrColumns,
        undefined,
        (progressStep, rawChunk) => {
          if (!tasks[taskId]) return
          if (rawChunk) {
            tasks[taskId].streamText = (tasks[taskId].streamText || '') + rawChunk
          }
          if (progressStep) {
            tasks[taskId].progress = progressStep
            if (!tasks[taskId].logs) tasks[taskId].logs = []
            const timeStr = new Date().toLocaleTimeString('zh-CN', { hour12: false })
            const lastLog = tasks[taskId].logs[tasks[taskId].logs.length - 1]
            if (!lastLog || !lastLog.includes(progressStep)) {
              tasks[taskId].logs.push(`[${timeStr}] ${progressStep}`)
            }
          }
        }
      )

      if (!tasks[taskId]) return // 任务已被人工中止/取消

      if (!ocrRes.success) {
        const timeStr = new Date().toLocaleTimeString('zh-CN', { hour12: false })
        const errLog = `识别失败: ${ocrRes.error || '大模型响应异常'}`
        if (!tasks[taskId].logs) tasks[taskId].logs = []
        tasks[taskId].logs.push(`[${timeStr}] ${errLog}`)
        tasks[taskId].status = 'error'
        tasks[taskId].message = ocrRes.error || '识别失败'
        return
      }

      // 修复：ocrRes.data 结构为 { items: [...], remarks: "" }
      // 应只取 items 数组存入 task.result，否则前端取到的是对象导致 length 为 undefined→0 行
      const remarksList = Array.isArray(ocrRes.data?.remarks)
        ? ocrRes.data.remarks
        : (ocrRes.data?.remarks ? [String(ocrRes.data.remarks)] : []);

      const savePayload = {
        columns: ocrColumns,
        items: Array.isArray(ocrRes.data?.items) ? ocrRes.data.items : [],
        remarks: remarksList,
      }

      const itemCount = Array.isArray(ocrRes.data?.items) ? ocrRes.data.items.length : 0
      const doneStep = `识别成功！完成 ${files.length} 张图片解析，合并提取 ${itemCount} 行表格数据。`
      const timeStr = new Date().toLocaleTimeString('zh-CN', { hour12: false })

      if (!tasks[taskId].logs) tasks[taskId].logs = []
      tasks[taskId].logs.push(`[${timeStr}] ${doneStep}`)

      tasks[taskId].status = 'done'
      tasks[taskId].columns = ocrColumns
      tasks[taskId].progress = doneStep
      tasks[taskId].result = savePayload
    } catch (err: any) {
      if (!tasks[taskId]) return
      tasks[taskId].status = 'error'
      tasks[taskId].message = err.message || String(err)
    }
  })()
})

app.get('/api/tasks/:task_id', (req, res) => {
  const { task_id } = req.params
  const task = tasks[task_id]
  if (!task) {
    return res.status(404).json({ detail: '任务不存在' })
  }
  res.json(task)
})

app.delete('/api/tasks/:task_id', (req, res) => {
  const { task_id } = req.params
  if (tasks[task_id]) {
    delete tasks[task_id]
  }
  res.json({ success: true })
})

// ===== 协议定价表 API (Rate Cards) =====
app.get('/api/ratecards', async (req, res) => {
  try {
    const files = await fs.readdir(getRatecardDir())
    const ratecardNames = files.filter(f => f.endsWith('.json')).map(f => f.replace('.json', ''))
    res.json(ratecardNames)
  } catch {
    res.json([])
  }
})

app.post('/api/ratecards/:name', async (req, res) => {
  const { name } = req.params
  const filePath = path.join(getRatecardDir(), `${name}.json`)
  try {
    const payload = { columns: PRESET_COLUMNS, items: [] }
    await fs.writeFile(filePath, JSON.stringify(payload, null, 2), 'utf-8')
    res.json({ success: true })
  } catch (err: any) {
    res.status(500).json({ detail: err.message })
  }
})

app.get('/api/ratecards/:name', async (req, res) => {
  const { name } = req.params
  const filePath = path.join(getRatecardDir(), `${name}.json`)
  try {
    const content = await fs.readFile(filePath, 'utf-8')
    res.json(JSON.parse(content))
  } catch {
    res.json({ columns: PRESET_COLUMNS, items: [] })
  }
})

app.patch('/api/ratecards/:name/rename', async (req, res) => {
  const { name } = req.params
  let { new_name } = req.body
  if (!new_name || typeof new_name !== 'string') {
    return res.status(400).json({ detail: '新定价表名称不能为空' })
  }
  new_name = new_name.trim().replace(/\.json$/, '')
  const oldPath = path.join(getRatecardDir(), `${name}.json`)
  const newPath = path.join(getRatecardDir(), `${new_name}.json`)

  try {
    if (!fsSync.existsSync(oldPath)) {
      return res.status(404).json({ detail: '原定价表不存在' })
    }
    if (name !== new_name && fsSync.existsSync(newPath)) {
      return res.status(400).json({ detail: '已存在同名的定价表' })
    }
    await fs.rename(oldPath, newPath)

    // 自动更新所有项目 settings.json 中引用的 ratecard_name
    try {
      const projEntries = await fs.readdir(getProjectsDir(), { withFileTypes: true })
      for (const entry of projEntries) {
        if (entry.isDirectory()) {
          const sPath = path.join(getProjectsDir(), entry.name, 'settings.json')
          if (fsSync.existsSync(sPath)) {
            const sContent = await fs.readFile(sPath, 'utf-8')
            const projSettings = JSON.parse(sContent)
            if (
              projSettings.ratecard_name === `${name}.json` ||
              projSettings.ratecard_name === name
            ) {
              projSettings.ratecard_name = `${new_name}.json`
              await fs.writeFile(sPath, JSON.stringify(projSettings, null, 2), 'utf-8')
            }
          }
        }
      }
    } catch {}

    res.json({ success: true, name: new_name })
  } catch (err: any) {
    res.status(500).json({ detail: err.message })
  }
})

app.delete('/api/ratecards/:name', async (req, res) => {
  const { name } = req.params
  const cleanName = name.replace(/\.json$/, '')
  const filePath = path.join(getRatecardDir(), `${cleanName}.json`)
  try {
    if (fsSync.existsSync(filePath)) {
      await fs.unlink(filePath)
    }
    res.json({ success: true })
  } catch (err: any) {
    res.status(500).json({ detail: err.message })
  }
})

// 定价表导入预览 (解析 Excel/CSV 表头、自动填充合并单元格组名与示例数据)
app.post('/api/ratecards/:name/import-preview', upload.single('file'), async (req, res) => {
  const file = req.file
  if (!file) {
    return res.status(400).json({ detail: '请选择文件' })
  }

  try {
    const workbook = XLSX.read(file.buffer, { type: 'buffer' })
    const sheetName = workbook.SheetNames[0]
    const sheet = workbook.Sheets[sheetName]

    // 处理 Excel 合并单元格 (sheet['!merges'])，将左上角单元格的值填充到合并矩形区域内的所有单元格
    if (sheet && sheet['!merges'] && Array.isArray(sheet['!merges'])) {
      sheet['!merges'].forEach((range: any) => {
        const startCellRef = XLSX.utils.encode_cell(range.s)
        const startCell = sheet[startCellRef]
        if (!startCell) return
        for (let R = range.s.r; R <= range.e.r; ++R) {
          for (let C = range.s.c; C <= range.e.c; ++C) {
            if (R === range.s.r && C === range.s.c) continue
            const cellRef = XLSX.utils.encode_cell({ r: R, c: C })
            sheet[cellRef] = { ...startCell }
          }
        }
      })
    }

    const jsonData: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1 })

    if (jsonData.length === 0) {
      return res.status(400).json({ detail: '文件内容为空' })
    }

    // 第一行作为原始表头（去除首尾空格和感叹号/问号修饰符）
    const rawHeaders: string[] = (jsonData[0] || []).map((h, idx) => {
      const str = String(h ?? '').trim().replace(/^[!?]/, '')
      return str || `未命名列_${idx + 1}`
    })

    // 去重处理：确保每个原始列头名称绝对唯一，防止 React 渲染 Key 冲突及错位
    const headerCounts: Record<string, number> = {}
    const headers: string[] = rawHeaders.map(h => {
      if (!headerCounts[h]) {
        headerCounts[h] = 1
        return h
      } else {
        headerCounts[h]++
        return `${h}_${headerCounts[h]}`
      }
    })

    const sampleRows: Record<string, string>[] = []
    const allRows: Record<string, string>[] = []

    // 用于组名列向下继承填充 (Forward Fill)
    const lastGroupValues: Record<number, string> = {}

    for (let i = 1; i < jsonData.length; i++) {
      const row = jsonData[i]
      if (!row || row.length === 0) continue
      const item: Record<string, string> = {}
      let hasVal = false

      headers.forEach((h, colIdx) => {
        let val = String(row[colIdx] ?? '').trim()
        
        // 如果当前列值为空，且属于前面的组名列，尝试从上一行非空组名向下继承
        if (!val && colIdx < 3 && lastGroupValues[colIdx]) {
          val = lastGroupValues[colIdx]
        }

        if (val) {
          hasVal = true
          lastGroupValues[colIdx] = val
        }
        item[h] = val
      })

      if (hasVal) {
        allRows.push(item)
        if (sampleRows.length < 5) sampleRows.push(item)
      }
    }

    res.json({ headers, sampleRows, allRows })
  } catch (err: any) {
    res.status(500).json({ detail: '解析 Excel/CSV 文件失败: ' + err.message })
  }
})

// 确认导入保存定价表 JSON（保存纯净的标准 columns 和 items 数据）
app.post('/api/ratecards/:name/import', async (req, res) => {
  const { name } = req.params
  const { headers, items, mapping } = req.body

  if (!Array.isArray(items) || !mapping) {
    return res.status(400).json({ detail: '缺少导入数据或列映射配置' })
  }

  try {
    // 过滤出被合法映射到 10 项内置标准列名的集合
    const usedPresetCols = Array.from(new Set(Object.values(mapping).filter((v: any) => Boolean(v) && PRESET_COLUMNS.includes(v as any)))) as string[]
    const targetColumns = PRESET_COLUMNS.filter(col => usedPresetCols.includes(col))

    if (targetColumns.length === 0) {
      return res.status(400).json({ detail: '未选择任何有效的系统内置标准列映射' })
    }

    const mappedItems: Record<string, string>[] = []

    for (const rawItem of items) {
      const newItem: Record<string, string> = {}
      let hasValue = false

      for (const [rawHeader, presetCol] of Object.entries(mapping)) {
        if (presetCol && typeof presetCol === 'string' && PRESET_COLUMNS.includes(presetCol as any)) {
          const val = String(rawItem[rawHeader] ?? '').trim()
          if (val) hasValue = true
          newItem[presetCol] = val
        }
      }

      if (hasValue) {
        mappedItems.push(newItem)
      }
    }

    const filePath = path.join(getRatecardDir(), `${name}.json`)
    const payload = {
      columns: targetColumns,
      items: mappedItems,
    }

    await fs.writeFile(filePath, JSON.stringify(payload, null, 2), 'utf-8')
    res.json({ success: true, columns: payload.columns, count: mappedItems.length })
  } catch (err: any) {
    res.status(500).json({ detail: err.message })
  }
})

// ===== 协议定价表匹配搜索 API =====
app.post('/api/match', async (req, res) => {
  const { ratecard_name, queries, limit = 5 } = req.body
  const result: Record<string, any[]> = {}

  if (!ratecard_name || !Array.isArray(queries)) {
    return res.json(result)
  }

  const normalizedFileName = ratecard_name.endsWith('.json') ? ratecard_name : `${ratecard_name}.json`
  const filePath = path.join(getRatecardDir(), normalizedFileName)

  let ratecardItems: Record<string, string>[] = []
  if (fsSync.existsSync(filePath)) {
    try {
      const content = await fs.readFile(filePath, 'utf-8')
      const data = JSON.parse(content)
      ratecardItems = Array.isArray(data.items) ? data.items : []
    } catch {}
  }

  for (const q of queries) {
    const term = String(q || '').trim().toLowerCase()
    if (!term) {
      result[q] = []
      continue
    }

    // 对比项名 (项目名称)
    const matches: { item: Record<string, string>; score: number }[] = []
    for (const item of ratecardItems) {
      const itemName = String(item['项目名称'] || '').trim().toLowerCase()
      if (!itemName) continue

      let score = 0
      if (itemName === term) {
        score = 100
      } else if (itemName.includes(term) || term.includes(itemName)) {
        score = 85
      } else {
        // 计算简单字符交集比例
        let matchCount = 0
        for (const char of term) {
          if (itemName.includes(char)) matchCount++
        }
        const ratio = matchCount / Math.max(term.length, itemName.length)
        if (ratio > 0.4) score = Math.round(ratio * 80)
      }

      if (score > 30) {
        matches.push({ item, score })
      }
    }

    matches.sort((a, b) => b.score - a.score)
    result[q] = matches.slice(0, limit).map(m => [m.item['项目名称'], m.score, m.item])
  }

  res.json(result)
})

// ===== 模板 API =====
app.get('/api/templates', async (req, res) => {
  try {
    const files = await fs.readdir(getTemplateDir())
    res.json(files.filter(f => f.endsWith('.xlsx')))
  } catch {
    res.json([])
  }
})

app.post('/api/templates', upload.single('file'), async (req, res) => {
  const file = req.file
  if (!file) return res.status(400).json({ detail: '请上传模板文件' })
  const filename = file.originalname
  try {
    await fs.writeFile(path.join(getTemplateDir(), filename), file.buffer)
    res.json({ success: true, filename })
  } catch (err: any) {
    res.status(500).json({ detail: err.message })
  }
})

app.delete('/api/templates/:filename', async (req, res) => {
  const { filename } = req.params
  try {
    await fs.unlink(path.join(getTemplateDir(), filename))
    res.json({ success: true })
  } catch (err: any) {
    res.status(500).json({ detail: err.message })
  }
})

// 智能定位托管的前端网页静态产物（兼容 EXE 运行、CWD 变动及 CLI 模式）
function findClientDistPath(): string {
  const exeDir = path.dirname(process.execPath)
  const candidates = [
    path.join(exeDir, 'client'),
    path.join(exeDir, 'dist', 'client'),
    path.join(process.cwd(), 'dist', 'client'),
    path.join(process.cwd(), 'client'),
    path.join(__dirname, '..', 'client'),
    path.join(__dirname, 'client')
  ]

  for (const candidate of candidates) {
    if (fsSync.existsSync(path.join(candidate, 'index.html'))) {
      return candidate
    }
  }
  return path.join(process.cwd(), 'dist', 'client')
}

const clientDistPath = findClientDistPath()

if (fsSync.existsSync(clientDistPath)) {
  app.use(express.static(clientDistPath))
}

app.get('*', (req, res, next) => {
  if (req.path.startsWith('/api')) return next()
  const indexPath = path.join(clientDistPath, 'index.html')
  if (fsSync.existsSync(indexPath)) {
    res.sendFile(indexPath)
  } else {
    // 开发模式下未找到静态打包产物时，自动跳转到 Vite 开发前端服务 (5173 端口)
    res.redirect('http://localhost:5173')
  }
})

export function startServerInstance(desiredPort?: number): Promise<{ port: number; url: string; server: any }> {
  return new Promise(async (resolve, reject) => {
    // 动态初始化数据目录结构
    await initDirs()
    // 未指定端口时默认设为 0 (由系统自动分配未占用的随机可用端口)
    const targetPort = desiredPort ?? (process.env.PORT ? Number(process.env.PORT) : 0)

    const listenOnPort = (port: number) => {
      const server = app.listen(port, async () => {
        const address = server.address()
        const actualPort = typeof address === 'object' && address ? address.port : port
        const url = `http://localhost:${actualPort}`
        console.log(`CarrotMRO 全栈服务已在端口 ${actualPort} 成功启动: ${url}`)

        if (process.argv.includes('--open')) {
          try {
            await open(url)
            console.log(`已在默认浏览器打开: ${url}`)
          } catch (err) {
            console.error('无法自动打开浏览器:', err)
          }
        }

        resolve({ port: actualPort, url, server })
      })

      server.on('error', (err: any) => {
        if (err.code === 'EADDRINUSE' && targetPort !== 0) {
          console.warn(`⚠️ 端口 ${port} 已被占用，自动尝试递增端口 ${port + 1}...`)
          listenOnPort(port + 1)
        } else {
          reject(err)
        }
      })
    }

    listenOnPort(targetPort)
  })
}

export { app }

if (process.env.NODE_ENV !== 'test' && !process.env.CARROTMRO_MANUAL_START) {
  startServerInstance(process.env.PORT ? Number(process.env.PORT) : 3000).catch((err) => {
    console.error('CarrotMRO 启动失败:', err)
  })
}

