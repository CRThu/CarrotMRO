import express from 'express'
import cors from 'cors'
import path from 'path'
import { fileURLToPath } from 'url'
import fs from 'fs/promises'
import fsSync from 'fs'
import open from 'open'
import multer from 'multer'
import * as XLSX from 'xlsx'
import { loadSettings, getSettings, saveSettings } from './services/settings.js'
import { runOcrWithLlm, testLlmConnection, ImageInput } from './services/llm.js'

const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)

const app = express()
const PORT = process.env.PORT || 3000
const upload = multer({ storage: multer.memoryStorage() })

app.use(cors())
app.use(express.json({ limit: '50mb' }))

// 数据目录（支持本地 JSON 文件管理）
const DATA_DIR = path.join(process.cwd(), 'data')
const PROJECTS_DIR = path.join(DATA_DIR, 'projects')
const RATECARD_DIR = path.join(DATA_DIR, 'ratecard')
const TEMPLATE_DIR = path.join(DATA_DIR, 'template')

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
  }
> = {}

// 初始化数据目录与自动加载持久化配置
async function initDirs() {
  try {
    await fs.mkdir(PROJECTS_DIR, { recursive: true })
    await fs.mkdir(RATECARD_DIR, { recursive: true })
    await fs.mkdir(TEMPLATE_DIR, { recursive: true })
    await loadSettings()
    console.log('数据目录及 settings.json 成功自动加载！')
  } catch (err) {
    console.error('初始化数据目录/设置失败:', err)
  }
}
initDirs()

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
    const entries = await fs.readdir(PROJECTS_DIR, { withFileTypes: true })
    const projectNames = entries.filter(e => e.isDirectory()).map(e => e.name)
    res.json(projectNames)
  } catch (err) {
    res.json([])
  }
})

app.post('/api/projects/:name', async (req, res) => {
  const { name } = req.params
  const projPath = path.join(PROJECTS_DIR, name)
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
    }
    await fs.writeFile(settingsPath, JSON.stringify(settingsData, null, 2), 'utf-8')
    res.json({ success: true, name })
  } catch (err: any) {
    res.status(500).json({ detail: err.message })
  }
})

app.get('/api/projects/:name', async (req, res) => {
  const { name } = req.params
  const settingsPath = path.join(PROJECTS_DIR, name, 'settings.json')
  const legacyConfigPath = path.join(PROJECTS_DIR, name, 'project.json')

  let settings = {
    name,
    created_at: new Date().toISOString(),
    ratecard_name: null as string | null,
    template_name: null as string | null,
    ocr_columns: ['项目名称', '单位', '数量', '不含税单价', '说明'],
    quotation_columns: PRESET_COLUMNS,
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
      await fs.mkdir(path.join(PROJECTS_DIR, name), { recursive: true })
      await fs.writeFile(settingsPath, JSON.stringify(settings, null, 2), 'utf-8')
    }
  } catch {}

  res.json(settings)
})

app.patch('/api/projects/:name/settings', async (req, res) => {
  const { name } = req.params
  const projPath = path.join(PROJECTS_DIR, name)
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
  const settingsPath = path.join(PROJECTS_DIR, name, 'settings.json')
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
  const settingsPath = path.join(PROJECTS_DIR, name, 'settings.json')
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
  const projPath = path.join(PROJECTS_DIR, name)
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
  const projPath = path.join(PROJECTS_DIR, name)
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
  const filePath = path.join(PROJECTS_DIR, name, file)
  try {
    const content = await fs.readFile(filePath, 'utf-8')
    res.json(JSON.parse(content))
  } catch (err: any) {
    res.status(500).json({ detail: err.message })
  }
})

app.put('/api/projects/:name/quotations/:file', async (req, res) => {
  const { name, file } = req.params
  const filePath = path.join(PROJECTS_DIR, name, file)
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
  const filePath = path.join(PROJECTS_DIR, name, file)
  try {
    await fs.unlink(filePath)
    res.json({ success: true })
  } catch (err: any) {
    res.status(500).json({ detail: err.message })
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
  const settingsPath = path.join(PROJECTS_DIR, name, 'settings.json')
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
        (progressStep) => {
          if (!tasks[taskId]) return
          tasks[taskId].progress = progressStep
          if (!tasks[taskId].logs) tasks[taskId].logs = []
          const timeStr = new Date().toLocaleTimeString('zh-CN', { hour12: false })
          const lastLog = tasks[taskId].logs[tasks[taskId].logs.length - 1]
          if (!lastLog || !lastLog.includes(progressStep)) {
            tasks[taskId].logs.push(`[${timeStr}] ${progressStep}`)
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
      const savePayload = {
        columns: ocrColumns,
        items: Array.isArray(ocrRes.data?.items) ? ocrRes.data.items : [],
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
    const files = await fs.readdir(RATECARD_DIR)
    const ratecardNames = files.filter(f => f.endsWith('.json')).map(f => f.replace('.json', ''))
    res.json(ratecardNames)
  } catch {
    res.json([])
  }
})

app.post('/api/ratecards/:name', async (req, res) => {
  const { name } = req.params
  const filePath = path.join(RATECARD_DIR, `${name}.json`)
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
  const filePath = path.join(RATECARD_DIR, `${name}.json`)
  try {
    const content = await fs.readFile(filePath, 'utf-8')
    res.json(JSON.parse(content))
  } catch {
    res.json({ columns: PRESET_COLUMNS, items: [] })
  }
})

// 定价表导入预览 (解析 Excel/CSV 表头与示例数据)
app.post('/api/ratecards/:name/import-preview', upload.single('file'), async (req, res) => {
  const file = req.file
  if (!file) {
    return res.status(400).json({ detail: '请选择文件' })
  }

  try {
    const workbook = XLSX.read(file.buffer, { type: 'buffer' })
    const sheetName = workbook.SheetNames[0]
    const sheet = workbook.Sheets[sheetName]
    const jsonData: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1 })

    if (jsonData.length === 0) {
      return res.status(400).json({ detail: '文件内容为空' })
    }

    // 第一行作为原始表头（去除首尾空格和感叹号/问号修饰符）
    const rawHeaders: string[] = (jsonData[0] || []).map(h => String(h || '').trim().replace(/^[!?]/, ''))
    const headers = rawHeaders.filter(h => h.length > 0)

    const sampleRows: Record<string, string>[] = []
    const allRows: Record<string, string>[] = []

    for (let i = 1; i < jsonData.length; i++) {
      const row = jsonData[i]
      if (!row || row.length === 0) continue
      const item: Record<string, string> = {}
      let hasVal = false
      headers.forEach((h, colIdx) => {
        const val = String(row[colIdx] ?? '').trim()
        if (val) hasVal = true
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

// 最终确认映射导入保存定价表 JSON
app.post('/api/ratecards/:name/import', async (req, res) => {
  const { name } = req.params
  const { headers, items, mapping } = req.body

  if (!Array.isArray(items) || !mapping) {
    return res.status(400).json({ detail: '缺少导入数据或映射配置' })
  }

  try {
    // 确定被映射出来的目标列名集合
    const usedPresetCols = Array.from(new Set(Object.values(mapping).filter((v: any) => Boolean(v)))) as string[]
    const targetColumns = PRESET_COLUMNS.filter(col => usedPresetCols.includes(col))

    const mappedItems: Record<string, string>[] = []

    for (const rawItem of items) {
      const newItem: Record<string, string> = {}
      let hasValue = false

      for (const [rawHeader, presetCol] of Object.entries(mapping)) {
        if (presetCol && typeof presetCol === 'string') {
          const val = String(rawItem[rawHeader] ?? '').trim()
          if (val) hasValue = true
          newItem[presetCol] = val
        }
      }

      // 如果有有效字段且项目名称不为空，保留该行
      if (hasValue && (newItem['项目名称'] || Object.keys(newItem).length > 0)) {
        mappedItems.push(newItem)
      }
    }

    const filePath = path.join(RATECARD_DIR, `${name}.json`)
    const payload = {
      columns: targetColumns.length > 0 ? targetColumns : PRESET_COLUMNS,
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

  const filePath = path.join(RATECARD_DIR, `${ratecard_name}.json`)
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
    const files = await fs.readdir(TEMPLATE_DIR)
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
    await fs.writeFile(path.join(TEMPLATE_DIR, filename), file.buffer)
    res.json({ success: true, filename })
  } catch (err: any) {
    res.status(500).json({ detail: err.message })
  }
})

app.delete('/api/templates/:filename', async (req, res) => {
  const { filename } = req.params
  try {
    await fs.unlink(path.join(TEMPLATE_DIR, filename))
    res.json({ success: true })
  } catch (err: any) {
    res.status(500).json({ detail: err.message })
  }
})

// 托管包内打包好的前端网页静态产物
const clientDistPath = fsSync.existsSync(path.join(process.cwd(), 'dist', 'client'))
  ? path.join(process.cwd(), 'dist', 'client')
  : path.join(__dirname, '..', 'client')

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

app.listen(PORT, async () => {
  const url = `http://localhost:${PORT}`
  console.log(`CarrotMRO 全栈服务已启动: ${url}`)
  
  if (process.argv.includes('--open')) {
    try {
      await open(url)
      console.log(`已在默认浏览器打开: ${url}`)
    } catch (err) {
      console.error('无法自动打开浏览器:', err)
    }
  }
})
