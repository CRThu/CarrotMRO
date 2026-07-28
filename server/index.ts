import express from 'express'
import cors from 'cors'
import path from 'path'
import { fileURLToPath } from 'url'
import fs from 'fs/promises'
import fsSync from 'fs'
import open from 'open'
import multer from 'multer'
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
const tasks: Record<string, { status: 'processing' | 'done' | 'error'; file?: string; columns?: string[]; result?: any; message?: string }> = {}

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

// 全局预制列
const PRESET_COLUMNS = [
  '序号', '项目名称', '规格型号', '单位', '数量', '综合单价', '合价',
  '品牌', '产地', '备注', '税率'
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

// API 占位/功能路由定义
app.get('/api/preset-columns', (req, res) => {
  res.json(PRESET_COLUMNS)
})

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
    const configPath = path.join(projPath, 'project.json')
    const configData = {
      name,
      created_at: new Date().toISOString(),
      ratecard_name: null,
      template_name: null,
      selected_columns: PRESET_COLUMNS.slice(0, 5)
    }
    await fs.writeFile(configPath, JSON.stringify(configData, null, 2), 'utf-8')
    res.json({ success: true, name })
  } catch (err: any) {
    res.status(500).json({ detail: err.message })
  }
})

app.get('/api/projects/:name', async (req, res) => {
  const { name } = req.params
  const configPath = path.join(PROJECTS_DIR, name, 'project.json')
  try {
    const content = await fs.readFile(configPath, 'utf-8')
    res.json(JSON.parse(content))
  } catch {
    res.json({
      name,
      ratecard_name: null,
      template_name: null,
      selected_columns: PRESET_COLUMNS.slice(0, 5)
    })
  }
})

app.get('/api/projects/:name/columns', async (req, res) => {
  const { name } = req.params
  const configPath = path.join(PROJECTS_DIR, name, 'project.json')
  try {
    const content = await fs.readFile(configPath, 'utf-8')
    const config = JSON.parse(content)
    res.json({
      available_columns: PRESET_COLUMNS,
      selected_columns: config.selected_columns || PRESET_COLUMNS.slice(0, 5),
      column_mappings: config.column_mappings || { ocr: {}, ratecard: {}, quotation: {} }
    })
  } catch {
    res.json({
      available_columns: PRESET_COLUMNS,
      selected_columns: PRESET_COLUMNS.slice(0, 5),
      column_mappings: { ocr: {}, ratecard: {}, quotation: {} }
    })
  }
})

app.patch('/api/projects/:name/columns', async (req, res) => {
  const { name } = req.params
  const { columns, scope, column_mapping } = req.body
  const configPath = path.join(PROJECTS_DIR, name, 'project.json')
  try {
    let config: any = {}
    if (fsSync.existsSync(configPath)) {
      config = JSON.parse(await fs.readFile(configPath, 'utf-8'))
    }
    if (columns) config.selected_columns = columns
    if (scope && column_mapping) {
      if (!config.column_mappings) config.column_mappings = { ocr: {}, ratecard: {}, quotation: {} }
      config.column_mappings[scope] = column_mapping
    }
    await fs.writeFile(configPath, JSON.stringify(config, null, 2), 'utf-8')
    res.json({ success: true, columns: config.selected_columns })
  } catch (err: any) {
    res.status(500).json({ detail: err.message })
  }
})

// ===== OCR 图片异步识别 API =====
app.post('/api/projects/:name/ocr', upload.array('files'), async (req, res) => {
  const { name } = req.params
  const files = (req.files as Express.Multer.File[]) || []

  if (files.length === 0) {
    return res.status(400).json({ detail: '请上传至少一张图片' })
  }

  // 获取项目配置的识别列
  const configPath = path.join(PROJECTS_DIR, name, 'project.json')
  let selectedColumns = PRESET_COLUMNS.slice(0, 5)
  if (fsSync.existsSync(configPath)) {
    try {
      const config = JSON.parse(await fs.readFile(configPath, 'utf-8'))
      if (config.selected_columns && config.selected_columns.length > 0) {
        selectedColumns = config.selected_columns
      }
    } catch {}
  }

  const taskId = `task_${Date.now()}_${Math.random().toString(36).substr(2, 5)}`
  tasks[taskId] = { status: 'processing' }
  res.json({ task_id: taskId });

  // 异步执行识别逻辑
  (async () => {

    try {
      const imageInputs: ImageInput[] = files.map(f => ({
        buffer: f.buffer,
        mimeType: f.mimetype,
      }))

      const ocrRes = await runOcrWithLlm(imageInputs, selectedColumns)

      if (!ocrRes.success) {
        tasks[taskId] = { status: 'error', message: ocrRes.error || '识别失败' }
        return
      }

      // 保存 ocr 文件结果
      const projPath = path.join(PROJECTS_DIR, name)
      await fs.mkdir(projPath, { recursive: true })
      const existingFiles = await fs.readdir(projPath)
      const ocrFiles = existingFiles.filter(f => f.startsWith('ocr-') && f.endsWith('.json'))
      const fileIndex = ocrFiles.length + 1
      const filename = `ocr-${fileIndex}.json`

      const savePayload = {
        columns: selectedColumns,
        data: ocrRes.data,
      }

      await fs.writeFile(path.join(projPath, filename), JSON.stringify(savePayload, null, 2), 'utf-8')

      tasks[taskId] = {
        status: 'done',
        file: filename,
        columns: selectedColumns,
        result: savePayload,
      }
    } catch (err: any) {
      tasks[taskId] = { status: 'error', message: err.message || String(err) }
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

app.get('/api/projects/:name/ocr-files', async (req, res) => {
  const { name } = req.params
  const projPath = path.join(PROJECTS_DIR, name)
  try {
    const files = await fs.readdir(projPath)
    res.json({ files: files.filter(f => f.startsWith('ocr-') && f.endsWith('.json')) })
  } catch {
    res.json({ files: [] })
  }
})

app.get('/api/projects/:name/ocr-files/:file', async (req, res) => {
  const { name, file } = req.params
  const filePath = path.join(PROJECTS_DIR, name, file)
  try {
    const content = await fs.readFile(filePath, 'utf-8')
    res.json(JSON.parse(content))
  } catch (err: any) {
    res.status(500).json({ detail: err.message })
  }
})

app.put('/api/projects/:name/ocr-files/:file', async (req, res) => {
  const { name, file } = req.params
  const filePath = path.join(PROJECTS_DIR, name, file)
  try {
    await fs.writeFile(filePath, JSON.stringify(req.body, null, 2), 'utf-8')
    res.json({ success: true })
  } catch (err: any) {
    res.status(500).json({ detail: err.message })
  }
})

app.delete('/api/projects/:name/ocr-files/:file', async (req, res) => {
  const { name, file } = req.params
  const filePath = path.join(PROJECTS_DIR, name, file)
  try {
    await fs.unlink(filePath)
    res.json({ success: true })
  } catch (err: any) {
    res.status(500).json({ detail: err.message })
  }
})

app.get('/api/ratecards', async (req, res) => {
  try {
    const files = await fs.readdir(RATECARD_DIR)
    const ratecardNames = files.filter(f => f.endsWith('.json')).map(f => f.replace('.json', ''))
    res.json(ratecardNames)
  } catch {
    res.json([])
  }
})

app.get('/api/templates', async (req, res) => {
  try {
    const files = await fs.readdir(TEMPLATE_DIR)
    res.json(files.filter(f => f.endsWith('.xlsx')))
  } catch {
    res.json([])
  }
})

app.post('/api/match', (req, res) => {
  const { queries } = req.body
  const result: Record<string, any[]> = {}
  if (Array.isArray(queries)) {
    for (const q of queries) {
      result[q] = [
        [`样例匹配项 A (${q})`, 95],
        [`样例匹配项 B (${q})`, 82]
      ]
    }
  }
  res.json(result)
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
