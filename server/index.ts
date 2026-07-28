import express from 'express'
import cors from 'cors'
import path from 'path'
import { fileURLToPath } from 'url'
import fs from 'fs/promises'
import fsSync from 'fs'
import open from 'open'

const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)

const app = express()
const PORT = process.env.PORT || 3000

app.use(cors())
app.use(express.json({ limit: '50mb' }))

// 数据目录（支持本地 JSON 文件管理）
const DATA_DIR = path.join(process.cwd(), 'data')
const PROJECTS_DIR = path.join(DATA_DIR, 'projects')
const RATECARD_DIR = path.join(DATA_DIR, 'ratecard')
const TEMPLATE_DIR = path.join(DATA_DIR, 'template')

// 初始化数据目录
async function initDirs() {
  try {
    await fs.mkdir(PROJECTS_DIR, { recursive: true })
    await fs.mkdir(RATECARD_DIR, { recursive: true })
    await fs.mkdir(TEMPLATE_DIR, { recursive: true })
  } catch (err) {
    console.error('初始化数据目录失败:', err)
  }
}
initDirs()

// 全局预制列
const PRESET_COLUMNS = [
  '序号', '项目名称', '规格型号', '单位', '数量', '综合单价', '合价',
  '品牌', '产地', '备注', '税率'
]

// API 占位符路由定义
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

app.get('/api/projects/:name/columns', (req, res) => {
  res.json({
    available_columns: PRESET_COLUMNS,
    selected_columns: PRESET_COLUMNS.slice(0, 5),
    column_mappings: { ocr: {}, ratecard: {}, quotation: {} }
  })
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
const clientDistPath = path.join(__dirname, '..', 'client')
app.use(express.static(clientDistPath))

app.get('*', (req, res, next) => {
  if (req.path.startsWith('/api')) return next()
  const indexPath = path.join(clientDistPath, 'index.html')
  if (fsSync.existsSync(indexPath)) {
    res.sendFile(indexPath)
  } else {
    res.send(`<h1>CarrotMRO 全栈服务已就绪</h1><p>端口: ${PORT}</p>`)
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
