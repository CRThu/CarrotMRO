import { app, BrowserWindow } from 'electron'
import path from 'path'
import { fileURLToPath } from 'url'
import { startServerInstance } from '../server/index.js'

const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)

// 在模块最顶层（在任何导入之前）准确锁定 PORTABLE_EXECUTABLE_DIR
const isPackagedExecutable = !path.basename(process.execPath).toLowerCase().includes('electron')
if (isPackagedExecutable && !process.env.PORTABLE_EXECUTABLE_DIR) {
  process.env.PORTABLE_EXECUTABLE_DIR = path.dirname(process.execPath)
}

let mainWindow: BrowserWindow | null = null

// 动态启动 Express REST API 服务（支持动态端口获取）
async function startServer(): Promise<string> {
  try {
    process.env.CARROTMRO_MANUAL_START = 'true'
    if (!process.env.PORTABLE_EXECUTABLE_DIR && app.isPackaged) {
      process.env.PORTABLE_EXECUTABLE_DIR = path.dirname(process.execPath)
    }
    const res = await startServerInstance()
    return res.url
  } catch (err) {
    console.error('Electron 主进程拉起 Express 服务异常:', err)
  }
  return 'http://localhost:3000'
}

async function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1366,
    height: 868,
    minWidth: 1024,
    minHeight: 700,
    title: 'CarrotMRO 综合管理系统',
    autoHideMenuBar: true,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
    },
  })

  // 隐藏系统默认原生菜单，保持洁净界面
  mainWindow.setMenu(null)

  // 启动服务获取实际监听成功的 URL
  const appUrl = await startServer()

  mainWindow.webContents.on('did-fail-load', () => {
    setTimeout(() => {
      mainWindow?.loadURL(appUrl)
    }, 500)
  })

  mainWindow.loadURL(appUrl)

  mainWindow.on('closed', () => {
    mainWindow = null
  })
}

app.whenReady().then(createWindow)

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit()
  }
})

app.on('activate', () => {
  if (mainWindow === null) {
    createWindow()
  }
})
