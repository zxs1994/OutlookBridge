const { app, BrowserWindow } = require('electron')
const path = require('path')
const createOutlookMailMac = require('./mac')
const createOutlookMailWindows = require('./win')

const isMac = process.platform === 'darwin'
const isWin = process.platform === 'win32'
const isDev = !app.isPackaged

let mainWindow = null
let firstProtocolArg = null

function createMainWindow() {
  mainWindow = new BrowserWindow({
    width: 600,
    height: 400,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
    }
  })

  mainWindow.loadFile(path.join(__dirname, 'renderer/index.html'))
  mainWindow.setMenuBarVisibility(false)

  mainWindow.webContents.on('did-fail-load', (event, code, desc) => {
    console.error('❌ 页面加载失败:', code, desc)
  })

  mainWindow.webContents.on('did-finish-load', () => {
    if (firstProtocolArg) {
      handleProtocol(firstProtocolArg)
      firstProtocolArg = null
    } else {
      logToWindow('👋 欢迎使用 OutlookBridge！请通过浏览器中的 outlookbridge:// 协议触发邮件发送。')
    }
  })
}

function logToWindow(message) {
  if (mainWindow && mainWindow.webContents) {
    mainWindow.webContents.send('log', message)
  }
  console.log(message)
}

console.log('启动参数:', process.argv)

if (!isDev) {
  if (!app.isDefaultProtocolClient('outlookbridge')) {
    const protocolArgs = isWin && process.argv[1] ? [path.resolve(process.argv[1])] : undefined
    app.setAsDefaultProtocolClient('outlookbridge', app.getPath('exe'), protocolArgs)
  }
} else {
  console.log('[开发模式] 请使用 OUTLOOKBRIDGE_URL 环境变量模拟 outlookbridge:// 协议')
}

if (isMac) {
  app.on('open-url', (event, urlStr) => {
    event.preventDefault()
    handleProtocol(urlStr)
  })
}

const gotLock = app.requestSingleInstanceLock()
if (!gotLock) {
  app.quit()
} else {
  app.on('second-instance', (event, commandLine) => {
    const protocolArg = commandLine.find(arg => arg.startsWith('outlookbridge://'))
    if (protocolArg) handleProtocol(protocolArg)

    if (mainWindow) {
      if (mainWindow.isMinimized()) mainWindow.restore()
      mainWindow.focus()
    }
  })

  app.whenReady().then(() => {
    const protocolArg = process.argv.find(arg => arg.startsWith('outlookbridge://'))
    if (protocolArg && !isDev) {
      firstProtocolArg = protocolArg // 页面准备好后再处理
    }

    createMainWindow()

    if (isDev && process.env.OUTLOOKBRIDGE_URL?.startsWith('outlookbridge://')) {
      handleProtocol(process.env.OUTLOOKBRIDGE_URL)
    }
  })
}

function handleProtocol(urlStr) {
  try {
    const rawUrl = decodeURIComponent(urlStr)
    logToWindow('[协议处理] URL:' + rawUrl)
    const url = new URL(rawUrl)
    const params = Object.fromEntries(url.searchParams.entries())

    if (!params.email) {
      logToWindow('❌ 缺少 email 参数')
      return
    }

    const fn = isMac ? createOutlookMailMac : isWin ? createOutlookMailWindows : null
    if (!fn) {
      logToWindow('❌ 当前系统不支持发送 Outlook 邮件')
      return
    }

    fn({
      to: params.email,
      subject: params.subject || '无主题',
      body: params.body || '',
      attachments: params.attachments?.trim()
        ? params.attachments.split(',').map(s => s.trim())
        : []
    }, logToWindow)
  } catch (err) {
    logToWindow(`❌ 协议处理失败: ${err.message}`)
  }
}

process.on('uncaughtException', (err) => {
  logToWindow('💥 未捕获异常: ' + err.message)
})