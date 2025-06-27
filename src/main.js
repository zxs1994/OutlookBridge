const { app, BrowserWindow } = require('electron')
const path = require('path')
const createOutlookMailMac = require('./mac')
const createOutlookMailWindows = require('./win')

const isMac = process.platform === 'darwin'
const isWin = process.platform === 'win32'
const isDev = !app.isPackaged

let mainWindow = null

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
}

function logToWindow(message) {
  if (mainWindow && mainWindow.webContents) {
    mainWindow.webContents.send('log', message)
  }
  console.log(message)
}

// 🟡 日志打印启动参数
logToWindow('启动参数:' + JSON.stringify(process.argv))

// ✅ 注册协议（仅生产）
if (!isDev) {
  if (!app.isDefaultProtocolClient('outlookbridge')) {
    const protocolArgs = isWin && process.argv[1] ? [path.resolve(process.argv[1])] : undefined
    app.setAsDefaultProtocolClient(
      'outlookbridge',
      process.execPath,
      protocolArgs
    )
  }
} else {
  logToWindow('[开发模式] 请使用 OUTLOOKBRIDGE_URL 环境变量模拟 outlookbridge:// 协议')
}

// ✅ Mac 上 open-url 事件
if (isMac) {
  app.on('open-url', (event, urlStr) => {
    event.preventDefault()
    handleProtocol(urlStr)
  })
}

// ✅ 防止多开，second-instance 接收参数
const gotLock = app.requestSingleInstanceLock()
if (!gotLock) {
  app.quit()
} else {
  app.on('second-instance', (event, commandLine) => {
    logToWindow('[second-instance] 参数:' + JSON.stringify(commandLine))
    const protocolArg = commandLine.find(arg => arg.startsWith('outlookbridge://'))
    if (protocolArg) handleProtocol(protocolArg)
  })
}

// ✅ App 准备好后处理首次启动的参数
app.whenReady().then(() => {
  createMainWindow()
  const protocolArg = process.argv.find(arg => arg.startsWith('outlookbridge://'))

  // ✅ 如果通过协议启动，等待 second-instance 处理，不在主进程重复调用
  if (protocolArg && !isDev) {
    logToWindow('[首次启动] 收到协议参数，等待 second-instance 处理')
    return
  }

  if (isDev) {
    const fn = isMac ? createOutlookMailMac : isWin ? createOutlookMailWindows : null
    fn({
      to: 'xusheng94@qq.com',
      subject: '无主题',
      body: '123',
      // attachments: ['https://pic.netbian.com/uploads/allimg/250121/231514-173747251455f8.jpg']
    }, logToWindow)
  }
})

/**
 * 统一处理 outlookbridge:// 协议
 * @param {string} urlStr
 */
function handleProtocol(urlStr) {
  try {
    const rawUrl = decodeURIComponent(urlStr)
    logToWindow('[协议处理] URL:' + rawUrl)
    const url = new URL(rawUrl)
    const params = Object.fromEntries(url.searchParams.entries())

    if (!params.email) {
      // 如果没有 email 参数，弹出提示框
      logToWindow('❌ 缺少 email 参数')
      return
    }

    const fn = isMac ? createOutlookMailMac : isWin ? createOutlookMailWindows : null
    if (!fn) {
      // 如果不是 Mac 或 Windows，提示不支持
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
    // 捕获协议处理中的错误
    logToWindow(`❌ 协议处理失败: ${err.message}`)
  }
}


// 捕获未处理异常
process.on('uncaughtException', (err) => {
  logToWindow('💥 未捕获异常:' + err)
})