const { app } = require('electron')
const { execSync } = require('child_process')
const fs = require('fs')
const os = require('os')
const path = require('path')

const isMac = process.platform === 'darwin'
const isWin = process.platform === 'win32'
const isDev = !app.isPackaged

// 🟡 日志打印启动参数
console.log('启动参数:', process.argv)

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
  console.log('[开发模式] 请使用 OUTLOOKBRIDGE_URL 环境变量模拟 outlookbridge:// 协议')
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
    console.log('[second-instance] 参数:', commandLine)
    const protocolArg = commandLine.find(arg => arg.startsWith('outlookbridge://'))
    if (protocolArg) handleProtocol(protocolArg)
  })
}

// ✅ App 准备好后处理首次启动的参数
app.whenReady().then(() => {
  const protocolArg = process.argv.find(arg => arg.startsWith('outlookbridge://'))

  // ✅ 如果通过协议启动，等待 second-instance 处理，不在主进程重复调用
  if (protocolArg && !isDev) {
    console.log('[首次启动] 收到协议参数，等待 second-instance 处理')
    return
  }

  if (isDev && process.env.OUTLOOKBRIDGE_URL?.startsWith('outlookbridge://')) {
    console.log('[开发模拟] 处理协议:', process.env.OUTLOOKBRIDGE_URL)
    handleProtocol(process.env.OUTLOOKBRIDGE_URL)
  }
})

/**
 * 统一处理 outlookbridge:// 协议
 * @param {string} urlStr
 */
function handleProtocol(urlStr) {
  try {
    const rawUrl = decodeURIComponent(urlStr)
    console.log('[协议处理] URL:', rawUrl)
    const url = new URL(rawUrl)
    const params = Object.fromEntries(url.searchParams.entries())

    if (!params.email) {
      console.error('❌ 缺少 email 参数')
      return
    }

    const fn = isMac ? createOutlookMailMac : isWin ? createOutlookMailWindows : null
    if (!fn) {
      console.error('❌ 当前系统不支持发送 Outlook 邮件')
      return
    }

    fn({
      to: params.email,
      subject: params.subject || '无主题',
      body: params.body || '',
      attachments: params.attachments?.trim()
        ? params.attachments.split(',').map(s => s.trim())
        : []
    })
  } catch (err) {
    console.error('❌ 协议处理失败:', err)
  }
}

/**
 * Windows 使用 PowerShell + COM 创建 Outlook 邮件
 */
function createOutlookMailWindows({ to, subject, body, attachments }) {
  try {
    const psScript = `
      $outlook = New-Object -ComObject Outlook.Application
      $mail = $outlook.CreateItem(0)
      $mail.To = ${JSON.stringify(to)}
      $mail.Subject = ${JSON.stringify(subject)}
      $mail.HTMLBody = ${JSON.stringify(body)}
      ${attachments.map(p => `$mail.Attachments.Add(${JSON.stringify(p)})`).join('\n')}
      $mail.Display()
      Start-Sleep -Milliseconds 300
      $shell = New-Object -ComObject WScript.Shell
      $shell.AppActivate("Outlook")
    `.trim()

    const encoded = Buffer.from(psScript, 'utf16le').toString('base64')

    execSync(`powershell -WindowStyle Hidden -NoProfile -EncodedCommand ${encoded}`, {
      stdio: 'ignore',
      windowsHide: true,
    })

    console.log('✅ 成功调用 Windows Outlook 并激活窗口')
  } catch (err) {
    console.error('❌ 调用 Outlook 出错:', err)
  }
}

/**
 * macOS 使用 AppleScript 调用 Outlook 创建邮件
 */
function createOutlookMailMac({ to, subject, body, attachments }) {
  try {
    const escapeAppleScriptString = str =>
      str.replace(/\\/g, '\\\\').replace(/"/g, '\\"')

    const asScript = `
      tell application "Microsoft Outlook"
        set newMessage to make new outgoing message with properties {subject:"${escapeAppleScriptString(subject)}", content:"${escapeAppleScriptString(body)}"}
        make new recipient at newMessage with properties {email address:{name:"", address:"${escapeAppleScriptString(to)}"}}
        ${attachments.map(filePath =>
          `make new attachment at newMessage with properties {file:(POSIX file "${escapeAppleScriptString(filePath)}")}`
        ).join('\n')}
        open newMessage
        activate
      end tell
    `
    const tmpFile = path.join(os.tmpdir(), 'outlook_temp.scpt')
    fs.writeFileSync(tmpFile, asScript)
    execSync(`osascript "${tmpFile}"`, { stdio: 'ignore' })
    console.log('✅ 成功调用 macOS Outlook')
  } catch (err) {
    console.error('❌ macOS 调用 Outlook 失败:', err)
  }
}

// 捕获未处理异常
process.on('uncaughtException', (err) => {
  console.error('💥 未捕获异常:', err)
})