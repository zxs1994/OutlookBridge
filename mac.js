const { execSync } = require('child_process')
const fs = require('fs')
const os = require('os')
const path = require('path')

/**
 * macOS 使用 AppleScript 调用 Outlook 创建邮件
 */
function createOutlookMailMac({ to, subject, body, attachments }) {
  try {
    const escapeAppleScriptString = str =>
      str.replace(/\\/g, '\\\\').replace(/"/g, '\\"')

    const tempDir = path.join(os.tmpdir(), 'outlookbridge_attachments')
    if (!fs.existsSync(tempDir)) fs.mkdirSync(tempDir, { recursive: true })

    const localPaths = attachments.map((url, i) => {
      const filename = `file_${i}_${Date.now()}.${url.split('.').pop().split('?')[0] || 'tmp'}`
      const filePath = path.join(tempDir, filename)
      execSync(`curl -sSL "${url}" -o "${filePath}"`)
      return filePath
    })

    const asScript = `
      tell application "Microsoft Outlook"
        set newMessage to make new outgoing message with properties {subject:"${escapeAppleScriptString(subject)}", content:"${escapeAppleScriptString(body)}"}
        make new recipient at newMessage with properties {email address:{name:"", address:"${escapeAppleScriptString(to)}"}}
        ${localPaths.map(filePath =>
          `make new attachment at newMessage with properties {file:(POSIX file "${escapeAppleScriptString(filePath)}")}`
        ).join('\n')}
        open newMessage
        activate
      end tell
    `
    const tmpFile = path.join(os.tmpdir(), 'outlook_temp.scpt')
    fs.writeFileSync(tmpFile, asScript)
    execSync(`osascript "${tmpFile}"`, { stdio: 'ignore' })
    console.log('✅ 成功调用 macOS Outlook，附件为 URL 下载')
  } catch (err) {
    console.error('❌ macOS 调用 Outlook 失败:', err)
  }
}

module.exports = createOutlookMailMac