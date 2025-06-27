const { execSync } = require('child_process')
const fs = require('fs')
const path = require('path')
const os = require('os')

function createOutlookMailMac({ to, subject, body, attachments }, logToWindow, onSuccess) {
  try {
    attachments = Array.isArray(attachments) ? attachments : []

    const escapeAppleScriptString = str =>
      str.replace(/\\/g, '\\\\').replace(/"/g, '\\"')

    const tempDir = path.join(os.tmpdir(), 'outlookbridge_attachments')
    if (!fs.existsSync(tempDir)) fs.mkdirSync(tempDir, { recursive: true })

    // 下载附件
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
        return "success"
      end tell
    `

    const tmpFile = path.join(os.tmpdir(), 'outlook_temp.scpt')
    fs.writeFileSync(tmpFile, asScript)

    const result = execSync(`osascript "${tmpFile}"`, { encoding: 'utf8' }).trim()

    logToWindow(`✅ macOS Outlook 调用成功，返回: ${result}`)
    if (onSuccess) onSuccess()
  } catch (err) {
    logToWindow(`❌ macOS 调用 Outlook 失败: ${err.message}`)
  }
}

module.exports = createOutlookMailMac