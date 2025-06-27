const { execSync, spawn } = require('child_process')
const fs = require('fs')
const os = require('os')
const path = require('path')

/**
 * 检测 Outlook 安装路径
 * @returns {string|null} Outlook.exe 的完整路径，如果未找到则返回 null
 */
function detectOutlookExePath() {
	// 所有可能的 Office 版本
	const officeVersions = [
		{ reg: '16.0', dir: 'Office16' }, // Office 2016/2019/365
		{ reg: '15.0', dir: 'Office15' }, // Office 2013
		{ reg: '14.0', dir: 'Office14' }, // Office 2010
		{ reg: '12.0', dir: 'Office12' }, // Office 2007
	]

	// 所有可能的安装基础路径
	const basePaths = [
		'C:\\Program Files\\Microsoft Office',
		'C:\\Program Files (x86)\\Microsoft Office',
		process.env.LOCALAPPDATA + '\\Microsoft\\Office', // 用户安装路径
		process.env.ProgramW6432 + '\\Microsoft Office', // 64位系统的程序路径
		process.env['ProgramFiles(x86)'] + '\\Microsoft Office', // 32位路径
	].filter(Boolean) // 过滤掉可能的undefined路径

	// 可能的子路径模式
	const subPaths = [
		'root\\{dir}', // Click-to-Run 安装
		'{dir}', // 传统安装
		'{reg}\\Outlook', // 用户安装路径
	]

	const exeName = 'OUTLOOK.EXE'

	// 检查所有可能的组合
	for (const base of basePaths) {
		for (const { dir, reg } of officeVersions) {
			for (const sub of subPaths) {
				const template = sub.replace('{dir}', dir).replace('{reg}', reg)

				const fullPath = path.join(base, template, exeName)

				try {
					if (fs.existsSync(fullPath)) {
						return fullPath
					}
				} catch (err) {
					console.debug(`检查路径 ${fullPath} 失败: ${err.message}`)
				}
			}
		}
	}

	return null
}

function createOutlookMailWindows({ to, subject, body, attachments = [] }, logToWindow) {
	logToWindow('🧪 开始调用 createOutlookMailWindows')

	const tempDir = path.join(os.tmpdir(), 'outlookbridge_attachments')
	if (!fs.existsSync(tempDir)) fs.mkdirSync(tempDir, { recursive: true })

	try {
		const outlookPath = detectOutlookExePath()
		if (!outlookPath) throw new Error('未找到 Outlook 安装路径')
		logToWindow(`🔍 找到 Outlook 路径: ${outlookPath}`)

		// ✅ 多附件走 COM 方式
		if (attachments.length > 1) {
			logToWindow(`📎 多附件模式，共 ${attachments.length} 个附件，准备下载并添加`)

			const downloadStatements = attachments
				.map((url, i) => {
					const filename = `file_${i}_${Date.now()}.${
						url.split('.').pop().split('?')[0] || 'tmp'
					}`
					const fullPath = path.join(tempDir, filename).replace(/\\/g, '\\\\')
					return `
          $wc = New-Object System.Net.WebClient
          $wc.DownloadFile("${url}", "${fullPath}")
          $mail.Attachments.Add("${fullPath}")
        `
				})
				.join('\n')

			const psScript = `
        $outlook = New-Object -ComObject Outlook.Application
        $mail = $outlook.CreateItem(0)
        $mail.To = ${JSON.stringify(to)}
        $mail.Subject = ${JSON.stringify(subject)}
        $mail.HTMLBody = ${JSON.stringify(body)}

        ${downloadStatements}

        $mail.Display()

        $shell = New-Object -ComObject WScript.Shell
        for ($i = 0; $i -lt 10; $i++) {
          Start-Sleep -Milliseconds 500
          if ($shell.AppActivate("Outlook")) { break }
        }
      `.trim()

			const encoded = Buffer.from(psScript, 'utf16le').toString('base64')
			execSync(
				`powershell -WindowStyle Hidden -NoProfile -EncodedCommand ${encoded}`,
				{
					stdio: 'ignore',
					windowsHide: true,
				}
			)

			logToWindow('✅ 多附件邮件已成功调用 Outlook')
			return
		}

		// ✅ 单附件或无附件，使用 outlook.exe 启动
		let downloadedFilePath = null
		if (attachments.length === 1) {
			logToWindow('📎 单附件模式，准备下载附件')

			const url = attachments[0]
			const ext = url.split('.').pop().split('?')[0] || 'tmp'
			const filename = `file_${Date.now()}.${ext}`
			downloadedFilePath = path.join(tempDir, filename)
			execSync(
				`powershell -Command "(New-Object Net.WebClient).DownloadFile('${url}', '${downloadedFilePath.replace(
					/\\/g,
					'\\\\'
				)}')"`
			)
			logToWindow(`✅ 附件下载完成，路径: ${downloadedFilePath}`)
		}

		// ✅ 构建 mailto 链接
		const mailtoParams = [
			`to=${encodeURIComponent(to)}`,
			`subject=${encodeURIComponent(subject)}`,
			`body=${encodeURIComponent(body)}`,
		].join('&')
		const cmd = downloadedFilePath
			? `"${outlookPath}" /a "${downloadedFilePath}" /m "${to}?${mailtoParams}"`
			: `"${outlookPath}" /c ipm.note /m "mailto:${to}?${mailtoParams}"`

		logToWindow(`📧 调用 Outlook 命令: ${cmd}`)
		execSync(cmd)
		logToWindow('✅ 邮件窗口已成功打开')
	} catch (err) {
		logToWindow(`❌ 调用 Outlook 出错：${err.message}`)
	}
}

module.exports = createOutlookMailWindows
