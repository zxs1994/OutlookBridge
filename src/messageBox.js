const { execSync } = require('child_process')
// ✅ Windows 弹框函数
// 弹出消息框（阻塞）
// ✅ 显示弹窗
function showMessageBox(message, title = 'Outlook Bridge') {
	const script = `Add-Type -AssemblyName PresentationFramework;[System.Windows.MessageBox]::Show('${message.replace(
		/'/g,
		"''"
	)}', '${title.replace(/'/g, "''")}')`
	execSync(`powershell -Command "${script}"`)
}

module.exports = showMessageBox