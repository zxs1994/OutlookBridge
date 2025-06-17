const { spawn } = require('child_process')
const path = require('path')

// 提取 outlookbridge:// 协议参数
const protocolArg = process.argv.find(arg => arg.startsWith('outlookbridge://'))

// 过滤非法参数，仅保留 main.js 作为入口
const args = [path.join(__dirname, 'main.js')]

if (protocolArg) {
  process.env.OUTLOOKBRIDGE_URL = protocolArg
}

const electronPath = path.join(__dirname, 'node_modules', '.bin', process.platform === 'win32' ? 'electron.cmd' : 'electron')

const child = spawn(electronPath, args, {
  stdio: 'inherit',
  env: process.env
})