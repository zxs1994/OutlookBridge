const { app } = require('electron');
const { execSync } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');

const isMac = process.platform === 'darwin';
const isWin = process.platform === 'win32';

// 注册自定义协议（开发环境可能无效，需打包后使用）
if (!app.isDefaultProtocolClient('outlookbridge')) {
  app.setAsDefaultProtocolClient('outlookbridge');
}

// Mac 上必须在 ready 前监听 open-url
app.on('open-url', (event, urlStr) => {
  event.preventDefault();
  handleProtocol(urlStr);
});

app.whenReady().then(() => {
  // macOS 正常流程：通过 open-url 接收
  // Windows 在首次启动时可能通过 process.argv 收到 URL
  if (isWin && process.argv.length >= 2) {
    const protocolArg = process.argv.find(arg => arg.startsWith('outlookbridge://'));
    if (protocolArg) handleProtocol(protocolArg);
  }
});

/**
 * 处理 outlookbridge:// 协议
 * @param {string} urlStr 
 */
function handleProtocol(urlStr) {
  try {
    const url = new URL(urlStr);
    const params = Object.fromEntries(url.searchParams.entries());

    if (!params.email) {
      console.error('缺少 email 参数');
      return;
    }

    const fn = isMac ? createOutlookMailMac : isWin ? createOutlookMailWindows : null;
    if (!fn) {
      console.error('当前系统不支持发送 Outlook 邮件');
      return;
    }

    fn({
      to: params.email,
      subject: params.subject || '无主题',
      body: params.body || '',
      attachments: params.attachments ? params.attachments.split(',') : []
    });

  } catch (err) {
    console.error('处理协议失败:', err);
  }
}

/**
 * Windows 下使用 PowerShell + COM 创建 Outlook 邮件
 */
function createOutlookMailWindows({ to, subject, body, attachments }) {
  const psScript = `
    $outlook = New-Object -ComObject Outlook.Application
    $mail = $outlook.CreateItem(0)
    $mail.To = ${JSON.stringify(to)}
    $mail.Subject = ${JSON.stringify(subject)}
    $mail.Body = ${JSON.stringify(body)}
    $mail.HTMLBody = ${JSON.stringify(body)}
    ${attachments.map(p => `$mail.Attachments.Add(${JSON.stringify(p)})`).join('\n')}
    $mail.Display()
  `;
  execSync(`powershell -Command "${psScript}"`, { stdio: 'ignore' });
}

/**
 * macOS 下使用 AppleScript 调用 Outlook 创建邮件
 */
function createOutlookMailMac({ to, subject, body, attachments }) {
  const asScript = `
  tell application "Microsoft Outlook"
    set newMessage to make new outgoing message with properties {subject:"${subject}", content:"${body}"}
    make new recipient at newMessage with properties {email address:{name:"", address:"${to}"}}
    ${attachments.map(filePath => `make new attachment at newMessage with properties {file:(POSIX file "${filePath}")}` ).join('\n')}
    open newMessage
    activate
  end tell
`;
  console.log('生成的 AppleScript:', asScript);
  const tmpFile = path.join(os.tmpdir(), 'outlook_temp.scpt');
  fs.writeFileSync(tmpFile, asScript);
  execSync(`osascript "${tmpFile}"`, { stdio: 'ignore' });
}

// 捕获未处理异常
process.on('uncaughtException', (err) => {
  console.error('未捕获异常:', err);
});