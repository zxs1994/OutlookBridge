# OutlookBridge 协议调试工具

该程序的核心功能是通过自定义协议实现邮件的快速创建和发送。

由于 `mailto:` 协议不支持添加附件，本程序通过注册自定义协议 `outlookbridge://` 实现带附件的邮件调用。

这是一个用于测试和开发 `outlookbridge://` 自定义协议的 Electron 应用，可让浏览器调用 Outlook 新建带附件的邮件，支持在开发环境中模拟协议调用，并提供可视化页面辅助测试。

---

## 📦 安装依赖

```bash
yarn install
```

---

## 🧪 开发调试

```bash
yarn dev
```

> ✅ 会自动读取 `.env` 文件中的变量，例如 `OUTLOOKBRIDGE_URL`，用于调试自定义协议调用。

`.env` 示例：

```env
OUTLOOKBRIDGE_URL=outlookbridge://?email=test@example.com&subject=%E6%B5%8B%E8%AF%95%E4%B8%BB%E9%A2%98&body=%E4%BD%A0%E5%A5%BD&attachments=https%3A%2F%2Fexample.com%2Ffile.jpg
```

---

## 🛠️ 构建打包

> 🖥️ 本程序支持打包生成 macOS 和 Windows 平台的客户端，可独立运行并注册 `outlookbridge://` 协议。

根据当前操作系统选择打包命令：

- **macOS 构建**：
  ```bash
  yarn build:mac
  ```

- **Windows 构建**：
  ```bash
  yarn build:win
  ```

---

## ✉️ 协议调用测试页面（test.html）

1. 打开浏览器，打开 `test.html` 文件（可直接拖入浏览器）；
2. 填写收件人邮箱、主题、正文、附件路径；
3. 点击【点击唤起邮件】，测试是否能调用 `outlookbridge://` 协议；
4. 控制台会输出生成的协议 URL。

页面示意：

```html
outlookbridge://?email=xxx@example.com&subject=测试&body=你好&attachments=https://example.com/xxx.jpg
```

---

## 💡 注意事项

- windows 32位系统未测试
- windows环境下你的Outlook版本需要支持 COM 接口调用, 可以双击用test文件夹下的test.vbs测试是否支持
- 在开发模式下，会从 `.env` 中读取 `OUTLOOKBRIDGE_URL` 自动进行模拟调用；
- 渲染进程通过 `window.electronAPI` 接收参数，可在页面中 `logToWindow()` 输出内容；
- 所有 URL 参数必须经过 `encodeURIComponent` 编码。

---

如有bug，请联系维护者更新。
xusheng94@qq.com
