# Outlook Mail Extractor

从 GuJumpgate 中拆出的独立 Outlook / Microsoft Graph 邮件提取工具。它可以通过 `client_id` + `refresh_token` 换取访问令牌，读取 Inbox / Junk 邮件，并按发件人、主题、关键字等条件提取 6 位验证码。

[![Deploy to Cloudflare](https://deploy.workers.cloudflare.com/button)](https://deploy.workers.cloudflare.com/?url=https://github.com/Z-WICK/outlook-mail-extractor/tree/main)

上面的按钮用于一键部署 Cloudflare Worker 版本；服务器版 Node CLI 和 UI 仍然保留，服务器部署不会被替换。按钮部署完成后，如需绑定自定义域名或配置 Cloudflare Access，请在 Cloudflare 控制台中继续配置。

## 环境要求

- Node.js 18+
- 一个可用的 Microsoft OAuth `client_id`
- 对应账号的 `refresh_token`

工具不依赖第三方 npm 包，直接使用 Node 原生 `fetch`。

## 安装与测试

```bash
cd /Users/wick/Documents/vibe_codeing/cli/tools/codex-re/outlook-mail-extractor
npm test
```

本地直接执行：

```bash
node bin/outlook-mail-extractor.js --help
```

启动本地 UI：

```bash
npm run ui
```

默认地址：

```text
http://127.0.0.1:8787
```

或临时链接为命令：

```bash
npm link
outlook-mail-extractor --help
outlook-mail-extractor-ui --help
```

## 配置方式

推荐用环境变量传敏感信息：

```bash
export OUTLOOK_CLIENT_ID="your-client-id"
export OUTLOOK_REFRESH_TOKEN="your-refresh-token"
```

也可以使用命令参数：

```bash
node bin/outlook-mail-extractor.js code \
  --client-id "your-client-id" \
  --refresh-token "your-refresh-token"
```

## 提取验证码

```bash
node bin/outlook-mail-extractor.js code \
  --mailbox INBOX,Junk \
  --sender openai.com \
  --subject login \
  --keyword code \
  --top 10 \
  --max-retries 3 \
  --retry-delay-ms 10000 \
  --pretty
```

输出示例：

```json
{
  "command": "code",
  "code": "123456",
  "emailTimestamp": 1779920400000,
  "messageId": "message-id",
  "sender": "noreply@example.com",
  "subject": "Login code",
  "mailbox": "INBOX",
  "nextRefreshToken": "new-refresh-token-if-returned",
  "message": {}
}
```

## 拉取邮件列表

```bash
node bin/outlook-mail-extractor.js messages \
  --mailbox INBOX,Junk \
  --top 10 \
  --pretty
```

CLI 输出会保留 `nextRefreshToken`，但不会直接输出 Microsoft token 响应中的 `access_token`。

## 本地 UI

```bash
npm run ui -- --host 127.0.0.1 --port 8787
```

页面提供：

- 快捷导入 `邮箱----密码----client_id----令牌` 格式，支持批量每行一个，解析后自动保存并填入第一条账号的 `Client ID` 和 `Refresh Token`
- 邮箱池持久化到当前浏览器 `localStorage`，支持搜索、切换和删除已导入邮箱
- `INBOX` / `Junk` 邮箱夹选择
- 发件人、主题、正文关键字、排除验证码过滤
- 验证码提取与邮件列表两个结果视图
- “本次会话记住配置”开关，仅使用浏览器 `sessionStorage`

## Cloudflare Workers 部署

仓库包含独立的 Worker 入口 `src/worker.mjs` 和 `wrangler.jsonc`。原来的 CLI 和 Node UI 仍然可以继续本地运行。

先安装 Wrangler 并检查配置：

```bash
npm install
npm run check
npm run dev
```

确认本地 Worker 正常后部署：

```bash
npm run deploy
```

Worker 使用 Microsoft Graph / Outlook 的 HTTPS API，不需要常驻进程或服务器文件。邮箱池和 refresh token 仍只保存在浏览器本地，并随当前请求发送；Worker 不会把账号写入 Cloudflare 存储。

生产环境建议把 Worker 绑定在受 Cloudflare Access 保护的自定义域名下，避免公开 API 被第三方滥用。不要把 Microsoft refresh token 写入 `wrangler.jsonc` 或提交到 GitHub。

快捷导入里的账号只会写入当前浏览器的 `localStorage`，不会写入服务器文件，也不会通过 `/api/accounts` 暴露共享邮箱池。换浏览器、换设备或清理站点数据后，需要重新导入。

UI 服务默认监听本机地址，部署时可以通过 Caddy 等网关反代。API 响应不会返回 Microsoft `access_token`。如果接口返回了新的 `refresh_token`，页面会在结果中显示其存在状态，JSON API 会保留 `nextRefreshToken` 供你更新配置。

## 作为库使用

```js
const {
  fetchMicrosoftMailboxMessages,
  fetchMicrosoftVerificationCode,
} = require('./src/microsoft-email.js');

const result = await fetchMicrosoftVerificationCode({
  clientId: process.env.OUTLOOK_CLIENT_ID,
  refreshToken: process.env.OUTLOOK_REFRESH_TOKEN,
  mailboxes: ['INBOX', 'Junk'],
  senderFilters: ['openai.com'],
  subjectFilters: ['login'],
  top: 10,
});

console.log(result.code);
```

## 常用参数

- `--mailbox`：支持 `INBOX`、`Junk`、`junk email`，可逗号分隔或重复传入。
- `--filter-after`：只匹配指定时间之后的邮件，支持 Unix 毫秒时间戳或可解析日期字符串。
- `--exclude-code`：忽略旧验证码，适合重发验证码场景。
- `--code-pattern`：自定义验证码正则，优先使用第一个捕获组。
- `--verbose`：把 Graph / Outlook token 策略回退过程输出到 stderr。
