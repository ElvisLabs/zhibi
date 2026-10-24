# 执笔

新版 Outlook for Mac 里的 AI 回信助手。Outlook 工具栏右侧多一个钢笔尖按钮，点开是一个聊天面板：选中邮件，说一句要点，它读完原邮件和附件，写出排好版的回复；满意后点一下，Outlook 的答复窗口打开，正文已经填好，你改完自己点发送。

- **回复有结构**：结论先行、逐条答复、状态表格，行动项写清楚谁、做什么、哪天之前。缺的事实标成【待确认】，不替你编。
- **格式跟随 Outlook**：字体、字号、颜色取自你最近发出的邮件，签名用 Outlook 的默认签名，也可以在设置里自定义。
- **读附件**：Word（连修订和批注）、Excel、PPT、PDF、文本、作为附件的邮件。图片和扫描件只给文件名。
- **分析不进邮件**：要它总结附件、提醒注意事项或者直接提问时，另给一段「分析」，只给你看。
- **接着改**：不满意就再说一句（「更简短」「改成英文」），在上一版基础上改。
- **只读**：只读邮箱，从不发送邮件，也不在服务器上存草稿。

## 安装

```bash
brew install --cask elvislabs/tap/zhibi
```

装好后打开「执笔」，按提示在「系统设置 → 隐私与安全性 → 辅助功能」里允许它。之后它会在登录时自动启动，右键按钮可以关掉。

### 依赖

| 依赖 | 用来做什么 | 怎么装 |
|---|---|---|
| 新版 Outlook for Mac（中文界面） | 按钮和面板挂在它的窗口上 | |
| [outlook-cli](https://github.com/fatecannotbealtered/outlook-cli)（Python 版） | 通过 EWS 只读邮件和附件 | `pipx install git+https://github.com/fatecannotbealtered/outlook-cli.git`，再按它的说明登录 Exchange 邮箱 |
| [Claude Code](https://claude.com/claude-code) | 写回复（`claude -p`，不带任何工具） | 装好后在终端运行一次 `claude` 登录；也可以在面板「设置」里改用 Anthropic API Key |

- npm 版 outlook-cli 是打包好的二进制，执笔调用不了，需要 pip、pipx 或 uv 装的 Python 版。
- 读 Excel、PDF 附件还要装几个库：`pipx inject outlook-cli openpyxl xlrd pypdf`。Word、PPT 和文本不需要。

## 使用

1. 在 Outlook 里选中一封邮件，点工具栏右侧的钢笔尖按钮，或按 ⌃⌥R。
2. 写要点，或者点「按原文拟一版」「同意，并给出时间」「逐条回复」「礼貌地婉拒」。
3. 草稿出来后，不满意直接再说一句；满意了点「打开全部答复窗口」。

- 面板跟着当前选中的邮件走：换一封邮件再发要点，就是给新邮件拟稿。
- 右上角齿轮是设置：AI 认证方式、API 地址、模型、回复格式、签名和结束语，带实时预览。
- 右键钢笔尖按钮：停止生成、打开输出目录（草稿）、查看日志、登录时自动启动、退出。
- 也可以从 Raycast、快捷指令调用：`open "outlookai://reply"` 打开面板，`open -g "outlookai://reply?points=同意&to=sender"` 直接发一条要点。

## 数据和隐私

- 选中邮件的正文、收件人和附件文字会发给 Claude，走你本机 Claude Code 的登录，或者你在设置里填的 API。
- 设置在 `~/Library/Application Support/Zhibi/config.json`，草稿在同目录的 `drafts/`，日志在 `~/Library/Logs/Zhibi/`。API Key 存在钥匙串里。
- 卸载：`brew uninstall --cask zhibi`。连设置和草稿一起删：`brew uninstall --zap --cask zhibi`。

## 已知限制

- 只支持新版 Outlook for Mac 的中文界面：打开回复窗口时按的是「答复 / 全部答复」按钮和「编辑 → 粘贴」菜单。
- 应用没有经过 Apple 公证，Homebrew 安装时会去掉隔离属性。ad-hoc 签名每个版本都不同，升级后如果按钮不出现，到「辅助功能」里把执笔删掉，再打开执笔重新允许。
- 按钮和面板的位置靠辅助功能接口读 Outlook 的窗口；判断 Outlook 在哪个桌面用的是系统私有的只读接口。

## 在 Claude Code 里用

仓库本身也是一个 Claude Code skill（`SKILL.md`）。clone 到 `~/.claude/skills/outlook-mail-writer`，在 Claude Code 里说「回复这封邮件」「归纳今天的未读」，它按同一套写作规范起草，再在 Outlook 里打开答复窗口。

## 开发

```bash
app/build.sh             # 编译，装到 ~/Applications/执笔.app 并重启
app/build.sh --release   # universal、ad-hoc 签名，打包到 dist/
app/publish.sh           # 按 app/Info.plist 的版本号打 tag、发 GitHub Release、更新 Homebrew cask
```

- `app/Zhibi.swift`：按钮、面板、跟随 Outlook 窗口和桌面切换。
- `app/pane.html`：面板页面（WKWebView）。
- `scripts/`：查邮件、读附件、调用 Claude、渲染 HTML、打开答复窗口。build 时拷进 app 包。
- `references/`：写作规范和回复模板，也会放进提示词。

## License

MIT
