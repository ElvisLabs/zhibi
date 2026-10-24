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

第一次使用：

1. 打开「执笔」，按提示在「系统设置 → 隐私与安全性 → 辅助功能」里允许它。之后它会在登录时自动启动，右键按钮可以关掉。
2. 在 Outlook 里点工具栏右侧的执笔按钮，再点面板右上角的齿轮，在「邮箱账号」里填邮箱和密码，点「登录并测试」。服务器一般留空，自动发现不了时再填，例如 `mail.company.com`。
3. 装好 [Claude Code](https://claude.com/claude-code) 并在终端运行一次 `claude` 登录。也可以在设置里改用 Anthropic API Key，但仍然要装 Claude Code。

读邮件用的 [outlook-cli](https://github.com/fatecannotbealtered/outlook-cli)（含 exchangelib）和读 Excel、PDF 附件的库都打包在 app 里，自带一份 Python，不用另外安装，所以 app 有 100 MB 左右。只需要新版 Outlook for Mac（中文界面）和 Claude Code。

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
- 设置在 `~/Library/Application Support/Zhibi/config.json`，草稿在同目录的 `drafts/`，日志在 `~/Library/Logs/Zhibi/`。
- 邮箱账号按 outlook-cli 的方式保存：账号和服务器在 `~/.outlook-cli/config.json`，密码在钥匙串里。用过 outlook-cli 的话会直接沿用它的登录，第一次读邮件时钥匙串会问是否允许执笔里的 python3 读取「outlook-cli」的密码，点「始终允许」。
- Anthropic API Key 也存在钥匙串里。
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
app/build.sh --release   # arm64 和 x86_64 各打一个 ad-hoc 签名的 zip 到 dist/
app/publish.sh           # 按 app/Info.plist 的版本号打 tag，发 GitHub Release 草稿，在 Apple 芯片上冒烟测试通过后发布，再更新 Homebrew cask
```

- `app/build_python.sh`：打包进 app 的 Python（[python-build-standalone](https://github.com/astral-sh/python-build-standalone) + outlook-cli + 附件读取库），按目标架构下载 wheel，所以 Intel Mac 也能打 arm64 的包。

- `app/Zhibi.swift`：按钮、面板、跟随 Outlook 窗口和桌面切换。
- `app/pane.html`：面板页面（WKWebView）。
- `scripts/`：查邮件、读附件、调用 Claude、渲染 HTML、打开答复窗口。build 时拷进 app 包。
- `references/`：写作规范和回复模板，也会放进提示词。

## License

MIT。app 里打包的 Python 和各个库保留各自的许可证，见 `Contents/Resources/python` 里的 `LICENSE.txt` 和各包的 `*.dist-info`。
