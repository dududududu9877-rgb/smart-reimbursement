# Smart Reimbursement · 智能报销系统

A local-first desktop/web expense organizer for Chinese-language reimbursement workflows. Review receipts, manage expense sheets and export Excel reports with attached images. AI extraction uses a provider configured by the user.

这是从个人实际使用的报销工具整理出的首个公开版本，适合个人在自己的电脑上运行。项目刚开始公开维护，尚无公开用户量、下载量或效果评测数据。

## 功能

- 图片/PDF 票据管理，付款凭证、水单与审批材料整理。
- 用户自行配置兼容 Chat Completions 的模型接口，辅助识别并人工核对字段。
- SQLite 本地保存报销单、明细与配置；支持多模型配置。
- Excel 导出，附付款/水单截图，提供重新制作的通用空白模板。
- 本地统计和自然语言数据查询。
- 可选飞书事件网关（需要单独配置、部署及访问控制）。

## 快速开始

需要 Python 3.10 或更高版本。

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python serve.py
```

Windows 激活环境使用 `.venv\Scripts\activate`。启动后在浏览器打开 `http://127.0.0.1:8765`。按 Ctrl+C 结束浏览器模式服务。macOS 也可以在安装依赖后双击 `start.command`。

原生桌面窗口是可选功能，另安装 `requirements-desktop.txt`；未安装 pywebview 时自动使用浏览器。仓库仅发布源码，未提供签名或公证的安装包。打包脚本和 spec 是原项目的 macOS 打包起点，需要在目标机器验证。

首次启动会创建空数据库，进入设置修改显示姓名。如需 AI 识别，在模型设置中填入自己的 endpoint、model 和 API key；纯手动录入、管理和导出不需要 API key。不同服务商对图片、PDF、语音及模型名称的支持有差异，界面的预设只是可编辑示例。

## 示例与导出

`records.example.json` 全部是虚构信息。可以在命令行导出该示例：

```bash
python -c "import json; from export_expense import export; export('10', records_override=json.load(open('records.example.json', encoding='utf-8')))"
```

`template.xlsx` 是为公开版本重新生成的空白模板，不含原单位的表格或真实票据。当前导出模板每份最多 12 条；超过时返回 `truncated` 提示，请分单导出。AI 结果必须人工核对金额、日期、编号及凭证。

## 数据与隐私

- 源码运行时，`data.db`、`invoices/`、`exports/` 等保存在项目目录；打包桌面版本保存在 `~/.reimbursement/`。
- 这些本地文件、日志、密钥配置和备份均被 `.gitignore` 排除。不要把真实票据或数据库提交到仓库、Issue 或 PR。
- API key 存在本地 SQLite 中，当前未做静态加密。请使用个人可信电脑并保护好本地目录。
- 主动使用 AI 功能时，相关图片、文字或查询上下文会发送到你配置的服务商。应用前端还会从 CDN 加载 ECharts 和 PDF.js；本地保存不等于完全离线。
- 本应用是单用户本地工具，没有多用户登录和完整的生产安全防护，请保持监听 `127.0.0.1`，不要直接开放到互联网或不可信局域网。
- 未匹配的 HTTP 路径返回 404，避免静态目录暴露数据库；启动不会终止占用端口的其他进程。

## 可选飞书网关

复制 `feishu_config.example.json` 为 `feishu_config.json` 并填写自己的应用信息。网关运行 `python feishu_gateway.py`，默认回调端口 9900、报销服务地址 `http://127.0.0.1:8765`。需要配置飞书事件订阅、消息权限和验证字段；加密事件还需要 `cryptography` 依赖。配置中的空白名单代表所有调用者，应按实际需要设置允许的用户。网关未包含完整生产部署方案，请先在受控测试环境验证。

## 开发与贡献

核心文件：`serve.py`（本地 HTTP API/数据库/AI）、`index.html`（界面）、`export_expense.py`（Excel）、`feishu_gateway.py`（可选飞书）。

欢迎提交可复现的问题和改进。Issue 请说明系统、Python 版本、复现步骤和已脱敏错误，不要附真实客户、发票或 API key。

基础验证：

```bash
python -m unittest -v test_smoke
```

后续方向：改善导出容量提示、提升字段验证和自动化测试覆盖、完善跨平台打包及本地凭证保护。这些是计划，尚未宣称完成。

## License

MIT。第三方库和 CDN 资源适用其各自许可证。
