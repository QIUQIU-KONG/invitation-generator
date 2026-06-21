# 安全与隐私说明

本项目用于演示邀请函自动化流程。请不要把真实员工、护照、飞书或公司内部资料提交到公开仓库。

## 不应提交的文件

- `.env`
- `.mcp.json`
- `knowledge_base.db`
- `generated.json`
- `output/`
- 生产环境 Excel 或 CSV
- 飞书登录二维码、截图或导出的认证文件
- 真实公司 Word 模板或已生成函件

## API 与密钥

飞书配置应通过本机环境变量提供：

```text
FEISHU_BASE_TOKEN
FEISHU_TABLE_ID
LARK_CLI_PATH
FEISHU_ADMIN_ID
```

公开仓库只能保留空的 `.env.example`。如果真实 token 曾经提交到 GitHub，请立即在飞书后台轮换或吊销该 token，并清理 Git 历史。

当前脚本调用 `lark-cli` 时会把 `FEISHU_BASE_TOKEN` 作为命令参数传入。它不会写入仓库，但在脚本运行期间，管理员或本机进程查看工具可能看到该参数。生产环境如果对本机进程可见性要求很高，建议改为 lark-cli 支持的本地认证配置或更安全的密钥读取方式。

## 发布前检查

建议上传前运行：

```bash
python -m unittest discover -s tests
python -m py_compile generate_docs.py init_database.py run_feishu_gui.py tests/test_generate_docs.py
```

并人工检查：

- 数据库是否只有演示数据。
- Excel/JSON 是否只有演示数据。
- Word 模板正文和元数据是否已脱敏。
- `.gitignore` 是否仍然忽略本地运行状态和敏感文件。
