# 邀请函生成工具

从飞书多维表格或本地 Excel 文件生成英文邀请函和答复函（`.docx`）。

这是公开演示版本，仅包含演示数据和脱敏模板。请不要提交真实员工数据库、已生成的函件、飞书令牌、二维码或生产 Excel 文件。

## 当前状态

- 可从飞书多维表格读取出差记录，也可从本地 Excel 读取。
- 可通过 SQLite 知识库补全姓名、部门、职位、性别等员工信息。
- 可由操作人员在本次记录中手动选择 `模板类型`，不再强依赖知识库部门。
- 可按部门/模板分组生成邀请函和答复函。
- 可用 `generated.json` 记录生成历史，避免重复生成未变化的数据。
- 可通过 Windows 桌面启动器给非技术用户使用。

## 近期更新

- 新增 `模板类型` 字段，支持在每次出差记录中手动选择模板。
- 新增 MIT License。
- 新增 `unittest` 单元测试，覆盖飞书解析、模板选择、指纹变更和护照号规范化。
- 更新演示 Excel/JSON，让公开示例能展示 `模板类型` 的用法。
- 清理 Word 模板元数据，避免携带本机用户名等编辑痕迹。

完整更新说明见 [CHANGELOG.md](CHANGELOG.md)。

## 公开发布说明

当前目录适合作为公开演示项目上传 GitHub。仓库内只保留：

- 演示知识库：`knowledge_base.example.db`
- 演示 Excel：`sample_data.xlsx`
- 演示飞书数据：`sample_feishu_data.json`
- 脱敏 Word 模板：`模板文件/`

以下文件应只保存在本地，已经写入 `.gitignore`：

- `knowledge_base.db`
- `generated.json`
- `output/`
- 生产环境使用的 `*.xlsx`
- `.env`
- `.mcp.json`
- 飞书二维码图片

上传前请再次确认：

- `.env` 不存在或未被提交。
- `knowledge_base.db` 未被提交，只提交 `knowledge_base.example.db`。
- `generated.json` 未被提交。
- `output/` 中没有生成的真实函件。
- 飞书二维码、截图、真实 Excel、真实 Word 模板没有进入仓库。
- Word/Excel 元数据不含真实作者、公司、账号或内部路径。

## 安装

安装依赖：

```bash
pip install -r requirements.txt
```

从演示数据库复制一份工作知识库：

```bash
copy knowledge_base.example.db knowledge_base.db
```

如果使用飞书模式，请将 `.env.example` 复制为自己的环境配置文件，或手动设置以下环境变量：

```text
FEISHU_BASE_TOKEN=
FEISHU_TABLE_ID=
LARK_CLI_PATH=
FEISHU_ADMIN_ID=
```

`FEISHU_ADMIN_ID` 是可选项。设置后，脚本可以向该飞书用户发送提醒消息。

## 数据字段

飞书或 Excel 出差记录建议包含以下字段：

```text
工号
护照号
出发时间
返程时间
模板类型
```

`模板类型` 用于手动指定本次记录使用哪个模板，可填写：

```text
项目组
运营组
人事组
供应商
```

模板选择优先级：

```text
本次记录的 模板类型 > 知识库里的部门
```

如果 `模板类型` 留空，程序会回退到知识库里的部门自动匹配模板。如果两者都无法匹配，预览时会提示警告并跳过该记录。

## 使用方法

仅预览，不生成文件：

```bash
python generate_docs.py --dry-run
```

确认后生成：

```bash
python generate_docs.py
```

无需确认，直接生成：

```bash
python generate_docs.py --yes
```

明确指定本地 Excel 文件：

```bash
python generate_docs.py --excel sample_data.xlsx
```

使用演示飞书 JSON：

```bash
python generate_docs.py --feishu sample_feishu_data.json --dry-run
```

强制重新生成：

```bash
python generate_docs.py --force --yes
```

## Windows 桌面启动器

在 Windows 上可双击：

```text
打开邀请函自动化.vbs
```

如果 VBS 被系统策略拦截，可改用：

```text
打开邀请函自动化.bat
```

窗口中包含：

- 预览检查
- 正式生成
- 强制重生成
- 打开输出文件夹
- 打开飞书

## 目录结构

```text
模板文件/
  项目组/
    邀请函.docx
    答复函.docx
  运营组/
    邀请函.docx
    答复函.docx
  人事组/
    邀请函.docx
    答复函.docx
output/
```

如果要启用 `供应商` 模板，请新增：

```text
模板文件/
  供应商/
    邀请函.docx
    答复函.docx
```

## 本地检查

上传或交付前建议运行：

```bash
python -m py_compile generate_docs.py init_database.py run_feishu_gui.py
python -m unittest discover -s tests
python generate_docs.py --feishu sample_feishu_data.json --dry-run
python generate_docs.py --excel sample_data.xlsx --dry-run
```

注意：运行前需要先准备 `knowledge_base.db`。

## 测试

项目使用 Python 标准库 `unittest`，无需额外测试依赖：

```bash
python -m unittest discover -s tests
```

当前测试覆盖飞书 JSON 解析、`模板类型` 优先级、模板变更指纹和护照号规范化。

## 许可证

本项目使用 MIT License，详见 [LICENSE](LICENSE)。
