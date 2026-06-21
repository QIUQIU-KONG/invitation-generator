---
name: invitation
description: >
  从飞书多维表格、本地 Excel 或演示 JSON 出差数据生成英文邀请函和答复函（.docx），
  通过 SQLite 知识库补全人员基础信息，并支持本次记录手动选择模板类型。触发词：
  邀请函、invitation、生成邀请函、答复函、签证函、visa letter、飞书出差函件自动化。
---

# Invitation Letter Generator

使用此 skill 处理邀请函生成项目，核心脚本通常是 `generate_docs.py`，知识库是 `knowledge_base.db`，演示库是 `knowledge_base.example.db`。

## 工作流

1. 确认项目目录包含 `generate_docs.py`、`模板文件/`、`requirements.txt` 和演示数据。
2. 如果缺少 `knowledge_base.db`，先从 `knowledge_base.example.db` 复制一份工作库。
3. 优先运行预览检查，确认人员、模板、日期、护照号和缺失项。
4. 预览无误后再正式生成 Word 文件。
5. 生成结果写入 `output/`，生成历史写入 `generated.json`。

## 数据来源

支持三种输入：

```powershell
python generate_docs.py --dry-run
python generate_docs.py --excel sample_data.xlsx --dry-run
python generate_docs.py --feishu sample_feishu_data.json --dry-run
```

飞书模式需要环境变量：

```text
FEISHU_BASE_TOKEN
FEISHU_TABLE_ID
LARK_CLI_PATH
FEISHU_ADMIN_ID
```

`FEISHU_ADMIN_ID` 可选，用于发送提醒消息。

## 字段规则

飞书或 Excel 出差记录建议包含：

| 字段 | 说明 |
| --- | --- |
| 工号 | 关联 SQLite 知识库的员工 ID |
| 护照号 | 本次记录护照号，优先级高于知识库 |
| 出发时间 | 实际出发日期 |
| 返程时间 | 实际返程日期 |
| 模板类型 | 本次手动选择的模板，可为空 |

模板选择优先级：

```text
本次记录的 模板类型 > 知识库里的 department
```

`模板类型` 支持：

```text
项目组
运营组
人事组
供应商
```

如果 `模板类型` 为空，使用知识库里的 `department` 自动匹配。匹配失败时，预览应提示并跳过该记录。

## 知识库

SQLite 表 `employees` 至少包含：

| 字段 | 说明 |
| --- | --- |
| emp_id | 工号，主键 |
| name | 中文姓名 |
| department | 默认部门/模板兜底来源 |
| position | 中文职位 |
| gender | 性别，男/女 |
| passport | 护照号，可为空 |

护照号优先级：

```text
飞书/Excel 本次记录 > 知识库
```

如果飞书有护照号而知识库为空或不同，脚本可同步更新知识库。如果同一护照号属于其他员工，应警告并跳过。

## 生成规则

日期计算：

```text
信函出差开始日期 = 出发时间 - 1 天
信函出差结束日期 = 返程时间 + 1 天
ISSUE_DATE = 信函出差开始日期 - 1 个月
REPLY_DATE = ISSUE_DATE + 2 天
```

同一模板/部门下多人合并到同一组文件，日期范围取最早开始日期和最晚结束日期。

姓名和称谓：

- 邀请函使用大写拼音姓名。
- 答复函按性别生成 `Mr.` 或 `Ms.`。
- 职位通过脚本内 `POSITION_MAP` 翻译；未命中时保留原文并警告。

## 操作准则

- 先用 `--dry-run` 预览，不要直接生成。
- 不要把真实 `knowledge_base.db`、`.env`、`generated.json`、`output/` 或生产 Excel 提交到 GitHub。
- 检查公开版本时，要确认只包含演示姓名、示例护照号、脱敏模板和空环境变量。
- 修改模板选择逻辑时，要同时更新 README、演示 JSON 和预览输出，避免用户不知道该填什么字段。
- 修改脚本后至少运行：

```powershell
python -m py_compile generate_docs.py init_database.py run_feishu_gui.py
python generate_docs.py --feishu sample_feishu_data.json --dry-run
python generate_docs.py --excel sample_data.xlsx --dry-run
```
