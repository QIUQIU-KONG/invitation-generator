---
name: invitation
description: >
  从 Excel 出差数据生成英文邀请函和答复函（.docx），通过 SQLite 知识库自动检索人员基础信息。
  触发词：邀请函、invitation、生成邀请函、答复函、签证函、visa letter。
---

# Invitation Letter Generator

从出差数据自动生成英文邀请函 + 答复函。支持两种数据源：本地 Excel 和飞书多维表格。

## 架构

```
Excel / 飞书多维表格（工号、出发时间、返程时间、护照号）
       ↓
SQLite 知识库（工号 → 姓名、部门、职位、性别）
       ↓
generate_docs.py（业务规则引擎）
       ↓
output/*.docx（按部门分组的邀请函 + 答复函）
```

## 两种生成模式

### 模式 A：Excel 本地文件

```powershell
# 在项目根目录下执行
python generate_docs.py path/to/travel_data.xlsx
```

### 模式 B：飞书多维表格（无需 Excel）

```powershell
cd /path/to/invitation-generator
python generate_docs.py
```

> 飞书模式需预先设置环境变量：`FEISHU_BASE_TOKEN`、`FEISHU_TABLE_ID`、`LARK_CLI_PATH`

## 业务规则

### 输入格式

**Excel 模式（4 列）：**

| 列名 | 说明 | 示例 |
|------|------|------|
| 工号 | 纯数字，关联知识库 | 1 → 001 |
| 出发时间 | 出差开始日期 | 2026-06-15 |
| 返程时间 | 出差结束日期 | 2026-06-20 |
| 护照号 | 护照号码 | E12345678 |

**飞书模式：** 多维表格包含相同 4 个字段，无需手动维护 Excel。

### 知识库 (SQLite)

| 字段 | 类型 | 说明 |
|------|------|------|
| emp_id | TEXT (主键) | 工号，3 位数字 |
| name | TEXT | 中文姓名 |
| department | TEXT | 项目组 / 运营组 / 人事组 |
| position | TEXT | 中文职位 |
| gender | TEXT | 男 / 女 |
| passport | TEXT | 护照号（可选） |

### 生成规则

| 类别 | 规则 |
|------|------|
| 模板选择 | 项目组→项目组模板, 人事组→人事组模板, 运营组→运营组模板；未知部门跳过并警告 |
| 增量生成 | 按"工号_出发日期"去重，只生成新增人员 |
| 日期计算 | ①信函出发日=出发−1天, ②信函返回日=返回+1天, ③ISSUE_DATE=出发日−1个月, ④REPLY_DATE=ISSUE_DATE+2天 |
| 日期格式 | `Month D, YYYY`（不带前导零，如 `June 9, 2026`） |
| 姓名转换 | 中文→大写拼音，姓和名空格分隔（自建 3000+ 汉字映射表） |
| 职位翻译 | 中文职位翻译为英文（50+ 映射表），无法翻译则保留原文 |
| 称谓 | 邀请函：直接填姓名拼音；答复函：男→Mr.，女→Ms. |
| 输出命名 | `<部门>_<函类型>_<姓名1>_<姓名2>.docx` |
| 字体 | Century Gothic，12pt（小四） |
| 多人同部门 | 合并为一封信，取最早出发+最晚返回，CC 行列全员 |

## 使用方式

### 飞书模式（日常使用）

1. HR 在飞书多维表格填写出差记录
2. 在 Claude Code 中输入 `/invitation` 或"生成邀请函"
3. 脚本自动读取飞书表格 → 查知识库 → 生成 .docx
4. 生成完成后自动回写飞书状态"已生成"

### Excel 模式（批量处理）

```powershell
cd /path/to/invitation-generator
python generate_docs.py demo_data.xlsx
```

### 一键演示

```powershell
cd /path/to/invitation-generator
python demo.py
```

## 知识库管理

```powershell
# 初始化知识库
python init_database.py 员工花名册.xlsx

# 查看模板结构
cd tools; python read_templates.py
```

## 依赖

```
pip install openpyxl python-docx pandas python-dateutil lxml
```

飞书模式额外需要：
```
npm install -g @larksuite/lark-cli
```
