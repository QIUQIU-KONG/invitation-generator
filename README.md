# Invitation Letter Generator

从出差人员数据自动生成英文邀请函和答复函（`.docx`），面向马来西亚签证场景。

> 🎯 将 HR 手工 30 分钟/批次的签证函工作压缩到 3 秒，零出错。

---

## 效果演示

```
python demo.py
```

```
Excel 出差表（工号+日期+护照号）
        ↓  知识库自动匹配
SQLite（姓名+部门+职位+性别）
        ↓  业务规则引擎
输出 output/*.docx（邀请函 + 答复函，按部门分组）
```

---

## 两种使用模式

| | Excel 模式 | 飞书模式 |
|---|---|---|
| 适用场景 | 月初批量处理 | 日常流水增量 |
| 数据入口 | 本地 `.xlsx` | 飞书多维表格 |
| 人员信息 | 知识库自动查询 | 知识库自动查询 |
| 运行命令 | `python generate_docs.py xxx.xlsx` | `python generate_docs.py` |

> 两种模式的**核心生成管线完全一致**，只是数据入口不同。人员基础信息（姓名/部门/职位/性别）无论哪种模式都从知识库查询，Excel 只需填工号。

---

## 快速开始

```bash
# 1. 安装依赖
pip install -r requirements.txt

# 2. 一键演示
python demo.py
```

### Excel 模式（三步）

```bash
# 第 1 步：初始化知识库（只需做一次）
python init_database.py 员工花名册.xlsx

# 第 2 步：准备出差 Excel（4 列）
# 工号 | 出发时间 | 返程时间 | 护照号

# 第 3 步：生成
python generate_docs.py 出差表.xlsx
```

### 飞书模式

```bash
setx FEISHU_BASE_TOKEN "YOUR_BASE_TOKEN"
setx FEISHU_TABLE_ID "YOUR_TABLE_ID"
setx LARK_CLI_PATH "C:\path\to\lark-cli.cmd"
python generate_docs.py
```

---

## 目录结构

```
invitation-generator/
├── README.md
├── requirements.txt
├── .gitignore
├── demo.py                   # 一键演示
├── demo_data.xlsx            # 演示用出差数据
├── generate_docs.py          # 主生成脚本
├── init_database.py          # 知识库初始化
├── pinyin_data.py            # 汉字拼音映射（3000+ 字）
├── knowledge_base.db         # SQLite 知识库（含 6 名 Demo 员工）
├── generated.json            # 增量追踪（自动生成）
├── 模板文件/                  # Word 模板（行政可直接修改）
│   ├── 项目组/（邀请函.docx + 答复函.docx）
│   ├── 运营组/（邀请函.docx + 答复函.docx）
│   └── 人事组/（邀请函.docx + 答复函.docx）
├── output/                   # 生成结果输出
└── tools/                    # 辅助工具
    ├── simplify_excel.py     # Excel 列名精简
    ├── update_excel_ids.py   # 工号批量更新
    ├── read_templates.py     # 模板结构查看
    └── check_sync.py         # 规则一致性检查
```

---

## 核心业务规则

### 部门 → 模板映射

| 数据中部门 | 模板目录 |
|-----------|----------|
| 项目组 / 项目 | `模板文件/项目组/` |
| 运营组 / 运营 | `模板文件/运营组/` |
| 人事组 / 人事 | `模板文件/人事组/` |

### 日期计算

| 步骤 | 规则 |
|------|------|
| 信函出发日 | 实际日期 − 1 天 |
| 信函返回日 | 实际日期 + 1 天 |
| ISSUE_DATE | 出发日 − 1 个月 |
| REPLY_DATE | ISSUE_DATE + 2 天 |

> 多人同部门取最早出发 + 最晚返回。使用 `dateutil.relativedelta` 正确处理跨月。

### 称谓与姓名

| 函件 | 规则 | 示例 |
|------|------|------|
| 邀请函 | 姓名拼音 | WANG XIAOMING |
| 答复函 | Mr./Ms. + 拼音 | Mr. WANG XIAOMING |

### 拼音转换

自建 3000+ 汉字映射表，不依赖第三方拼音库。未收录生僻字打印警告并保留原文。

### 职位翻译

50+ 中英映射（`POSITION_MAP`），未匹配项保留原文作为 fallback。

### 生成规范

- 字体：Century Gothic
- 字号：12pt（小四）
- 多人同部门合并为一封信，CC 行列全员

---

## 技术栈

| 组件 | 选型 | 理由 |
|------|------|------|
| Excel | openpyxl | 跨平台，无需 Office |
| Word | python-docx | 操作 OOXML，保留原模板样式 |
| 数据处理 | pandas | 分组聚合、日期解析 |
| 日期 | python-dateutil | `relativedelta` 跨月跨年准确 |
| 知识库 | SQLite | 零配置，单文件 |
| 飞书 | lark-cli | 飞书官方 CLI |

## 设计决策

**为什么用 Word 占位符而不是模板引擎？**
行政人员可直接用 Word 修改模板（调格式、改措辞），代码只做占位符替换。

**为什么知识库用 SQLite？**
人员信息相对稳定，与每次变化的出差数据分离，减少录入错误。

**为什么提供两种数据入口？**
Excel 适合月初大批量；飞书适合日常增量。核心管线完全复用。

## 许可

MIT
