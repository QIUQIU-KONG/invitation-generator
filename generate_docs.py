# -*- coding: utf-8 -*-
"""
从飞书多维表格读取出差数据，查询 SQLite 知识库，生成邀请函和答复函 (.docx)

用法:
    python generate_docs.py                          # 从飞书表格读取
    python generate_docs.py path/to/your_file.xlsx   # 从本地 Excel 读取

依赖:
    pip install openpyxl python-docx pandas python-dateutil
"""
import sys
import os
import json
import subprocess
import sqlite3
from dateutil.relativedelta import relativedelta
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from pinyin_data import PINYIN_DATA

# === 配置 ===
DEFAULT_EXCEL = r"path/to/your/travel_data.xlsx"
TEMPLATE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "模板文件")
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "knowledge_base.db")
TRACKING_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "generated.json")

# 飞书多维表格配置
FEISHU_BASE_TOKEN = os.environ.get("FEISHU_BASE_TOKEN", "")
FEISHU_TABLE_ID = os.environ.get("FEISHU_TABLE_ID", "")
LARK_CLI = os.environ.get("LARK_CLI_PATH", "lark-cli")

DEPT_TEMPLATE_MAP = {
    "项目组": "项目组", "项目": "项目组",
    "人事组": "人事组", "人事": "人事组",
    "运营组": "运营组", "运营": "运营组",
}

# === 字体配置 ===
FONT_NAME = 'Century Gothic'
FONT_SIZE_PT = 12  # 小四

# === 职位翻译表 ===
POSITION_MAP = {
    "经理": "Manager", "总经理": "General Manager", "副总经理": "Deputy General Manager",
    "总监": "Director", "主管": "Supervisor", "助理": "Assistant",
    "工程师": "Engineer", "高级工程师": "Senior Engineer", "初级工程师": "Junior Engineer",
    "技术员": "Technician", "操作员": "Operator", "分析师": "Analyst",
    "顾问": "Consultant", "专员": "Specialist", "协调员": "Coordinator",
    "主任": "Chief", "组长": "Team Leader", "副组长": "Deputy Team Leader",
    "秘书": "Secretary", "翻译": "Translator", "会计": "Accountant",
    "设计师": "Designer", "项目经理": "Project Manager",
    "质检员": "Quality Inspector", "采购员": "Purchaser", "销售": "Sales",
    "实习生": "Intern", "技师": "Technician", "切割技师": "Cutting Technician",
    "多线切割技师": "Multi-wire Cutting Technician", "中级操作员": "Intermediate Operator",
    "体系工程师": "System Engineer", "副经理": "Deputy Manager",
    "设备工程师": "Equipment Engineer",
    "资深基建工程师": "Senior Infrastructure Engineer",
    "项目副总监": "Deputy Project Director",
    "资深制程巡检员": "Senior Process Inspector",
    "资深测试技术工程师": "Senior Test Technology Engineer",
    "过程质量工程师": "Process Quality Engineer",
    "制造技术副经理": "Deputy Manufacturing Technology Manager",
    "电镀工程师": "Plating Engineer",
    "PME工程师": "PME Engineer",
    "线长": "Line Leader",
}


# === 飞书通知 ===
def notify_feishu(title, message):
    """通过飞书发送通知消息"""
    import subprocess
    try:
        content = f"[{title}]\n{message}"
        result = subprocess.run(
            [LARK_CLI, "im", "+message-send",
             "--receive-id-type", "open_id",
             "--receive-id", os.environ.get("FEISHU_OPEN_ID", "YOUR_OPEN_ID"),
             "--msg-type", "text",
             "--content", content,
             "--as", "bot"],
            capture_output=True, text=True, timeout=10
        )
        if result.returncode == 0:
            print("  [OK] 飞书通知已发送")
        else:
            print(f"  [警告] 飞书通知发送失败: {result.stderr}")
    except Exception as e:
        print(f"  [警告] 飞书通知异常: {e}")


def update_feishu_status(record_id, status):
    """更新飞书表格中的生成状态"""
    if not record_id:
        return
    try:
        result = subprocess.run(
            [LARK_CLI, "base", "+record-upsert",
             "--base-token", FEISHU_BASE_TOKEN,
             "--table-id", FEISHU_TABLE_ID,
             "--record-id", record_id,
             "--json", json.dumps({"生成状态": status}, ensure_ascii=False),
             "--as", "user"],
            capture_output=True, text=True, timeout=10,
            encoding='utf-8', errors='replace'
        )
        if result.returncode == 0:
            print(f"  [OK] {record_id}: 已标记为{status}")
        else:
            print(f"  [WARN] {record_id}: 状态更新失败 - {result.stderr[:200]}")
    except Exception as e:
        print(f"  [警告] 飞书状态更新失败: {e}")


# === 数据读取 ===
def read_feishu_table(base_token, table_id):
    """从飞书多维表格读取出差数据"""
    cmd = [
        LARK_CLI, "base", "+record-list",
        "--base-token", base_token,
        "--table-id", table_id,
        "--as", "user"
    ]
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=30, encoding='utf-8', errors='replace')
        if result.returncode != 0:
            return None
        records = []
        lines = result.stdout.strip().split('\n')
        data_start = False
        for line in lines:
            line = line.strip()
            if line.startswith('|') and not data_start:
                data_start = True
                continue
            if data_start and line.startswith('|') and not line.startswith('| ---'):
                if 'Meta:' in line:
                    continue
                if '记录 ID' in line or 'record_id' in line.lower():
                    continue
                cells = [c.strip() for c in line.split('|')[1:-1]]
                if len(cells) >= 5:
                    records.append({
                        'record_id': cells[0],
                        '工号': cells[1].zfill(3),
                        '护照号': cells[2] if cells[2] else 'N/A',
                        '出发时间': cells[3],
                        '返程时间': cells[4]
                    })
        return records
    except Exception:
        return None


def query_employee(emp_id):
    """从 SQLite 知识库查询人员信息"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT name, department, position, gender, passport FROM employees WHERE emp_id = ?", (emp_id,))
    result = cursor.fetchone()
    conn.close()
    if result:
        return {"name": result[0], "department": result[1], "position": result[2], "gender": result[3], "passport": result[4] if result[4] else ""}
    return None


def query_employees_by_ids(emp_ids):
    """批量查询人员信息"""
    if not emp_ids:
        return {}
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    placeholders = ','.join(['?' for _ in emp_ids])
    cursor.execute(f"SELECT emp_id, name, department, position, gender FROM employees WHERE emp_id IN ({placeholders})", emp_ids)
    results = cursor.fetchall()
    conn.close()
    return {row[0]: {"name": row[1], "department": row[2], "position": row[3], "gender": row[4]} for row in results}


# === 工具函数 ===
def set_run_font(run, font_name=FONT_NAME, size_pt=FONT_SIZE_PT):
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        from docx.oxml import OxmlElement
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), font_name)


def set_paragraph_font(para, font_name=FONT_NAME, size_pt=FONT_SIZE_PT):
    for run in para.runs:
        set_run_font(run, font_name, size_pt)


def to_pinyin_upper(name_zh):
    name = str(name_zh).strip()
    parts = []
    i = 0
    while i < len(name):
        if i + 1 < len(name) and name[i:i+2] in PINYIN_DATA:
            parts.append(PINYIN_DATA[name[i:i+2]])
            i += 2
        elif name[i] in PINYIN_DATA:
            parts.append(PINYIN_DATA[name[i]])
            i += 1
        elif '一' <= name[i] <= '鿿':
            parts.append(name[i])
            print(f"  [警告] 未收录拼音: {name[i]} (来自 '{name}')")
            i += 1
        else:
            parts.append(name[i])
            i += 1
    if len(parts) >= 2:
        return f"{parts[0].upper()} {''.join(parts[1:]).upper()}"
    return "".join(parts).upper()


def fmt_date(dt):
    return f"{dt.strftime('%B')} {dt.day}, {dt.year}"


def fmt_short_date(dt):
    return f"{dt.strftime('%B')} {dt.day}, {dt.year}"


def resolve_dept(dept_raw, warnings):
    dept = str(dept_raw).strip()
    if dept in DEPT_TEMPLATE_MAP:
        return DEPT_TEMPLATE_MAP[dept]
    for key, val in DEPT_TEMPLATE_MAP.items():
        if dept.startswith(key) or key.startswith(dept):
            return val
    warnings.append(dept)
    return None


def title_with_mr(name_zh, gender):
    py = to_pinyin_upper(name_zh)
    title = "Mr." if str(gender).strip() == "男" else "Ms."
    return f"{title} {py}"


def translate_pos(pos):
    """翻译职位，无法翻译则保留原文"""
    if not isinstance(pos, str) or not pos.strip():
        return "N/A"
    pos = pos.strip()
    if pos in POSITION_MAP:
        return POSITION_MAP[pos]
    for cn, en in POSITION_MAP.items():
        if cn in pos:
            return en
    return pos


# === 文档生成 ===
def fill_personnel(doc, people, is_reply=False):
    def _build_inv_blocks(ppl):
        blocks = []
        for p in ppl:
            name = to_pinyin_upper(p["name"])
            blocks.append([f"Name: {name}", f"Passport Number: {p['passport']}", f"Title: {p['position']}"])
        return blocks

    def _build_rep_lines(ppl):
        lines = []
        for p in ppl:
            name = title_with_mr(p["name"], p["gender"])
            lines.append(f"{name}, Position: {p['position']}, Passport No.: {p['passport']}")
        return lines

    if is_reply:
        return _fill_personnel_reply(doc, people, _build_rep_lines(people))
    else:
        return _fill_personnel_inv(doc, people, _build_inv_blocks(people))


def _fill_personnel_inv(doc, people, blocks):
    name_idxs, passport_idxs, pos_idxs = [], [], []
    for i, p in enumerate(doc.paragraphs):
        if "{{NAME}}" in p.text: name_idxs.append(i)
        if "{{PASSPORT}}" in p.text: passport_idxs.append(i)
        if "{{POSITION}}" in p.text: pos_idxs.append(i)

    slot_count = min(len(name_idxs), len(passport_idxs), len(pos_idxs))
    if slot_count == 0:
        return doc

    filled_count = min(slot_count, len(people))
    for s in range(filled_count):
        _replace_para_text(doc.paragraphs[name_idxs[s]], blocks[s][0])
        _replace_para_text(doc.paragraphs[passport_idxs[s]], blocks[s][1])
        _replace_para_text(doc.paragraphs[pos_idxs[s]], blocks[s][2])

    if filled_count < slot_count:
        to_delete = []
        for s in range(filled_count, slot_count):
            to_delete.extend([name_idxs[s], passport_idxs[s], pos_idxs[s]])
        for idx in sorted(to_delete, reverse=True):
            doc.paragraphs[idx]._element.getparent().remove(doc.paragraphs[idx]._element)

    if len(people) > slot_count:
        last_idx = max(name_idxs[filled_count - 1], passport_idxs[filled_count - 1], pos_idxs[filled_count - 1])
        last_para = doc.paragraphs[last_idx]
        for extra_p in people[slot_count:]:
            for line in [f"Name: {to_pinyin_upper(extra_p['name'])}", f"Passport Number: {extra_p['passport']}", f"Title: {extra_p['position']}"]:
                new_elem = last_para._element.makeelement(last_para._element.tag, last_para._element.attrib)
                last_para._element.addnext(new_elem)
                last_para = type(doc.paragraphs[0])(new_elem, doc.paragraphs[0]._parent)
                last_para.text = line
                if slot_count > 0:
                    last_para.style = doc.paragraphs[name_idxs[0]].style
                set_paragraph_font(last_para)
    return doc


def _fill_personnel_reply(doc, people, lines):
    name_idxs = [i for i, p in enumerate(doc.paragraphs) if "{{NAME}}" in p.text]
    if not name_idxs:
        return doc

    body_idxs = [i for i in name_idxs if "CC:" not in doc.paragraphs[i].text]
    cc_idxs = [i for i in name_idxs if "CC:" in doc.paragraphs[i].text]

    if cc_idxs:
        cc_names = ", ".join(title_with_mr(p["name"], p["gender"]) for p in people)
        for cc_i in cc_idxs:
            _replace_para_text(doc.paragraphs[cc_i], f"CC: {cc_names}")

    filled_count = min(len(body_idxs), len(lines))
    for s in range(filled_count):
        _replace_para_text(doc.paragraphs[body_idxs[s]], lines[s])

    # 删除多余的占位符槽位（模板槽位数 > 实际人数）
    if len(body_idxs) > filled_count:
        to_delete = []
        for s in range(filled_count, len(body_idxs)):
            to_delete.append(body_idxs[s])
        for idx in sorted(to_delete, reverse=True):
            doc.paragraphs[idx]._element.getparent().remove(doc.paragraphs[idx]._element)

    if len(people) > len(body_idxs):
        last_para = doc.paragraphs[body_idxs[-1]]
        for extra_line in lines[len(body_idxs):]:
            new_elem = last_para._element.makeelement(last_para._element.tag, last_para._element.attrib)
            last_para._element.addnext(new_elem)
            last_para = type(doc.paragraphs[0])(new_elem, doc.paragraphs[0]._parent)
            last_para.text = extra_line
            set_paragraph_font(last_para)
    return doc


def _replace_para_text(para, new_text):
    if para.runs:
        para.runs[0].text = new_text
        for r in para.runs[1:]:
            r.text = ""
    else:
        para.text = new_text
    set_paragraph_font(para)


def cleanup_date_artifacts(para):
    import re
    full_text = ''.join(r.text for r in para.runs)
    full_text = re.sub(r'2026[, ]+2026', '2026', full_text)
    if para.runs:
        para.runs[0].text = full_text
        for run in para.runs[1:]:
            run.text = ''


def replace_non_personnel(doc, replacements):
    for para in doc.paragraphs:
        for key, val in replacements.items():
            if key in para.text and "{{NAME}}" not in key and "{{PASSPORT}}" not in key and "{{POSITION}}" not in key:
                for run in para.runs:
                    if key in run.text:
                        run.text = run.text.replace(key, val)
                        set_run_font(run)
        cleanup_date_artifacts(para)
    return doc


def generate_for_dept(dept_folder, dept_label, records, names_str=""):
    inv_path = os.path.join(TEMPLATE_DIR, dept_folder, "邀请函.docx")
    rep_path = os.path.join(TEMPLATE_DIR, dept_folder, "答复函.docx")
    if not os.path.exists(inv_path) or not os.path.exists(rep_path):
        print(f"  [跳过] 模板不存在: {dept_folder}")
        return

    travel_start = records["adj_depart"].min()
    travel_end = records["adj_return"].max()
    issue_date = travel_start - relativedelta(months=1)
    reply_date = issue_date + pd.Timedelta(days=2)
    if travel_start == travel_end:
        travel_period = fmt_short_date(travel_start)
    else:
        travel_period = f"{fmt_short_date(travel_start)} to {fmt_short_date(travel_end)}"

    people = [{"name": row["name"], "passport": row["passport"], "position": row["position"], "gender": row["gender"]} for _, row in records.iterrows()]

    year_str = str(travel_start.year)
    common_replacements = {
        "{{ISSUE_DATE}}": fmt_date(issue_date),
        "{{TRAVEL_DATE}}": travel_period,
        "{{TravelPeriod}}": f"{travel_period}, {year_str}",
        "{{INVITE_PERSONNEL_LIST}}": "",
    }
    rep_extra = {"{{REPLY_DATE}}": fmt_date(reply_date)}

    inv_filename = f"{dept_label}_邀请函_{names_str}.docx" if names_str else f"{dept_label}_邀请函.docx"
    rep_filename = f"{dept_label}_答复函_{names_str}.docx" if names_str else f"{dept_label}_答复函.docx"

    inv_doc = Document(inv_path)
    replace_non_personnel(inv_doc, common_replacements)
    fill_personnel(inv_doc, people, is_reply=False)
    inv_out = os.path.join(OUTPUT_DIR, inv_filename)
    inv_doc.save(inv_out)
    print(f"  [OK] 邀请函: {inv_out}")

    rep_doc = Document(rep_path)
    replace_non_personnel(rep_doc, {**common_replacements, **rep_extra})
    fill_personnel(rep_doc, people, is_reply=True)
    rep_out = os.path.join(OUTPUT_DIR, rep_filename)
    rep_doc.save(rep_out)
    print(f"  [OK] 答复函: {rep_out}")


# === 主流程 ===
def main():
    excel_path = None
    feishu_data_file = None

    # Parse arguments
    i = 1
    while i < len(sys.argv):
        if sys.argv[i] == '--feishu' and i + 1 < len(sys.argv):
            feishu_data_file = sys.argv[i + 1]
            i += 2
        elif sys.argv[i] == '--excel' and i + 1 < len(sys.argv):
            excel_path = sys.argv[i + 1]
            i += 2
        elif not sys.argv[i].startswith('--'):
            excel_path = sys.argv[i]
            i += 1
        else:
            i += 1

    if not os.path.exists(DB_PATH):
        print(f"[错误] 知识库不存在: {DB_PATH}")
        print("请先运行 init_database.py 初始化知识库")
        sys.exit(1)

    # If Feishu data file is provided, use it directly
    if feishu_data_file:
        with open(feishu_data_file, 'r', encoding='utf-8') as f:
            raw_data = json.load(f)
        _process_feishu_data(raw_data)
        return

    # If --excel is specified, force Excel mode
    if excel_path:
        if not os.path.exists(excel_path):
            print(f"[错误] 文件不存在: {excel_path}")
            sys.exit(1)
        _process_excel(excel_path)
        return

    # Default: read from Feishu table directly
    if not FEISHU_BASE_TOKEN or not FEISHU_TABLE_ID:
        print("[提示] 请先设置环境变量:")
        print("  setx FEISHU_BASE_TOKEN \"YOUR_BASE_TOKEN\"")
        print("  setx FEISHU_TABLE_ID \"YOUR_TABLE_ID\"")
        print("  setx LARK_CLI_PATH \"C:\\path\\to\\lark-cli.cmd\"")
        print("  设置后重新打开终端即可生效")
        sys.exit(1)

    records = read_feishu_table(FEISHU_BASE_TOKEN, FEISHU_TABLE_ID)
    if records:
        raw_data = []
        for d in records:
            raw_data.append({
                'emp_id': d['工号'],
                'passport': d['护照号'] if d['护照号'] != 'N/A' else '',
                'depart_date': d['出发时间'],
                'return_date': d['返程时间'],
                'record_id': d.get('record_id', '')
            })
        _process_feishu_data(raw_data)
    else:
        print("飞书表格无数据或读取失败")
        # Fallback to Excel
        if os.path.exists(DEFAULT_EXCEL):
            print(f"尝试使用本地 Excel: {DEFAULT_EXCEL}")
            _process_excel(DEFAULT_EXCEL)
        else:
            print("[错误] 无可用数据源")
            sys.exit(1)


def _process_feishu_data(raw_data):
    """处理飞书数据"""
    print(f"飞书表格共 {len(raw_data)} 条记录")
    print("查询知识库...")

    records = []
    warnings = []

    for d in raw_data:
        emp_id = d.get('emp_id', d.get('工号', ''))
        if not emp_id.startswith('00'):
            emp_id = emp_id.zfill(3)

        emp_info = query_employee(emp_id)
        if not emp_info:
            print(f"  [警告] 工号 {emp_id} 不在知识库中，跳过")
            continue

        # 护照号优先级: 飞书表格 > 知识库
        feishu_passport = d.get('passport', d.get('护照号', ''))
        kb_passport = emp_info.get('passport', '')

        if feishu_passport:
            passport = feishu_passport
        elif kb_passport:
            passport = kb_passport
        else:
            passport = ''
            warnings.append(f"工号 {emp_id} ({emp_info['name']}) 护照号为空，请补充")

        records.append({
            '工号': emp_id,
            'name': emp_info['name'],
            'department': emp_info['department'],
            'position': translate_pos(emp_info['position']),
            'gender': emp_info['gender'],
            'passport': passport,
            '出发时间': d.get('depart_date', d.get('出发时间', '')),
            '返程时间': d.get('return_date', d.get('返程时间', '')),
            'record_id': d.get('record_id', '')
        })

    # 输出警告 + 飞书通知
    if warnings:
        print("\n=== 护照号缺失警告 ===")
        for w in warnings:
            print(f"  [警告] {w}")
        # 通过飞书通知用户
        notify_feishu("护照号缺失提醒", "\n".join(warnings))

    if not records:
        print("[错误] 无有效数据")
        sys.exit(1)

    df = pd.DataFrame(records)
    df['depart_date'] = pd.to_datetime(df['出发时间'])
    df['return_date'] = pd.to_datetime(df['返程时间'])
    _process_dataframe(df)

def _process_dataframe(df):
    """处理 DataFrame 数据，生成邀请函"""
    warnings = []
    df['__dept'] = df['department'].apply(lambda x: resolve_dept(x, warnings))
    df['adj_depart'] = pd.to_datetime(df['depart_date']) - pd.Timedelta(days=1)
    df['adj_return'] = pd.to_datetime(df['return_date']) + pd.Timedelta(days=1)

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # 增量追踪（按 工号_出发日期 去重）
    generated_ids = set()
    if os.path.exists(TRACKING_FILE):
        with open(TRACKING_FILE, 'r', encoding='utf-8') as f:
            generated_ids = set(json.load(f))

    new_depts = {}
    tracking_keys = []  # 记录本次生成的追踪 key 和 record_id
    for _, row in df.iterrows():
        tracking_key = f"{row['工号']}_{row['出发时间']}"
        if tracking_key not in generated_ids:
            dept = row['__dept']
            if dept not in new_depts:
                new_depts[dept] = []
            new_depts[dept].append(row)
            tracking_keys.append({
                'key': tracking_key,
                'record_id': row.get('record_id', ''),
                'emp_id': row['工号']
            })

    if not new_depts:
        print("\n没有新的人员需要生成，所有文件已是最新。")
        return

    total_new = sum(len(persons) for persons in new_depts.values())
    print(f"\n检测到 {total_new} 个新增人员:")
    for dept, persons in new_depts.items():
        names = [row['name'] for row in persons]
        print(f"  [{dept}] {len(persons)} 人: {', '.join(names)}")

    print()

    for dept_folder, persons in new_depts.items():
        dept_df = pd.DataFrame(persons)
        print(f"--- {dept_folder} ({len(persons)} 人) ---")
        person_names = [to_pinyin_upper(row['name']) for _, row in dept_df.iterrows()]
        names_str = '_'.join(person_names)
        generate_for_dept(dept_folder, dept_folder, dept_df, names_str)

    # 更新本地追踪文件
    new_generated = generated_ids | {k['key'] for k in tracking_keys}
    new_generated_str = sorted(list(new_generated))
    with open(TRACKING_FILE, 'w', encoding='utf-8') as f:
        json.dump(new_generated_str, f, ensure_ascii=False)

    # 写回飞书表格标记"已生成"
    print("\n更新飞书表格生成状态...")
    for k in tracking_keys:
        update_feishu_status(k['record_id'], '已生成')

    if warnings:
        print(f"\n[警告] 以下部门无匹配模板，已跳过: {', '.join(set(warnings))}")

    print(f"\n完成! 文件输出至: {OUTPUT_DIR}")
    print(f"已生成 {len(tracking_keys)} 人，已追踪 {len(new_generated_str)} 人")


def _process_excel(excel_path):
    """从本地 Excel 处理数据"""
    df = pd.read_excel(excel_path, engine="openpyxl")
    df.columns = [str(c).replace("\n", "").strip() for c in df.columns]
    print(f"读取 Excel: {excel_path}")

    emp_ids = []
    for _, row in df.iterrows():
        emp_id = row['工号']
        if isinstance(emp_id, (int, float)):
            emp_ids.append(str(int(emp_id)).zfill(3))
        else:
            emp_ids.append(str(emp_id).strip().zfill(3))
    employees = query_employees_by_ids(emp_ids)

    merged_data = []
    for _, row in df.iterrows():
        emp_id_raw = row['工号']
        emp_id = str(int(emp_id_raw)).zfill(3) if isinstance(emp_id_raw, (int, float)) else str(emp_id_raw).strip().zfill(3)
        emp_info = employees.get(emp_id)
        if not emp_info:
            print(f"  [警告] 工号 {emp_id} 在知识库中不存在，跳过")
            continue
        merged_data.append({
            "工号": emp_id, "name": emp_info["name"], "department": emp_info["department"],
            "position": translate_pos(emp_info["position"]), "gender": emp_info["gender"],
            "passport": str(row['护照号']).strip() if pd.notna(row.get('护照号')) else "N/A",
            "出发时间": row['出发时间'], "返程时间": row['返程时间'],
            "depart_date": row['出发时间'], "return_date": row['返程时间']
        })

    if not merged_data:
        print("[错误] 无有效数据")
        sys.exit(1)

    df_merged = pd.DataFrame(merged_data)
    _process_dataframe(df_merged)


if __name__ == "__main__":
    main()
