# -*- coding: utf-8 -*-
"""
初始化 SQLite 知识库
从 Excel 导入人员基础信息到 knowledge_base.db

用法:
    python init_database.py                          # 使用默认 Excel 路径
    python init_database.py path/to/your_file.xlsx   # 指定 Excel 文件
"""

import sys
import os
import sqlite3
import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')

# === 配置 ===
DEFAULT_EXCEL = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sample_data.xlsx")
DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "knowledge_base.db")


def create_database():
    """创建数据库和表"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS employees (
            emp_id TEXT PRIMARY KEY,
            name TEXT NOT NULL,
            department TEXT NOT NULL,
            position TEXT NOT NULL,
            gender TEXT NOT NULL,
            passport TEXT DEFAULT ''
        )
    ''')

    conn.commit()
    return conn


def import_from_excel(conn, excel_path):
    """从 Excel 导入人员数据（按工号去重）"""
    print(f"读取 Excel: {excel_path}")

    df = pd.read_excel(excel_path, engine="openpyxl")
    df.columns = [str(c).replace("\n", "").strip() for c in df.columns]

    # 检查必要的列
    required_cols = ['名字', '部门Dept.', 'positition', '出差人员性别']
    for col in required_cols:
        if col not in df.columns:
            print(f"[错误] 找不到列: {col}")
            return False

    cursor = conn.cursor()

    # 获取现有最大工号
    cursor.execute("SELECT MAX(CAST(emp_id AS INTEGER)) FROM employees")
    result = cursor.fetchone()
    max_id = result[0] if result[0] else 0

    # 获取已有工号集合（用于去重）
    cursor.execute("SELECT emp_id FROM employees")
    existing_ids = {row[0] for row in cursor.fetchall()}

    # 检查 Excel 是否有工号列
    has_emp_id = '序号' in df.columns or '工号' in df.columns

    inserted = 0
    skipped = 0

    for _, row in df.iterrows():
        name = str(row['名字']).strip()
        department = str(row['部门Dept.']).strip()
        position = str(row['positition']).strip() if pd.notna(row['positition']) else 'N/A'
        gender = str(row['出差人员性别']).strip()

        # 护照号（如果有的话）
        passport_col = next((c for c in df.columns if '护照' in c), None)
        passport = str(row[passport_col]).strip() if passport_col and pd.notna(row.get(passport_col)) else ''

        # 确定工号：优先使用 Excel 中的工号，否则自动生成
        if has_emp_id and '工号' in df.columns:
            emp_id = str(row['工号']).strip()
        elif has_emp_id and '序号' in df.columns:
            emp_id = str(row['序号']).strip().zfill(3)
        else:
            max_id += 1
            emp_id = str(max_id).zfill(3)

        # 按工号去重
        if emp_id in existing_ids:
            skipped += 1
            continue

        existing_ids.add(emp_id)

        # 插入数据
        cursor.execute('''
            INSERT OR REPLACE INTO employees (emp_id, name, department, position, gender, passport)
            VALUES (?, ?, ?, ?, ?, ?)
        ''', (emp_id, name, department, position, gender, passport))

        inserted += 1
        print(f"  插入: {emp_id} - {name} ({department})")

    conn.commit()

    print(f"\n导入完成:")
    print(f"  插入: {inserted} 人")
    print(f"  跳过: {skipped} 人（工号已存在）")

    return True


def show_database(conn):
    """显示数据库内容"""
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM employees ORDER BY emp_id")
    rows = cursor.fetchall()

    print(f"\n=== 知识库内容 ({len(rows)} 人) ===")
    header = '{:<8} {:<12} {:<10} {:<22} {:<6} {:<14}'.format('工号', '姓名', '部门', '职位', '性别', '护照号')
    print(header)
    print('-' * 72)

    for row in rows:
        passport = row[5] if len(row) > 5 and row[5] else '-'
        line = '{:<8} {:<12} {:<10} {:<22} {:<6} {:<14}'.format(row[0], row[1], row[2], row[3], row[4], passport)
        print(line)


def main():
    excel_path = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_EXCEL

    if not os.path.exists(excel_path):
        print(f"[错误] Excel 文件不存在: {excel_path}")
        sys.exit(1)

    print(f"创建数据库: {DB_PATH}")
    conn = create_database()

    if import_from_excel(conn, excel_path):
        show_database(conn)

    conn.close()
    print(f"\n完成! 数据库保存在: {DB_PATH}")


if __name__ == "__main__":
    main()
