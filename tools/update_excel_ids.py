# -*- coding: utf-8 -*-
"""
更新简化 Excel 中的工号，使用知识库中的工号

用法:
    python update_excel_ids.py
"""

import sys
import os
import sqlite3
import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')

# === 配置 ===
EXCEL_PATH = r"path/to/your/weekly_data_simple.xlsx"
ORIGINAL_EXCEL = r"path/to/your/weekly_data.Excel.xlsx"
DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "knowledge_base.db")


def main():
    if not os.path.exists(DB_PATH):
        print(f"[错误] 知识库不存在: {DB_PATH}")
        sys.exit(1)

    if not os.path.exists(ORIGINAL_EXCEL):
        print(f"[错误] 原始 Excel 不存在: {ORIGINAL_EXCEL}")
        sys.exit(1)

    # 读取原始 Excel
    print(f"读取原始 Excel: {ORIGINAL_EXCEL}")
    df_orig = pd.read_excel(ORIGINAL_EXCEL, engine="openpyxl")
    df_orig.columns = [str(c).replace("\n", "").strip() for c in df_orig.columns]

    # 读取简化 Excel
    print(f"读取简化 Excel: {EXCEL_PATH}")
    df_simple = pd.read_excel(EXCEL_PATH, engine="openpyxl")
    df_simple.columns = [str(c).replace("\n", "").strip() for c in df_simple.columns]

    # 从知识库查询姓名到工号的映射
    print("从知识库查询姓名到工号的映射...")
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT emp_id, name FROM employees")
    name_to_id = {row[1]: row[0] for row in cursor.fetchall()}
    conn.close()

    print(f"知识库中有 {len(name_to_id)} 人")

    # 更新简化 Excel 中的工号
    print("\n更新工号...")
    updated = 0
    not_found = []

    # 创建新的工号列
    new_ids = []
    for i, row in df_simple.iterrows():
        # 从原始 Excel 获取姓名
        if i < len(df_orig):
            name = str(df_orig.iloc[i]['名字']).strip()
            if name in name_to_id:
                new_ids.append(name_to_id[name])
                updated += 1
            else:
                new_ids.append(str(row['工号']))
                not_found.append(name)
        else:
            new_ids.append(str(row['工号']))

    # 更新工号列
    df_simple['工号'] = new_ids

    # 保存更新后的简化 Excel
    df_simple.to_excel(EXCEL_PATH, index=False, engine="openpyxl")

    print(f"\n更新完成:")
    print(f"  更新: {updated} 人")
    if not_found:
        print(f"  未找到: {', '.join(not_found)}")

    print(f"\n简化 Excel 前5行:")
    print(df_simple.head().to_string())


if __name__ == "__main__":
    main()
