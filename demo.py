# -*- coding: utf-8 -*-
"""
邀请函自动生成器 - 一键演示脚本

面试时只需运行: python demo.py
"""

import os, sys, json, time, subprocess

# 确保 Windows GBK 终端不报编码错
try:
    sys.stdout.reconfigure(encoding='utf-8')
except:
    pass

REPO = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, REPO)

DIVIDER = "=" * 56
THIN = "-" * 56


def pause(msg="按回车继续..."):
    input(f"\n  {msg}")


def step(title, func):
    """运行一个演示步骤"""
    print(f"\n{DIVIDER}")
    print(f"  {title}")
    print(DIVIDER)
    func()
    print()


def main():
    print(f"\n{'█' * 56}")
    print(f"█  {'邀请函自动生成器 · Invitation Letter Generator':<51} █")
    print(f"█  {'Demo 演示':<53} █")
    print(f"{'█' * 56}")

    # ── Step 1: 展示知识库 ──
    step("Step 1/4  查看知识库（6 名员工）", lambda: show_database())

    # ── Step 2: 展示出差数据 ──
    step("Step 2/4  查看出差数据（demo_data.xlsx）", lambda: show_excel())

    # ── Step 3: 生成邀请函 ──
    step("Step 3/4  自动生成邀请函 + 答复函", lambda: generate())

    # ── Step 4: 展示输出 ──
    step("Step 4/4  查看生成结果", lambda: show_output())

    # ── 完成 ──
    print(f"{DIVIDER}")
    print(f"  [OK] 演示完成！")
    print(f"  输出目录: {os.path.join(REPO, 'output')}")
    print(f"{DIVIDER}")
    print()

    if os.name == "nt":
        try:
            os.startfile(os.path.join(REPO, "output"))
            print("  [>] 已自动打开 output 文件夹")
        except:
            pass


def show_database():
    import sqlite3
    conn = sqlite3.connect(os.path.join(REPO, "knowledge_base.db"))
    c = conn.cursor()
    c.execute("SELECT * FROM employees ORDER BY emp_id")
    rows = c.fetchall()
    conn.close()

    hdr = f"  {'工号':<6} {'姓名':<8} {'部门':<8} {'职位':<14} {'性别':<4} {'护照号':<14}"
    print(hdr)
    print(f"  {THIN[2:]}")
    for r in rows:
        print(f"  {r[0]:<6} {r[1]:<8} {r[2]:<8} {r[3]:<14} {r[4]:<4} {r[5]:<14}")
    print(f"\n  [*] 知识库存储员工基础信息（姓名/部门/职位/性别）")
    print(f"  [*] 出差 Excel 只填工号，其余字段自动查询")


def show_excel():
    import pandas as pd
    df = pd.read_excel(os.path.join(REPO, "demo_data.xlsx"), engine="openpyxl")

    # 关联知识库查询姓名
    import sqlite3
    conn = sqlite3.connect(os.path.join(REPO, "knowledge_base.db"))
    c = conn.cursor()
    c.execute("SELECT emp_id, name FROM employees")
    name_map = {row[0]: row[1] for row in c.fetchall()}
    conn.close()

    hdr = f"  {'工号':<6} {'姓名':<8} {'出发时间':<12} {'返程时间':<12} {'护照号':<14}"
    print(hdr)
    print(f"  {THIN[2:]}")
    for _, row in df.iterrows():
        eid = str(int(row["工号"])).zfill(3)
        name = name_map.get(eid, "?")
        print(f"  {eid:<6} {name:<8} {str(row['出发时间']):<12} {str(row['返程时间']):<12} {row['护照号']:<14}")

    print(f"\n  [*] 这张表只填了 4 列：工号、出发时间、返程时间、护照号")
    print(f"  [*] 姓名/部门/职位/性别 → 全部从知识库自动匹配")


def generate():
    # 清空旧输出和追踪文件
    out_dir = os.path.join(REPO, "output")
    for f in os.listdir(out_dir):
        os.remove(os.path.join(out_dir, f))
    tracking = os.path.join(REPO, "generated.json")
    if os.path.exists(tracking):
        os.remove(tracking)

    from generate_docs import _process_excel as gen
    gen(os.path.join(REPO, "demo_data.xlsx"))


def show_output():
    out_dir = os.path.join(REPO, "output")
    files = sorted([f for f in os.listdir(out_dir) if f.endswith(".docx")])
    if not files:
        print("  (无文件)")
        return

    # 按部门分组展示
    from collections import defaultdict
    groups = defaultdict(list)
    for f in files:
        dept = f.split("_")[0]
        groups[dept].append(f)

    for dept, flist in groups.items():
        print(f"  > {dept}")
        for f in sorted(flist):
            size_kb = os.path.getsize(os.path.join(out_dir, f)) / 1024
            print(f"     ✓ {f}  ({size_kb:.0f} KB)")

    total = len(files)
    dept_count = len(groups)
    print(f"\n  [*] 共生成 {total} 个文件，覆盖 {dept_count} 个部门")
    print(f"  [*] 每个部门各 1 份邀请函 + 1 份答复函")


if __name__ == "__main__":
    main()
