# -*- coding: utf-8 -*-
"""
简化 Excel 文件，只保留必要的列
工号、出发时间、返程时间、护照号

用法:
    python simplify_excel.py                          # 使用默认 Excel 路径
    python simplify_excel.py path/to/your_file.xlsx   # 指定 Excel 文件
"""

import sys
import os
import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')

# === 配置 ===
DEFAULT_EXCEL = r"path/to/your/weekly_data.Excel.xlsx"
OUTPUT_EXCEL = r"path/to/your/weekly_data_simple.xlsx"


def simplify_excel(excel_path, output_path):
    """简化 Excel 文件"""
    print(f"读取 Excel: {excel_path}")

    df = pd.read_excel(excel_path, engine="openpyxl")
    df.columns = [str(c).replace("\n", "").strip() for c in df.columns]

    # 显示原始列名
    print(f"\n原始列名 ({len(df.columns)} 列):")
    for i, col in enumerate(df.columns):
        print(f"  {i}: {col}")

    # 检查必要的列
    required_cols = {
        '名字': '工号',  # 需要转换为工号
        '出发时间Depart date': '出发时间',
        '返程时间Back date': '返程时间',
        '护照号': '护照号'
    }

    missing_cols = []
    for col in required_cols.keys():
        if col not in df.columns:
            missing_cols.append(col)

    if missing_cols:
        print(f"\n[错误] 找不到列: {', '.join(missing_cols)}")
        return False

    # 创建简化的 DataFrame
    df_simple = pd.DataFrame()

    # 复制必要的列
    df_simple['出发时间'] = df['出发时间Depart date']
    df_simple['返程时间'] = df['返程时间Back date']
    df_simple['护照号'] = df['护照号']

    # 工号列需要后续从知识库查询填充
    # 暂时使用序号或留空
    if '序号' in df.columns:
        df_simple.insert(0, '工号', df['序号'])
    else:
        df_simple.insert(0, '工号', range(1, len(df_simple) + 1))

    # 保存简化后的 Excel
    df_simple.to_excel(output_path, index=False, engine="openpyxl")

    print(f"\n简化后的列名 ({len(df_simple.columns)} 列):")
    for i, col in enumerate(df_simple.columns):
        print(f"  {i}: {col}")

    print(f"\n前5行数据:")
    print(df_simple.head().to_string())

    print(f"\n保存到: {output_path}")
    print(f"总行数: {len(df_simple)}")

    return True


def main():
    excel_path = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_EXCEL

    if not os.path.exists(excel_path):
        print(f"[错误] Excel 文件不存在: {excel_path}")
        sys.exit(1)

    simplify_excel(excel_path, OUTPUT_EXCEL)
    print(f"\n完成!")


if __name__ == "__main__":
    main()
