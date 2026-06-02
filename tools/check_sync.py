# -*- coding: utf-8 -*-
"""
检查 SKILL.md 和 CLAUDE.md 的业务规则是否一致
用法: python check_sync.py
"""

import sys
import re

sys.stdout.reconfigure(encoding='utf-8')

SKILL_PATH = r"~/.claude/skills/invitation/SKILL.md"
CLAUDE_PATH = r"/path/to/CLAUDE.md"

# 要检查的关键规则
RULES_TO_CHECK = [
    "部门映射",
    "日期计算",
    "日期格式",
    "职位处理",
    "字体",
    "字号",
    "Century Gothic",
    "Month D, YYYY",
    "小四",
    "12pt",
    "知识库",
    "SQLite",
]


def read_file(path):
    with open(path, "r", encoding="utf-8") as f:
        return f.read()


def check_rule_in_content(rule, content):
    return rule in content


def main():
    print("=== SKILL.md 与 CLAUDE.md 同步检查 ===\n")

    skill_content = read_file(SKILL_PATH)
    claude_content = read_file(CLAUDE_PATH)

    all_ok = True

    for rule in RULES_TO_CHECK:
        in_skill = check_rule_in_content(rule, skill_content)
        in_claude = check_rule_in_content(rule, claude_content)

        if in_skill and in_claude:
            status = "[OK] 两边都有"
        elif in_skill and not in_claude:
            status = "[WARN] CLAUDE.md 缺失"
            all_ok = False
        elif not in_skill and in_claude:
            status = "[WARN] SKILL.md 缺失"
            all_ok = False
        else:
            status = "[FAIL] 两边都缺失"
            all_ok = False

        print(f"  {rule}: {status}")

    print("\n" + "=" * 40)

    if all_ok:
        print("[OK] 所有规则同步正常！")
    else:
        print("[WARN] 存在不一致，请检查并修复！")


if __name__ == "__main__":
    main()
