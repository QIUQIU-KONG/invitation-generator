# -*- coding: utf-8 -*-
"""Small desktop launcher for the Feishu invitation workflow."""

from __future__ import annotations

import os
import subprocess
import sys
import threading
import tkinter as tk
import winreg
from pathlib import Path
from tkinter import messagebox, scrolledtext


ROOT = Path(__file__).resolve().parent
SCRIPT = ROOT / "generate_docs.py"
OUTPUT_DIR = ROOT / "output"
FEISHU_EXE = Path(os.environ.get("LOCALAPPDATA", "")) / "Feishu" / "Feishu.exe"


def load_user_environment(env: dict[str, str]) -> dict[str, str]:
    """Pick up user-level variables set by setx without requiring a restart."""
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment") as key:
            for name in ("FEISHU_BASE_TOKEN", "FEISHU_TABLE_ID", "LARK_CLI_PATH"):
                try:
                    value, _ = winreg.QueryValueEx(key, name)
                except FileNotFoundError:
                    continue
                if value:
                    env[name] = str(value)
    except OSError:
        pass
    return env


def decode_output(data: bytes) -> str:
    if not data:
        return ""
    for encoding in ("utf-8", "gb18030", "gbk", "cp936"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            pass
    return data.decode("utf-8", errors="replace")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("邀请函自动化")
        self.geometry("900x620")
        self.minsize(760, 520)
        self.configure(bg="#f6f7f9")

        self.status = tk.StringVar(value="先在飞书填写出差记录，然后点击“预览检查”。")
        self._build_ui()

    def _build_ui(self):
        header = tk.Frame(self, bg="#f6f7f9")
        header.pack(fill="x", padx=18, pady=(16, 8))

        title = tk.Label(
            header,
            text="邀请函自动化",
            font=("Microsoft YaHei UI", 18, "bold"),
            bg="#f6f7f9",
            fg="#1f2937",
        )
        title.pack(anchor="w")

        subtitle = tk.Label(
            header,
            text="飞书填写完成后，先预览检查，再正式生成邀请函和答复函。",
            font=("Microsoft YaHei UI", 10),
            bg="#f6f7f9",
            fg="#4b5563",
        )
        subtitle.pack(anchor="w", pady=(4, 0))

        actions = tk.Frame(self, bg="#f6f7f9")
        actions.pack(fill="x", padx=18, pady=8)

        self.preview_btn = self._button(actions, "预览检查", lambda: self.run_workflow(["--dry-run"]))
        self.preview_btn.pack(side="left", padx=(0, 8))

        self.generate_btn = self._button(actions, "正式生成", lambda: self.run_workflow(["--yes"]))
        self.generate_btn.pack(side="left", padx=8)

        self.force_btn = self._button(actions, "强制重生成", lambda: self.confirm_force())
        self.force_btn.pack(side="left", padx=8)

        self.output_btn = self._button(actions, "打开输出文件夹", self.open_output)
        self.output_btn.pack(side="left", padx=8)

        self.feishu_btn = self._button(actions, "打开飞书", self.open_feishu)
        self.feishu_btn.pack(side="left", padx=8)

        status = tk.Label(
            self,
            textvariable=self.status,
            font=("Microsoft YaHei UI", 10),
            bg="#f6f7f9",
            fg="#374151",
            anchor="w",
        )
        status.pack(fill="x", padx=18, pady=(6, 4))

        self.log = scrolledtext.ScrolledText(
            self,
            wrap="word",
            font=("Consolas", 10),
            bg="#ffffff",
            fg="#111827",
            relief="solid",
            borderwidth=1,
        )
        self.log.pack(fill="both", expand=True, padx=18, pady=(4, 18))
        self.log.insert("end", "等待操作...\n")
        self.log.configure(state="disabled")

    def _button(self, parent, text, command):
        return tk.Button(
            parent,
            text=text,
            command=command,
            font=("Microsoft YaHei UI", 10),
            bg="#2563eb",
            fg="#ffffff",
            activebackground="#1d4ed8",
            activeforeground="#ffffff",
            relief="flat",
            padx=14,
            pady=8,
            cursor="hand2",
        )

    def set_buttons(self, enabled: bool):
        state = "normal" if enabled else "disabled"
        for button in (self.preview_btn, self.generate_btn, self.force_btn, self.output_btn, self.feishu_btn):
            button.configure(state=state)

    def append_log(self, text: str):
        self.log.configure(state="normal")
        self.log.insert("end", text)
        self.log.see("end")
        self.log.configure(state="disabled")

    def replace_log(self, text: str):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.insert("end", text)
        self.log.configure(state="disabled")

    def run_workflow(self, args: list[str]):
        if not SCRIPT.exists():
            messagebox.showerror("找不到脚本", f"没有找到：{SCRIPT}")
            return
        self.set_buttons(False)
        self.status.set("正在处理，请稍等...")
        self.replace_log("开始运行...\n\n")

        def worker():
            cmd = [sys.executable, str(SCRIPT), *args]
            try:
                env = os.environ.copy()
                env = load_user_environment(env)
                env["PYTHONIOENCODING"] = "utf-8"
                proc = subprocess.run(
                    cmd,
                    cwd=str(ROOT),
                    capture_output=True,
                    env=env,
                )
                output = decode_output(proc.stdout)
                if proc.stderr:
                    output += "\n[错误信息]\n" + decode_output(proc.stderr)
                code = proc.returncode
            except Exception as exc:
                output = str(exc)
                code = 1

            def done():
                self.append_log(output or "(没有输出)\n")
                self.set_buttons(True)
                if code == 0:
                    self.status.set("完成。请查看上方日志或打开输出文件夹。")
                else:
                    self.status.set("没有完成。请查看日志里的提示。")

            self.after(0, done)

        threading.Thread(target=worker, daemon=True).start()

    def confirm_force(self):
        if messagebox.askyesno("确认强制重生成", "这会忽略历史记录，重新生成飞书中可生成的记录。继续吗？"):
            self.run_workflow(["--force", "--yes"])

    def open_output(self):
        OUTPUT_DIR.mkdir(exist_ok=True)
        os.startfile(str(OUTPUT_DIR))

    def open_feishu(self):
        if FEISHU_EXE.exists():
            subprocess.Popen([str(FEISHU_EXE)])
        else:
            messagebox.showinfo("未找到飞书", "没有找到本机飞书程序，请从开始菜单打开飞书。")


if __name__ == "__main__":
    App().mainloop()
