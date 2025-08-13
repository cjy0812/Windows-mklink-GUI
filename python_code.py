import os
import subprocess
import sys
import json
import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from pathlib import Path
from ttkbootstrap.tooltip import ToolTip

PREF_FILE = Path.home() / ".winmklink_prefs.json"

LANG_TEXTS = {
    "en": {
        "title": "Winmklink GUI Tool",
        "target": "Target Path:",
        "link": "Link Path:",
        "link_type": "Link Type:",
        "choose": "Browse",
        "lang": "Language:",
        "create": "Create Link",
        "link_types": ["Symbolic Link (File)", "Symbolic Link (Directory)", "Hard Link (File)", "Junction (Directory)"],
        "output": "Output:",
        "cmd_preview": "Command Preview:",
        "admin_run": "Run as Admin",
        "tip_target": "Select the existing target file or directory",
        "tip_link": "Choose where to create the link (will suggest target's name)",
        "tip_link_type": "Select the type of link to create",
        "tip_create": "Execute the mklink command",
        "tip_admin": "Run the command in an elevated command prompt",
        "tip_lang": "Switch application language",
    },
    "zh": {
        "title": "Winmklink 图形工具",
        "target": "目标路径：",
        "link": "链接路径：",
        "link_type": "链接类型：",
        "choose": "浏览",
        "lang": "语言：",
        "create": "创建链接",
        "link_types": ["符号链接（文件）", "符号链接（目录）", "硬链接（文件）", "目录联接（Junction）"],
        "output": "输出：",
        "cmd_preview": "命令预览：",
        "admin_run": "以管理员运行",
        "tip_target": "选择已存在的目标文件或目录",
        "tip_link": "选择链接位置（默认建议目标名称）",
        "tip_link_type": "选择要创建的链接类型",
        "tip_create": "执行 mklink 命令",
        "tip_admin": "在管理员命令提示符中运行命令",
        "tip_lang": "切换应用语言",
    }
}

class WinMkLinkApp(ttk.Window):
    def __init__(self):
        lang_pref = self.load_lang_pref()
        super().__init__(themename="cosmo")
        self.lang = lang_pref if lang_pref in LANG_TEXTS else "zh"
        self.texts = LANG_TEXTS[self.lang]
        self.title(self.texts["title"])
        self.geometry("650x520")

        self.target_var = tk.StringVar()
        self.link_var = tk.StringVar()
        self.link_type_var = tk.StringVar(value=self.texts["link_types"][0])
        self.cmd_preview_var = tk.StringVar()

        # Language selection
        lang_frame = ttk.Frame(self)
        lang_frame.pack(fill=X, pady=5, padx=10)
        ttk.Label(lang_frame, text=self.texts["lang"]).pack(side=LEFT)
        self.lang_combo = ttk.Combobox(lang_frame, values=["中文", "English"], width=10)
        self.lang_combo.set("中文" if self.lang == "zh" else "English")
        self.lang_combo.bind("<<ComboboxSelected>>", self.change_lang)
        self.lang_combo.pack(side=LEFT, padx=5)
        ToolTip(self.lang_combo, self.texts["tip_lang"])

        # Target path
        target_frame = ttk.Frame(self)
        target_frame.pack(fill=X, pady=5, padx=10)
        ttk.Label(target_frame, text=self.texts["target"]).pack(side=LEFT)
        self.target_entry = ttk.Entry(target_frame, textvariable=self.target_var, width=50)
        self.target_entry.pack(side=LEFT, padx=5)
        ttk.Button(target_frame, text=self.texts["choose"], command=self.browse_target).pack(side=LEFT)
        ToolTip(self.target_entry, self.texts["tip_target"])

        # Link path
        link_frame = ttk.Frame(self)
        link_frame.pack(fill=X, pady=5, padx=10)
        ttk.Label(link_frame, text=self.texts["link"]).pack(side=LEFT)
        self.link_entry = ttk.Entry(link_frame, textvariable=self.link_var, width=50)
        self.link_entry.pack(side=LEFT, padx=5)
        ttk.Button(link_frame, text=self.texts["choose"], command=self.browse_link).pack(side=LEFT)
        ToolTip(self.link_entry, self.texts["tip_link"])

        # Link type
        type_frame = ttk.Frame(self)
        type_frame.pack(fill=X, pady=5, padx=10)
        ttk.Label(type_frame, text=self.texts["link_type"]).pack(side=LEFT)
        self.type_combo = ttk.Combobox(type_frame, values=self.texts["link_types"], textvariable=self.link_type_var, width=30)
        self.type_combo.pack(side=LEFT, padx=5)
        self.type_combo.bind("<<ComboboxSelected>>", self.update_cmd_preview)
        ToolTip(self.type_combo, self.texts["tip_link_type"])

        # Command preview
        cmd_frame = ttk.Frame(self)
        cmd_frame.pack(fill=X, pady=5, padx=10)
        ttk.Label(cmd_frame, text=self.texts["cmd_preview"]).pack(anchor=W)
        self.cmd_entry = ttk.Entry(cmd_frame, textvariable=self.cmd_preview_var, width=80, state="readonly")
        self.cmd_entry.pack(fill=X, pady=2)

        # Buttons
        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill=X, pady=10, padx=10)
        create_btn = ttk.Button(btn_frame, text=self.texts["create"], bootstyle=SUCCESS, command=self.create_link)
        create_btn.pack(side=LEFT, padx=5)
        ToolTip(create_btn, self.texts["tip_create"])

        admin_btn = ttk.Button(btn_frame, text=self.texts["admin_run"], bootstyle=WARNING, command=self.run_as_admin)
        admin_btn.pack(side=LEFT, padx=5)
        ToolTip(admin_btn, self.texts["tip_admin"])

        # Output
        output_frame = ttk.LabelFrame(self, text=self.texts["output"])
        output_frame.pack(fill=BOTH, expand=YES, padx=10, pady=5)
        self.output_text = tk.Text(output_frame, wrap="word", height=12)
        self.output_text.pack(fill=BOTH, expand=YES)

        self.update_cmd_preview()

    def browse_target(self):
        link_type = self.link_type_var.get()
        # 判断链接类型来选择文件或目录选择器
        if "目录" in link_type or "Directory" in link_type or "Junction" in link_type:
            path = filedialog.askdirectory(title=self.texts["target"])
        else:
            path = filedialog.askopenfilename(title=self.texts["target"])
        if path:
            self.target_var.set(path)
            self.update_cmd_preview()

    def browse_link(self):
        target_path = self.target_var.get()
        initialdir = os.path.dirname(target_path) if target_path else os.getcwd()
        # 取目标文件/目录的名称做默认文件名建议
        suggested_name = ""
        if target_path:
            suggested_name = os.path.basename(target_path)
        # 让用户选择链接的完整路径（保存对话框），默认文件名为目标名
        path = filedialog.asksaveasfilename(title=self.texts["link"], initialdir=initialdir, initialfile=suggested_name)
        if path:
            self.link_var.set(path)
            self.update_cmd_preview()

    def update_cmd_preview(self, event=None):
        target = self.target_var.get()
        link = self.link_var.get()
        link_type = self.link_type_var.get()
        if not target or not link:
            self.cmd_preview_var.set("")
            return
        cmd = ["mklink"]
        if "目录" in link_type or "Directory" in link_type:
            if "符号" in link_type or "Symbolic" in link_type:
                cmd.append("/D")
            elif "Junction" in link_type:
                cmd.append("/J")
        elif "硬" in link_type or "Hard" in link_type:
            cmd.append("/H")
        cmd.append(f'"{link}"')
        cmd.append(f'"{target}"')
        self.cmd_preview_var.set(" ".join(cmd))

    def create_link(self):
        cmd = self.cmd_preview_var.get()
        if not cmd:
            messagebox.showwarning(self.texts["title"], "Please fill target and link paths.")
            return
        try:
            # 指定 gbk 编码防止中文乱码
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True, encoding="gbk")
            self.output_text.insert("end", result.stdout)
            if result.stderr:
                self.output_text.insert("end", result.stderr)
            self.output_text.see("end")
        except Exception as e:
            messagebox.showerror(self.texts["title"], str(e))

    def run_as_admin(self):
        cmd = self.cmd_preview_var.get()
        if not cmd:
            messagebox.showwarning(self.texts["title"], "Please fill target and link paths.")
            return
        try:
            # powershell命令提升权限执行cmd命令
            subprocess.run(
                f'powershell -Command "Start-Process cmd -ArgumentList \'/k {cmd}\' -Verb RunAs"',
                shell=True)
        except Exception as e:
            messagebox.showerror(self.texts["title"], str(e))

    def change_lang(self, event=None):
        sel = self.lang_combo.get()
        self.lang = "zh" if sel == "中文" else "en"
        self.save_lang_pref(self.lang)
        python = sys.executable
        os.execl(python, python, *sys.argv)

    def load_lang_pref(self):
        if PREF_FILE.exists():
            try:
                with open(PREF_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    return data.get("lang", "zh")
            except Exception:
                return "zh"
        return "zh"

    def save_lang_pref(self, lang):
        try:
            with open(PREF_FILE, "w", encoding="utf-8") as f:
                json.dump({"lang": lang}, f)
        except Exception:
            pass


if __name__ == "__main__":
    app = WinMkLinkApp()
    app.mainloop()
