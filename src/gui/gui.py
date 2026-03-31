import calendar
from datetime import date, timedelta

import customtkinter as ctk
from tkinter import filedialog
import yaml
import threading
import sys
import io
import os

from src.analysis import start_analysis_flow
from src.core.config import get_base_dir, load_config

THEME_DATA = {
    "Dark": {"bg": "#0d1117", "card": "#161b22", "border": "#30363d", "text": "#f0f6fc", "input_bg": "#010409",
             "icon_color": "#8b949e"},
    "Light": {"bg": "#ffffff", "card": "#f6f8fa", "border": "#d0d7de", "text": "#1f2328", "input_bg": "#ffffff",
              "icon_color": "#6e7781"}
}


class ModernAnalysisGUI(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Data Engine v6.4")
        self.geometry("480x600")
        self.resizable(False, False)

        # 确保 config.yaml 始终相对于工程根目录定位 (兼容 EXE)
        base_dir = get_base_dir()
        self.config_path = os.path.join(base_dir, 'config.yaml')
        self.config = self.init_config()

        # --- 路径自愈逻辑：分发给他人后，自动修正失效的绝对路径 ---
        if 'file_config' in self.config:
            for key in ['input_file', 'output_file', 'template_file']:
                path = self.config['file_config'].get(key, "")
                if path and not os.path.exists(path):
                    # 尝试在 EXE 同级目录找同名文件
                    filename = os.path.basename(path)
                    guess = os.path.join(base_dir, filename)
                    if os.path.exists(guess):
                        self.config['file_config'][key] = os.path.abspath(guess)
        # ---------------------------------------------------

        self.appearance_mode = self.config.get('theme', 'Dark')
        ctk.set_appearance_mode(self.appearance_mode)

        self.setup_ui()
        self.apply_theme_styles()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def init_config(self):
        # --- 核心逻辑：计算上个月的区间 ---
        today = date.today()
        # 计算上个月的年份和月份
        first_day_of_this_month = today.replace(day=1)
        last_day_of_last_month = first_day_of_this_month - timedelta(days=1)

        last_month_year = last_day_of_last_month.year
        last_month_num = last_day_of_last_month.month

        # 获取上个月第一天和最后一天
        start_date_str = f"{last_month_year}-{last_month_num:02d}-01"
        _, last_day_num = calendar.monthrange(last_month_year, last_month_num)
        end_date_str = f"{last_month_year}-{last_month_num:02d}-{last_day_num:02d}"
        # --------------------------------

        default = {
            'file_config': {'input_file': '', 'output_file': ''},
            'analysis_period': {
                'start_date': start_date_str,  # 动态生成的上月1号
                'end_date': end_date_str  # 动态生成的上月月底
            },
            'theme': 'Dark',
            'execution_mode': 'all'
        }

        # 如果配置文件不存在，创建并使用上述动态默认值
        if not os.path.exists(self.config_path):
            with open(self.config_path, 'w', encoding='utf-8') as f:
                yaml.dump(default, f)
            return default

        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                loaded = yaml.safe_load(f) or default
                # 即使文件存在，如果里面的日期是空的，也可以选择是否覆盖为上月
                # 这里我们尊重用户上次保存的值，如果用户没改过，初始化时就用上面的计算值
                return loaded
        except:
            return default

    def setup_ui(self):
        self.header = ctk.CTkFrame(self, fg_color="transparent", height=35)
        self.header.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(self.header, text="Data Engine", font=("Inter", 16, "bold")).pack(side="left")
        self.theme_btn = ctk.CTkButton(self.header, text="", width=30, height=30, fg_color="transparent",
                                       command=self.toggle_theme)
        self.theme_btn.pack(side="right")

        self.main_container = ctk.CTkFrame(self, fg_color="transparent")
        self.main_container.pack(fill="x", padx=20)

        # 输入
        self.file_card = self.create_card("DATA SOURCE (EXCEL)")
        self.input_entry = ctk.CTkEntry(self.file_card, height=28, border_width=1)
        self.input_entry.insert(0, self.config['file_config'].get('input_file', ''))
        self.input_entry.pack(side="left", expand=True, fill="x", padx=(10, 5), pady=8)
        ctk.CTkButton(self.file_card, text="Browse", width=50, height=28, command=lambda: self.browse('file')).pack(
            side="right", padx=(0, 10))

        # 输出 (带自动补全文件名)
        self.out_card = self.create_card("OUTPUT PATH (AUTO-NAMED)")
        self.out_entry = ctk.CTkEntry(self.out_card, height=28, border_width=1)
        self.out_entry.insert(0, self.config['file_config'].get('output_file', ''))
        self.out_entry.pack(side="left", expand=True, fill="x", padx=(10, 5), pady=8)
        ctk.CTkButton(self.out_card, text="Select", width=50, height=28, command=lambda: self.browse('dir')).pack(
            side="right", padx=(0, 10))

        # 时间
        self.date_card = self.create_card("DATE RANGE")
        self.start_date = ctk.CTkEntry(self.date_card, height=28, border_width=1)
        self.start_date.insert(0, self.config['analysis_period']['start_date'])
        self.start_date.pack(side="left", expand=True, fill="x", padx=(10, 2), pady=8)
        ctk.CTkLabel(self.date_card, text="–").pack(side="left")
        self.end_date = ctk.CTkEntry(self.date_card, height=28, border_width=1)
        self.end_date.insert(0, self.config['analysis_period']['end_date'])
        self.end_date.pack(side="left", expand=True, fill="x", padx=(2, 10), pady=8)

        # 日志
        self.log_output = ctk.CTkTextbox(self, corner_radius=6, border_width=1, font=("Consolas", 10))
        self.log_output.pack(fill="both", expand=True, padx=20, pady=(5, 10))

        # 运行按钮
        self.run_btn = ctk.CTkButton(self, text="Run Analysis", font=("Inter", 13, "bold"), height=36,
                                     command=self.start_thread)
        self.run_btn.pack(fill="x", padx=20, pady=(0, 15))

    def create_card(self, title):
        card = ctk.CTkFrame(self.main_container, corner_radius=8, border_width=1)
        card.pack(fill="x", pady=3)
        ctk.CTkLabel(card, text=title, font=("Inter", 9, "bold"), text_color="#6e7781").pack(anchor="w", padx=10,
                                                                                             pady=(4, 0))
        return card

    def apply_theme_styles(self):
        c = THEME_DATA[self.appearance_mode]
        self.configure(fg_color=c["bg"])
        self.theme_btn.configure(text="🌙" if self.appearance_mode == "Dark" else "☀️", text_color=c["icon_color"])
        self.run_btn.configure(fg_color=c["card"], border_color=c["border"], text_color=c["text"])
        self.log_output.configure(fg_color=c["card"], border_color=c["border"], text_color="#8b949e")

    def toggle_theme(self):
        self.appearance_mode = "Light" if self.appearance_mode == "Dark" else "Dark"
        ctk.set_appearance_mode(self.appearance_mode)
        self.apply_theme_styles()

    def browse(self, mode):
        if mode == 'file':
            f = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
            if f:
                self.input_entry.delete(0, "end")
                self.input_entry.insert(0, f)
        else:
            d = filedialog.askdirectory()
            if d:
                # 自动拼接文件名并强制转换为绝对路径，解决相对路径失效问题
                full_path = os.path.abspath(os.path.join(d, "数据清洗汇总结果.xlsx"))
                self.out_entry.delete(0, "end")
                self.out_entry.insert(0, full_path)

    def start_thread(self):
        self.save_config()
        threading.Thread(target=self.run_logic, daemon=True).start()

    def run_logic(self):
        self.run_btn.configure(state="disabled", text="Processing...")
        self.log_output.delete("1.0", "end")

        class Redirection:
            def __init__(self, widget): self.widget = widget

            def write(self, s):
                self.widget.insert("end", s)
                self.widget.see("end")

            def flush(self): pass

        old_stdout = sys.stdout
        sys.stdout = Redirection(self.log_output)

        try:
            start_analysis_flow(self.config)
        except Exception as e:
            print(f"\n[GUI ERROR]: {e}")
        finally:
            sys.stdout = old_stdout
            self.run_btn.configure(state="normal", text="Run Analysis")

    def save_config(self):
        # 在保存前，将 entry 中的路径尝试转换为绝对路径（如果用户手动输入了相对路径）
        in_file = self.input_entry.get().strip()
        out_file = self.out_entry.get().strip()
        
        if in_file and not os.path.isabs(in_file):
            in_file = os.path.abspath(in_file)
        if out_file and not os.path.isabs(out_file):
            out_file = os.path.abspath(out_file)

        self.config.update({
            'file_config': {'input_file': in_file, 'output_file': out_file},
            'analysis_period': {'start_date': self.start_date.get(), 'end_date': self.end_date.get()},
            'theme': self.appearance_mode
        })
        with open(self.config_path, 'w', encoding='utf-8') as f: yaml.dump(self.config, f)

    def on_closing(self):
        self.save_config()
        self.destroy()


if __name__ == "__main__":
    app = ModernAnalysisGUI()
    app.mainloop()