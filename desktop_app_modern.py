import csv
import json
import os
import re
import subprocess
import sys
import threading
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk
import openpyxl


APP_TITLE = "Microreview"
QUESTION_COUNT = 8
REQUIRED_COUNT = 3
OPTIONAL_MAX_SCORE = 25
DEFAULT_PATTERN = "{学号}_第{题号}题"
SUPPORTED_EXTENSIONS = {
    ".pdf", ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx",
    ".png", ".jpg", ".jpeg", ".txt", ".zip", ".rar"
}

QUESTION_KEYWORDS = {
    1: ["爱国主义教育实践", "爱国主义", "爱国"],
    2: ["志愿服务实践", "志愿服务", "志愿"],
    3: ["劳动与生活实践", "劳动实践", "劳动"],
    4: ["管理服务实践", "管理服务", "管理"],
    5: ["勤工助学实践", "勤工助学"],
    6: ["应急救护实践", "应急救护"],
    7: ["跨文化交流实践", "跨文化交流"],
    8: ["参与国家重大活动实践", "国家重大活动"],
}
QUESTION_TITLES = {
    1: "爱国主义教育实践",
    2: "志愿服务实践",
    3: "劳动与生活实践",
    4: "管理服务实践",
    5: "勤工助学实践",
    6: "应急救护实践",
    7: "跨文化交流实践",
    8: "参与国家重大活动实践",
}
COLORS = {
    "bg": "#F7F0E6",
    "card": "#FFF8EC",
    "card_soft": "#F3E8D8",
    "shadow": "#D8C9B4",
    "line": "#E2D3BE",
    "text": "#202020",
    "muted": "#746A5E",
    "primary": "#D87A45",
    "primary_hover": "#BF6635",
    "accent": "#1F1F1F",
    "accent_soft": "#EFE2D0",
    "warning_soft": "#FCE7D5",
    "warning_text": "#B85F2D",
    "danger": "#B94A3B",
}

FONT_BODY = ("FangSong_GB2312", 14)
FONT_SMALL = ("FangSong_GB2312", 12)
FONT_TITLE = ("FangSong_GB2312", 23, "bold")
FONT_SECTION = ("FangSong_GB2312", 17, "bold")
FONT_BUTTON = ("Comic Sans MS", 13, "bold")
FONT_LATIN = ("Comic Sans MS", 14)
FONT_LATIN_SMALL = ("Comic Sans MS", 12)
FONT_LATIN_SECTION = ("Comic Sans MS", 17, "bold")
FONT_SCORE = ("Comic Sans MS", 34, "bold")


def data_file() -> Path:
    base = os.environ.get("APPDATA")
    root = Path(base) / "StudentReviewApp" if base else Path.home() / ".student_review_app"
    root.mkdir(parents=True, exist_ok=True)
    return root / "review_data.json"


class SoftCard(ctk.CTkFrame):
    def __init__(self, master, radius=22, padding=0, **kwargs):
        super().__init__(master, fg_color=COLORS["bg"], corner_radius=0)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.shadow = ctk.CTkFrame(self, fg_color=COLORS["shadow"], corner_radius=radius)
        self.shadow.grid(row=0, column=0, sticky="nsew", padx=(3, 0), pady=(4, 0))
        self.card = ctk.CTkFrame(
            self,
            fg_color=kwargs.pop("fg_color", COLORS["card"]),
            corner_radius=radius,
            border_width=kwargs.pop("border_width", 1),
            border_color=kwargs.pop("border_color", COLORS["line"]),
        )
        self.card.grid(row=0, column=0, sticky="nsew", padx=(0, 3), pady=(0, 4))
        if padding:
            self.card.configure()


class ModernButton(ctk.CTkButton):
    def __init__(self, master, pulse_color=None, **kwargs):
        self.base_color = kwargs.get("fg_color", COLORS["primary"])
        self.pulse_color = pulse_color or COLORS["accent"]
        kwargs.setdefault("height", 38)
        super().__init__(
            master,
            corner_radius=18,
            font=FONT_BUTTON,
            hover=True,
            **kwargs,
        )

    def _clicked(self, event=None):
        try:
            original = self.cget("fg_color")
            self.configure(fg_color=self.pulse_color)
            self.after(110, lambda: self.configure(fg_color=original))
        finally:
            return super()._clicked(event)


class ReviewApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")
        self.title(APP_TITLE)
        self.geometry("1420x840")
        self.minsize(1180, 740)
        self.configure(fg_color=COLORS["bg"])

        self.students = []
        self.scores = {}
        self.current_index = 0
        self.material_folder = ""
        self.file_pattern = DEFAULT_PATTERN
        self.material_files = []
        self.material_index = {}
        self.material_report = self.empty_material_report()
        self.pending_review_ids = set()
        self.visible_student_indexes = []
        self.student_item_frames = {}
        self.student_item_summary_labels = {}
        self.material_indexing = False
        self.material_scan_token = 0
        self.expanded_grades = set()
        self.grade_state_initialized = False
        self.question_widgets = {}
        self.save_after_id = None
        self.search_after_id = None
        self.pending_roster_refresh_needed = False

        self.protocol("WM_DELETE_WINDOW", self.on_close)

        self.load_data()
        self.show_splash()

    def show_splash(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.splash_frame = ctk.CTkFrame(self, fg_color=COLORS["bg"], corner_radius=0)
        self.splash_frame.grid(row=0, column=0, sticky="nsew")

        shadow = ctk.CTkFrame(self.splash_frame, fg_color=COLORS["shadow"], corner_radius=36, width=760, height=450)
        shadow.place(relx=0.5, rely=0.5, anchor="center", x=8, y=10)
        card = ctk.CTkFrame(
            self.splash_frame,
            fg_color=COLORS["card"],
            corner_radius=36,
            border_width=1,
            border_color=COLORS["line"],
            width=760,
            height=450,
        )
        card.place(relx=0.5, rely=0.5, anchor="center")
        card.grid_propagate(False)
        card.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(card, text="MICROREVIEW", font=("Comic Sans MS", 34, "bold"), text_color=COLORS["text"]).grid(row=0, column=0, pady=(46, 8))
        ctk.CTkLabel(card, text="材料审阅工作台", font=FONT_SECTION, text_color=COLORS["muted"]).grid(row=1, column=0, pady=(0, 28))

        welcome = ctk.CTkFrame(card, fg_color="transparent")
        welcome.grid(row=2, column=0, pady=(0, 8))
        self.splash_welcome_cn = ctk.CTkLabel(welcome, text="", font=FONT_TITLE, text_color=COLORS["text"])
        self.splash_welcome_cn.grid(row=0, column=0)
        self.splash_welcome_en = ctk.CTkLabel(welcome, text="", font=("Comic Sans MS", 25, "bold"), text_color=COLORS["primary"])
        self.splash_welcome_en.grid(row=0, column=1, padx=(8, 0))

        self.splash_work = ctk.CTkLabel(card, text="", font=FONT_TITLE, text_color=COLORS["text"])
        self.splash_work.grid(row=3, column=0, pady=(12, 2))
        signature_box = ctk.CTkFrame(card, fg_color="transparent")
        signature_box.grid(row=4, column=0, padx=120, pady=(6, 18), sticky="e")
        self.splash_signature = ctk.CTkLabel(signature_box, text="", font=FONT_TITLE, text_color=COLORS["primary"], anchor="e")
        self.splash_signature.grid(row=0, column=0, sticky="e")
        self.splash_date = ctk.CTkLabel(signature_box, text="", font=("Comic Sans MS", 23, "bold"), text_color=COLORS["primary"], anchor="e")
        self.splash_date.grid(row=1, column=0, sticky="e", pady=(4, 0))

        self.splash_progress = ctk.CTkProgressBar(card, width=560, height=14, corner_radius=10, fg_color=COLORS["line"], progress_color=COLORS["primary"])
        self.splash_progress.grid(row=6, column=0, pady=(8, 8))
        self.splash_progress.set(0)
        ctk.CTkLabel(card, text="Loading...", font=FONT_LATIN_SMALL, text_color=COLORS["muted"]).grid(row=7, column=0)

        self.splash_items = [
            (self.splash_welcome_cn, "欢迎老师来到", 0.00, 0.28, 70),
            (self.splash_welcome_en, "MICROREVIEW!", 0.28, 0.50, 60),
            (self.splash_work, "工作愉快！", 0.50, 0.68, 85),
            (self.splash_signature, "----张乐岩", 0.68, 0.84, 75),
            (self.splash_date, "2026.7夏", 0.84, 1.00, 80),
        ]
        self.after(350, lambda: self.animate_splash_item(0, 0))

    def animate_splash_item(self, item_index, char_index):
        if item_index >= len(self.splash_items):
            self.splash_progress.set(1)
            self.after(3000, self.finish_splash)
            return

        label, text, start_progress, end_progress, delay = self.splash_items[item_index]
        if char_index <= len(text):
            label.configure(text=text[:char_index])
            span = end_progress - start_progress
            ratio = char_index / max(len(text), 1)
            self.splash_progress.set(start_progress + span * ratio)
            self.after(delay, lambda: self.animate_splash_item(item_index, char_index + 1))
        else:
            self.after(260, lambda: self.animate_splash_item(item_index + 1, 0))

    def finish_splash(self):
        self.splash_frame.destroy()
        self.grid_columnconfigure(0, weight=0, minsize=0)
        self.grid_rowconfigure(0, weight=0)
        self.build_ui()
        self.refresh_all()
    def build_ui(self):
        self.grid_columnconfigure(0, weight=0, minsize=252)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure(2, weight=0, minsize=286)
        self.grid_rowconfigure(1, weight=1)

        top = ctk.CTkFrame(self, fg_color=COLORS["card"], corner_radius=0, height=76, border_width=0)
        top.grid(row=0, column=0, columnspan=3, sticky="ew")
        top.grid_propagate(False)
        top.grid_columnconfigure(0, weight=1)
        title_box = ctk.CTkFrame(top, fg_color="transparent")
        title_box.grid(row=0, column=0, padx=(22, 12), pady=10, sticky="w")
        ctk.CTkLabel(title_box, text=APP_TITLE, font=FONT_LATIN_SECTION, text_color=COLORS["text"], anchor="w").grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(title_box, text="导入学生名单，选择材料文件夹，按学生逐题查看并评分。", font=FONT_SMALL, text_color=COLORS["muted"], anchor="w").grid(row=1, column=0, sticky="w")
        actions = ctk.CTkFrame(top, fg_color="transparent")
        actions.grid(row=0, column=1, padx=(8, 18), pady=12, sticky="e")
        self.toolbar_button(actions, "导入学生CSV", self.import_students, 0, secondary=True)
        self.toolbar_button(actions, "下载模板", self.export_template, 1, secondary=True)
        self.toolbar_button(actions, "读取记录", self.load_saved_records, 2, secondary=True)
        self.toolbar_button(actions, "保存记录", self.save_records, 3, secondary=True)
        self.toolbar_button(actions, "材料匹配报告", self.open_material_report, 4, secondary=True)
        self.toolbar_button(actions, "选择材料文件夹", self.choose_material_folder, 5, secondary=True)
        self.toolbar_button(actions, "导出成绩表", self.export_scores, 6)
        self.folder_label = ctk.CTkLabel(top, text="未选择材料文件夹", font=FONT_SMALL, text_color=COLORS["muted"], anchor="e")
        self.folder_label.grid(row=1, column=1, padx=(8, 22), pady=(0, 4), sticky="e")

        left = ctk.CTkFrame(self, fg_color=COLORS["card"], corner_radius=0, border_width=1, border_color=COLORS["line"], width=252)
        left.grid(row=1, column=0, sticky="nsew")
        left.grid_propagate(False)
        left.grid_rowconfigure(3, weight=1)
        left.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(left, text="学生列表", font=FONT_SECTION, text_color=COLORS["text"]).grid(row=0, column=0, padx=(18, 4), pady=(22, 8), sticky="w")
        self.student_count_label = ctk.CTkLabel(left, text="0 人", font=FONT_LATIN_SMALL, text_color=COLORS["muted"])
        self.student_count_label.grid(row=0, column=1, padx=(4, 18), pady=(22, 8), sticky="e")
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self.schedule_student_search())
        self.search_entry = ctk.CTkEntry(
            left, textvariable=self.search_var, placeholder_text="搜索姓名或学号",
            height=38, corner_radius=8, border_color=COLORS["line"], font=FONT_LATIN
        )
        self.search_entry.grid(row=1, column=0, columnspan=2, padx=18, pady=(8, 14), sticky="ew")
        self.pending_only_var = tk.BooleanVar(value=False)
        self.pending_only_switch = ctk.CTkSwitch(
            left,
            text="只看待批改学生",
            variable=self.pending_only_var,
            command=self.on_pending_filter_changed,
            font=FONT_BODY,
            progress_color=COLORS["primary"],
            button_color=COLORS["primary"],
            button_hover_color=COLORS["primary_hover"],
        )
        self.pending_only_switch.grid(row=2, column=0, columnspan=2, padx=18, pady=(0, 12), sticky="w")
        self.student_list_frame = ctk.CTkScrollableFrame(left, fg_color=COLORS["card"], corner_radius=0)
        self.student_list_frame.grid(row=3, column=0, columnspan=2, padx=10, pady=(0, 12), sticky="nsew")
        self.student_list_frame.grid_columnconfigure(0, weight=1)

        center = ctk.CTkFrame(self, fg_color=COLORS["bg"], corner_radius=0)
        center.grid(row=1, column=1, sticky="nsew")
        center.grid_columnconfigure(0, weight=1)
        center.grid_rowconfigure(2, weight=1)

        review_bar = ctk.CTkFrame(center, fg_color=COLORS["card"], corner_radius=0, height=92, border_width=0)
        review_bar.grid(row=0, column=0, sticky="ew")
        review_bar.grid_propagate(False)
        review_bar.grid_columnconfigure(0, weight=1)
        self.current_label = ctk.CTkLabel(review_bar, text="请先导入学生名单", font=FONT_SECTION, text_color=COLORS["text"], anchor="w")
        self.current_label.grid(row=0, column=0, padx=22, pady=(12, 2), sticky="ew")
        nav = ctk.CTkFrame(review_bar, fg_color="transparent")
        nav.grid(row=0, column=1, rowspan=2, padx=18, pady=14, sticky="e")
        self.nav_button(nav, "上一位", self.prev_student, 0, secondary=True)
        self.nav_button(nav, "下一位同学", self.next_student, 1)
        self.pattern_var = tk.StringVar(value="新规则：文件名包含学号 + 对应实践关键词即可自动匹配；旧命名规则仍兼容。")
        self.pattern_entry = ctk.CTkEntry(review_bar, textvariable=self.pattern_var, height=34, corner_radius=8, font=FONT_LATIN_SMALL, border_color=COLORS["line"], state="disabled")
        self.pattern_entry.grid(row=1, column=0, padx=22, pady=(2, 10), sticky="ew")
        self.pattern_entry.bind("<FocusOut>", lambda _e: self.update_pattern())
        self.pattern_entry.bind("<Return>", lambda _e: self.update_pattern())

        self.question_frame = ctk.CTkScrollableFrame(center, fg_color=COLORS["bg"], corner_radius=0)
        self.question_frame.grid(row=2, column=0, sticky="nsew", padx=18, pady=18)
        self.question_frame.grid_columnconfigure(0, weight=1)

        right = ctk.CTkFrame(self, fg_color=COLORS["card"], corner_radius=0, border_width=1, border_color=COLORS["line"], width=286)
        right.grid(row=1, column=2, sticky="nsew")
        right.grid_propagate(False)
        right.grid_columnconfigure(0, weight=1)
        right.grid_rowconfigure(6, weight=1)
        ctk.CTkLabel(right, text="成绩记录", font=FONT_SECTION, text_color=COLORS["text"]).grid(row=0, column=0, padx=18, pady=(22, 6), sticky="w")
        self.score_label = ctk.CTkLabel(right, text="0", font=FONT_SCORE, text_color=COLORS["primary"])
        self.score_label.grid(row=1, column=0, padx=18, sticky="w")
        self.completion_label = ctk.CTkLabel(right, text="必修 0/3 · 选修 0/5", font=FONT_LATIN, text_color=COLORS["muted"])
        self.completion_label.grid(row=2, column=0, padx=18, pady=(0, 10), sticky="w")
        ctk.CTkLabel(right, text="●  必修：第 1-3 题", font=FONT_LATIN_SMALL, text_color=COLORS["warning_text"]).grid(row=3, column=0, padx=18, pady=(2, 0), sticky="w")
        ctk.CTkLabel(right, text="●  选修：第 4-8 题，每题 25 分", font=FONT_LATIN_SMALL, text_color=COLORS["accent"]).grid(row=4, column=0, padx=18, pady=(2, 10), sticky="w")
        ModernButton(right, text="清空本地数据", command=self.clear_data, fg_color="#FFF3F2", hover_color="#FFE3E0", text_color=COLORS["danger"], width=246, height=36).grid(row=5, column=0, padx=18, pady=(0, 12), sticky="ew")
        self.summary_frame = ctk.CTkScrollableFrame(right, fg_color=COLORS["card"], corner_radius=0)
        self.summary_frame.grid(row=6, column=0, padx=10, pady=(0, 12), sticky="nsew")
        self.summary_frame.grid_columnconfigure(0, weight=1)
    def toolbar_button(self, master, text, command, column, secondary=False, danger=False):
        if danger:
            fg, hover, text_color = "#FFF3F2", "#FFE3E0", COLORS["danger"]
        elif secondary:
            fg, hover, text_color = "#F7FAFE", "#EAF2FF", COLORS["text"]
        else:
            fg, hover, text_color = COLORS["primary"], COLORS["primary_hover"], "#FFFFFF"
        ModernButton(master, text=text, command=command, fg_color=fg, hover_color=hover, text_color=text_color, width=96, height=36).grid(row=0, column=column, padx=5, pady=0)

    def nav_button(self, master, text, command, column, secondary=False):
        fg = "#F7FAFE" if secondary else COLORS["primary"]
        hover = "#EAF2FF" if secondary else COLORS["primary_hover"]
        color = COLORS["text"] if secondary else "#FFFFFF"
        ModernButton(master, text=text, command=command, fg_color=fg, hover_color=hover, text_color=color, width=106, height=36).grid(row=0, column=column, padx=4, pady=0)

    def load_data(self):
        path = data_file()
        if not path.exists():
            return
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            self.students = data.get("students", [])
            self.scores = data.get("scores", {})
            self.current_index = int(data.get("current_index", 0))
            self.material_folder = data.get("material_folder", "")
            self.file_pattern = data.get("file_pattern", DEFAULT_PATTERN)
        except Exception:
            messagebox.showwarning(APP_TITLE, "本地保存数据读取失败，将使用空数据启动。")

    def save_data(self):
        payload = {
            "students": self.students,
            "scores": self.scores,
            "current_index": self.current_index,
            "material_folder": self.material_folder,
            "file_pattern": self.file_pattern,
        }
        data_file().write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        self.save_after_id = None

    def schedule_save_data(self, delay=700):
        if self.save_after_id:
            self.after_cancel(self.save_after_id)
        self.save_after_id = self.after(delay, self.save_data)

    def on_close(self):
        if self.save_after_id:
            self.after_cancel(self.save_after_id)
            self.save_data()
        self.destroy()

    def save_records(self):
        self.sync_material_submissions()
        self.save_data()
        self.refresh_all()
        messagebox.showinfo(APP_TITLE, "当前学生名单、材料路径和批改记录已保存。")

    def load_saved_records(self):
        path = filedialog.askopenfilename(
            title="\u8bfb\u53d6\u5386\u53f2\u6210\u7ee9\u8868",
            filetypes=[("\u6210\u7ee9\u8868", "*.csv *.xlsx"), ("CSV \u8868\u683c", "*.csv"), ("Excel \u8868\u683c", "*.xlsx"), ("\u6240\u6709\u6587\u4ef6", "*.*")],
        )
        if not path:
            return
        try:
            rows = self.read_score_table(path)
            students, scores = self.parse_score_records(rows)
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"\u8bfb\u53d6\u5931\u8d25\uff1a{exc}")
            return
        if not students:
            messagebox.showwarning(APP_TITLE, "\u8868\u683c\u4e2d\u6ca1\u6709\u8bc6\u522b\u5230\u6709\u6548\u5b66\u751f\u8bb0\u5f55\u3002")
            return
        self.students = students
        self.scores = scores
        self.current_index = 0
        self.material_folder = ""
        self.material_files = []
        self.material_index = {}
        self.material_report = self.empty_material_report()
        self.pending_review_ids = set()
        self.expanded_grades = set()
        self.grade_state_initialized = False
        self.save_data()
        self.refresh_all()
        messagebox.showinfo(APP_TITLE, f"\u5df2\u4ece\u6210\u7ee9\u8868\u8bfb\u53d6 {len(students)} \u540d\u5b66\u751f\u7684\u6279\u6539\u8bb0\u5f55\u3002\n\u73b0\u5728\u53ef\u4ee5\u91cd\u65b0\u9009\u62e9\u6750\u6599\u6587\u4ef6\u5939\u7ee7\u7eed\u6279\u6539\u3002")

    def ensure_record(self, student_id):
        if student_id not in self.scores:
            self.scores[student_id] = {
                "questions": [{"submitted": False, "score": "", "note": ""} for _ in range(QUESTION_COUNT)]
            }
        return self.scores[student_id]

    def current_student(self):
        if not self.students:
            return None
        self.current_index = max(0, min(self.current_index, len(self.students) - 1))
        return self.students[self.current_index]

    def import_students(self):
        path = filedialog.askopenfilename(title="\u9009\u62e9\u5b66\u751f\u540d\u5355 CSV", filetypes=[("CSV \u6587\u4ef6", "*.csv"), ("\u6240\u6709\u6587\u4ef6", "*.*")])
        if not path:
            return
        try:
            rows = self.read_csv(path)
            students = self.parse_students(rows)
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"\u5bfc\u5165\u5931\u8d25\uff1a{exc}")
            return
        if not students:
            messagebox.showwarning(APP_TITLE, "\u6ca1\u6709\u8bc6\u522b\u5230\u6709\u6548\u5b66\u751f\u3002")
            return
        before_count = len(self.students)
        added_count, updated_count = self.merge_students(students)
        if before_count == 0 and self.students:
            self.current_index = 0
        else:
            self.current_index = max(0, min(self.current_index, len(self.students) - 1))
        self.expanded_grades = set()
        self.grade_state_initialized = True
        self.save_data()
        self.refresh_all()
        messagebox.showinfo(APP_TITLE, f"\u5df2\u5408\u5e76\u5bfc\u5165 {len(students)} \u6761\u540d\u5355\u8bb0\u5f55\u3002\n\u65b0\u589e {added_count} \u4eba\uff0c\u66f4\u65b0 {updated_count} \u4eba\uff1b\u5f53\u524d\u5171 {len(self.students)} \u4eba\u3002")

    def merge_students(self, incoming_students):
        existing_by_id = {student["id"]: student for student in self.students}
        added_count = 0
        updated_count = 0
        for student in incoming_students:
            sid = student["id"]
            if sid in existing_by_id:
                existing_by_id[sid]["name"] = student["name"]
                updated_count += 1
            else:
                self.students.append(student)
                existing_by_id[sid] = student
                added_count += 1
            self.ensure_record(sid)
        return added_count, updated_count

    def read_csv(self, path):
        for encoding in ("utf-8-sig", "gbk", "utf-16"):
            try:
                with open(path, newline="", encoding=encoding) as file:
                    return list(csv.reader(file))
            except UnicodeError:
                continue
        raise ValueError("无法读取 CSV 文件。")

    def cell_text(self, value):
        if value is None:
            return ""
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return str(value).strip()

    def read_score_table(self, path):
        suffix = Path(path).suffix.lower()
        if suffix == ".xlsx":
            workbook = openpyxl.load_workbook(path, data_only=True, read_only=True)
            sheet = workbook.active
            return [[self.cell_text(cell) for cell in row] for row in sheet.iter_rows(values_only=True)]
        return self.read_csv(path)

    def parse_score_records(self, rows):
        if len(rows) < 2:
            return [], {}
        headers = [self.cell_text(cell).lstrip("\\ufeff") for cell in rows[0]]
        lookup = {header: index for index, header in enumerate(headers)}
        student_id_key = "\u5b66\u53f7"
        name_key = "\u59d3\u540d"
        if student_id_key not in lookup or name_key not in lookup:
            raise ValueError("\u6210\u7ee9\u8868\u8868\u5934\u5fc5\u987b\u5305\u542b\u201c\u5b66\u53f7\u201d\u548c\u201c\u59d3\u540d\u201d\u3002")

        students = []
        scores = {}
        seen = set()
        for row in rows[1:]:
            sid = self.cell_text(row[lookup[student_id_key]]) if len(row) > lookup[student_id_key] else ""
            name = self.cell_text(row[lookup[name_key]]) if len(row) > lookup[name_key] else ""
            if not sid or not name or sid in seen:
                continue
            questions = []
            for number in range(1, QUESTION_COUNT + 1):
                done_key = f"\u7b2c{number}\u9898\u5b8c\u6210"
                score_key = f"\u7b2c{number}\u9898\u6210\u7ee9"
                note_key = f"\u7b2c{number}\u9898\u5907\u6ce8"
                done = self.cell_text(row[lookup[done_key]]) if done_key in lookup and len(row) > lookup[done_key] else ""
                score = self.cell_text(row[lookup[score_key]]) if score_key in lookup and len(row) > lookup[score_key] else ""
                note = self.cell_text(row[lookup[note_key]]) if note_key in lookup and len(row) > lookup[note_key] else ""
                submitted = done in {"\u662f", "1", "true", "True", "TRUE", "\u5df2\u5b8c\u6210", "\u5df2\u63d0\u4ea4"} or bool(score or note)
                questions.append({"submitted": submitted, "score": score, "note": note})
            students.append({"name": name, "id": sid})
            scores[sid] = {"questions": questions}
            seen.add(sid)
        return students, scores

    def parse_students(self, rows):
        if len(rows) < 2:
            return []
        headers = [cell.strip().lstrip("\ufeff").lower() for cell in rows[0]]
        name_idx = next((i for i, h in enumerate(headers) if h in {"姓名", "name", "学生姓名"}), -1)
        id_idx = next((i for i, h in enumerate(headers) if h in {"学号", "studentid", "student_id", "id", "编号"}), -1)
        if name_idx == -1 or id_idx == -1:
            raise ValueError("CSV 表头必须包含“姓名”和“学号”两列。")
        result, seen = [], set()
        for row in rows[1:]:
            name = row[name_idx].strip() if len(row) > name_idx else ""
            sid = row[id_idx].strip() if len(row) > id_idx else ""
            if name and sid and sid not in seen:
                result.append({"name": name, "id": sid})
                seen.add(sid)
        return result

    def export_template(self):
        path = filedialog.asksaveasfilename(title="保存名单模板", defaultextension=".csv", initialfile="学生名单模板.csv", filetypes=[("CSV 文件", "*.csv")])
        if not path:
            return
        with open(path, "w", newline="", encoding="utf-8-sig") as file:
            writer = csv.writer(file)
            writer.writerow(["姓名", "学号"])
            writer.writerow(["张三", "20260001"])
            writer.writerow(["李四", "20260002"])
        messagebox.showinfo(APP_TITLE, "名单模板已保存。")

    def choose_material_folder(self):
        folder = filedialog.askdirectory(title="选择材料文件夹")
        if folder:
            self.material_folder = folder
            self.refresh_all(reindex_materials=True)
            self.save_data()

    def sync_material_submissions(self):
        if not self.students or not self.material_index:
            return
        for student in self.students:
            record = self.ensure_record(student["id"])
            matched_questions = self.material_index.get(student["id"], {})
            for index, question in enumerate(record["questions"]):
                if index + 1 in matched_questions:
                    question["submitted"] = True

    def empty_material_report(self):
        return {
            "folder": "",
            "supported_files": 0,
            "matched_files": 0,
            "matched_items": 0,
            "id_without_question": [],
            "question_without_id": [],
            "unrecognized": [],
            "duplicates": [],
        }

    def index_material_files(self):
        files, material_index, report = self.build_material_index(self.material_folder, self.students)
        self.material_files = files
        self.material_index = material_index
        self.material_report = report

    def build_material_index(self, folder, students):
        material_files = []
        material_index = {}
        report = self.empty_material_report()
        report["folder"] = folder or ""
        if not folder or not Path(folder).exists():
            return material_files, material_index, report
        student_ids = {self.normalize_filename(student["id"]): student["id"] for student in students}
        non_numeric_ids = {sid: original for sid, original in student_ids.items() if not sid.isdigit()}
        for root, _dirs, files in os.walk(folder):
            for name in files:
                path = Path(root) / name
                if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
                    continue
                material_files.append(path)
                report["supported_files"] += 1
                normalized_name = self.normalize_filename(path.name)
                candidate_ids = []
                for token in re.findall(r"\d{4,}", normalized_name):
                    if token in student_ids:
                        candidate_ids.append(student_ids[token])
                if non_numeric_ids:
                    for sid, original in non_numeric_ids.items():
                        if sid and sid in normalized_name:
                            candidate_ids.append(original)
                question_numbers = [
                    number for number in range(1, QUESTION_COUNT + 1)
                    if self.match_question_keywords(path.name, number)
                ]
                if not candidate_ids and not question_numbers:
                    report["unrecognized"].append(str(path))
                    continue
                if candidate_ids and not question_numbers:
                    report["id_without_question"].append(str(path))
                    continue
                if question_numbers and not candidate_ids:
                    report["question_without_id"].append(str(path))
                    continue
                matched_this_file = False
                for original_id in set(candidate_ids):
                    for number in question_numbers:
                        existing = material_index.setdefault(original_id, {}).get(number)
                        if existing:
                            report["duplicates"].append(f"{original_id} 第{number}题：{existing.name} / {path.name}")
                            continue
                        material_index.setdefault(original_id, {})[number] = path
                        report["matched_items"] += 1
                        matched_this_file = True
                if matched_this_file:
                    report["matched_files"] += 1
        return material_files, material_index, report

    def start_material_indexing(self, folder):
        self.material_scan_token += 1
        token = self.material_scan_token
        self.material_indexing = True
        self.material_files = []
        self.material_index = {}
        self.material_report = self.empty_material_report()
        self.rebuild_pending_review_cache()
        self.folder_label.configure(text="\u6b63\u5728\u540e\u53f0\u5339\u914d\u6750\u6599\uff0c\u8bf7\u7a0d\u5019...")
        students_snapshot = [dict(student) for student in self.students]

        def worker():
            try:
                files, index, report = self.build_material_index(folder, students_snapshot)
                self.after(0, lambda: self.finish_material_indexing(token, folder, files, index, report, None))
            except Exception as exc:
                self.after(0, lambda e=exc: self.finish_material_indexing(token, folder, [], {}, self.empty_material_report(), e))

        threading.Thread(target=worker, daemon=True).start()

    def finish_material_indexing(self, token, folder, files, material_index, report, error):
        if token != self.material_scan_token:
            return
        self.material_indexing = False
        if error:
            messagebox.showerror(APP_TITLE, f"\u6750\u6599\u5339\u914d\u5931\u8d25\uff1a{error}")
            self.folder_label.configure(text="\u6750\u6599\u5339\u914d\u5931\u8d25")
            return
        self.material_folder = folder
        self.material_files = files
        self.material_index = material_index
        self.material_report = report
        self.sync_material_submissions()
        self.rebuild_pending_review_cache()
        self.save_data()
        self.refresh_student_list()
        self.refresh_current()
        self.refresh_questions()
        self.refresh_summary()
        self.folder_label.configure(text=f"{folder}\uff08{len(files)} \u4e2a\u6587\u4ef6\uff0c\u540e\u53f0\u5339\u914d\u5b8c\u6210\uff09")

    def refresh_all(self, reindex_materials=False):
        if self.material_folder and reindex_materials:
            self.start_material_indexing(self.material_folder)
        self.rebuild_pending_review_cache()
        self.refresh_student_list()
        self.refresh_current()
        self.refresh_questions()
        self.refresh_summary()
        text = "\u672a\u9009\u62e9\u6750\u6599\u6587\u4ef6\u5939"
        if self.material_folder:
            suffix = "\u540e\u53f0\u5339\u914d\u4e2d..." if self.material_indexing else f"{len(self.material_files)} \u4e2a\u6587\u4ef6"
            text = f"{self.material_folder}\uff08{suffix}\uff09"
        self.folder_label.configure(text=text)

    def grade_key(self, student):
        sid = str(student.get("id", "")).strip()
        prefix = sid[:2]
        if len(prefix) == 2 and prefix.isdigit():
            return f"20{prefix}级"
        return "其他"

    def student_filter_state(self):
        keyword = self.search_var.get().strip().lower() if hasattr(self, "search_var") else ""
        pending_var = getattr(self, "pending_only_var", None)
        pending_only = bool(pending_var.get()) if pending_var else False
        return keyword, pending_only

    def grouped_student_indexes(self, keyword=None, pending_only=None):
        if keyword is None or pending_only is None:
            keyword, pending_only = self.student_filter_state()
        grouped = {}
        visible_count = 0
        for index, student in enumerate(self.students):
            name = student["name"].lower()
            sid = student["id"].lower()
            if keyword and keyword not in name and keyword not in sid:
                continue
            if pending_only and not self.is_pending_review(student):
                continue
            grade = self.grade_key(student)
            grouped.setdefault(grade, []).append((index, student))
            visible_count += 1
        return grouped, visible_count

    def ordered_student_indexes(self, keyword=None, pending_only=None):
        grouped, _visible_count = self.grouped_student_indexes(keyword, pending_only)
        ordered = []
        for grade in sorted(grouped.keys(), reverse=True):
            ordered.extend(index for index, _student in grouped[grade])
        return ordered

    def next_index_in_order(self, order, current_index, wrap=False):
        if not order:
            return None
        if current_index not in order:
            return order[0]
        position = order.index(current_index)
        if position + 1 < len(order):
            return order[position + 1]
        return order[0] if wrap else None

    def previous_index_in_order(self, order, current_index, wrap=False):
        if not order:
            return None
        if current_index not in order:
            return order[-1]
        position = order.index(current_index)
        if position > 0:
            return order[position - 1]
        return order[-1] if wrap else None

    def refresh_student_list(self):
        self.pending_roster_refresh_needed = False
        for widget in self.student_list_frame.winfo_children():
            widget.destroy()
        self.visible_student_indexes = []
        self.student_item_frames = {}
        self.student_item_summary_labels = {}
        keyword, pending_only = self.student_filter_state()
        grouped, visible_count = self.grouped_student_indexes(keyword, pending_only)
        if pending_only:
            self.student_count_label.configure(text=f"{visible_count}/{len(self.students)} \u4eba")
        else:
            self.student_count_label.configure(text=f"{len(self.students)} \u4eba")

        if keyword:
            visible_grades = set(grouped.keys())
        else:
            visible_grades = self.expanded_grades

        row = 0
        for grade in sorted(grouped.keys(), reverse=True):
            students = grouped[grade]
            expanded = grade in visible_grades
            arrow = "\u25be" if expanded else "\u25b8"
            header = ctk.CTkButton(
                self.student_list_frame,
                text=f"{arrow}  {grade}    {len(students)}\u4eba",
                command=lambda g=grade: self.toggle_grade(g),
                anchor="w",
                height=40,
                corner_radius=14,
                fg_color="#F3F7FD",
                hover_color="#EAF2FF",
                text_color=COLORS["text"],
                font=FONT_LATIN_SECTION,
                border_width=1,
                border_color=COLORS["line"],
            )
            header.grid(row=row, column=0, sticky="ew", pady=(4, 3), padx=2)
            row += 1

            if not expanded:
                continue

            for index, student in students:
                self.visible_student_indexes.append(index)
                active = index == self.current_index
                item = ctk.CTkButton(
                    self.student_list_frame,
                    text=f"{student['name']}  {student['id']}",
                    command=lambda idx=index: self.select_student(idx),
                    anchor="w",
                    height=34,
                    corner_radius=10,
                    fg_color="#EAF2FF" if active else COLORS["card"],
                    hover_color="#F1F7FF",
                    text_color=COLORS["text"],
                    font=FONT_LATIN,
                    border_width=1,
                    border_color=COLORS["primary"] if active else "#EEF2F7",
                )
                item.grid(row=row, column=0, sticky="ew", pady=2, padx=(12, 2))
                self.student_item_frames[index] = item
                row += 1

    def on_pending_filter_changed(self):
        self.pending_roster_refresh_needed = False
        self.rebuild_pending_review_cache()
        if self.pending_only_var.get():
            order = self.ordered_student_indexes()
            pending_index = order[0] if order else None
            if pending_index is not None:
                self.expanded_grades.add(self.grade_key(self.students[pending_index]))
                self.select_student(pending_index)
                return
        self.refresh_student_list()

    def schedule_student_search(self):
        if self.search_after_id:
            self.after_cancel(self.search_after_id)
        self.search_after_id = self.after(220, self.refresh_student_list)

    def update_student_row_style(self, index, active):
        frame = self.student_item_frames.get(index)
        if not frame:
            return
        frame.configure(fg_color="#EAF2FF" if active else COLORS["card"], border_color=COLORS["primary"] if active else "#EEF2F7")

    def update_student_row_summary(self, index):
        # The student list intentionally shows only name and ID for speed.
        return

    def toggle_grade(self, grade):
        if grade in self.expanded_grades:
            self.expanded_grades.remove(grade)
        else:
            self.expanded_grades.add(grade)
        self.refresh_student_list()

    def hover_card(self, frame, entering, active):
        if active:
            return
        frame.configure(fg_color="#F1F7FF" if entering else COLORS["card_soft"])

    def select_student(self, index):
        previous_index = self.current_index
        self.current_index = index
        self.schedule_save_data()
        self.update_student_row_style(previous_index, False)
        self.update_student_row_style(index, True)
        self.animate_current_card()
        self.refresh_current()
        self.refresh_questions()
        self.refresh_summary()

    def animate_current_card(self):
        self.current_label.configure(text_color=COLORS["primary"])
        self.after(150, lambda: self.current_label.configure(text_color=COLORS["text"]))

    def refresh_current(self):
        student = self.current_student()
        if not student:
            self.current_label.configure(text="请先导入学生名单")
        else:
            self.current_label.configure(text=f"{student['name']}    学号：{student['id']}")

    def refresh_questions(self):
        for widget in self.question_frame.winfo_children():
            widget.destroy()
        self.question_widgets = {}
        student = self.current_student()
        if not student:
            empty = ctk.CTkFrame(self.question_frame, fg_color="transparent")
            empty.grid(row=0, column=0, pady=(230, 0), sticky="n")
            ctk.CTkLabel(empty, text="先导入学生名单", font=FONT_TITLE, text_color=COLORS["text"]).grid(row=0, column=0, pady=(0, 10))
            ctk.CTkLabel(empty, text="CSV 至少包含两列：姓名、学号。导入后选择材料文件夹即可开始批阅。", font=FONT_BODY, text_color=COLORS["muted"]).grid(row=1, column=0)
            return
        record = self.ensure_record(student["id"])
        visible_questions = self.visible_questions_for_student(student, record)
        if not visible_questions:
            empty = ctk.CTkFrame(self.question_frame, fg_color="transparent")
            empty.grid(row=0, column=0, pady=(210, 0), sticky="n")
            ctk.CTkLabel(empty, text="暂未匹配到本次提交材料", font=FONT_TITLE, text_color=COLORS["text"]).grid(row=0, column=0, pady=(0, 10))
            ctk.CTkLabel(empty, text="选择材料文件夹后，系统会按“学号 + 实践关键词”自动显示可批改题目；已有历史成绩也会保留显示。", font=FONT_BODY, text_color=COLORS["muted"]).grid(row=1, column=0)
            return
        for row, i in enumerate(visible_questions):
            self.add_question_card(row, i, record["questions"][i])

    def add_question_card(self, row_index, index, question):
        number = index + 1
        matched = self.find_file(number)
        required = number <= REQUIRED_COUNT
        frame = SoftCard(self.question_frame, radius=24)
        frame.grid(row=row_index, column=0, sticky="ew", pady=8, padx=2)
        frame.card.grid_columnconfigure(3, weight=1)
        tag_bg = COLORS["warning_soft"] if required else COLORS["accent_soft"]
        tag_color = COLORS["warning_text"] if required else COLORS["accent"]
        ctk.CTkLabel(frame.card, text=f"{number}. {QUESTION_TITLES[number]}", font=FONT_SECTION, text_color=COLORS["text"]).grid(row=0, column=0, padx=(16, 8), pady=(14, 4), sticky="w")
        ctk.CTkLabel(frame.card, text="必修" if required else "选修 · 25分", font=FONT_LATIN_SMALL, text_color=tag_color, fg_color=tag_bg, corner_radius=12, padx=10, pady=3).grid(row=0, column=1, pady=(14, 4), sticky="w")
        file_text = f"已匹配：{matched.name}" if matched else "未匹配材料"
        ctk.CTkLabel(frame.card, text=file_text, font=FONT_LATIN_SMALL, text_color=COLORS["muted"], anchor="w").grid(row=1, column=0, columnspan=4, padx=16, pady=(0, 8), sticky="ew")
        ModernButton(frame.card, text="查看材料", command=lambda n=number: self.open_material(n), fg_color="#F7FAFE", hover_color="#EAF2FF", text_color=COLORS["text"], width=96).grid(row=0, column=4, padx=16, pady=(14, 4))

        submitted = tk.BooleanVar(value=bool(question.get("submitted")))
        score = tk.StringVar(value=str(question.get("score", "")))
        note = tk.StringVar(value=question.get("note", ""))
        self.question_widgets[index] = (submitted, score, note)
        check = ctk.CTkCheckBox(frame.card, text="已提交/完成", variable=submitted, command=lambda i=index: self.update_question(i), font=FONT_BODY, corner_radius=8)
        check.grid(row=2, column=0, padx=16, pady=(4, 14), sticky="w")
        ctk.CTkLabel(frame.card, text="得分", font=FONT_BODY, text_color=COLORS["text"]).grid(row=2, column=1, padx=(4, 6), pady=(4, 14))
        score_entry = ctk.CTkEntry(frame.card, textvariable=score, width=88, height=36, corner_radius=16, font=FONT_LATIN, border_color=COLORS["line"])
        score_entry.grid(row=2, column=2, padx=(0, 12), pady=(4, 14), sticky="w")
        score_entry.bind("<FocusOut>", lambda _e, i=index: self.update_question(i))
        score_entry.bind("<Return>", lambda _e, i=index: self.update_question(i))
        note_entry = ctk.CTkEntry(frame.card, textvariable=note, height=36, corner_radius=16, font=FONT_LATIN, placeholder_text="备注，可留空", border_color=COLORS["line"])
        note_entry.grid(row=2, column=3, columnspan=2, padx=(0, 16), pady=(4, 14), sticky="ew")
        note_entry.bind("<FocusOut>", lambda _e, i=index: self.update_question(i))
        note_entry.bind("<Return>", lambda _e, i=index: self.update_question(i))

    def refresh_summary(self):
        for widget in self.summary_frame.winfo_children():
            widget.destroy()
        student = self.current_student()
        if not student:
            self.score_label.configure(text="0")
            self.completion_label.configure(text="必修 0/3 · 选修 0/5")
            ctk.CTkLabel(self.summary_frame, text="暂无学生。", font=FONT_BODY, text_color=COLORS["muted"]).grid(row=0, column=0, padx=12, pady=12)
            return
        record = self.ensure_record(student["id"])
        required, optional = self.completion(student["id"])
        total = self.total_score(student["id"])
        self.score_label.configure(text=f"{total:g}")
        self.completion_label.configure(text=f"必修 {required}/3 · 选修 {optional}/5")
        for i, q in enumerate(record["questions"]):
            n = i + 1
            row = ctk.CTkFrame(self.summary_frame, fg_color=COLORS["card"], corner_radius=16)
            row.grid(row=i, column=0, sticky="ew", padx=4, pady=5)
            row.grid_columnconfigure(1, weight=1)
            label = "必修" if n <= REQUIRED_COUNT else "选修"
            done = "是" if q.get("submitted") else "否"
            ctk.CTkLabel(row, text=f"第 {n} 题", font=FONT_LATIN, text_color=COLORS["text"]).grid(row=0, column=0, padx=10, pady=8)
            ctk.CTkLabel(row, text=f"{label} · 完成 {done}", font=FONT_LATIN_SMALL, text_color=COLORS["muted"]).grid(row=0, column=1, padx=6, pady=8, sticky="w")
            ctk.CTkLabel(row, text=str(q.get("score", "")), font=FONT_LATIN, text_color=COLORS["primary"]).grid(row=0, column=2, padx=10, pady=8)

    def update_question(self, index):
        student = self.current_student()
        if not student or index not in self.question_widgets:
            return
        was_pending = self.is_pending_review(student)
        submitted, score_var, note = self.question_widgets[index]
        score = score_var.get().strip()
        if score:
            try:
                value = max(0, float(score))
                if index + 1 > REQUIRED_COUNT:
                    value = min(OPTIONAL_MAX_SCORE, value)
                score = f"{value:g}"
                score_var.set(score)
            except ValueError:
                messagebox.showwarning(APP_TITLE, "得分必须是数字。")
                score = ""
                score_var.set("")
        record = self.ensure_record(student["id"])
        record["questions"][index] = {"submitted": bool(submitted.get()), "score": score, "note": note.get().strip()}
        self.set_pending_review_status(student)
        is_pending = self.is_pending_review(student)
        self.schedule_save_data()
        self.update_student_row_summary(self.current_index)
        self.refresh_summary()
        if getattr(self, "pending_only_var", None) and self.pending_only_var.get() and was_pending != is_pending:
            self.pending_roster_refresh_needed = True

    def prev_student(self):
        if self.students:
            order = list(self.visible_student_indexes) if self.visible_student_indexes else self.ordered_student_indexes()
            previous_index = self.previous_index_in_order(order, self.current_index, wrap=False)
            if previous_index is None:
                return
            self.expanded_grades.add(self.grade_key(self.students[previous_index]))
            self.select_student(previous_index)
            self.refresh_student_list()

    def next_student(self):
        if self.students:
            order = list(self.visible_student_indexes) if self.visible_student_indexes else self.ordered_student_indexes()
            next_index = self.next_index_in_order(order, self.current_index, wrap=False)
            if next_index is None:
                if self.pending_roster_refresh_needed:
                    self.refresh_student_list()
                message = "当前已经是最后一位待批改学生。" if getattr(self, "pending_only_var", None) and self.pending_only_var.get() else "当前已经是最后一位学生。"
                messagebox.showinfo(APP_TITLE, message)
                return
            self.expanded_grades.add(self.grade_key(self.students[next_index]))
            self.select_student(next_index)
            self.refresh_student_list()

    def update_pattern(self):
        self.file_pattern = DEFAULT_PATTERN
        self.save_data()
        self.refresh_questions()

    def build_keyword(self, question_number, student=None):
        student = student or self.current_student()
        if not student:
            return ""
        return (
            self.file_pattern
            .replace("{姓名}", student["name"])
            .replace("{学号}", student["id"])
            .replace("{题号}", str(question_number))
        ).lower()

    def normalize_filename(self, value):
        return str(value).lower().replace(" ", "").replace("　", "").replace("_", "").replace("-", "").replace("－", "")

    def match_question_keywords(self, filename, question_number):
        normalized = self.normalize_filename(filename)
        return any(self.normalize_filename(keyword) in normalized for keyword in QUESTION_KEYWORDS.get(question_number, []))

    def has_question_history(self, question):
        return bool(question.get("submitted") or str(question.get("score", "")).strip() or str(question.get("note", "")).strip())

    def calculate_pending_review(self, student):
        record = self.ensure_record(student["id"])
        matched_numbers = set(self.material_index.get(student["id"], {}).keys())
        for index, question in enumerate(record["questions"]):
            number = index + 1
            visible = number in matched_numbers or self.has_question_history(question)
            if visible and not str(question.get("score", "")).strip():
                return True
        return False

    def rebuild_pending_review_cache(self):
        self.pending_review_ids = {
            student["id"]
            for student in self.students
            if self.calculate_pending_review(student)
        }

    def set_pending_review_status(self, student):
        sid = student["id"]
        if self.calculate_pending_review(student):
            self.pending_review_ids.add(sid)
        else:
            self.pending_review_ids.discard(sid)

    def is_pending_review(self, student):
        return student["id"] in self.pending_review_ids

    def visible_questions_for_student(self, student, record):
        visible = []
        for index, question in enumerate(record["questions"]):
            number = index + 1
            if self.find_file_for_student(student, number) or self.has_question_history(question):
                visible.append(index)
        return visible

    def find_file_for_student(self, student, question_number):
        if not student:
            return None
        indexed = self.material_index.get(student["id"], {}).get(question_number)
        if indexed:
            return indexed

        sid = self.normalize_filename(student["id"])
        for path in self.material_files:
            normalized_name = self.normalize_filename(path.name)
            if sid in normalized_name and self.match_question_keywords(path.name, question_number):
                self.material_index.setdefault(student["id"], {})[question_number] = path
                return path

        keyword = self.build_keyword(question_number, student)
        matched = next((path for path in self.material_files if keyword and keyword in path.name.lower()), None)
        if matched:
            self.material_index.setdefault(student["id"], {})[question_number] = matched
        return matched

    def find_file(self, question_number):
        return self.find_file_for_student(self.current_student(), question_number)

    def open_material(self, question_number):
        if not self.material_folder:
            messagebox.showwarning(APP_TITLE, "请先选择材料文件夹。")
            return
        matched = self.find_file(question_number)
        if not matched:
            student_id = self.current_student()["id"]
            messagebox.showwarning(APP_TITLE, f"未找到匹配文件：学号 {student_id} + 第{question_number}题关键词\n请确认文件名包含学生学号，以及该题对应的实践关键词。")
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(matched)
            elif sys.platform == "darwin":
                subprocess.run(["open", str(matched)], check=False)
            else:
                subprocess.run(["xdg-open", str(matched)], check=False)
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"打开材料失败：{exc}")

    def total_score(self, student_id):
        total = 0.0
        for q in self.ensure_record(student_id)["questions"]:
            try:
                total += float(q.get("score") or 0)
            except ValueError:
                pass
        return total

    def completion(self, student_id):
        record = self.ensure_record(student_id)
        required = sum(1 for q in record["questions"][:REQUIRED_COUNT] if q.get("submitted"))
        optional = sum(1 for q in record["questions"][REQUIRED_COUNT:] if q.get("submitted"))
        return required, optional

    def material_report_text(self):
        report = self.material_report or self.empty_material_report()
        lines = [
            "材料匹配报告",
            "",
            f"材料文件夹：{report.get('folder') or self.material_folder or '未选择'}",
            f"支持格式文件数：{report.get('supported_files', 0)}",
            f"成功匹配文件数：{report.get('matched_files', 0)}",
            f"成功匹配题目数：{report.get('matched_items', 0)}",
            f"识别到学号但未识别题目：{len(report.get('id_without_question', []))}",
            f"识别到题目但未匹配学生：{len(report.get('question_without_id', []))}",
            f"完全无法识别：{len(report.get('unrecognized', []))}",
            f"重复匹配：{len(report.get('duplicates', []))}",
            "",
        ]

        def append_section(title, items):
            lines.append(title)
            if not items:
                lines.append("  无")
            else:
                limit = 300
                for item in items[:limit]:
                    lines.append(f"  - {item}")
                if len(items) > limit:
                    lines.append(f"  ... 还有 {len(items) - limit} 条未显示")
            lines.append("")

        append_section("识别到学号但未识别题目的文件", report.get("id_without_question", []))
        append_section("识别到题目但未匹配学生的文件", report.get("question_without_id", []))
        append_section("完全无法识别的文件", report.get("unrecognized", []))
        append_section("重复匹配记录", report.get("duplicates", []))
        return "\n".join(lines)

    def open_material_report(self):
        if self.material_indexing:
            messagebox.showinfo(APP_TITLE, "材料仍在后台匹配中，请稍后再查看报告。")
            return
        if not self.material_folder and not self.material_files:
            messagebox.showwarning(APP_TITLE, "请先选择材料文件夹。")
            return
        window = ctk.CTkToplevel(self)
        window.title("材料匹配报告")
        window.geometry("980x680")
        window.configure(fg_color=COLORS["bg"])
        window.grid_columnconfigure(0, weight=1)
        window.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(window, text="材料匹配报告", font=FONT_TITLE, text_color=COLORS["text"]).grid(row=0, column=0, padx=22, pady=(20, 8), sticky="w")
        box = ctk.CTkTextbox(window, font=FONT_BODY, fg_color=COLORS["card"], text_color=COLORS["text"], border_color=COLORS["line"], border_width=1, corner_radius=18)
        box.grid(row=1, column=0, padx=18, pady=(0, 18), sticky="nsew")
        box.insert("1.0", self.material_report_text())
        box.configure(state="disabled")

    def export_scores(self):
        if not self.students:
            messagebox.showwarning(APP_TITLE, "请先导入学生名单。")
            return
        filename = f"学生材料批阅成绩_{datetime.now().strftime('%Y%m%d_%H%M')}.csv"
        path = filedialog.asksaveasfilename(title="导出成绩表", defaultextension=".csv", initialfile=filename, filetypes=[("CSV 表格", "*.csv")])
        if not path:
            return
        headers = ["学号", "姓名"]
        headers += [f"第{i}题完成" for i in range(1, QUESTION_COUNT + 1)]
        headers += [f"第{i}题成绩" for i in range(1, QUESTION_COUNT + 1)]
        headers += [f"第{i}题备注" for i in range(1, QUESTION_COUNT + 1)]
        headers += ["必修完成数", "选修完成数", "总成绩"]
        with open(path, "w", newline="", encoding="utf-8-sig") as file:
            writer = csv.writer(file)
            writer.writerow(headers)
            for student in self.students:
                record = self.ensure_record(student["id"])
                required, optional = self.completion(student["id"])
                row = [student["id"], student["name"]]
                row += ["是" if q.get("submitted") else "否" for q in record["questions"]]
                row += [q.get("score", "") for q in record["questions"]]
                row += [q.get("note", "") for q in record["questions"]]
                row += [required, optional, f"{self.total_score(student['id']):g}"]
                writer.writerow(row)
        messagebox.showinfo(APP_TITLE, f"成绩表已导出：\n{path}")

    def clear_data(self):
        if not messagebox.askyesno(APP_TITLE, "确定清空本地保存的学生和评分数据吗？此操作不会删除材料文件。"):
            return
        self.students = []
        self.scores = {}
        self.current_index = 0
        self.material_folder = ""
        self.material_files = []
        self.material_index = {}
        self.material_report = self.empty_material_report()
        self.pending_review_ids = set()
        self.file_pattern = DEFAULT_PATTERN
        self.pattern_var.set("新规则：文件名包含学号 + 对应实践关键词即可自动匹配；旧命名规则仍兼容。")
        self.save_data()
        self.refresh_all()


if __name__ == "__main__":
    ReviewApp().mainloop()




























