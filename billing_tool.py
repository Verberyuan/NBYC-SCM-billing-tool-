# -*- coding: utf-8 -*-
"""
账单处理工具 —— 图形界面入口（v4.1）

本版新增
--------
* **燃油费率记忆**：填过一次之后自动记住，下次处理别的账单直接沿用，
  只有遇到还没填过的周才会再问你（见 rate_store.py）。
* 从源账单**反推燃油费率**，一键填入按周对话框，省得去官网一个个抄。
* 选择主文件后自动识别客户编码。
* 主文件 + 多个补充文件，支持拖拽（需 tkinterdnd2，未装则用按钮）。
"""

import os
import queue
import sys
import threading
import traceback
from datetime import date, datetime
from decimal import Decimal
from typing import Dict, List, Optional, Tuple

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

import billing_core as core
import customer_config as cfgmod
import fuel_rates as fr
import rate_store
from money import parse_rate, format_rate

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except Exception:
    DND_FILES = None
    TkinterDnD = None
    DND_AVAILABLE = False


APP_TITLE = "账单处理工具"
APP_VERSION = "v4.0.0"

FUEL_MODE_LABELS = {
    "none": "不计燃油（该客户无需费率）",
    "schedule": "按燃油费率表计算",
    "source": "直接使用源文件燃油金额",
}

FUEL_SOURCE_MEMORY = "memory"
FUEL_SOURCE_UNIFORM = "uniform"
FUEL_SOURCE_WEEKLY = "weekly"
FUEL_SOURCE_CSV = "csv"


def resource_path(relative_path: str) -> str:
    if getattr(sys, "frozen", False):
        base_path = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


def parse_dropped_paths(data: str) -> List[str]:
    paths, buf, in_brace = [], "", False
    for ch in data:
        if ch == "{":
            in_brace, buf = True, ""
        elif ch == "}":
            in_brace = False
            if buf:
                paths.append(buf)
            buf = ""
        elif ch == " " and not in_brace:
            if buf:
                paths.append(buf)
            buf = ""
        else:
            buf += ch
    if buf:
        paths.append(buf)
    return [p for p in paths if p.lower().endswith((".xlsx", ".xls", ".xlsm"))]


class BillingToolApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"{APP_TITLE} {APP_VERSION}")
        self.root.geometry("880x860")
        self.root.minsize(800, 680)

        codes = cfgmod.list_customer_codes() or ["DEFAULT"]
        self.customer_var = tk.StringVar(value=codes[0])
        self.input_path_var = tk.StringVar()
        self.output_path_var = tk.StringVar()
        self.month_var = tk.StringVar(value=self._default_month())
        self.fuel_source_var = tk.StringVar(value=FUEL_SOURCE_MEMORY)
        self.fuel_rate_var = tk.StringVar()
        self.fuel_mode_text_var = tk.StringVar()
        self.memory_text_var = tk.StringVar()
        self.rate_csv_var = tk.StringVar()
        self.expected_count_var = tk.StringVar()
        self.formula_func_var = tk.StringVar(value="SUM（普通求和）")
        self.other_mode_var = tk.StringVar(value="自动识别（推荐）")
        self.remove_empty_var = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(value="请选择主账单文件（会自动识别客户编码）")

        self.supplement_paths: List[str] = []
        self._output_path_auto = True
        self._processing = False

        self._log_queue: queue.Queue = queue.Queue()
        self._worker_thread: Optional[threading.Thread] = None
        self._worker_result = None
        self._worker_error: Optional[BaseException] = None

        self._build_widgets()
        self._refresh_memory_label()
        self._on_customer_change()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ------------------------------------------------------------------
    # 界面
    # ------------------------------------------------------------------
    def _default_month(self) -> str:
        today = date.today()
        return f"{today.year}-{today.month:02d}"

    def _build_widgets(self):
        pad = {"padx": 10, "pady": 5}

        # ---- 文件 ----
        frm_file = ttk.LabelFrame(self.root, text="第一步：选择账单文件（主文件 1 个 + 补充文件若干）")
        frm_file.pack(fill="x", **pad)

        row_main = ttk.Frame(frm_file)
        row_main.pack(fill="x", padx=10, pady=(8, 4))
        ttk.Label(row_main, text="主文件（整月）：", width=14).pack(side="left")
        ttk.Entry(row_main, textvariable=self.input_path_var, state="readonly").pack(
            side="left", fill="x", expand=True, padx=(0, 6))
        self.browse_input_btn = ttk.Button(row_main, text="浏览...", command=self.browse_input_file)
        self.browse_input_btn.pack(side="left")

        row_sup = ttk.Frame(frm_file)
        row_sup.pack(fill="x", padx=10, pady=(4, 4))
        ttk.Label(row_sup, text="补充文件（周）：", width=14).pack(side="left", anchor="n")

        list_frame = ttk.Frame(row_sup)
        list_frame.pack(side="left", fill="x", expand=True, padx=(0, 6))
        self.supplement_list = tk.Listbox(list_frame, height=4, selectmode="extended")
        self.supplement_list.pack(side="left", fill="x", expand=True)
        sb = ttk.Scrollbar(list_frame, orient="vertical", command=self.supplement_list.yview)
        sb.pack(side="left", fill="y")
        self.supplement_list.configure(yscrollcommand=sb.set)

        btn_col = ttk.Frame(row_sup)
        btn_col.pack(side="left", anchor="n")
        self.add_sup_btn = ttk.Button(btn_col, text="添加...", command=self.add_supplement_files, width=10)
        self.add_sup_btn.pack(pady=(0, 4))
        ttk.Button(btn_col, text="移除选中", command=self.remove_selected_supplements, width=10).pack(pady=(0, 4))
        ttk.Button(btn_col, text="清空", command=self.clear_supplements, width=10).pack()

        hint = ("提示：可以把 Excel 直接拖进窗口，第一个作主文件、其余作补充文件。"
                if DND_AVAILABLE else
                "提示：如需拖拽文件，请先安装 tkinterdnd2（pip install tkinterdnd2）。")
        ttk.Label(frm_file, text=hint, foreground="#666666").pack(anchor="w", padx=10, pady=(0, 6))
        if DND_AVAILABLE:
            self._enable_drag_drop()

        # ---- 客户与账期 ----
        frm_cust = ttk.LabelFrame(self.root, text="第二步：客户与账期")
        frm_cust.pack(fill="x", **pad)

        row = ttk.Frame(frm_cust)
        row.pack(fill="x", padx=10, pady=(8, 4))
        ttk.Label(row, text="客户编码：", width=14).pack(side="left")
        ttk.Combobox(row, textvariable=self.customer_var, state="readonly", width=12,
                     values=cfgmod.list_customer_codes() or ["DEFAULT"]).pack(side="left")
        self.customer_var.trace_add("write", lambda *_: self._on_customer_change())
        ttk.Label(row, text="账单月份：").pack(side="left", padx=(20, 4))
        ttk.Entry(row, textvariable=self.month_var, width=12).pack(side="left")
        ttk.Label(row, text="格式 YYYY-MM").pack(side="left", padx=6)

        row_b = ttk.Frame(frm_cust)
        row_b.pack(fill="x", padx=10, pady=(0, 4))
        ttk.Label(row_b, text="", width=14).pack(side="left")
        ttk.Label(row_b, textvariable=self.fuel_mode_text_var, foreground="#005599").pack(side="left")

        row_c = ttk.Frame(frm_cust)
        row_c.pack(fill="x", padx=10, pady=(0, 8))
        ttk.Label(row_c, text="系统订单件数：", width=14).pack(side="left")
        ttk.Entry(row_c, textvariable=self.expected_count_var, width=12).pack(side="left")
        ttk.Label(row_c, text="可选，用于核对本账期计入的订单行数").pack(side="left", padx=8)

        # ---- 燃油费率 ----
        self.frm_fuel = ttk.LabelFrame(self.root, text="第三步：燃油费率")
        self.frm_fuel.pack(fill="x", **pad)

        mem_row = ttk.Frame(self.frm_fuel)
        mem_row.pack(fill="x", padx=10, pady=(8, 2))
        self.rb_memory = ttk.Radiobutton(
            mem_row, text="使用已记忆的费率（推荐）", value=FUEL_SOURCE_MEMORY,
            variable=self.fuel_source_var, command=self._on_fuel_source_change)
        self.rb_memory.pack(side="left")
        ttk.Label(mem_row, textvariable=self.memory_text_var, foreground="#005599").pack(side="left", padx=10)
        ttk.Button(mem_row, text="管理已记忆费率...", command=self.manage_memory, width=18).pack(side="right")

        uni_row = ttk.Frame(self.frm_fuel)
        uni_row.pack(fill="x", padx=10, pady=2)
        self.rb_uniform = ttk.Radiobutton(
            uni_row, text="统一费率：", value=FUEL_SOURCE_UNIFORM,
            variable=self.fuel_source_var, command=self._on_fuel_source_change)
        self.rb_uniform.pack(side="left")
        self.fuel_rate_entry = ttk.Entry(uni_row, textvariable=self.fuel_rate_var, width=12)
        self.fuel_rate_entry.pack(side="left", padx=6)
        ttk.Label(uni_row, text="支持 25% / 25 / 0.25，应用到本账期所有费率周").pack(side="left")

        wk_row = ttk.Frame(self.frm_fuel)
        wk_row.pack(fill="x", padx=10, pady=2)
        self.rb_weekly = ttk.Radiobutton(
            wk_row, text="按周分别填写（点开始后弹窗，可从源账单一键反推）", value=FUEL_SOURCE_WEEKLY,
            variable=self.fuel_source_var, command=self._on_fuel_source_change)
        self.rb_weekly.pack(side="left")

        csv_row = ttk.Frame(self.frm_fuel)
        csv_row.pack(fill="x", padx=10, pady=(2, 8))
        self.rb_csv = ttk.Radiobutton(
            csv_row, text="导入费率表：", value=FUEL_SOURCE_CSV,
            variable=self.fuel_source_var, command=self._on_fuel_source_change)
        self.rb_csv.pack(side="left")
        ttk.Entry(csv_row, textvariable=self.rate_csv_var, state="readonly").pack(
            side="left", fill="x", expand=True, padx=6)
        self.rate_csv_btn = ttk.Button(csv_row, text="选择CSV...", command=self.browse_rate_csv, width=12)
        self.rate_csv_btn.pack(side="left")

        # ---- 高级 ----
        frm_adv = ttk.LabelFrame(self.root, text="高级选项（无特殊需求可保持默认）")
        frm_adv.pack(fill="x", **pad)

        row4 = ttk.Frame(frm_adv)
        row4.pack(fill="x", padx=10, pady=(8, 4))
        ttk.Label(row4, text="汇总函数：", width=14).pack(side="left")
        ttk.Combobox(row4, textvariable=self.formula_func_var, state="readonly", width=20,
                     values=["SUM（普通求和）", "SUBTOTAL_109（忽略隐藏行）"]).pack(side="left")
        ttk.Label(row4, text="未知sheet：").pack(side="left", padx=(20, 4))
        ttk.Combobox(row4, textvariable=self.other_mode_var, state="readonly", width=16,
                     values=["自动识别（推荐）", "逐个询问", "全部跳过"]).pack(side="left")

        row5 = ttk.Frame(frm_adv)
        row5.pack(fill="x", padx=10, pady=(0, 8))
        ttk.Checkbutton(row5, text="自动删除「仅有表头、无数据」的空sheet（汇总/仓租/出库订单/核账表除外）",
                        variable=self.remove_empty_var).pack(side="left")

        # ---- 输出 ----
        frm_out = ttk.LabelFrame(self.root, text="第四步：输出文件")
        frm_out.pack(fill="x", **pad)
        ttk.Entry(frm_out, textvariable=self.output_path_var, state="readonly").pack(
            side="left", fill="x", expand=True, padx=(10, 6), pady=8)
        self.browse_output_btn = ttk.Button(frm_out, text="另存为...", command=self.browse_output_file)
        self.browse_output_btn.pack(side="left", padx=(0, 10), pady=8)

        # ---- 操作 ----
        frm_action = ttk.Frame(self.root)
        frm_action.pack(fill="x", padx=10, pady=(4, 6))
        self.start_btn = ttk.Button(frm_action, text="一键生成结果", command=self.on_start)
        self.start_btn.pack(side="left", ipadx=16, ipady=6)
        self.open_folder_btn = ttk.Button(frm_action, text="打开输出文件夹",
                                          command=self.open_output_folder, state="disabled")
        self.open_folder_btn.pack(side="left", padx=10)
        ttk.Label(frm_action, textvariable=self.status_var, foreground="#333333").pack(side="left", padx=16)

        # ---- 日志 ----
        frm_log = ttk.LabelFrame(self.root, text="处理日志")
        frm_log.pack(fill="both", expand=True, **pad)
        self.log_text = scrolledtext.ScrolledText(frm_log, state="disabled", font=("Consolas", 9), wrap="word")
        self.log_text.pack(fill="both", expand=True, padx=6, pady=6)

    def _enable_drag_drop(self):
        def on_drop(event):
            paths = parse_dropped_paths(event.data)
            if not paths:
                return
            if not self.input_path_var.get():
                self._set_input_path(paths[0])
                paths = paths[1:]
            self._add_supplements(paths)

        try:
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind("<<Drop>>", on_drop)
            self.supplement_list.drop_target_register(DND_FILES)
            self.supplement_list.dnd_bind(
                "<<Drop>>", lambda e: self._add_supplements(parse_dropped_paths(e.data)))
        except Exception:
            pass

    # ------------------------------------------------------------------
    # 客户 / 燃油联动
    # ------------------------------------------------------------------
    def _current_config(self):
        return cfgmod.get_customer_config(self.customer_var.get())

    def _carrier(self) -> str:
        return self._current_config()["fuel"].get("carrier") or "FedEx"

    def _refresh_memory_label(self):
        periods = rate_store.load_periods(self._carrier())
        self.memory_text_var.set(rate_store.describe_memory(periods))

    def _on_customer_change(self):
        config = self._current_config()
        mode = config["fuel"]["mode"]
        self.fuel_mode_text_var.set(f"该客户燃油模式：{FUEL_MODE_LABELS.get(mode, mode)}")

        needs_rate = (mode == "schedule")
        for widget in (self.rb_memory, self.rb_uniform, self.rb_weekly, self.rb_csv,
                       self.fuel_rate_entry, self.rate_csv_btn):
            try:
                widget.configure(state="normal" if needs_rate else "disabled")
            except Exception:
                pass
        self._refresh_memory_label()
        if needs_rate:
            self._on_fuel_source_change()

        if self._output_path_auto and self.input_path_var.get():
            self._set_auto_output_path(self.input_path_var.get())

    def _on_fuel_source_change(self):
        source = self.fuel_source_var.get()
        try:
            self.fuel_rate_entry.configure(state="normal" if source == FUEL_SOURCE_UNIFORM else "disabled")
            self.rate_csv_btn.configure(state="normal" if source == FUEL_SOURCE_CSV else "disabled")
        except Exception:
            pass

    # ------------------------------------------------------------------
    # 文件
    # ------------------------------------------------------------------
    def browse_input_file(self):
        path = filedialog.askopenfilename(
            title="选择主账单Excel文件（整月）",
            filetypes=[("Excel 文件", "*.xlsx *.xls *.xlsm *.XLSX"), ("所有文件", "*.*")])
        if path:
            self._set_input_path(path)

    def _set_input_path(self, path: str):
        self.input_path_var.set(path)
        self.status_var.set("已选择主文件，可继续添加补充文件")
        detected = self._detect_customer(path)
        if detected:
            self.customer_var.set(detected)
            self.status_var.set(f"已自动识别客户编码：{detected}")
        if self._output_path_auto:
            self._set_auto_output_path(path)

    def _detect_customer(self, path: str) -> str:
        """先看文件名，再看【汇总】里的客户编号。"""
        name = os.path.basename(path).upper()
        for code in cfgmod.list_customer_codes():
            if code in name:
                return code
        try:
            values = core.read_summary_template_values(path)
            return cfgmod.detect_customer_code([values.get("客户编号", "")])
        except Exception:
            return ""

    def add_supplement_files(self):
        paths = filedialog.askopenfilenames(
            title="选择补充账单文件（可多选）",
            filetypes=[("Excel 文件", "*.xlsx *.xls *.xlsm *.XLSX"), ("所有文件", "*.*")])
        self._add_supplements(list(paths))

    def _add_supplements(self, paths: List[str]):
        added = 0
        for path in paths:
            if not path or path in self.supplement_paths:
                continue
            if os.path.abspath(path) == os.path.abspath(self.input_path_var.get() or ""):
                continue
            self.supplement_paths.append(path)
            self.supplement_list.insert("end", os.path.basename(path))
            added += 1
        if added:
            self.status_var.set(f"已添加 {added} 个补充文件，共 {len(self.supplement_paths)} 个")

    def remove_selected_supplements(self):
        for idx in sorted(self.supplement_list.curselection(), reverse=True):
            self.supplement_list.delete(idx)
            del self.supplement_paths[idx]

    def clear_supplements(self):
        self.supplement_list.delete(0, "end")
        self.supplement_paths.clear()

    def browse_rate_csv(self):
        path = filedialog.askopenfilename(
            title="选择燃油费率表 CSV", filetypes=[("CSV 文件", "*.csv"), ("所有文件", "*.*")])
        if path:
            self.rate_csv_var.set(path)

    def _set_auto_output_path(self, input_path: str):
        folder = os.path.dirname(input_path)
        stem, _ext = os.path.splitext(os.path.basename(input_path))
        for suffix in ("_已清洗", "_已处理"):
            if stem.endswith(suffix):
                stem = stem[: -len(suffix)]
                break
        self.output_path_var.set(os.path.join(folder, f"{stem}_已处理.xlsx"))
        self._output_path_auto = True

    def browse_output_file(self):
        default_dir = os.path.dirname(self.output_path_var.get()) or os.path.dirname(self.input_path_var.get()) or "."
        default_name = os.path.basename(self.output_path_var.get()) or "已处理结果.xlsx"
        path = filedialog.asksaveasfilename(
            title="选择结果保存位置", initialdir=default_dir, initialfile=default_name,
            defaultextension=".xlsx", filetypes=[("Excel 文件", "*.xlsx")])
        if path:
            self.output_path_var.set(path)
            self._output_path_auto = False

    def open_output_folder(self):
        out_path = self.output_path_var.get()
        folder = os.path.dirname(out_path) or "."
        try:
            if sys.platform.startswith("win"):
                os.startfile(folder)  # noqa: S606
            elif sys.platform == "darwin":
                os.system(f'open "{folder}"')
            else:
                os.system(f'xdg-open "{folder}"')
        except Exception:
            messagebox.showinfo("提示", f"结果文件位于：\n{out_path}")

    # ------------------------------------------------------------------
    # 日志（线程安全）
    # ------------------------------------------------------------------
    def append_log(self, text: str):
        self._log_queue.put(str(text))

    def _drain_log_queue(self):
        lines = []
        try:
            while True:
                lines.append(self._log_queue.get_nowait())
        except queue.Empty:
            pass
        if not lines:
            return
        self.log_text.configure(state="normal")
        self.log_text.insert("end", "\n".join(lines) + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def clear_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def _ask_on_main_thread(self, func, *args):
        box = {"result": None, "error": None}
        done = threading.Event()

        def runner():
            try:
                box["result"] = func(*args)
            except BaseException as exc:  # noqa: BLE001
                box["error"] = exc
            finally:
                done.set()

        self.root.after(0, runner)
        done.wait()
        if box["error"] is not None:
            raise box["error"]
        return box["result"]

    # ------------------------------------------------------------------
    # 校验
    # ------------------------------------------------------------------
    def _parse_month(self, raw: str):
        parts = raw.strip().split("-")
        if len(parts) != 2 or not parts[0].isdigit() or not parts[1].isdigit():
            raise ValueError("账单月份格式不正确，请填写为 YYYY-MM，例如 2026-07。")
        year, month = int(parts[0]), int(parts[1])
        if not (1 <= month <= 12):
            raise ValueError("账单月份中的“月”应为 1-12。")
        if not (2000 <= year <= 2100):
            raise ValueError("账单月份中的“年”看起来不正确，请检查。")
        return year, month

    def _parse_expected_count(self, raw: str):
        raw = raw.strip()
        if not raw:
            return None
        if not raw.isdigit():
            raise ValueError("系统订单件数应为整数，如需跳过核对请留空。")
        return int(raw)

    def _formula_function_value(self) -> str:
        return "SUBTOTAL_109" if "SUBTOTAL" in self.formula_func_var.get() else "SUM"

    def _other_sheet_mode_value(self) -> str:
        text = self.other_mode_var.get()
        if "跳过" in text:
            return "skip"
        if "询问" in text:
            return "ask"
        return "auto"

    # ------------------------------------------------------------------
    # 主流程
    # ------------------------------------------------------------------
    def on_start(self):
        if self._processing:
            return

        input_path = self.input_path_var.get().strip()
        if not input_path:
            messagebox.showwarning("提示", "请先选择主账单Excel文件。")
            return
        if not os.path.exists(input_path):
            messagebox.showerror("错误", "所选主文件不存在，请重新选择。")
            return

        output_path = self.output_path_var.get().strip()
        if not output_path:
            messagebox.showwarning("提示", "请设置结果输出路径。")
            return
        if os.path.abspath(output_path) == os.path.abspath(input_path):
            messagebox.showerror("错误", "输出文件不能与主文件相同，请更换保存路径。")
            return

        try:
            year, month = self._parse_month(self.month_var.get())
        except ValueError as e:
            messagebox.showerror("账单月份格式错误", str(e))
            return

        try:
            expected_count = self._parse_expected_count(self.expected_count_var.get())
        except ValueError as e:
            messagebox.showerror("系统订单件数格式错误", str(e))
            return

        config = self._current_config()
        fuel_mode = config["fuel"]["mode"]
        fuel_source = self.fuel_source_var.get() if fuel_mode == "schedule" else None
        uniform_rate: Optional[Decimal] = None

        if fuel_source == FUEL_SOURCE_UNIFORM:
            raw = self.fuel_rate_var.get().strip()
            if not raw:
                messagebox.showwarning("提示", "请填写统一燃油费率，例如 25%。")
                return
            try:
                uniform_rate = parse_rate(raw)
            except Exception:
                messagebox.showerror("燃油费率格式错误", "请填写 25%、25 或 0.25 这样的格式。")
                return
        elif fuel_source == FUEL_SOURCE_CSV:
            if not self.rate_csv_var.get().strip():
                messagebox.showwarning("提示", "请先选择燃油费率表 CSV 文件。")
                return

        self._start_worker(
            input_path=input_path, output_path=output_path, year=year, month=month,
            customer_code=self.customer_var.get(), fuel_source=fuel_source,
            uniform_rate=uniform_rate, expected_count=expected_count,
        )

    def _start_worker(self, input_path, output_path, year, month, customer_code,
                      fuel_source, uniform_rate, expected_count):
        self._processing = True
        for btn in (self.start_btn, self.open_folder_btn, self.browse_input_btn,
                    self.browse_output_btn, self.add_sup_btn):
            btn.configure(state="disabled")
        self.status_var.set("处理中，请稍候...")
        self.clear_log()

        self._log_queue = queue.Queue()
        self._worker_result = None
        self._worker_error = None

        params = dict(
            input_path=input_path, output_path=output_path, year=year, month=month,
            customer_code=customer_code, fuel_source=fuel_source,
            uniform_rate=uniform_rate, expected_count=expected_count,
            supplements=list(self.supplement_paths),
            rate_csv=self.rate_csv_var.get().strip() or None,
            formula_function=self._formula_function_value(),
            other_sheet_mode=self._other_sheet_mode_value(),
            remove_empty=self.remove_empty_var.get(),
        )
        self._worker_thread = threading.Thread(target=self._worker, kwargs=params, daemon=True)
        self._worker_thread.start()
        self.root.after(100, self._poll_worker)

    def _worker(self, input_path, output_path, year, month, customer_code, fuel_source,
                uniform_rate, expected_count, supplements, rate_csv, formula_function,
                other_sheet_mode, remove_empty):
        try:
            config = cfgmod.get_customer_config(customer_code)
            carrier = config["fuel"].get("carrier") or "FedEx"

            fuel_periods = None
            weekly_callback = None

            if config["fuel"]["mode"] == "schedule":
                if fuel_source == FUEL_SOURCE_MEMORY:
                    fuel_periods = self._resolve_from_memory(
                        input_path, supplements, customer_code, year, month, carrier)
                elif fuel_source == FUEL_SOURCE_WEEKLY:
                    def weekly_callback(week_ranges):
                        defaults = self._infer_defaults(input_path, customer_code, year, month)
                        return self._ask_on_main_thread(
                            self._weekly_rate_dialog, week_ranges, defaults, None)

            self._worker_result = core.run_pipeline(
                main_path=input_path,
                supplement_paths=supplements,
                year=year, month=month, output_path=output_path,
                customer_code=customer_code,
                uniform_fuel_rate=uniform_rate,
                fuel_periods=fuel_periods,
                fuel_rate_csv=rate_csv if fuel_source == FUEL_SOURCE_CSV else None,
                weekly_rate_callback=weekly_callback,
                expected_order_count=expected_count,
                formula_function=formula_function,
                other_sheet_mode=other_sheet_mode,
                remove_empty_sheets=remove_empty,
                column_choice_callback=lambda s, c: self._ask_on_main_thread(
                    self._choose_column_dialog, s, c),
                order_count_mismatch_callback=lambda a, e: self._ask_on_main_thread(
                    self._confirm_order_count_mismatch, a, e),
                precheck_callback=lambda r: self._ask_on_main_thread(self._precheck_dialog, r),
                log=self.append_log,
            )

            # ---- 记忆本次用到的费率，下次直接沿用 ----
            used = self._worker_result.get("fuel_periods") or []
            if used:
                path = rate_store.remember(used)
                self.append_log(f"已记忆本次燃油费率（{len(used)} 个区间）：{path}")
        except BaseException as exc:  # noqa: BLE001
            self._worker_error = exc

    def _resolve_from_memory(self, input_path, supplements, customer_code, year, month, carrier):
        """用记忆里的费率；缺哪几周就只问哪几周。"""
        periods = rate_store.load_periods(carrier)
        self.append_log(f"读取已记忆费率：{rate_store.describe_memory(periods)}")

        week_ranges = core.preview_week_ranges(input_path, supplements, customer_code, year, month)
        missing = rate_store.missing_weeks(periods, week_ranges, carrier)

        if not missing:
            self.append_log("已记忆的费率已覆盖本账期全部费率周，无需再填写。")
            return periods

        self.append_log(f"本账期有 {len(missing)} 个费率周尚未记忆，需要补填。")
        defaults = self._infer_defaults(input_path, customer_code, year, month)
        filled = self._ask_on_main_thread(self._weekly_rate_dialog, missing, defaults, periods)
        if filled is None:
            raise core.PipelineCancelled("用户取消了燃油费率的补填。")

        new_periods = [
            fr.FuelPeriod(start=s, end=e, rate=filled[s], carrier=carrier)
            for s, e in missing if s in filled
        ]
        return periods + new_periods

    def _infer_defaults(self, input_path, customer_code, year, month) -> Dict[date, Decimal]:
        """从源账单反推每周实际费率，用作弹窗预填值。"""
        try:
            inferred = core.infer_rates_from_file(input_path, customer_code, year, month)
        except Exception as exc:  # noqa: BLE001
            self.append_log(f"（反推燃油费率失败，将留空由人工填写：{exc}）")
            return {}
        result = {}
        for start, _end, rate, count, note in inferred:
            if rate is not None and count > 0:
                result[start] = rate
                if str(note).startswith("★"):
                    self.append_log(f"  反推提示 {start:%Y-%m-%d}：{note}")
        return result

    def _poll_worker(self):
        self._drain_log_queue()
        if self._worker_thread is not None and self._worker_thread.is_alive():
            self.root.after(100, self._poll_worker)
            return
        self._drain_log_queue()
        self._finish()

    def _finish(self):
        err = self._worker_error
        try:
            if err is None:
                self._show_success(self._worker_result)
            elif isinstance(err, core.PipelineCancelled):
                self.status_var.set("已取消")
                self.append_log(f"处理已取消：{err}")
                self._drain_log_queue()
            elif isinstance(err, core.OrderCountMismatchError):
                self.status_var.set("已取消：订单件数核对未通过")
                self.append_log("处理已取消：订单件数核对未通过，且用户选择不继续。")
                self._drain_log_queue()
            elif isinstance(err, core.BillingProcessError):
                self.status_var.set("处理失败")
                self.append_log(f"[错误] {err}")
                self._drain_log_queue()
                messagebox.showerror("处理失败", str(err))
            else:
                self.status_var.set("处理失败")
                detail = "".join(traceback.format_exception(type(err), err, err.__traceback__))
                self.append_log(f"[未预期的错误] {err}\n{detail}")
                self._drain_log_queue()
                messagebox.showerror("处理失败", f"发生未预期的错误：\n{err}\n\n可复制日志内容反馈给开发者。")
        finally:
            self._processing = False
            self._worker_thread = None
            for btn in (self.start_btn, self.browse_input_btn, self.browse_output_btn, self.add_sup_btn):
                btn.configure(state="normal")
            self._refresh_memory_label()

    def _on_close(self):
        if self._processing:
            if not messagebox.askyesno("正在处理中", "账单正在处理中，现在关闭会中断处理。\n\n确定要关闭吗？"):
                return
        self.root.destroy()

    def _show_success(self, result):
        self.open_folder_btn.configure(state="normal")
        stats = result.get("stats", {})
        issues = result.get("issue_counts", {})
        n = len(result["summary_items"])
        self.status_var.set(f"处理完成！共 {n} 项费用已汇总。")

        issue_text = "\n".join(f"  {k}：{v} 条" for k, v in issues.items() if v)
        flagged = [r for r in (result.get("reconciliation") or []) if str(r["说明"]).startswith("★")]
        flag_text = "\n".join(
            f"  {r['项目']}：账单 {r['账单原值']} / 明细 {r['明细重算值']}，差额 {r['差额']}"
            for r in flagged)
        removed = result.get("removed_empty_sheets") or []
        removed_text = f"\n已自动删除空sheet（{len(removed)}个）：{'、'.join(removed)}" if removed else ""

        messagebox.showinfo(
            "处理完成",
            f"账单处理已完成！\n\n结果文件：\n{result['output_path']}\n\n"
            f"客户：{result['customer_code']}\n"
            f"计入订单行数：{stats.get('计算订单行数')}（独立订单 {stats.get('独立订单编号数')} 个）\n"
            f"跨月剔除：{stats.get('跨月订单行数')} 行；规则剔除：{stats.get('按规则剔除行数')} 行\n"
            f"重算费用总计：{stats.get('重算费用总计')}\n"
            + (f"\n★账单与明细不符（可能漏单）：\n{flag_text}\n" if flag_text else "")
            + (f"\n需关注的异常：\n{issue_text}\n" if issue_text else "")
            + f"\n共汇总 {n} 项费用，明细见「汇总」「核对汇总」「逐单核对」。{removed_text}",
        )

    # ------------------------------------------------------------------
    # 对话框
    # ------------------------------------------------------------------
    def _precheck_dialog(self, report) -> bool:
        if not report.has_warning:
            return True

        dlg = tk.Toplevel(self.root)
        dlg.title("运行前检查结果")
        dlg.transient(self.root)
        dlg.grab_set()

        ttk.Label(dlg, text="以下为运行前检查结果，请确认无误后继续：",
                  justify="left").pack(padx=16, pady=(16, 8), anchor="w")
        text = scrolledtext.ScrolledText(dlg, width=100, height=22, font=("Consolas", 9), wrap="word")
        text.pack(padx=16, fill="both", expand=True)
        text.insert("end", report.as_text())
        text.configure(state="disabled")

        result = {"value": False}

        def on_continue():
            result["value"] = True
            dlg.destroy()

        btns = ttk.Frame(dlg)
        btns.pack(pady=12)
        ttk.Button(btns, text="继续计算", command=on_continue, width=12).pack(side="left", padx=6)
        ttk.Button(btns, text="中止", command=dlg.destroy, width=12).pack(side="left", padx=6)

        dlg.protocol("WM_DELETE_WINDOW", dlg.destroy)
        self.root.wait_window(dlg)
        return result["value"]

    def _confirm_order_count_mismatch(self, actual: int, expected: int) -> bool:
        return messagebox.askyesno(
            "订单件数核对不一致",
            f"系统中订单件数为 {expected}，本账期实际计入 {actual} 行，相差 {actual - expected} 条。\n\n"
            f"是否仍要继续生成结果？\n（选择“否”将中止本次处理）")

    def _weekly_rate_dialog(self, week_ranges, defaults: Optional[Dict[date, Decimal]],
                            known_periods) -> Optional[Dict[date, Decimal]]:
        """按周填写费率。defaults 为从源账单反推出来的预填值。"""
        result = {"value": None}
        defaults = defaults or {}

        dlg = tk.Toplevel(self.root)
        dlg.title("填写燃油费率")
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.resizable(False, False)

        ttk.Label(
            dlg,
            text=(f"以下 {len(week_ranges)} 个费率周还没有燃油费率，请填写"
                  f"（支持 25% / 25 / 0.25）。\n"
                  f"区间按承运商官网口径的自然周，未做任何日期平移。\n"
                  f"带底色的数值是程序从本份账单反推出来的参考值，请与 FedEx 官网核对后再确认。"),
            justify="left",
        ).pack(padx=16, pady=(16, 8), anchor="w")

        rows_frame = ttk.Frame(dlg)
        rows_frame.pack(padx=16, pady=4, fill="x")

        entry_vars: List[Tuple[date, tk.StringVar]] = []
        for idx, (start, end) in enumerate(week_ranges, start=1):
            row = ttk.Frame(rows_frame)
            row.pack(fill="x", pady=2)
            ttk.Label(row, text=f"第{idx}周（{start:%Y-%m-%d} 至 {end:%Y-%m-%d}）：", width=34).pack(side="left")
            prefill = format_rate(defaults[start]) if start in defaults else ""
            var = tk.StringVar(value=prefill)
            entry = tk.Entry(row, textvariable=var, width=12)
            if prefill:
                entry.configure(background="#FFF7D6")
            entry.pack(side="left")
            if start in defaults:
                ttk.Label(row, text="（源账单反推）", foreground="#888888").pack(side="left", padx=6)
            entry_vars.append((start, var))

        error_label = ttk.Label(dlg, text="", foreground="#cc0000")
        error_label.pack(padx=16, anchor="w")

        def fill_all():
            first = entry_vars[0][1].get().strip()
            if not first:
                error_label.configure(text="请先填好第一周，再点「全部同第一周」。")
                return
            for _start, var in entry_vars[1:]:
                var.set(first)
            error_label.configure(text="")

        def on_confirm():
            rates: Dict[date, Decimal] = {}
            for start, var in entry_vars:
                raw = var.get().strip()
                if not raw:
                    error_label.configure(text="请填写每一周的燃油费率，不能留空。")
                    return
                try:
                    rates[start] = parse_rate(raw)
                except Exception:
                    error_label.configure(text=f"「{raw}」格式不正确，请填写 25%、25 或 0.25。")
                    return
            result["value"] = rates
            dlg.destroy()

        btn_frame = ttk.Frame(dlg)
        btn_frame.pack(pady=(6, 16))
        ttk.Button(btn_frame, text="全部同第一周", command=fill_all, width=14).pack(side="left", padx=6)
        ttk.Button(btn_frame, text="确定并记忆", command=on_confirm, width=14).pack(side="left", padx=6)
        ttk.Button(btn_frame, text="取消", command=dlg.destroy, width=10).pack(side="left", padx=6)

        dlg.protocol("WM_DELETE_WINDOW", dlg.destroy)
        self.root.wait_window(dlg)
        return result["value"]

    def manage_memory(self):
        """查看 / 编辑 / 删除已记忆的燃油费率。"""
        carrier = self._carrier()
        dlg = tk.Toplevel(self.root)
        dlg.title("管理已记忆的燃油费率")
        dlg.transient(self.root)
        dlg.grab_set()

        ttk.Label(dlg, text=f"记忆文件：{rate_store.store_path()}",
                  foreground="#666666").pack(padx=16, pady=(14, 6), anchor="w")

        cols = ("carrier", "start", "end", "rate")
        tree = ttk.Treeview(dlg, columns=cols, show="headings", height=14)
        for col, title, width in zip(cols, ("承运商", "开始日期", "结束日期", "燃油费率"),
                                     (100, 130, 130, 110)):
            tree.heading(col, text=title)
            tree.column(col, width=width, anchor="center")
        tree.pack(padx=16, fill="both", expand=True)

        def reload_tree():
            tree.delete(*tree.get_children())
            for p in rate_store.load_periods():
                tree.insert("", "end", values=(p.carrier, f"{p.start:%Y-%m-%d}",
                                               f"{p.end:%Y-%m-%d}", format_rate(p.rate)))

        reload_tree()

        add_frame = ttk.LabelFrame(dlg, text="新增 / 覆盖一个区间")
        add_frame.pack(fill="x", padx=16, pady=10)
        v_start, v_end, v_rate = tk.StringVar(), tk.StringVar(), tk.StringVar()
        for label, var, width in (("开始日期", v_start, 14), ("结束日期", v_end, 14), ("费率", v_rate, 10)):
            ttk.Label(add_frame, text=f"{label}：").pack(side="left", padx=(8, 2), pady=8)
            ttk.Entry(add_frame, textvariable=var, width=width).pack(side="left")

        msg = ttk.Label(dlg, text="", foreground="#cc0000")
        msg.pack(padx=16, anchor="w")

        def add_period():
            try:
                start = datetime.strptime(v_start.get().strip(), "%Y-%m-%d").date()
                end = datetime.strptime(v_end.get().strip(), "%Y-%m-%d").date()
                rate = parse_rate(v_rate.get().strip())
            except Exception:
                msg.configure(text="请按 2026-07-06 的格式填日期，费率填 25% 或 0.25。")
                return
            if start > end:
                msg.configure(text="开始日期不能晚于结束日期。")
                return
            rate_store.remember([fr.FuelPeriod(start=start, end=end, rate=rate, carrier=carrier)])
            msg.configure(text="")
            reload_tree()
            self._refresh_memory_label()

        def delete_selected():
            selected = {tree.item(i)["values"][1] for i in tree.selection()}
            if not selected:
                return
            kept = [p for p in rate_store.load_periods() if f"{p.start:%Y-%m-%d}" not in selected]
            rate_store.replace_all(kept)
            reload_tree()
            self._refresh_memory_label()

        def clear_all():
            if messagebox.askyesno("确认", "确定要清空全部已记忆的燃油费率吗？", parent=dlg):
                rate_store.clear()
                reload_tree()
                self._refresh_memory_label()

        def export_csv():
            path = filedialog.asksaveasfilename(
                parent=dlg, title="导出费率表", defaultextension=".csv",
                initialfile="fuel_rates.csv", filetypes=[("CSV 文件", "*.csv")])
            if path:
                fr.write_periods_template_csv(path, rate_store.load_periods())
                messagebox.showinfo("已导出", f"费率表已导出到：\n{path}", parent=dlg)

        btns = ttk.Frame(dlg)
        btns.pack(pady=(4, 14))
        ttk.Button(btns, text="添加/覆盖", command=add_period, width=12).pack(side="left", padx=5)
        ttk.Button(btns, text="删除选中", command=delete_selected, width=12).pack(side="left", padx=5)
        ttk.Button(btns, text="清空记忆", command=clear_all, width=12).pack(side="left", padx=5)
        ttk.Button(btns, text="导出CSV", command=export_csv, width=12).pack(side="left", padx=5)
        ttk.Button(btns, text="关闭", command=dlg.destroy, width=10).pack(side="left", padx=5)

        dlg.protocol("WM_DELETE_WINDOW", dlg.destroy)
        self.root.wait_window(dlg)

    def _choose_column_dialog(self, sheet_name: str, candidates):
        result = {"value": None}

        dlg = tk.Toplevel(self.root)
        dlg.title(f"选择费用列 - {sheet_name}")
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.resizable(False, False)

        ttk.Label(dlg, text=f"在 sheet【{sheet_name}】中识别到多个可能的费用列，\n请选择需要汇总的列：",
                  justify="left").pack(padx=16, pady=(16, 8), anchor="w")

        var = tk.IntVar(value=-1)
        from openpyxl.utils import get_column_letter
        for idx, (col, header) in enumerate(candidates):
            ttk.Radiobutton(dlg, text=f"{get_column_letter(col)} 列：{header}",
                            variable=var, value=idx).pack(anchor="w", padx=32, pady=2)
        ttk.Radiobutton(dlg, text="跳过该 sheet（不汇总）", variable=var, value=-1).pack(
            anchor="w", padx=32, pady=(2, 10))

        def on_confirm():
            idx = var.get()
            if idx is not None and idx >= 0:
                result["value"] = candidates[idx]
            dlg.destroy()

        btn_frame = ttk.Frame(dlg)
        btn_frame.pack(pady=(0, 14))
        ttk.Button(btn_frame, text="确定", command=on_confirm, width=10).pack(side="left", padx=6)
        ttk.Button(btn_frame, text="跳过此表", command=dlg.destroy, width=10).pack(side="left", padx=6)

        dlg.protocol("WM_DELETE_WINDOW", dlg.destroy)
        self.root.wait_window(dlg)
        return result["value"]


def main():
    root = TkinterDnD.Tk() if DND_AVAILABLE else tk.Tk()
    try:
        root.iconbitmap(resource_path("billing_tool.ico"))
    except Exception:
        pass
    BillingToolApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
