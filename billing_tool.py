# -*- coding: utf-8 -*-
"""
账单处理工具 —— 图形界面入口

功能：选择原始账单Excel -> 填写账单月份 -> 填写燃油费率 -> 一键生成结果
    （内部自动完成：按月筛选、整行去重、AI列空值/0值剔除、仓租/入库详情/出库订单/
     其他费用sheet求和、生成「汇总」sheet）

打包为exe：见同目录下 build_exe.bat（需在 Windows + 已安装 Python 环境下运行一次）。
"""

import os
import queue
import sys
import threading
import traceback
from datetime import date
from typing import Optional

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

import billing_core as core

def resource_path(relative_path: str) -> str:
    """获取开发环境或 PyInstaller 打包后的资源路径。"""
    if getattr(sys, "frozen", False):
        base_path = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))

    return os.path.join(base_path, relative_path)



APP_TITLE = "账单处理工具"
APP_VERSION = "v3.0.0"


class BillingToolApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(f"{APP_TITLE} {APP_VERSION}")
        self.root.geometry("760x680")
        self.root.minsize(680, 560)

        self.input_path_var = tk.StringVar()
        self.output_path_var = tk.StringVar()
        self.month_var = tk.StringVar(value=self._default_month())
        self.fuel_rate_var = tk.StringVar()
        self.weekly_mode_var = tk.BooleanVar(value=False)
        self.expected_count_var = tk.StringVar()
        self.formula_func_var = tk.StringVar(value="SUM（普通求和）")
        self.other_mode_var = tk.StringVar(value="自动识别（推荐）")
        self.status_var = tk.StringVar(value="请选择需要处理的账单Excel文件")

        # 运行模式：full=清洗+正式处理；clean=仅清洗；formal=仅正式处理
        self.run_mode_var = tk.StringVar(value="full")
        # 自动删除“仅有表头”的空sheet
        self.remove_empty_var = tk.BooleanVar(value=True)
        # 完整流程下，是否额外留档一份清洗后的文件
        self.save_cleaned_var = tk.BooleanVar(value=False)

        self._output_path_auto = True  # 输出路径是否仍由程序自动推导（未被用户手动改过）
        self._processing = False

        # ---- 后台线程相关 ----
        self._log_queue: queue.Queue = queue.Queue()
        self._worker_thread: Optional[threading.Thread] = None
        self._worker_result = None
        self._worker_error: Optional[BaseException] = None

        self._build_widgets()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ------------------------------------------------------------------
    # 界面搭建
    # ------------------------------------------------------------------
    def _default_month(self) -> str:
        today = date.today()
        return f"{today.year}-{today.month:02d}"

    def _build_widgets(self):
        pad = {"padx": 10, "pady": 6}

        # ---- 文件选择 ----
        frm_file = ttk.LabelFrame(self.root, text="第一步：选择原始账单Excel文件")
        frm_file.pack(fill="x", **pad)

        entry_file = ttk.Entry(frm_file, textvariable=self.input_path_var, state="readonly")
        entry_file.pack(side="left", fill="x", expand=True, padx=(10, 6), pady=10)
        self.browse_input_btn = ttk.Button(frm_file, text="浏览...", command=self.browse_input_file)
        self.browse_input_btn.pack(side="left", padx=(0, 10), pady=10)

        # ---- 运行模式 ----
        frm_mode = ttk.LabelFrame(self.root, text="运行模式")
        frm_mode.pack(fill="x", **pad)

        mode_row = ttk.Frame(frm_mode)
        mode_row.pack(fill="x", padx=10, pady=(8, 2))
        ttk.Radiobutton(
            mode_row, text="完整流程（清洗 + 正式处理）", value="full",
            variable=self.run_mode_var, command=self._on_mode_change,
        ).pack(side="left")
        ttk.Radiobutton(
            mode_row, text="仅清洗（输出清洗后文件，供手动补充数据）", value="clean",
            variable=self.run_mode_var, command=self._on_mode_change,
        ).pack(side="left", padx=16)

        mode_row2 = ttk.Frame(frm_mode)
        mode_row2.pack(fill="x", padx=10, pady=(0, 4))
        ttk.Radiobutton(
            mode_row2, text="仅正式处理（对已清洗、已手动补充好的文件）", value="formal",
            variable=self.run_mode_var, command=self._on_mode_change,
        ).pack(side="left")

        mode_row3 = ttk.Frame(frm_mode)
        mode_row3.pack(fill="x", padx=10, pady=(0, 8))
        ttk.Checkbutton(
            mode_row3, text="自动删除「仅有表头、无数据」的空sheet（汇总/仓租/出库订单除外）",
            variable=self.remove_empty_var,
        ).pack(side="left")

        self.save_cleaned_chk = ttk.Checkbutton(
            mode_row3, text="同时留档清洗后文件", variable=self.save_cleaned_var,
        )
        self.save_cleaned_chk.pack(side="left", padx=16)

        # ---- 参数设置 ----
        frm_params = ttk.LabelFrame(self.root, text="第二步：填写处理参数")
        frm_params.pack(fill="x", **pad)

        row1 = ttk.Frame(frm_params)
        row1.pack(fill="x", padx=10, pady=(10, 4))
        ttk.Label(row1, text="账单月份：", width=16).pack(side="left")
        ttk.Entry(row1, textvariable=self.month_var, width=14).pack(side="left")
        ttk.Label(row1, text="格式 YYYY-MM，例如 2026-06").pack(side="left", padx=8)

        row2 = ttk.Frame(frm_params)
        row2.pack(fill="x", padx=10, pady=4)
        ttk.Label(row2, text="燃油费率：", width=16).pack(side="left")
        self.fuel_rate_entry = ttk.Entry(row2, textvariable=self.fuel_rate_var, width=14)
        self.fuel_rate_entry.pack(side="left")
        self.fuel_rate_hint_var = tk.StringVar(value="支持 18.5% / 18.5 / 0.185 等写法，应用到本月所有自然周")
        ttk.Label(row2, textvariable=self.fuel_rate_hint_var).pack(side="left", padx=8)

        row2b = ttk.Frame(frm_params)
        row2b.pack(fill="x", padx=10, pady=(0, 4))
        self.weekly_chk = ttk.Checkbutton(
            row2b, text="按自然周分别设置燃油费率（点击「一键生成结果」后会弹窗按周填写）",
            variable=self.weekly_mode_var, command=self._on_weekly_mode_toggle,
        )
        self.weekly_chk.pack(side="left", padx=(122, 0))

        row3 = ttk.Frame(frm_params)
        row3.pack(fill="x", padx=10, pady=(4, 10))
        ttk.Label(row3, text="系统中订单件数：", width=16).pack(side="left")
        ttk.Entry(row3, textvariable=self.expected_count_var, width=14).pack(side="left")
        ttk.Label(row3, text="可选，用于核对清洗后数据量是否一致").pack(side="left", padx=8)

        # ---- 高级选项（默认保持即可） ----
        frm_adv = ttk.LabelFrame(self.root, text="高级选项（无特殊需求可保持默认）")
        frm_adv.pack(fill="x", **pad)

        row4 = ttk.Frame(frm_adv)
        row4.pack(fill="x", padx=10, pady=(10, 4))
        ttk.Label(row4, text="汇总函数：", width=16).pack(side="left")
        ttk.Combobox(
            row4, textvariable=self.formula_func_var, state="readonly", width=20,
            values=["SUM（普通求和）", "SUBTOTAL_109（忽略隐藏行）"],
        ).pack(side="left")

        row5 = ttk.Frame(frm_adv)
        row5.pack(fill="x", padx=10, pady=(4, 10))
        ttk.Label(row5, text="未知sheet处理方式：", width=16).pack(side="left")
        ttk.Combobox(
            row5, textvariable=self.other_mode_var, state="readonly", width=20,
            values=["自动识别（推荐）", "逐个询问", "全部跳过"],
        ).pack(side="left")

        # ---- 输出设置 ----
        frm_out = ttk.LabelFrame(self.root, text="第三步：输出文件")
        frm_out.pack(fill="x", **pad)

        entry_out = ttk.Entry(frm_out, textvariable=self.output_path_var, state="readonly")
        entry_out.pack(side="left", fill="x", expand=True, padx=(10, 6), pady=10)
        self.browse_output_btn = ttk.Button(frm_out, text="另存为...", command=self.browse_output_file)
        self.browse_output_btn.pack(side="left", padx=(0, 10), pady=10)

        # ---- 操作按钮 ----
        frm_action = ttk.Frame(self.root)
        frm_action.pack(fill="x", padx=10, pady=(4, 6))

        self.start_btn = ttk.Button(frm_action, text="一键生成结果", command=self.on_start)
        self.start_btn.pack(side="left", ipadx=16, ipady=6)

        self.open_folder_btn = ttk.Button(frm_action, text="打开输出文件夹", command=self.open_output_folder, state="disabled")
        self.open_folder_btn.pack(side="left", padx=10)

        ttk.Label(frm_action, textvariable=self.status_var, foreground="#333333").pack(side="left", padx=16)

        # ---- 日志 ----
        frm_log = ttk.LabelFrame(self.root, text="处理日志")
        frm_log.pack(fill="both", expand=True, **pad)

        self.log_text = scrolledtext.ScrolledText(frm_log, state="disabled", font=("Consolas", 9), wrap="word")
        self.log_text.pack(fill="both", expand=True, padx=6, pady=6)

    def _on_mode_change(self):
        """切换运行模式时，联动调整可用控件、按钮文案与输出文件名。"""
        mode = self.run_mode_var.get()

        if mode == "clean":
            # 仅清洗：不需要燃油费率
            self._set_fuel_widgets_state("disabled")
            self.save_cleaned_chk.configure(state="disabled")
            self.start_btn.configure(text="生成清洗后文件")
            self.status_var.set("仅清洗模式：输出清洗后文件，可手动补充数据后再用「仅正式处理」")
        elif mode == "formal":
            # 仅正式处理：不需要订单件数核对（清洗阶段才用）
            self._set_fuel_widgets_state("normal")
            self.save_cleaned_chk.configure(state="disabled")
            self.start_btn.configure(text="生成正式处理结果")
            self.status_var.set("仅正式处理模式：请选择已清洗（可含手动补充数据）的文件")
        else:
            self._set_fuel_widgets_state("normal")
            self.save_cleaned_chk.configure(state="normal")
            self.start_btn.configure(text="一键生成结果")
            self.status_var.set("完整流程：清洗 + 正式处理")

        # 输出文件名后缀随模式变化
        if self._output_path_auto and self.input_path_var.get():
            self._set_auto_output_path(self.input_path_var.get())

    def _set_fuel_widgets_state(self, state: str):
        try:
            self.fuel_rate_entry.configure(state=state)
            self.weekly_chk.configure(state=state)
        except Exception:
            pass

    def _output_suffix(self) -> str:
        mode = self.run_mode_var.get()
        if mode == "clean":
            return "_已清洗"
        return "_已处理"

    def _on_weekly_mode_toggle(self):
        if self.weekly_mode_var.get():
            self.fuel_rate_hint_var.set("将作为每周默认值；点击「一键生成结果」后可按周分别调整")
        else:
            self.fuel_rate_hint_var.set("支持 18.5% / 18.5 / 0.185 等写法，应用到本月所有自然周")

    # ------------------------------------------------------------------
    # 文件选择相关
    # ------------------------------------------------------------------
    def browse_input_file(self):
        path = filedialog.askopenfilename(
            title="选择原始账单Excel文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")],
        )
        if not path:
            return
        self.input_path_var.set(path)
        self.status_var.set("已选择文件，请填写参数后点击「一键生成结果」")

        if self._output_path_auto:
            self._set_auto_output_path(path)

    def _set_auto_output_path(self, input_path: str):
        folder = os.path.dirname(input_path)
        stem, _ext = os.path.splitext(os.path.basename(input_path))
        # 避免在已清洗文件上再次处理时出现 "_已清洗_已处理" 这种叠加后缀
        for suffix in ("_已清洗", "_已处理"):
            if stem.endswith(suffix):
                stem = stem[: -len(suffix)]
                break
        auto_path = os.path.join(folder, f"{stem}{self._output_suffix()}.xlsx")
        self.output_path_var.set(auto_path)
        self._output_path_auto = True

    def browse_output_file(self):
        default_dir = os.path.dirname(self.output_path_var.get()) or os.path.dirname(self.input_path_var.get()) or "."
        default_name = os.path.basename(self.output_path_var.get()) or "已处理结果.xlsx"
        path = filedialog.asksaveasfilename(
            title="选择结果保存位置",
            initialdir=default_dir,
            initialfile=default_name,
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
        )
        if not path:
            return
        self.output_path_var.set(path)
        self._output_path_auto = False

    def open_output_folder(self):
        out_path = self.output_path_var.get()
        folder = os.path.dirname(out_path) or "."
        try:
            if sys.platform.startswith("win"):
                os.startfile(folder)  # noqa: S606  (Windows专用)
            elif sys.platform == "darwin":
                os.system(f'open "{folder}"')
            else:
                os.system(f'xdg-open "{folder}"')
        except Exception:
            messagebox.showinfo("提示", f"结果文件位于：\n{out_path}")

    # ------------------------------------------------------------------
    # 日志输出（线程安全）
    #   后台线程 -> 只往队列里放；主线程 -> 定时取出并刷到界面。
    #   tkinter 不是线程安全的，绝不能在后台线程里直接操作控件。
    # ------------------------------------------------------------------
    def append_log(self, text: str):
        """供后台线程调用：只入队，不碰界面。"""
        self._log_queue.put(str(text))

    def _drain_log_queue(self):
        """主线程调用：把队列里积压的日志一次性刷到界面（批量插入，避免频繁重绘）。"""
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

    # ------------------------------------------------------------------
    # 跨线程交互：后台线程请求主线程弹窗，并阻塞等待用户的选择结果
    # ------------------------------------------------------------------
    def _ask_on_main_thread(self, func, *args):
        """
        在主线程执行 func(*args) 并把结果返回给调用方（后台线程）。
        后台线程会阻塞到用户操作完成为止；主线程此时正在 mainloop 中，不会死锁。
        """
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
    # 输入校验
    # ------------------------------------------------------------------
    def _parse_month(self, raw: str):
        raw = raw.strip()
        parts = raw.split("-")
        if len(parts) != 2 or not parts[0].isdigit() or not parts[1].isdigit():
            raise ValueError("账单月份格式不正确，请填写为 YYYY-MM，例如 2026-06。")
        year, month = int(parts[0]), int(parts[1])
        if not (1 <= month <= 12):
            raise ValueError("账单月份中的“月”应为 1-12 之间的数字。")
        if year < 2000 or year > 2100:
            raise ValueError("账单月份中的“年”看起来不正确，请检查。")
        return year, month

    def _parse_expected_count(self, raw: str):
        raw = raw.strip()
        if not raw:
            return None
        if not raw.isdigit():
            raise ValueError("系统中订单件数应为整数，如需跳过核对请留空。")
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
            messagebox.showwarning("提示", "请先选择原始账单Excel文件。")
            return
        if not os.path.exists(input_path):
            messagebox.showerror("错误", "所选文件不存在，请重新选择。")
            return

        output_path = self.output_path_var.get().strip()
        if not output_path:
            messagebox.showwarning("提示", "请设置结果输出路径。")
            return

        try:
            year, month = self._parse_month(self.month_var.get())
        except ValueError as e:
            messagebox.showerror("账单月份格式错误", str(e))
            return

        fuel_rate_raw = self.fuel_rate_var.get().strip()
        weekly_mode = self.weekly_mode_var.get()
        run_mode = self.run_mode_var.get()

        default_rate_for_prefill: Optional[float] = None
        uniform_fuel_rate: Optional[float] = None

        if run_mode == "clean":
            # 仅清洗阶段用不到燃油费率，跳过校验
            weekly_mode = False
        elif weekly_mode:
            # 按周模式下，主界面的费率格子是可选的"每周默认值"，用于弹窗预填，允许为空
            if fuel_rate_raw:
                try:
                    default_rate_for_prefill = core.parse_rate(fuel_rate_raw)
                except Exception:
                    messagebox.showerror("燃油费率格式错误", "请填写例如 18.5%、18.5 或 0.185 这样的格式，或留空。")
                    return
        else:
            if not fuel_rate_raw:
                messagebox.showwarning("提示", "请填写燃油费率，例如 18.5%。")
                return
            try:
                uniform_fuel_rate = core.parse_rate(fuel_rate_raw)
            except Exception:
                messagebox.showerror("燃油费率格式错误", "请填写例如 18.5%、18.5 或 0.185 这样的格式。")
                return

        try:
            expected_count = self._parse_expected_count(self.expected_count_var.get())
        except ValueError as e:
            messagebox.showerror("系统订单件数格式错误", str(e))
            return

        if os.path.abspath(output_path) == os.path.abspath(input_path):
            messagebox.showerror("错误", "输出文件不能与原始文件相同，请更换保存路径，避免覆盖原始数据。")
            return

        self._run_pipeline(
            input_path=input_path,
            output_path=output_path,
            year=year,
            month=month,
            fuel_rate=uniform_fuel_rate,
            weekly_mode=weekly_mode,
            default_rate_for_prefill=default_rate_for_prefill,
            expected_count=expected_count,
            run_mode=run_mode,
        )

    def _run_pipeline(self, input_path, output_path, year, month, fuel_rate, weekly_mode,
                      default_rate_for_prefill, expected_count, run_mode="full"):
        self._processing = True
        self.start_btn.configure(state="disabled")
        self.open_folder_btn.configure(state="disabled")
        self.browse_input_btn.configure(state="disabled")
        self.browse_output_btn.configure(state="disabled")
        self.status_var.set("处理中，请稍候...（界面可正常操作，处理在后台进行）")
        self.clear_log()

        # 清空上一轮可能残留的日志与结果
        self._log_queue = queue.Queue()
        self._worker_result = None
        self._worker_error = None

        # 弹窗类回调统一包装：后台线程调用 -> 转到主线程弹窗 -> 结果回传
        weekly_rate_callback = None
        if weekly_mode:
            def weekly_rate_callback(week_ranges):
                return self._ask_on_main_thread(self._weekly_rate_dialog, week_ranges, default_rate_for_prefill)

        def column_choice_callback(sheet_name, candidates):
            return self._ask_on_main_thread(self._choose_column_dialog, sheet_name, candidates)

        def order_count_mismatch_callback(actual, expected):
            return self._ask_on_main_thread(self._confirm_order_count_mismatch, actual, expected)

        kwargs = dict(
            input_path=input_path,
            output_path=output_path,
            year=year,
            month=month,
            fuel_rate=fuel_rate,
            weekly_rate_callback=weekly_rate_callback,
            column_choice_callback=column_choice_callback,
            order_count_mismatch_callback=order_count_mismatch_callback,
            expected_count=expected_count,
            run_mode=run_mode,
            # 以下配置项必须在主线程读好再传入：
            # tkinter 的 Variable 同样不是线程安全的，不能在后台线程调用 .get()
            remove_empty=self.remove_empty_var.get(),
            save_cleaned=self.save_cleaned_var.get(),
            formula_function=self._formula_function_value(),
            other_sheet_mode=self._other_sheet_mode_value(),
        )

        self._worker_thread = threading.Thread(target=self._worker, kwargs=kwargs, daemon=True)
        self._worker_thread.start()
        self.root.after(100, lambda: self._poll_worker(run_mode))

    def _worker(self, input_path, output_path, year, month, fuel_rate, weekly_rate_callback,
                column_choice_callback, order_count_mismatch_callback, expected_count, run_mode,
                remove_empty, save_cleaned, formula_function, other_sheet_mode):
        """后台线程实际执行处理；只通过队列输出日志，不访问界面控件或 tkinter 变量。"""
        try:
            if run_mode == "clean":
                self._worker_result = core.run_clean_only(
                    input_path=input_path,
                    year=year,
                    month=month,
                    output_path=output_path,
                    expected_order_count=expected_count,
                    remove_empty_sheets=remove_empty,
                    order_count_mismatch_callback=order_count_mismatch_callback,
                    log=self.append_log,
                )
            elif run_mode == "formal":
                self._worker_result = core.run_formal_only(
                    input_path=input_path,
                    year=year,
                    month=month,
                    output_path=output_path,
                    fuel_rate=fuel_rate,
                    weekly_rate_callback=weekly_rate_callback,
                    formula_function=formula_function,
                    other_sheet_mode=other_sheet_mode,
                    remove_empty_sheets=remove_empty,
                    column_choice_callback=column_choice_callback,
                    log=self.append_log,
                )
            else:
                cleaned_output_path = None
                if save_cleaned:
                    stem, _ext = os.path.splitext(output_path)
                    if stem.endswith("_已处理"):
                        stem = stem[: -len("_已处理")]
                    cleaned_output_path = f"{stem}_已清洗.xlsx"

                self._worker_result = core.run_full_pipeline(
                    input_path=input_path,
                    year=year,
                    month=month,
                    fuel_rate=fuel_rate,
                    weekly_rate_callback=weekly_rate_callback,
                    output_path=output_path,
                    expected_order_count=expected_count,
                    formula_function=formula_function,
                    other_sheet_mode=other_sheet_mode,
                    remove_empty_sheets=remove_empty,
                    cleaned_output_path=cleaned_output_path,
                    column_choice_callback=column_choice_callback,
                    order_count_mismatch_callback=order_count_mismatch_callback,
                    log=self.append_log,
                )
        except BaseException as exc:  # noqa: BLE001  交给主线程统一提示
            self._worker_error = exc

    def _poll_worker(self, run_mode: str):
        """主线程定时轮询：刷新日志；后台跑完后统一收尾。"""
        self._drain_log_queue()

        if self._worker_thread is not None and self._worker_thread.is_alive():
            self.root.after(100, lambda: self._poll_worker(run_mode))
            return

        self._drain_log_queue()  # 收尾再取一次，确保最后几条日志不丢
        self._finish(run_mode)

    def _finish(self, run_mode: str):
        err = self._worker_error
        try:
            if err is None:
                self._show_success(self._worker_result, run_mode)
            elif isinstance(err, core.PipelineCancelled):
                self.status_var.set("已取消：未完成燃油费率填写")
                self.append_log("处理已取消：用户未完成按周燃油费率的填写。")
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
                messagebox.showerror("处理失败", f"发生未预期的错误：\n{err}\n\n可将日志区域内容复制反馈给开发者排查。")
        finally:
            self._processing = False
            self._worker_thread = None
            self.start_btn.configure(state="normal")
            self.browse_input_btn.configure(state="normal")
            self.browse_output_btn.configure(state="normal")

    def _on_close(self):
        """处理进行中直接关窗会丢失结果，先确认。"""
        if self._processing:
            if not messagebox.askyesno("正在处理中", "账单正在处理中，现在关闭会中断处理且不会生成结果。\n\n确定要关闭吗？"):
                return
        self.root.destroy()

    def _show_success(self, result, run_mode: str):
        removed = result.get("removed_empty_sheets") or []
        removed_text = f"\n已自动删除空sheet（{len(removed)}个）：{'、'.join(removed)}" if removed else ""

        self.open_folder_btn.configure(state="normal")

        if run_mode == "clean":
            self.status_var.set(f"清洗完成！出库订单 {result['actual_order_count']} 条。")
            messagebox.showinfo(
                "清洗完成",
                f"数据清洗已完成！\n\n清洗后文件：\n{result['output_path']}\n\n"
                f"清洗后出库订单：{result['actual_order_count']} 条{removed_text}\n\n"
                f"您可以在该文件中手动补充数据，然后切换到「仅正式处理」模式继续处理。",
            )
        else:
            n = len(result["summary_items"])
            self.status_var.set(f"处理完成！共 {n} 项费用已汇总。")
            count_text = ""
            if result.get("actual_order_count") is not None:
                count_text = f"清洗后出库订单：{result['actual_order_count']} 条\n"
            messagebox.showinfo(
                "处理完成",
                f"账单处理已完成！\n\n结果文件：\n{result['output_path']}\n\n"
                f"{count_text}共汇总 {n} 项费用，明细见「汇总」sheet。{removed_text}",
            )

    # ------------------------------------------------------------------
    # 交互回调：订单件数核对不一致 / 未知sheet多候选列选择
    # ------------------------------------------------------------------
    def _confirm_order_count_mismatch(self, actual: int, expected: int) -> bool:
        diff = actual - expected
        msg = (
            f"系统中订单件数为 {expected}，清洗后【出库订单】总数据量为 {actual}，相差 {diff} 条。\n\n"
            f"是否仍要继续生成正式处理结果？\n"
            f"（选择“否”将中止本次处理，请检查数据后重新运行）"
        )
        return messagebox.askyesno("订单件数核对不一致", msg)

    def _weekly_rate_dialog(self, week_ranges, default_rate: Optional[float]):
        """
        弹窗让用户按自然周分别填写燃油费率。
        week_ranges: [(start_date, end_date), ...]，均为 datetime.date
        default_rate: 主界面费率格子里的值（小数），用于预填，可为 None
        返回 {周一日期: 费率(小数)} 字典；用户取消则返回 None。
        """
        result = {"value": None}
        default_text = f"{default_rate * 100:.2f}%" if default_rate is not None else ""

        dlg = tk.Toplevel(self.root)
        dlg.title("按自然周填写燃油费率")
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.resizable(False, False)

        ttk.Label(
            dlg,
            text=f"本月出库订单覆盖以下 {len(week_ranges)} 个自然周，请分别填写燃油费率：\n（支持 18.5% / 18.5 / 0.185 等写法）",
            justify="left",
        ).pack(padx=16, pady=(16, 8), anchor="w")

        rows_frame = ttk.Frame(dlg)
        rows_frame.pack(padx=16, pady=4, fill="x")

        entry_vars = []
        for idx, (start, end) in enumerate(week_ranges, start=1):
            row = ttk.Frame(rows_frame)
            row.pack(fill="x", pady=2)
            ttk.Label(row, text=f"第{idx}周（{start:%Y-%m-%d} 至 {end:%Y-%m-%d}）：", width=32).pack(side="left")
            var = tk.StringVar(value=default_text)
            ttk.Entry(row, textvariable=var, width=12).pack(side="left")
            entry_vars.append((start, var))

        error_label = ttk.Label(dlg, text="", foreground="#cc0000")
        error_label.pack(padx=16, anchor="w")

        def on_confirm():
            rates = {}
            for start, var in entry_vars:
                raw = var.get().strip()
                if not raw:
                    error_label.configure(text="请填写每一周的燃油费率，不能留空。")
                    return
                try:
                    rates[start] = core.parse_rate(raw)
                except Exception:
                    error_label.configure(text=f"「{raw}」格式不正确，请填写例如 18.5%、18.5 或 0.185。")
                    return
            result["value"] = rates
            dlg.destroy()

        def on_cancel():
            result["value"] = None
            dlg.destroy()

        btn_frame = ttk.Frame(dlg)
        btn_frame.pack(pady=(6, 16))
        ttk.Button(btn_frame, text="确定", command=on_confirm, width=10).pack(side="left", padx=6)
        ttk.Button(btn_frame, text="取消", command=on_cancel, width=10).pack(side="left", padx=6)

        dlg.protocol("WM_DELETE_WINDOW", on_cancel)
        self.root.wait_window(dlg)
        return result["value"]

    def _choose_column_dialog(self, sheet_name: str, candidates):
        result = {"value": None}

        dlg = tk.Toplevel(self.root)
        dlg.title(f"选择费用列 - {sheet_name}")
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.resizable(False, False)

        ttk.Label(
            dlg,
            text=f"在 sheet【{sheet_name}】中识别到多个可能的费用列，\n请选择需要汇总的列：",
            justify="left",
        ).pack(padx=16, pady=(16, 8), anchor="w")

        var = tk.IntVar(value=-1)
        for idx, (col, header) in enumerate(candidates):
            from openpyxl.utils import get_column_letter
            ttk.Radiobutton(
                dlg, text=f"{get_column_letter(col)} 列：{header}", variable=var, value=idx,
            ).pack(anchor="w", padx=32, pady=2)
        ttk.Radiobutton(dlg, text="跳过该 sheet（不汇总）", variable=var, value=-1).pack(anchor="w", padx=32, pady=(2, 10))

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
    root = tk.Tk()

    try:
        root.iconbitmap(resource_path("billing_tool.ico"))
    except Exception:
        pass


    app = BillingToolApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
