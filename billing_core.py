# -*- coding: utf-8 -*-
"""
billing_core.py —— 账单处理核心逻辑

本模块把原来两份脚本（数据清洗 + 正式处理）合并、改写为不依赖命令行 input() 的函数，
供 GUI（billing_tool.py）调用，也可以单独 import 后用于自动化测试。

核心流程 run_full_pipeline()：
  第一阶段（清洗）：
    1. 按月份筛选【仓租】【出库订单】
    2. 整行去重
    3. 剔除【出库订单】AI列（订单处理费）为空/为0的记录
    4.（可选）与"系统中订单件数"核对
  第二阶段（正式处理）：
    2) 仓租 R列「实际费用」求和
    2) 入库详情 K列「费用」求和
    3) 出库订单：按统一燃油费率写入 AZ/AL/AY 公式并求和
    4) 其他 sheet（工单/索赔抵扣/核账补收/退件订单/退款详情/未知sheet）求和
    5) 生成/更新「汇总」sheet
"""

import calendar
import os
import re
import tempfile
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from typing import Callable, Dict, Iterable, List, Optional, Sequence, Tuple

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

# ------------------------- 基础配置 -------------------------

SHEET_WAREHOUSE_RENT = "仓租"
SHEET_INBOUND_DETAIL = "入库详情"
SHEET_INBOUND_ORDER = "入库订单"
SHEET_OUTBOUND_ORDER = "出库订单"
SHEET_FUEL_RATE = "燃油费率表"
SHEET_OTHER_SUMMARY = "其他费用汇总"
DEFAULT_MAIN_SUMMARY_SHEET = "汇总"

# ---- 防重复计算：某些 sheet 与另一个 sheet 是“汇总 与 明细”的关系，是同一笔钱 ----
# 例如【入库订单】是【入库详情】的汇总表，两者金额相同；
# 若两个都汇总进「汇总」sheet，费用总计会把这笔钱算两次。
# 格式：{被跳过的sheet: (已计入该金额的sheet, 跳过原因)}
# 注意：只有当“已计入的sheet”确实被处理过时才跳过；
#       否则（例如某客户只有【入库订单】而没有【入库详情】）仍照常汇总，避免漏算。
SHEETS_ALREADY_COUNTED = {
    SHEET_INBOUND_ORDER: (
        SHEET_INBOUND_DETAIL,
        "与【入库详情】为同一笔钱（汇总表 vs 明细表），已由【入库详情】计入，跳过以避免重复计算。",
    ),
}

AI_COLUMN_LETTER = "AI"          # 出库订单：订单处理费（清洗阶段用于判断是否为空，见下方说明）
RENT_DATE_COL = "时间"            # 仓租：判断月份用的字段
OUTBOUND_DATE_COL = "创建时间"     # 出库订单：判断月份用的字段

# ---- 正式处理阶段：全部改为“按表头文字定位列”，不再依赖固定列字母 ----
# 不同客户的账单列顺序/列数会不一样（例如有的客户没有“自有运单计费/复合订单处理费”等列），
# 因此这里只按表头名匹配，且只汇总实际存在的列。

# 仓租 / 入库详情 汇总列（按表头查找，找不到时回退到原固定列）
RENT_AMOUNT_HEADER = "实际费用"
RENT_AMOUNT_FALLBACK_LETTER = "R"
INBOUND_AMOUNT_HEADER = "费用"
INBOUND_AMOUNT_FALLBACK_LETTER = "K"

# 出库订单相关表头
OUTBOUND_DATE_HEADER = "创建时间"           # 用于匹配燃油费率所在自然周
OUTBOUND_LOGISTICS_HEADER = "物流费用"       # 会被重写为公式：-(燃油基数 + 实际燃油费)
OUTBOUND_TOTAL_HEADER = "费用总计"           # 会被重写为公式：处理费类 + 物流费用
OUTBOUND_FUEL_HEADER = "实际燃油费"          # 新建列：燃油基数 × 燃油费率

# 进入“费用总计”的处理费类列（存在几列就加几列）
OUTBOUND_PROCESSING_FEE_HEADERS = ["订单处理费", "自有运单计费", "复合订单处理费"]
# 计算燃油费的基数列（存在几列就加几列）；注意原始“燃油附加费”列不参与，燃油费按新费率重算
OUTBOUND_FREIGHT_BASE_HEADERS = [
    "住宅地址附加费", "基础运费", "操作附加费-尺寸", "偏远附加费", "超尺寸附加费", "操作附加费",
]

# 固定 sheet：只允许汇总指定表头，避免误算其他金额列
FIXED_OTHER_SHEET_RULES = {
    "工单": {"target_header": "费用总计", "display_name": "工单处理费"},
    "索赔抵扣": {"target_header": "索赔结果", "display_name": "索赔抵扣"},
    "核账补收": {"target_header": "补收金额", "display_name": "核账补收费"},
    "退件订单": {"target_header": "费用总计", "display_name": "退件订单"},
    "退款详情": {"target_header": "费用", "display_name": "退款详情"},
}

# 未知 sheet 的通用候选关键词
UNKNOWN_AMOUNT_KEYWORDS = [
    "费用总计", "总费用", "实际费用", "费用", "金额", "收费金额", "账单金额",
    "应收金额", "应付金额", "服务费", "操作费", "运费", "补收", "退款", "赔偿", "索赔",
]

# 汇总 sheet 的显示顺序；未列入的未知项目会排在后面
MAIN_SUMMARY_DISPLAY_ORDER = [
    "出库订单", "入库订单", "工单处理费", "索赔抵扣", "仓租费", "核账补收费", "退件订单", "退款详情",
]

LogFunc = Callable[[str], None]

# ---- “仅有表头的空sheet”自动清除 ----
# 以下 sheet 即使为空也绝不删除：
#   汇总 —— 第2~6行是客户编号/昵称/账单币种等关键信息，且是最终输出目标
#   仓租 / 出库订单 —— 清洗流程强制依赖，删掉会直接报“缺少必要的sheet”
#   （例如 IPQI 的【仓租】本身就是仅有表头，必须保留）
PROTECTED_SHEETS_FROM_REMOVAL = {
    DEFAULT_MAIN_SUMMARY_SHEET,
    SHEET_WAREHOUSE_RENT,
    SHEET_OUTBOUND_ORDER,
    SHEET_FUEL_RATE,
    SHEET_OTHER_SUMMARY,
}


def _noop_log(_text: str) -> None:
    return None


class OrderCountMismatchError(Exception):
    """系统订单件数与清洗后数据量不一致，且用户选择不继续。"""


class PipelineCancelled(Exception):
    """用户在交互环节（如按周填写燃油费率）主动取消。"""


class BillingProcessError(Exception):
    """账单处理过程中的可预期错误（会以友好文案展示给用户）。"""


@dataclass
class FeeSummaryItem:
    display_name: str
    sheet_name: str
    total_cell: str
    amount_header: str = ""
    note: str = ""


# ------------------------- 通用工具函数 -------------------------

def norm_text(value) -> str:
    if value is None:
        return ""
    return str(value).strip().replace("\n", "").replace("\r", "").replace(" ", "")


def quote_sheet_name(sheet_name: str) -> str:
    return "'" + sheet_name.replace("'", "''") + "'"


def is_blank(value) -> bool:
    return value is None or str(value).strip() == ""


def is_summary_or_archive_sheet(sheet_name: str, main_summary_sheet: str) -> bool:
    if sheet_name in {main_summary_sheet, SHEET_FUEL_RATE, SHEET_OTHER_SUMMARY}:
        return True
    archive_suffixes = [
        "_重复数据", "_月份外已删除", "_AI列为空已删除", "_已删除", "重复数据", "月份外已删除",
    ]
    return any(sheet_name.endswith(suffix) for suffix in archive_suffixes)


def find_header_col(ws: Worksheet, header_name: str, header_row: int = 1) -> Optional[int]:
    target = norm_text(header_name)
    for cell in ws[header_row]:
        if norm_text(cell.value) == target:
            return cell.column
    return None


def find_header_col_by_keywords(ws: Worksheet, keywords: Sequence[str], header_row: int = 1) -> List[Tuple[int, str]]:
    candidates = []
    for cell in ws[header_row]:
        header = norm_text(cell.value)
        if not header:
            continue
        if any(keyword in header for keyword in keywords):
            candidates.append((cell.column, header))
    return candidates


def find_last_data_row(ws: Worksheet, key_cols: Optional[Sequence[int]] = None, start_row: int = 2) -> int:
    total_labels = {"合计", "总计", "费用合计", "费用总计", "实际费用合计", "索赔结果合计", "补收金额合计"}
    max_row = ws.max_row

    for row in range(max_row, start_row - 1, -1):
        row_values = [norm_text(ws.cell(row, col).value) for col in range(1, min(ws.max_column, 8) + 1)]
        if any(v in total_labels or v.endswith("合计") for v in row_values):
            continue

        if key_cols:
            for col in key_cols:
                if col <= ws.max_column and not is_blank(ws.cell(row, col).value):
                    return row
        else:
            for col in range(1, ws.max_column + 1):
                if not is_blank(ws.cell(row, col).value):
                    return row

    return start_row - 1


def remove_existing_total_rows(ws: Worksheet, amount_col: int, label_texts: Sequence[str], header_row: int = 1) -> None:
    label_col = max(1, amount_col - 1)
    labels = {norm_text(x) for x in label_texts}
    initial_max_row = ws.max_row
    tail_start_row = max(header_row + 1, initial_max_row - 5)

    for row in range(initial_max_row, header_row, -1):
        label_value = norm_text(ws.cell(row, label_col).value)
        amount_formula = norm_text(ws.cell(row, amount_col).value)

        if label_value in labels:
            ws.delete_rows(row, 1)
            continue

        a_value = norm_text(ws.cell(row, 1).value)
        if a_value in labels:
            ws.delete_rows(row, 1)
            continue

        if (
            row >= tail_start_row
            and (amount_formula.startswith("=SUM(") or amount_formula.startswith("=SUBTOTAL("))
            and ("合计" in label_value or "总计" in label_value or "合计" in a_value or "总计" in a_value)
        ):
            ws.delete_rows(row, 1)


def make_sum_formula(col_letter: str, start_row: int, end_row: int, formula_function: str = "SUM") -> str:
    if end_row < start_row:
        return "=0"
    if formula_function.upper() == "SUBTOTAL_109":
        return f"=SUBTOTAL(109,{col_letter}{start_row}:{col_letter}{end_row})"
    return f"=SUM({col_letter}{start_row}:{col_letter}{end_row})"


def write_total_formula(
    ws: Worksheet,
    amount_col: int,
    total_label: str,
    formula_function: str = "SUM",
    start_row: int = 2,
    key_cols: Optional[Sequence[int]] = None,
) -> str:
    remove_existing_total_rows(ws, amount_col, [total_label])
    amount_letter = get_column_letter(amount_col)
    last_data_row = find_last_data_row(ws, key_cols=key_cols or [amount_col], start_row=start_row)
    total_row = max(start_row, last_data_row + 1)

    label_col = max(1, amount_col - 1)
    ws.cell(total_row, label_col).value = total_label
    ws.cell(total_row, amount_col).value = make_sum_formula(amount_letter, start_row, last_data_row, formula_function)
    ws.cell(total_row, label_col).font = Font(bold=True)
    ws.cell(total_row, amount_col).font = Font(bold=True)
    ws.cell(total_row, amount_col).number_format = '#,##0.00;[Red]-#,##0.00'
    return f"{amount_letter}{total_row}"


def parse_rate(raw: str) -> float:
    """
    支持输入：18.5% / 18.5 / 0.185，统一返回小数：0.185
    """
    text = str(raw).strip().replace("％", "%")
    if not text:
        raise ValueError("燃油费率不能为空")
    if text.endswith("%"):
        return float(text[:-1].strip()) / 100
    value = float(text)
    if value > 1:
        return value / 100
    return value


def coerce_excel_date(value) -> Optional[date]:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value

    text = str(value).strip()
    if not text:
        return None

    known_formats = (
        "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S", "%Y-%m-%d", "%Y/%m/%d",
        "%m/%d/%Y %H:%M:%S", "%m/%d/%Y",
    )
    for fmt in known_formats:
        try:
            return datetime.strptime(text[:19], fmt).date()
        except ValueError:
            pass

    try:
        from dateutil.parser import parse
        return parse(text).date()
    except Exception:
        return None


def build_week_ranges(year: int, month: int, extra_dates: Optional[Iterable[date]] = None) -> List[Tuple[date, date]]:
    month_start = date(year, month, 1)
    month_end = date(year, month, calendar.monthrange(year, month)[1])

    end_date = month_end
    if extra_dates:
        valid_dates = [d for d in extra_dates if d is not None]
        if valid_dates:
            end_date = max(end_date, max(valid_dates))

    first_monday = month_start - timedelta(days=month_start.weekday())
    last_sunday = end_date + timedelta(days=(6 - end_date.weekday()))

    ranges = []
    cur = first_monday
    while cur <= last_sunday:
        ranges.append((cur, cur + timedelta(days=6)))
        cur += timedelta(days=7)
    return ranges


def write_fuel_rate_sheet(wb, rates: Sequence[Tuple[date, date, float]]) -> None:
    if SHEET_FUEL_RATE in wb.sheetnames:
        del wb[SHEET_FUEL_RATE]

    ws = wb.create_sheet(SHEET_FUEL_RATE)
    headers = ["周序号", "开始日期", "结束日期", "燃油费率"]
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(1, col)
        cell.value = header
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    for row, (start, end, rate) in enumerate(rates, start=2):
        ws.cell(row, 1).value = row - 1
        ws.cell(row, 2).value = start
        ws.cell(row, 3).value = end
        ws.cell(row, 4).value = rate
        ws.cell(row, 2).number_format = "yyyy-mm-dd"
        ws.cell(row, 3).number_format = "yyyy-mm-dd"
        ws.cell(row, 4).number_format = "0.00%"

    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 15
    ws.column_dimensions["C"].width = 15
    ws.column_dimensions["D"].width = 12


# ------------------------- 第一阶段：数据清洗 -------------------------

def dataframe_is_empty(df: pd.DataFrame) -> bool:
    """
    判断一个 sheet 是否“仅有表头 / 完全没有有效数据”。
    整行全空的记录不算数据（有些导出文件会带若干空白行）。
    """
    if df is None or len(df.columns) == 0:
        return True
    return len(df.dropna(how="all")) == 0


def drop_header_only_sheets(
    all_sheets: Dict[str, pd.DataFrame],
    protected: Optional[Iterable[str]] = None,
    log: LogFunc = _noop_log,
) -> Tuple[Dict[str, pd.DataFrame], List[str]]:
    """
    删除“仅有表头、没有任何数据行”的 sheet（原来需要人工手动删）。

    保护名单内的 sheet 即使为空也保留，避免破坏必需结构。
    返回 (保留下来的sheets, 被删除的sheet名列表)。
    """
    protected_set = set(protected) if protected is not None else set(PROTECTED_SHEETS_FROM_REMOVAL)

    kept: Dict[str, pd.DataFrame] = {}
    removed: List[str] = []

    for name, df in all_sheets.items():
        if dataframe_is_empty(df):
            if name in protected_set:
                log(f"  [保留] 【{name}】仅有表头，但属于必需sheet，保留（合计将记为0）。")
                kept[name] = df
            else:
                removed.append(name)
        else:
            kept[name] = df

    if removed:
        log(f"[空sheet清除] 以下 {len(removed)} 个sheet仅有表头、无数据，已自动删除：{'、'.join(removed)}")
    else:
        log("[空sheet清除] 未发现仅有表头的空sheet。")

    return kept, removed


def worksheet_is_header_only(ws: Worksheet) -> bool:
    """openpyxl 层面判断 sheet 是否仅有表头（供“仅正式处理”模式使用，不做 pandas 往返以免破坏手工内容）。"""
    for row in range(2, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            if not is_blank(ws.cell(row, col).value):
                return False
    return True


def drop_header_only_worksheets(
    wb,
    protected: Optional[Iterable[str]] = None,
    log: LogFunc = _noop_log,
) -> List[str]:
    """在已加载的工作簿上直接删除仅有表头的 sheet，保留其余所有内容与格式。"""
    protected_set = set(protected) if protected is not None else set(PROTECTED_SHEETS_FROM_REMOVAL)
    removed: List[str] = []

    for name in list(wb.sheetnames):
        if name in protected_set:
            continue
        if is_summary_or_archive_sheet(name, DEFAULT_MAIN_SUMMARY_SHEET):
            continue
        if worksheet_is_header_only(wb[name]):
            del wb[name]
            removed.append(name)

    if removed:
        log(f"[空sheet清除] 以下 {len(removed)} 个sheet仅有表头、无数据，已自动删除：{'、'.join(removed)}")
    else:
        log("[空sheet清除] 未发现仅有表头的空sheet。")
    return removed


def get_column_name_by_letter(df: pd.DataFrame, letter: str) -> str:
    idx = column_index_from_string(letter) - 1
    if idx >= len(df.columns):
        raise BillingProcessError(f"列 {letter} 超出了sheet的实际列数，请检查表结构是否变化。")
    return df.columns[idx]


def filter_by_month(df: pd.DataFrame, date_col: str, year: int, month: int, sheet_name: str, log: LogFunc = _noop_log):
    if date_col not in df.columns:
        raise BillingProcessError(f"【{sheet_name}】sheet 中没有找到表头【{date_col}】，请检查表结构。")

    parsed_dates = pd.to_datetime(df[date_col], errors="coerce")
    is_target_month = (parsed_dates.dt.year == year) & (parsed_dates.dt.month == month)

    df_removed = df[~is_target_month].copy()
    df_clean = df[is_target_month].copy()

    if len(df_removed) > 0:
        df_removed.insert(0, "删除原因", f"非{year}-{month:02d}数据（按{date_col}判断）")
        log(f"[{sheet_name}] 非{year}-{month:02d}的数据共 {len(df_removed)} 条，已删除。")
    else:
        log(f"[{sheet_name}] 全部数据均属于{year}-{month:02d}，无需删除。")

    if parsed_dates.isna().any():
        n_invalid = parsed_dates.isna().sum()
        log(f"  提示：[{sheet_name}] 中有 {n_invalid} 条记录的{date_col}无法识别为日期，已一并计入删除范围。")

    return df_clean, df_removed


def dedup_sheet(df: pd.DataFrame, sheet_name: str, log: LogFunc = _noop_log):
    is_dup = df.duplicated(keep="first")
    df_dup = df[is_dup].copy()
    df_clean = df[~is_dup].copy()

    if len(df_dup) > 0:
        df_dup.insert(0, "删除原因", "重复数据")
        log(f"[{sheet_name}] 发现 {len(df_dup)} 条重复记录，已删除。")
    else:
        log(f"[{sheet_name}] 未发现重复记录。")

    return df_clean, df_dup


def remove_empty_ai_rows(df: pd.DataFrame, ai_column_letter: str = AI_COLUMN_LETTER, log: LogFunc = _noop_log):
    # 优先按表头「订单处理费」定位，找不到再回退到固定列字母（AI）
    processing_header = OUTBOUND_PROCESSING_FEE_HEADERS[0]  # "订单处理费"
    if processing_header in df.columns:
        ai_col = processing_header
    else:
        ai_col = get_column_name_by_letter(df, ai_column_letter)
        log(f"  [{SHEET_OUTBOUND_ORDER}] 未找到表头「{processing_header}」，回退使用固定列 {ai_column_letter} 判空。")
    ai_numeric = pd.to_numeric(df[ai_col], errors="coerce")

    is_empty = df[ai_col].isna() | (df[ai_col].astype(str).str.strip() == "")
    is_zero = ai_numeric == 0
    is_removed = is_empty | is_zero

    df_removed = df[is_removed].copy()
    df_clean = df[~is_removed].copy()

    if len(df_removed) > 0:
        reason = pd.Series("", index=df_removed.index)
        reason[df_removed[ai_col].isna() | (df_removed[ai_col].astype(str).str.strip() == "")] = f"AI列（{ai_col}）为空"
        reason[reason == ""] = f"AI列（{ai_col}）为0"
        df_removed.insert(0, "删除原因", reason)
        log(f"[{SHEET_OUTBOUND_ORDER}] AI列（{ai_col}）为空或为0的记录共 {len(df_removed)} 条，已删除。")
    else:
        log(f"[{SHEET_OUTBOUND_ORDER}] AI列（{ai_col}）无空值/0值记录。")

    return df_clean, df_removed


def clean_billing_data(all_sheets: Dict[str, pd.DataFrame], year: int, month: int, log: LogFunc = _noop_log) -> dict:
    """执行第一阶段清洗，返回 {"sheets": {...}, "outbound_clean_count": int}"""
    if SHEET_WAREHOUSE_RENT not in all_sheets or SHEET_OUTBOUND_ORDER not in all_sheets:
        raise BillingProcessError(f"文件中缺少必要的sheet：{SHEET_WAREHOUSE_RENT} 或 {SHEET_OUTBOUND_ORDER}")

    df_rent = all_sheets[SHEET_WAREHOUSE_RENT]
    df_rent_month, df_rent_month_removed = filter_by_month(df_rent, RENT_DATE_COL, year, month, SHEET_WAREHOUSE_RENT, log=log)

    df_outbound = all_sheets[SHEET_OUTBOUND_ORDER]
    df_outbound_month, df_outbound_month_removed = filter_by_month(df_outbound, OUTBOUND_DATE_COL, year, month, SHEET_OUTBOUND_ORDER, log=log)

    df_rent_clean, df_rent_dup = dedup_sheet(df_rent_month, SHEET_WAREHOUSE_RENT, log=log)
    df_outbound_clean, df_outbound_dup = dedup_sheet(df_outbound_month, SHEET_OUTBOUND_ORDER, log=log)

    df_outbound_clean, df_outbound_empty_ai = remove_empty_ai_rows(df_outbound_clean, log=log)

    log("========== 第一阶段清洗结果汇总 ==========")
    log(
        f"{SHEET_WAREHOUSE_RENT}：原始 {len(df_rent)} 条 -> "
        f"{year}-{month:02d}月份 {len(df_rent_month)} 条（月份外剔除 {len(df_rent_month_removed)} 条） -> "
        f"去重后 {len(df_rent_clean)} 条（重复剔除 {len(df_rent_dup)} 条）"
    )
    log(
        f"{SHEET_OUTBOUND_ORDER}：原始 {len(df_outbound)} 条 -> "
        f"{year}-{month:02d}月份 {len(df_outbound_month)} 条（月份外剔除 {len(df_outbound_month_removed)} 条） -> "
        f"去重后 {len(df_outbound_month) - len(df_outbound_dup)} 条（重复剔除 {len(df_outbound_dup)} 条） -> "
        f"最终 {len(df_outbound_clean)} 条（AI列为空/为0剔除 {len(df_outbound_empty_ai)} 条）"
    )

    output_sheets = dict(all_sheets)
    output_sheets[SHEET_WAREHOUSE_RENT] = df_rent_clean
    output_sheets[SHEET_OUTBOUND_ORDER] = df_outbound_clean

    if len(df_rent_month_removed) > 0:
        output_sheets[f"{SHEET_WAREHOUSE_RENT}_月份外已删除"] = df_rent_month_removed
    if len(df_outbound_month_removed) > 0:
        output_sheets[f"{SHEET_OUTBOUND_ORDER}_月份外已删除"] = df_outbound_month_removed
    if len(df_rent_dup) > 0:
        output_sheets[f"{SHEET_WAREHOUSE_RENT}_重复数据"] = df_rent_dup
    if len(df_outbound_dup) > 0:
        output_sheets[f"{SHEET_OUTBOUND_ORDER}_重复数据"] = df_outbound_dup
    if len(df_outbound_empty_ai) > 0:
        output_sheets[f"{SHEET_OUTBOUND_ORDER}_AI列为空已删除"] = df_outbound_empty_ai

    return {"sheets": output_sheets, "outbound_clean_count": len(df_outbound_clean)}


def compute_outbound_week_ranges(output_sheets: Dict[str, pd.DataFrame], year: int, month: int) -> List[Tuple[date, date]]:
    """
    基于清洗后的【出库订单】数据，计算本月实际覆盖到的自然周区间。
    供 GUI 在弹出"按周填写燃油费率"对话框前预先获知需要填写哪些周。
    """
    df = output_sheets.get(SHEET_OUTBOUND_ORDER)
    if df is None or OUTBOUND_DATE_COL not in df.columns:
        return build_week_ranges(year, month)

    dates: List[date] = []
    for value in df[OUTBOUND_DATE_COL]:
        d = coerce_excel_date(value)
        if d:
            dates.append(d)
    return build_week_ranges(year, month, dates)


# ------------------------- 第二部分：仓租、入库详情 -------------------------

def process_warehouse_rent(wb, formula_function: str, log: LogFunc = _noop_log) -> Optional[FeeSummaryItem]:
    if SHEET_WAREHOUSE_RENT not in wb.sheetnames:
        log(f"[跳过] 未找到 sheet：{SHEET_WAREHOUSE_RENT}")
        return None

    ws = wb[SHEET_WAREHOUSE_RENT]
    amount_col = find_header_col(ws, RENT_AMOUNT_HEADER)
    if amount_col:
        log(f"  [{SHEET_WAREHOUSE_RENT}] 定位到「{RENT_AMOUNT_HEADER}」列：{get_column_letter(amount_col)}")
    else:
        amount_col = column_index_from_string(RENT_AMOUNT_FALLBACK_LETTER)
        log(f"  [{SHEET_WAREHOUSE_RENT}] 未找到表头「{RENT_AMOUNT_HEADER}」，回退使用固定列 {RENT_AMOUNT_FALLBACK_LETTER}。请核对该sheet表头是否变化。")
    total_cell = write_total_formula(ws, amount_col=amount_col, total_label="实际费用合计", formula_function=formula_function, key_cols=[amount_col])
    log(f"[完成] {SHEET_WAREHOUSE_RENT}!{total_cell} 已写入实际费用汇总。")
    return FeeSummaryItem("仓租费", SHEET_WAREHOUSE_RENT, total_cell, RENT_AMOUNT_HEADER)


def process_inbound_detail(wb, formula_function: str, log: LogFunc = _noop_log) -> Optional[FeeSummaryItem]:
    if SHEET_INBOUND_DETAIL not in wb.sheetnames:
        log(f"[跳过] 未找到 sheet：{SHEET_INBOUND_DETAIL}")
        return None

    ws = wb[SHEET_INBOUND_DETAIL]
    amount_col = find_header_col(ws, INBOUND_AMOUNT_HEADER)
    if amount_col:
        log(f"  [{SHEET_INBOUND_DETAIL}] 定位到「{INBOUND_AMOUNT_HEADER}」列：{get_column_letter(amount_col)}")
    else:
        amount_col = column_index_from_string(INBOUND_AMOUNT_FALLBACK_LETTER)
        log(f"  [{SHEET_INBOUND_DETAIL}] 未找到表头「{INBOUND_AMOUNT_HEADER}」，回退使用固定列 {INBOUND_AMOUNT_FALLBACK_LETTER}。请核对该sheet表头是否变化。")
    total_cell = write_total_formula(ws, amount_col=amount_col, total_label="费用合计", formula_function=formula_function, key_cols=[amount_col])
    log(f"[完成] {SHEET_INBOUND_DETAIL}!{total_cell} 已写入费用汇总。")
    return FeeSummaryItem("入库订单", SHEET_INBOUND_DETAIL, total_cell, INBOUND_AMOUNT_HEADER)


# ------------------------- 第三部分：出库订单（统一燃油费率） -------------------------

def process_outbound_order(
    wb,
    formula_function: str,
    year: int,
    month: int,
    fuel_rate: Optional[float] = None,
    weekly_rates: Optional[Dict[date, float]] = None,
    log: LogFunc = _noop_log,
) -> Optional[FeeSummaryItem]:
    """
    燃油费率支持两种模式：
      - 统一费率：传 fuel_rate（一个小数，应用到本月覆盖的所有自然周）
      - 按自然周分别设置：传 weekly_rates（{周一日期: 费率} 字典），优先于 fuel_rate 生效
    """
    if SHEET_OUTBOUND_ORDER not in wb.sheetnames:
        log(f"[跳过] 未找到 sheet：{SHEET_OUTBOUND_ORDER}")
        return None

    ws = wb[SHEET_OUTBOUND_ORDER]

    # ---- 全部按表头文字定位列，兼容不同客户的列顺序/列数差异 ----
    date_col = find_header_col(ws, OUTBOUND_DATE_HEADER)
    logistics_col = find_header_col(ws, OUTBOUND_LOGISTICS_HEADER)
    total_col = find_header_col(ws, OUTBOUND_TOTAL_HEADER)

    proc_cols = [(h, c) for h in OUTBOUND_PROCESSING_FEE_HEADERS if (c := find_header_col(ws, h))]
    base_cols = [(h, c) for h in OUTBOUND_FREIGHT_BASE_HEADERS if (c := find_header_col(ws, h))]

    # 必要表头缺失则明确报错（友好提示到底缺了哪个）
    missing = []
    if not date_col:
        missing.append(OUTBOUND_DATE_HEADER)
    if not total_col:
        missing.append(OUTBOUND_TOTAL_HEADER)
    if not logistics_col:
        missing.append(OUTBOUND_LOGISTICS_HEADER)
    if not base_cols:
        missing.append("燃油基数列（住宅地址附加费/基础运费/操作附加费 等，至少需要一列）")
    if missing:
        raise BillingProcessError(
            f"【{SHEET_OUTBOUND_ORDER}】缺少必要表头：{'、'.join(missing)}。请核对该客户的账单列结构。"
        )

    # 新建“实际燃油费”列：优先放在“费用总计”右侧一列；若该位置已被占用则追加到末尾
    fuel_col = total_col + 1
    existing_header = norm_text(ws.cell(1, fuel_col).value)
    if existing_header and existing_header != norm_text(OUTBOUND_FUEL_HEADER):
        fuel_col = ws.max_column + 1

    # 写表头（物流费用/费用总计通常已存在，这里统一确保表头文字规范）
    ws.cell(1, fuel_col).value = OUTBOUND_FUEL_HEADER
    ws.cell(1, total_col).value = OUTBOUND_TOTAL_HEADER
    ws.cell(1, logistics_col).value = OUTBOUND_LOGISTICS_HEADER

    # 预先算好各列字母，供公式引用
    fuel_letter = get_column_letter(fuel_col)
    total_letter = get_column_letter(total_col)
    logistics_letter = get_column_letter(logistics_col)
    base_letters = [get_column_letter(c) for _h, c in base_cols]
    proc_letters = [get_column_letter(c) for _h, c in proc_cols]

    log(
        f"  [{SHEET_OUTBOUND_ORDER}] 列定位："
        f"创建时间={get_column_letter(date_col)}，物流费用={logistics_letter}，费用总计={total_letter}，"
        f"实际燃油费(新建)={fuel_letter}；"
        f"处理费类={'/'.join(f'{h}({get_column_letter(c)})' for h, c in proc_cols) or '无'}；"
        f"燃油基数={'/'.join(f'{h}({get_column_letter(c)})' for h, c in base_cols)}"
    )

    remove_existing_total_rows(ws, total_col, ["费用总计合计", "费用总计"])

    key_cols = [date_col, total_col, logistics_col] + [c for _h, c in proc_cols] + [c for _h, c in base_cols]
    last_data_row = find_last_data_row(ws, key_cols=key_cols, start_row=2)

    if last_data_row < 2:
        total_cell = write_total_formula(ws, total_col, "费用总计合计", formula_function, key_cols=[total_col])
        log(f"[提示] {SHEET_OUTBOUND_ORDER} 没有订单数据。")
        return FeeSummaryItem("出库订单", SHEET_OUTBOUND_ORDER, total_cell, OUTBOUND_TOTAL_HEADER, "无订单数据")

    outbound_dates: List[date] = []
    invalid_date_rows: List[int] = []
    for row in range(2, last_data_row + 1):
        order_date = coerce_excel_date(ws.cell(row, date_col).value)
        if order_date is None:
            invalid_date_rows.append(row)
        else:
            outbound_dates.append(order_date)

    if invalid_date_rows:
        preview = "、".join(str(row) for row in invalid_date_rows[:20])
        raise BillingProcessError(
            f"【出库订单】有 {len(invalid_date_rows)} 行无法解析【创建时间】，示例行号：{preview}。请先修正日期后重试。"
        )

    week_ranges = build_week_ranges(year, month, outbound_dates)

    if weekly_rates is not None:
        missing_weeks = [start for start, _end in week_ranges if start not in weekly_rates]
        if missing_weeks:
            detail = "、".join(
                f"{s:%Y-%m-%d}至{(s + timedelta(days=6)):%Y-%m-%d}" for s in missing_weeks
            )
            raise BillingProcessError(f"缺少以下自然周的燃油费率：{detail}")
        rates = [(start, end, weekly_rates[start]) for start, end in week_ranges]
    else:
        if fuel_rate is None:
            raise BillingProcessError("请提供燃油费率（统一费率 fuel_rate 或按周费率 weekly_rates 二选一）。")
        rates = [(start, end, fuel_rate) for start, end in week_ranges]

    write_fuel_rate_sheet(wb, rates)
    rate_by_week = {start: rate for start, _end, rate in rates}
    fallback_rate = fuel_rate if fuel_rate is not None else next(iter(rate_by_week.values()), 0.0)

    total_rows = last_data_row - 1
    # 大账单（数千行）写公式耗时较久，定期输出进度，避免用户以为程序卡死
    progress_step = max(500, total_rows // 10)
    log(f"  正在写入出库订单公式，共 {total_rows} 行...")

    for row in range(2, last_data_row + 1):
        order_date = coerce_excel_date(ws.cell(row, date_col).value)
        week_start = order_date - timedelta(days=order_date.weekday())
        rate = rate_by_week.get(week_start, fallback_rate)

        rate_text = format(rate, ".15g")
        # 燃油基数 = 实际存在的附加费列之和
        base_sum = "SUM(" + ",".join(f"{L}{row}" for L in base_letters) + ")"
        # 费用总计构成 = 处理费类各列 + 物流费用列
        total_components = proc_letters + [logistics_letter]
        total_sum = "SUM(" + ",".join(f"{L}{row}" for L in total_components) + ")"

        ws.cell(row, fuel_col).value = f"=ROUND({base_sum}*{rate_text},2)"
        ws.cell(row, logistics_col).value = f"=ROUND(({base_sum}+{fuel_letter}{row})*-1,2)"
        ws.cell(row, total_col).value = f"=ROUND({total_sum},2)"

        ws.cell(row, fuel_col).number_format = '#,##0.00;[Red]-#,##0.00'
        ws.cell(row, logistics_col).number_format = '#,##0.00;[Red]-#,##0.00'
        ws.cell(row, total_col).number_format = '#,##0.00;[Red]-#,##0.00'

        done = row - 1
        if done % progress_step == 0 and done < total_rows:
            log(f"    已写入 {done}/{total_rows} 行...")

    log(f"    已写入 {total_rows}/{total_rows} 行，完成。")

    total_row = last_data_row + 1
    label_col = max(1, total_col - 1)
    ws.cell(total_row, label_col).value = "费用总计合计"
    ws.cell(total_row, total_col).value = make_sum_formula(total_letter, 2, last_data_row, formula_function)
    ws.cell(total_row, label_col).font = Font(bold=True)
    ws.cell(total_row, total_col).font = Font(bold=True)
    ws.cell(total_row, total_col).number_format = '#,##0.00;[Red]-#,##0.00'
    ws.cell(total_row, fuel_col).value = None

    total_cell = f"{total_letter}{total_row}"
    if weekly_rates is not None:
        rate_desc = "、".join(f"{s:%m-%d}起 {r * 100:.2f}%" for s, _e, r in rates)
        log(f"[完成] {SHEET_OUTBOUND_ORDER}!{total_cell} 已写入费用总计汇总（按周费率：{rate_desc}）。")
    else:
        log(f"[完成] {SHEET_OUTBOUND_ORDER}!{total_cell} 已写入费用总计汇总（统一燃油费率 {fuel_rate * 100:.2f}%，覆盖 {len(week_ranges)} 个自然周）。")
    return FeeSummaryItem("出库订单", SHEET_OUTBOUND_ORDER, total_cell, OUTBOUND_TOTAL_HEADER)


# ------------------------- 第四部分：剩余 sheet -------------------------

ColumnChoiceCallback = Callable[[str, List[Tuple[int, str]]], Optional[Tuple[int, str]]]


def choose_candidate_column(
    ws: Worksheet,
    candidates: List[Tuple[int, str]],
    mode: str,
    column_choice_callback: Optional[ColumnChoiceCallback] = None,
    log: LogFunc = _noop_log,
) -> Optional[Tuple[int, str]]:
    if not candidates:
        return None
    if len(candidates) == 1:
        return candidates[0]
    if mode == "skip":
        return None

    if mode == "ask" and column_choice_callback is not None:
        return column_choice_callback(ws.title, candidates)

    # auto 模式，或 ask 模式但未提供交互回调时的兜底（保证一键运行不会卡住）
    for keyword in UNKNOWN_AMOUNT_KEYWORDS:
        for col, header in candidates:
            if keyword == header or keyword in header:
                return col, header
    return candidates[0]


def process_other_sheets(
    wb,
    formula_function: str,
    already_processed_sheets: Iterable[str],
    other_sheet_mode: str,
    main_summary_sheet: str,
    column_choice_callback: Optional[ColumnChoiceCallback] = None,
    log: LogFunc = _noop_log,
) -> List[FeeSummaryItem]:
    processed = set(already_processed_sheets)
    results: List[FeeSummaryItem] = []
    other_summary_rows = []

    for sheet_name in list(wb.sheetnames):
        if sheet_name in processed:
            continue
        if is_summary_or_archive_sheet(sheet_name, main_summary_sheet):
            continue

        # 防重复计算：与其他 sheet 是“汇总 vs 明细”关系的，跳过（但仅当对方确实已计入）
        if sheet_name in SHEETS_ALREADY_COUNTED:
            covering_sheet, reason = SHEETS_ALREADY_COUNTED[sheet_name]
            if covering_sheet in processed:
                log(f"[跳过] {sheet_name}：{reason}")
                other_summary_rows.append([sheet_name, "", "", "", "跳过", reason])
                continue
            else:
                log(
                    f"[注意] {sheet_name}：通常由【{covering_sheet}】计入，"
                    f"但本文件未处理【{covering_sheet}】，因此仍照常汇总本表，避免漏算。"
                )

        ws = wb[sheet_name]

        if sheet_name in FIXED_OTHER_SHEET_RULES:
            rule = FIXED_OTHER_SHEET_RULES[sheet_name]
            target_header = rule["target_header"]
            amount_col = find_header_col(ws, target_header)

            if not amount_col:
                msg = f"固定规则要求汇总「{target_header}」，但未找到该列，已跳过。"
                log(f"[跳过] {sheet_name}：{msg}")
                other_summary_rows.append([sheet_name, "", "", "", "跳过", msg])
                continue

            total_label = f"{target_header}合计"
            total_cell = write_total_formula(ws, amount_col=amount_col, total_label=total_label, formula_function=formula_function, key_cols=[amount_col])
            display_name = rule["display_name"]
            item = FeeSummaryItem(display_name, sheet_name, total_cell, target_header)
            results.append(item)
            other_summary_rows.append([sheet_name, display_name, target_header, total_cell, "已汇总", "固定规则"])
            log(f"[完成] {sheet_name}!{total_cell} 已按固定规则汇总「{target_header}」。")
            continue

        candidates = find_header_col_by_keywords(ws, UNKNOWN_AMOUNT_KEYWORDS)
        chosen = choose_candidate_column(ws, candidates, other_sheet_mode, column_choice_callback=column_choice_callback, log=log)

        if not chosen:
            msg = "未知 sheet 未找到可汇总费用列，或已选择跳过。"
            log(f"[跳过] {sheet_name}：{msg}")
            other_summary_rows.append([sheet_name, "", "", "", "跳过", msg])
            continue

        amount_col, header = chosen
        total_label = f"{header}合计"
        total_cell = write_total_formula(ws, amount_col=amount_col, total_label=total_label, formula_function=formula_function, key_cols=[amount_col])
        item = FeeSummaryItem(sheet_name, sheet_name, total_cell, header, "未知 sheet 通用识别")
        results.append(item)
        other_summary_rows.append([sheet_name, sheet_name, header, total_cell, "已汇总", "通用识别"])
        log(f"[完成] {sheet_name}!{total_cell} 已按通用规则汇总「{header}」。")

    write_other_fee_summary_sheet(wb, other_summary_rows)
    return results


def write_other_fee_summary_sheet(wb, rows: List[List[str]]) -> None:
    if SHEET_OTHER_SUMMARY in wb.sheetnames:
        del wb[SHEET_OTHER_SUMMARY]

    ws = wb.create_sheet(SHEET_OTHER_SUMMARY)
    headers = ["来源sheet", "汇总显示名称", "汇总列", "合计单元格", "处理状态", "备注"]
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(1, col)
        cell.value = header
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    for row_idx, row_data in enumerate(rows, start=2):
        for col_idx, value in enumerate(row_data, start=1):
            ws.cell(row_idx, col_idx).value = value

    widths = [18, 18, 18, 14, 12, 36]
    for col_idx, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width


# ------------------------- 第五部分：最终「汇总」sheet -------------------------

def sort_main_summary_items(items: Sequence[FeeSummaryItem]) -> List[FeeSummaryItem]:
    order = {name: idx for idx, name in enumerate(MAIN_SUMMARY_DISPLAY_ORDER)}

    def sort_key(item: FeeSummaryItem):
        return (order.get(item.display_name, 999), item.display_name)

    seen = set()
    unique_items = []
    for item in items:
        key = (item.display_name, item.sheet_name, item.total_cell)
        if key in seen:
            continue
        seen.add(key)
        unique_items.append(item)

    return sorted(unique_items, key=sort_key)


def find_or_create_main_summary_sheet(wb, summary_sheet_name: str):
    if summary_sheet_name in wb.sheetnames:
        return wb[summary_sheet_name]

    for ws in wb.worksheets:
        if norm_text(ws["A1"].value) == "类别" and norm_text(ws["B1"].value) == "金额":
            return ws

    ws = wb.create_sheet(summary_sheet_name, 0)
    ws["A1"] = "类别"
    ws["B1"] = "金额"
    for row, label in enumerate(["客户编号", "客户昵称", "公司名称", "账单币种", "账单起止日期"], start=2):
        ws.cell(row, 1).value = label
    return ws


def clear_main_summary_fee_area(ws: Worksheet, start_row: int = 7) -> None:
    max_row = max(ws.max_row, start_row)
    for row in range(start_row, max_row + 1):
        for col in range(1, 3):
            cell = ws.cell(row, col)
            cell.value = None
            cell.fill = PatternFill(fill_type=None)
            cell.font = Font(bold=False)
            cell.alignment = Alignment(horizontal="general")
            cell.number_format = "General"


def write_main_summary_sheet(wb, items: Sequence[FeeSummaryItem], summary_sheet_name: str = DEFAULT_MAIN_SUMMARY_SHEET, log: LogFunc = _noop_log) -> None:
    ws = find_or_create_main_summary_sheet(wb, summary_sheet_name)
    clear_main_summary_fee_area(ws, start_row=7)

    sorted_items = sort_main_summary_items(items)
    row = 7

    for item in sorted_items:
        ws.cell(row, 1).value = item.display_name
        ws.cell(row, 2).value = f"={quote_sheet_name(item.sheet_name)}!${re.sub(r'([A-Z]+)([0-9]+)', r'\1$\2', item.total_cell)}"
        ws.cell(row, 2).number_format = '#,##0.00;[Red]-#,##0.00'
        row += 1

    total_row = row
    ws.cell(total_row, 1).value = "费用总计"

    if total_row == 7:
        ws.cell(total_row, 2).value = "=0"
    else:
        ws.cell(total_row, 2).value = f"=SUM(B7:B{total_row - 1})"

    yellow_fill = PatternFill("solid", fgColor="FFFF00")
    ws.cell(total_row, 2).fill = yellow_fill
    ws.cell(total_row, 1).font = Font(bold=True)
    ws.cell(total_row, 2).font = Font(bold=True)
    ws.cell(total_row, 2).number_format = '#,##0.00;[Red]-#,##0.00'

    ws.column_dimensions["A"].width = max(ws.column_dimensions["A"].width or 0, 18)
    ws.column_dimensions["B"].width = max(ws.column_dimensions["B"].width or 0, 16)

    for r in range(1, total_row + 1):
        ws.cell(r, 1).alignment = Alignment(horizontal="left")
        ws.cell(r, 2).alignment = Alignment(horizontal="right")

    log(f"[完成] {ws.title}!A7:B{total_row} 已更新，前六行未改动。")


def configure_excel_recalculation(wb) -> None:
    try:
        wb.calculation.calcMode = "auto"
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
    except Exception:
        pass


# ------------------------- 总流程编排 -------------------------

OrderMismatchCallback = Callable[[int, int], bool]
WeeklyRateCallback = Callable[[List[Tuple[date, date]]], Optional[Dict[date, float]]]


def _read_all_sheets(input_path: str, log: LogFunc = _noop_log) -> Dict[str, pd.DataFrame]:
    log(f"正在读取文件：{input_path}")
    try:
        return pd.read_excel(input_path, sheet_name=None)
    except Exception as exc:
        raise BillingProcessError(f"无法读取 Excel 文件，请确认文件未被占用且格式正确。\n详细信息：{exc}") from exc


def _write_sheets_to_excel(output_sheets: Dict[str, pd.DataFrame], path: str, log: LogFunc = _noop_log) -> None:
    try:
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            for name, df in output_sheets.items():
                df.to_excel(writer, sheet_name=name[:31], index=False)
    except PermissionError as exc:
        raise BillingProcessError(
            f"无法写入文件：{path}\n该文件可能正被 Excel 打开，请关闭后重试。"
        ) from exc


def _save_workbook(wb, output_path: str, log: LogFunc = _noop_log) -> None:
    """保存工作簿；文件被占用/无权限时给出友好提示，而不是抛出原始 traceback。"""
    try:
        wb.save(output_path)
    except PermissionError as exc:
        raise BillingProcessError(
            f"无法保存结果文件：{output_path}\n"
            f"该文件可能正被 Excel 打开，请关闭该文件后重新运行。"
        ) from exc
    except OSError as exc:
        raise BillingProcessError(f"保存结果文件失败：{output_path}\n详细信息：{exc}") from exc


def _check_order_count(
    actual_order_count: int,
    expected_order_count: Optional[int],
    order_count_mismatch_callback: Optional[OrderMismatchCallback],
    log: LogFunc = _noop_log,
) -> bool:
    if expected_order_count is None:
        log("未填写系统中订单件数，跳过核对步骤。")
        return True

    if expected_order_count != actual_order_count:
        diff = actual_order_count - expected_order_count
        msg = f"核对不一致：系统订单件数为 {expected_order_count}，清洗后出库订单总数据量为 {actual_order_count}，相差 {diff} 条。"
        log(msg)
        matched = order_count_mismatch_callback(actual_order_count, expected_order_count) if order_count_mismatch_callback else False
        if not matched:
            raise OrderCountMismatchError(msg)
        log("用户确认继续，忽略件数差异，继续处理。")
        return True

    log(f"核对通过：系统订单件数（{expected_order_count}）与出库订单清洗后总数据量（{actual_order_count}）一致。")
    return True


def get_outbound_week_ranges_from_workbook(wb, year: int, month: int) -> List[Tuple[date, date]]:
    """从已加载的工作簿里算出【出库订单】覆盖的自然周（供“仅正式处理”模式按周填费率用）。"""
    if SHEET_OUTBOUND_ORDER not in wb.sheetnames:
        return build_week_ranges(year, month)

    ws = wb[SHEET_OUTBOUND_ORDER]
    date_col = find_header_col(ws, OUTBOUND_DATE_HEADER)
    if not date_col:
        return build_week_ranges(year, month)

    dates: List[date] = []
    for row in range(2, ws.max_row + 1):
        d = coerce_excel_date(ws.cell(row, date_col).value)
        if d:
            dates.append(d)
    return build_week_ranges(year, month, dates)


def _formal_process_workbook(
    wb,
    year: int,
    month: int,
    fuel_rate: Optional[float],
    weekly_rates: Optional[Dict[date, float]],
    formula_function: str,
    other_sheet_mode: str,
    column_choice_callback: Optional[ColumnChoiceCallback],
    log: LogFunc = _noop_log,
) -> List[FeeSummaryItem]:
    """第二阶段：在已加载的工作簿上完成正式处理（仓租/入库详情/出库订单/其他费用/汇总）。"""
    log("正在进行正式处理（仓租 / 入库详情 / 出库订单 / 其他费用 / 汇总）...")

    summary_items: List[FeeSummaryItem] = []
    processed_sheets = set()

    item = process_warehouse_rent(wb, formula_function, log=log)
    if item:
        summary_items.append(item)
        processed_sheets.add(item.sheet_name)

    item = process_inbound_detail(wb, formula_function, log=log)
    if item:
        summary_items.append(item)
        processed_sheets.add(item.sheet_name)

    item = process_outbound_order(
        wb, formula_function, year, month,
        fuel_rate=fuel_rate, weekly_rates=weekly_rates, log=log,
    )
    if item:
        summary_items.append(item)
        processed_sheets.add(item.sheet_name)

    other_items = process_other_sheets(
        wb,
        formula_function=formula_function,
        already_processed_sheets=processed_sheets,
        other_sheet_mode=other_sheet_mode,
        main_summary_sheet=DEFAULT_MAIN_SUMMARY_SHEET,
        column_choice_callback=column_choice_callback,
        log=log,
    )
    summary_items.extend(other_items)

    write_main_summary_sheet(wb, summary_items, DEFAULT_MAIN_SUMMARY_SHEET, log=log)
    configure_excel_recalculation(wb)
    return summary_items


def run_clean_only(
    input_path: str,
    year: int,
    month: int,
    output_path: str,
    expected_order_count: Optional[int] = None,
    remove_empty_sheets: bool = True,
    order_count_mismatch_callback: Optional[OrderMismatchCallback] = None,
    log: LogFunc = _noop_log,
) -> dict:
    """
    只做第一阶段清洗，并把清洗结果单独保存成文件。

    用途：先输出清洗后的文件，人工在里面补充/调整数据（例如手动新增表格），
    再用 run_formal_only() 对补充完的文件做正式处理。
    """
    all_sheets = _read_all_sheets(input_path, log=log)

    removed_sheets: List[str] = []
    if remove_empty_sheets:
        all_sheets, removed_sheets = drop_header_only_sheets(all_sheets, log=log)

    cleaned = clean_billing_data(all_sheets, year, month, log=log)
    output_sheets = cleaned["sheets"]
    actual_order_count = cleaned["outbound_clean_count"]

    _check_order_count(actual_order_count, expected_order_count, order_count_mismatch_callback, log=log)

    log("正在保存清洗后的文件...")
    _write_sheets_to_excel(output_sheets, output_path, log=log)
    log(f"清洗完成，已保存至：{output_path}")
    log("提示：可在该文件中手动补充数据后，用「仅正式处理」模式继续处理。")

    return {
        "stage": "clean_only",
        "actual_order_count": actual_order_count,
        "expected_order_count": expected_order_count,
        "removed_empty_sheets": removed_sheets,
        "output_path": output_path,
        "summary_items": [],
    }


def run_formal_only(
    input_path: str,
    year: int,
    month: int,
    output_path: str,
    fuel_rate: Optional[float] = None,
    weekly_rate_callback: Optional[WeeklyRateCallback] = None,
    formula_function: str = "SUM",
    other_sheet_mode: str = "auto",
    remove_empty_sheets: bool = False,
    column_choice_callback: Optional[ColumnChoiceCallback] = None,
    log: LogFunc = _noop_log,
) -> dict:
    """
    只做第二阶段正式处理，输入应为“已清洗（并可能已人工补充数据）”的文件。

    注意：本模式不会重新清洗，也不做 pandas 往返，
    直接在原工作簿上处理，以完整保留人工补充的内容与格式。
    """
    log(f"正在读取文件：{input_path}")
    try:
        wb = load_workbook(input_path)
    except Exception as exc:
        raise BillingProcessError(f"无法读取 Excel 文件，请确认文件未被占用且格式正确。\n详细信息：{exc}") from exc

    removed_sheets: List[str] = []
    if remove_empty_sheets:
        removed_sheets = drop_header_only_worksheets(wb, log=log)

    weekly_rates: Optional[Dict[date, float]] = None
    if weekly_rate_callback is not None:
        week_ranges_preview = get_outbound_week_ranges_from_workbook(wb, year, month)
        log(f"本月出库订单覆盖 {len(week_ranges_preview)} 个自然周，等待按周填写燃油费率...")
        weekly_rates = weekly_rate_callback(week_ranges_preview)
        if weekly_rates is None:
            raise PipelineCancelled("用户取消了按周燃油费率的填写。")
        log("已获取按周燃油费率，继续处理。")

    summary_items = _formal_process_workbook(
        wb, year, month, fuel_rate, weekly_rates,
        formula_function, other_sheet_mode, column_choice_callback, log=log,
    )

    _save_workbook(wb, output_path, log=log)
    log(f"处理完成，结果已保存至：{output_path}")

    return {
        "stage": "formal_only",
        "matched": True,
        "actual_order_count": None,
        "expected_order_count": None,
        "removed_empty_sheets": removed_sheets,
        "summary_items": summary_items,
        "output_path": output_path,
    }


def run_full_pipeline(
    input_path: str,
    year: int,
    month: int,
    output_path: str,
    fuel_rate: Optional[float] = None,
    weekly_rate_callback: Optional[WeeklyRateCallback] = None,
    expected_order_count: Optional[int] = None,
    formula_function: str = "SUM",
    other_sheet_mode: str = "auto",
    remove_empty_sheets: bool = True,
    cleaned_output_path: Optional[str] = None,
    column_choice_callback: Optional[ColumnChoiceCallback] = None,
    order_count_mismatch_callback: Optional[OrderMismatchCallback] = None,
    log: LogFunc = _noop_log,
) -> dict:
    """
    完整跑通"清洗 -> 正式处理 -> 保存"，中间用临时文件衔接两个阶段。

    燃油费率二选一：
      - fuel_rate：统一费率，应用到本月覆盖的所有自然周
      - weekly_rate_callback：按周填写。清洗完成后会用实际覆盖到的自然周区间
        调用该回调，回调返回 {周一日期: 费率} 字典；返回 None 表示用户取消。

    remove_empty_sheets：是否自动删除“仅有表头”的空sheet（保护名单除外）。
    cleaned_output_path：若指定，会额外保存一份清洗后的中间文件供留档/复核。

    出错时抛出 BillingProcessError / OrderCountMismatchError / PipelineCancelled，
    调用方（GUI）负责友好提示。
    """
    all_sheets = _read_all_sheets(input_path, log=log)

    removed_sheets: List[str] = []
    if remove_empty_sheets:
        all_sheets, removed_sheets = drop_header_only_sheets(all_sheets, log=log)

    cleaned = clean_billing_data(all_sheets, year, month, log=log)
    output_sheets = cleaned["sheets"]
    actual_order_count = cleaned["outbound_clean_count"]

    matched = _check_order_count(actual_order_count, expected_order_count, order_count_mismatch_callback, log=log)

    weekly_rates: Optional[Dict[date, float]] = None
    if weekly_rate_callback is not None:
        week_ranges_preview = compute_outbound_week_ranges(output_sheets, year, month)
        log(f"本月出库订单覆盖 {len(week_ranges_preview)} 个自然周，等待按周填写燃油费率...")
        weekly_rates = weekly_rate_callback(week_ranges_preview)
        if weekly_rates is None:
            raise PipelineCancelled("用户取消了按周燃油费率的填写。")
        log("已获取按周燃油费率，继续处理。")

    # 可选：额外留档一份清洗后的文件
    if cleaned_output_path:
        log("正在保存清洗后的中间文件...")
        _write_sheets_to_excel(output_sheets, cleaned_output_path, log=log)
        log(f"清洗后文件已保存至：{cleaned_output_path}")

    tmp_fd, tmp_path = tempfile.mkstemp(suffix=".xlsx")
    os.close(tmp_fd)

    try:
        log("正在生成清洗后的中间数据...")
        _write_sheets_to_excel(output_sheets, tmp_path, log=log)

        wb = load_workbook(tmp_path)
        summary_items = _formal_process_workbook(
            wb, year, month, fuel_rate, weekly_rates,
            formula_function, other_sheet_mode, column_choice_callback, log=log,
        )
        _save_workbook(wb, output_path, log=log)
        log(f"处理完成，结果已保存至：{output_path}")
    finally:
        try:
            os.remove(tmp_path)
        except OSError:
            pass

    return {
        "stage": "full",
        "matched": matched,
        "actual_order_count": actual_order_count,
        "expected_order_count": expected_order_count,
        "removed_empty_sheets": removed_sheets,
        "summary_items": summary_items,
        "output_path": output_path,
    }
