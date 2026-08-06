# -*- coding: utf-8 -*-
"""
money.py —— 金额与费率的 Decimal 统一处理

对应《代码修改建议》：
  12. 全程使用 Decimal，不要使用 float
  13. 每笔费用先按规则取整，再汇总

所有金额字段进入计算前都必须经过 to_decimal()，
所有中间结果都必须经过 round_money()，
禁止在任何地方出现 float(value) 的金额运算。
"""

from decimal import Decimal, ROUND_HALF_UP, ROUND_HALF_EVEN, ROUND_DOWN, InvalidOperation
from typing import Optional

ZERO = Decimal("0")

ROUNDING_MODES = {
    "ROUND_HALF_UP": ROUND_HALF_UP,
    "ROUND_HALF_EVEN": ROUND_HALF_EVEN,
    "ROUND_DOWN": ROUND_DOWN,
}

# 常见的“空值”文本，出现在金额列时按 0 处理而不是报错
_BLANK_TEXTS = {"", "-", "--", "—", "none", "null", "nan", "n/a", "na", "#n/a", "无"}

# 金额文本里允许出现并被剔除的符号
_STRIP_CHARS = ["¥", "￥", "$", "＄", ",", "，", " ", "\u3000", "\n", "\r", "\t"]


class MoneyParseError(ValueError):
    """金额字段里出现了无法解析为数字的文本（对应运行前检查：金额列出现文本）。"""

    def __init__(self, raw, field: str = ""):
        self.raw = raw
        self.field = field
        where = f"字段【{field}】" if field else "金额字段"
        super().__init__(f"{where}中的值「{raw}」不是有效数字，无法参与金额计算。")


def is_blank_value(value) -> bool:
    """判断是否为“空”（None / 空字符串 / pandas 的 NaN / NaT）。"""
    if value is None:
        return True
    # pandas / numpy 的 NaN：自身不等于自身
    if isinstance(value, float) and value != value:
        return True
    text = str(value).strip()
    if text.lower() in _BLANK_TEXTS:
        return True
    if text.lower() in {"nat", "nan"}:
        return True
    return False


def to_decimal(value, field: str = "", default: Optional[Decimal] = ZERO,
               strict: bool = True) -> Decimal:
    """
    把任意来源的值安全转成 Decimal。

    strict=True  ：无法解析时抛 MoneyParseError（用于金额列，配合运行前检查报异常）
    strict=False ：无法解析时返回 default（用于容错场景）
    """
    if isinstance(value, Decimal):
        return value
    if is_blank_value(value):
        return default if default is not None else ZERO
    if isinstance(value, bool):
        # 布尔值出现在金额列一定是数据问题
        if strict:
            raise MoneyParseError(value, field)
        return default if default is not None else ZERO
    if isinstance(value, int):
        return Decimal(value)
    if isinstance(value, float):
        # 关键：float 必须先转 str 再进 Decimal，避免二进制浮点误差被带进来
        return Decimal(str(value))

    text = str(value).strip()
    for ch in _STRIP_CHARS:
        text = text.replace(ch, "")

    negative_parens = text.startswith("(") and text.endswith(")")
    if negative_parens:
        text = text[1:-1]

    if text == "":
        return default if default is not None else ZERO

    try:
        result = Decimal(text)
    except InvalidOperation:
        if strict:
            raise MoneyParseError(value, field)
        return default if default is not None else ZERO

    return -result if negative_parens else result


def looks_like_money_text(value) -> bool:
    """金额列里出现的、无法解析成数字的“文本”，用于运行前数据质量检查。"""
    if is_blank_value(value):
        return False
    try:
        to_decimal(value, strict=True)
    except MoneyParseError:
        return True
    return False


def round_money(value, digits: int = 2, mode: str = "ROUND_HALF_UP") -> Decimal:
    """
    按业务规则取整。默认 2 位小数、四舍五入（ROUND_HALF_UP）。

    注意：Python 内置 round() 用的是银行家舍入（ROUND_HALF_EVEN），
    与财务口径不一致，因此这里必须显式指定 rounding。
    """
    dec = value if isinstance(value, Decimal) else to_decimal(value, strict=False)
    exp = Decimal(1).scaleb(-digits)  # digits=2 -> Decimal("0.01")
    return dec.quantize(exp, rounding=ROUNDING_MODES.get(mode, ROUND_HALF_UP))


def parse_rate(raw) -> Decimal:
    """
    燃油费率解析，统一返回小数形式的 Decimal。
    支持：18.5% / 18.5 / 0.185 / "18.5％"
    """
    if isinstance(raw, Decimal):
        return raw
    if isinstance(raw, (int, float)):
        value = Decimal(str(raw))
        return value / Decimal(100) if value > 1 else value

    text = str(raw).strip().replace("％", "%").replace(" ", "")
    if not text:
        raise ValueError("燃油费率不能为空")
    if text.endswith("%"):
        return to_decimal(text[:-1], field="燃油费率") / Decimal(100)

    value = to_decimal(text, field="燃油费率")
    # 0.185 视为 18.5%；18.5 也视为 18.5%
    return value / Decimal(100) if value > 1 else value


def format_rate(rate: Decimal) -> str:
    """把小数费率格式化为 18.50% 这样的展示文本。"""
    return f"{(rate * Decimal(100)).quantize(Decimal('0.01'))}%"


def decimal_to_float(value: Decimal) -> float:
    """仅在写入 Excel 单元格时使用：openpyxl 不接受 Decimal，需要转 float。

    此时数值已完成取整，转 float 不会再引入可见误差。
    """
    if value is None:
        return 0.0
    return float(value)
