# -*- coding: utf-8 -*-
"""
fuel_rates.py —— 燃油费率区间表

对应《代码修改建议》：
   9. FedEx 费率采用承运商官网周区间，不再由账单数据反推、不再整体后移一天
  10. 时区明确但不随意转换；燃油费率匹配哪个日期字段由配置决定
  11. 费率表必须检查覆盖、重叠、空档、非法费率

设计要点
--------
* 费率区间是**显式配置的起止日期**，不是程序猜出来的。
* 区间闭区间：start <= 日期 <= end。
* 校验在开始计算前执行，发现重叠/空档/未覆盖直接阻断，不允许“带病运行”。
"""

from __future__ import annotations

import csv
import os
from dataclasses import dataclass, field
from datetime import date, timedelta
from decimal import Decimal
from typing import Iterable, List, Optional, Sequence, Tuple

from money import parse_rate, format_rate

DEFAULT_TIMEZONE = "Asia/Shanghai"


class FuelRateError(Exception):
    """费率表本身有问题（重叠/空档/非法费率/未覆盖订单日期）。"""


@dataclass
class FuelPeriod:
    start: date
    end: date
    rate: Decimal
    carrier: str = "FedEx"
    service_group: str = ""
    timezone: str = DEFAULT_TIMEZONE

    def contains(self, day: date) -> bool:
        return self.start <= day <= self.end

    def describe(self) -> str:
        return f"{self.start:%Y-%m-%d} ~ {self.end:%Y-%m-%d}  {format_rate(self.rate)}"


# --------------------------------------------------------------------------
# 生成 / 载入费率区间
# --------------------------------------------------------------------------

def carrier_week_ranges(
    first_day: date,
    last_day: date,
    week_start_weekday: int = 0,
) -> List[Tuple[date, date]]:
    """
    生成覆盖 [first_day, last_day] 的承运商自然周区间（闭区间，7 天一段）。

    week_start_weekday: 0=周一（FedEx 中国区账单口径，例 2026-06-29 周一 ~ 2026-07-05 周日）
                        6=周日（部分承运商以周日为一周开始）

    注意：这里**不做任何 ±1 天的边界平移**。若承运商官网区间与自然周不同，
    请直接用 fuel_rates.csv 明确写死起止日期，不要靠程序推。
    """
    offset = (first_day.weekday() - week_start_weekday) % 7
    cur = first_day - timedelta(days=offset)
    ranges: List[Tuple[date, date]] = []
    while cur <= last_day:
        ranges.append((cur, cur + timedelta(days=6)))
        cur += timedelta(days=7)
    return ranges


def build_periods_from_weekly_rates(
    week_rates: Sequence[Tuple[date, date, Decimal]],
    carrier: str = "FedEx",
    service_group: str = "",
    timezone: str = DEFAULT_TIMEZONE,
) -> List[FuelPeriod]:
    return [
        FuelPeriod(start=s, end=e, rate=r, carrier=carrier,
                   service_group=service_group, timezone=timezone)
        for s, e, r in week_rates
    ]


def build_uniform_periods(
    first_day: date,
    last_day: date,
    rate: Decimal,
    carrier: str = "FedEx",
    week_start_weekday: int = 0,
    timezone: str = DEFAULT_TIMEZONE,
) -> List[FuelPeriod]:
    """统一费率：仍然按周切分区间，方便后续单独调整某一周。"""
    return [
        FuelPeriod(start=s, end=e, rate=rate, carrier=carrier, timezone=timezone)
        for s, e in carrier_week_ranges(first_day, last_day, week_start_weekday)
    ]


def _parse_date_text(text: str) -> date:
    text = str(text).strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y%m%d"):
        try:
            from datetime import datetime
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    raise FuelRateError(f"费率表中的日期「{text}」格式不正确，应为 YYYY-MM-DD。")


def load_periods_from_csv(path: str) -> List[FuelPeriod]:
    """
    从 CSV 载入费率表。表头（中英文均可识别）：
        carrier,service_group,start,end,rate,timezone
        承运商,服务类型,开始日期,结束日期,燃油费率,时区
    rate 支持 0.25 / 25% / 25 三种写法。
    """
    if not os.path.exists(path):
        raise FuelRateError(f"找不到燃油费率表文件：{path}")

    alias = {
        "carrier": "carrier", "承运商": "carrier",
        "service_group": "service_group", "服务类型": "service_group", "服务分组": "service_group",
        "start": "start", "开始日期": "start", "起始日期": "start", "开始": "start",
        "end": "end", "结束日期": "end", "截止日期": "end", "结束": "end",
        "rate": "rate", "燃油费率": "rate", "费率": "rate",
        "timezone": "timezone", "时区": "timezone",
    }

    periods: List[FuelPeriod] = []
    with open(path, "r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        for line_no, raw_row in enumerate(reader, start=2):
            row = {}
            for key, value in raw_row.items():
                if key is None:
                    continue
                mapped = alias.get(str(key).strip().lower()) or alias.get(str(key).strip())
                if mapped:
                    row[mapped] = value
            if not row.get("start") or not row.get("end"):
                continue
            try:
                periods.append(
                    FuelPeriod(
                        start=_parse_date_text(row["start"]),
                        end=_parse_date_text(row["end"]),
                        rate=parse_rate(row.get("rate", "")),
                        carrier=(row.get("carrier") or "FedEx").strip(),
                        service_group=(row.get("service_group") or "").strip(),
                        timezone=(row.get("timezone") or DEFAULT_TIMEZONE).strip(),
                    )
                )
            except Exception as exc:
                raise FuelRateError(f"费率表第 {line_no} 行解析失败：{exc}") from exc

    if not periods:
        raise FuelRateError(f"费率表 {path} 中没有读到任何有效区间。")
    return periods


def write_periods_template_csv(path: str, periods: Sequence[FuelPeriod]) -> None:
    """导出一份费率表模板，方便按承运商官网区间维护。"""
    with open(path, "w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.writer(fh)
        writer.writerow(["carrier", "service_group", "start", "end", "rate", "timezone"])
        for p in periods:
            writer.writerow([
                p.carrier, p.service_group,
                f"{p.start:%Y-%m-%d}", f"{p.end:%Y-%m-%d}",
                f"{p.rate}", p.timezone,
            ])


# --------------------------------------------------------------------------
# 校验（建议 11）
# --------------------------------------------------------------------------

def validate_periods(
    periods: Sequence[FuelPeriod],
    required_dates: Optional[Iterable[date]] = None,
    allow_gap: bool = False,
) -> List[str]:
    """
    返回问题清单（空列表 = 通过）。调用方决定是阻断还是仅告警。

    检查项：
      1. 费率为空 / 小于 0 / 明显异常（>100%）
      2. start > end
      3. 两个区间重叠
      4. 区间之间存在日期空档
      5. 是否覆盖全部应计订单日期
    """
    problems: List[str] = []
    if not periods:
        return ["燃油费率表为空，无法计算燃油费。"]

    for p in periods:
        if p.rate is None:
            problems.append(f"区间 {p.start:%Y-%m-%d}~{p.end:%Y-%m-%d} 的费率为空。")
        elif p.rate < 0:
            problems.append(f"区间 {p.start:%Y-%m-%d}~{p.end:%Y-%m-%d} 的费率为负数（{p.rate}）。")
        elif p.rate > 1:
            problems.append(
                f"区间 {p.start:%Y-%m-%d}~{p.end:%Y-%m-%d} 的费率为 {p.rate}（超过 100%），请确认是否填错。"
            )
        if p.start > p.end:
            problems.append(f"区间起止日期颠倒：{p.start:%Y-%m-%d} > {p.end:%Y-%m-%d}。")

    ordered = sorted(periods, key=lambda x: (x.start, x.end))
    for previous, current in zip(ordered, ordered[1:]):
        if current.start <= previous.end:
            problems.append(
                f"燃油费率区间重叠：{previous.start:%Y-%m-%d}~{previous.end:%Y-%m-%d} "
                f"与 {current.start:%Y-%m-%d}~{current.end:%Y-%m-%d}。"
            )
        elif not allow_gap and current.start > previous.end + timedelta(days=1):
            problems.append(
                f"燃油费率区间存在空档：{(previous.end + timedelta(days=1)):%Y-%m-%d} "
                f"至 {(current.start - timedelta(days=1)):%Y-%m-%d} 没有费率。"
            )

    if required_dates:
        uncovered = sorted({d for d in required_dates
                            if d is not None and not any(p.contains(d) for p in ordered)})
        if uncovered:
            preview = "、".join(f"{d:%Y-%m-%d}" for d in uncovered[:10])
            more = f" 等 {len(uncovered)} 个日期" if len(uncovered) > 10 else ""
            problems.append(f"以下订单日期未被任何费率区间覆盖：{preview}{more}。")

    return problems


def infer_rates_from_samples(
    samples: Sequence[Tuple[date, Decimal, Decimal]],
    week_ranges: Sequence[Tuple[date, date]],
) -> List[Tuple[date, date, Optional[Decimal], int, str]]:
    """
    从源账单反推每个周区间实际使用的燃油费率。

    samples: [(日期, 燃油基数, 源文件燃油附加费), ...]，基数或燃油为 0 的行请先剔除。
    返回 [(start, end, 推断费率或None, 样本数, 说明), ...]

    用途：承运商官网费率一时查不到时，可以先用这个看看源账单到底按多少算的；
    也可以用来验证"账单是否把费率整体后移了一天"——若某周区间内出现两种明显不同的
    费率，说明该周的边界与账单实际换档日不一致。
    """
    results = []
    for start, end in week_ranges:
        rates = []
        for day, base, fuel in samples:
            if day is None or base is None or base == 0 or fuel is None or fuel == 0:
                continue
            if start <= day <= end:
                rates.append((fuel / base).quantize(Decimal("0.0001")))
        if not rates:
            results.append((start, end, None, 0, "该区间无可用样本"))
            continue

        rates.sort()
        median = rates[len(rates) // 2]
        # 单笔燃油费经过 2 位取整，反推值会有小幅抖动，取中位数并归到 0.05% 档
        rounded = (median * Decimal(400)).quantize(Decimal("1")) / Decimal(400)
        spread = rates[-1] - rates[0]
        note = "样本一致" if spread <= Decimal("0.004") else (
            f"★区间内费率不一致（{rates[0]}~{rates[-1]}），该周边界可能与账单换档日不符")
        results.append((start, end, rounded, len(rates), note))
    return results


def find_rate(periods: Sequence[FuelPeriod], day: date,
              carrier: Optional[str] = None) -> Optional[Decimal]:
    """查某天适用的燃油费率；查不到返回 None（调用方按异常处理，不得默默按 0 计算）。"""
    if day is None:
        return None
    for p in periods:
        if carrier and p.carrier and p.carrier.strip().lower() != carrier.strip().lower():
            continue
        if p.contains(day):
            return p.rate
    return None


def describe_periods(periods: Sequence[FuelPeriod]) -> str:
    return "；".join(p.describe() for p in sorted(periods, key=lambda x: x.start))
