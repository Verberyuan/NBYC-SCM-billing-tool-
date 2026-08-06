# -*- coding: utf-8 -*-
"""
rate_store.py —— 燃油费率「记忆」功能

需求：第一次输入过燃油费率之后，之后处理别的表格可以直接沿用，不必重复输入。

做法
----
把费率区间存成一份 JSON（`fuel_rate_memory.json`），放在 exe 同目录；
若该目录不可写（例如装在 Program Files 下），自动退到用户目录
`%USERPROFILE%\\.billing_tool\\fuel_rate_memory.json`。

每次成功处理完账单后，本次用到的费率区间会自动写回记忆；
下次打开程序，界面上会显示「已记忆 N 个区间」，勾选即可直接使用。
覆盖不到的周会被单独列出来，只需要补填缺的那几周。
"""

from __future__ import annotations

import json
import os
import sys
import tempfile
from dataclasses import asdict
from datetime import date, datetime
from decimal import Decimal
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

from fuel_rates import FuelPeriod, DEFAULT_TIMEZONE

MEMORY_FILENAME = "fuel_rate_memory.json"
STORE_VERSION = 2


# ---------------------------------------------------------------------------
# 存储位置
# ---------------------------------------------------------------------------

def _app_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def _fallback_dir() -> str:
    home = os.path.expanduser("~")
    return os.path.join(home, ".billing_tool")


def _dir_writable(path: str) -> bool:
    try:
        os.makedirs(path, exist_ok=True)
        fd, tmp = tempfile.mkstemp(dir=path)
        os.close(fd)
        os.remove(tmp)
        return True
    except Exception:
        return False


def store_path() -> str:
    """返回记忆文件的完整路径（优先 exe 同目录，不可写时退到用户目录）。"""
    primary = _app_dir()
    if _dir_writable(primary):
        return os.path.join(primary, MEMORY_FILENAME)
    fallback = _fallback_dir()
    os.makedirs(fallback, exist_ok=True)
    return os.path.join(fallback, MEMORY_FILENAME)


# ---------------------------------------------------------------------------
# 读写
# ---------------------------------------------------------------------------

def _period_to_dict(p: FuelPeriod) -> Dict[str, str]:
    return {
        "carrier": p.carrier,
        "service_group": p.service_group,
        "start": p.start.strftime("%Y-%m-%d"),
        "end": p.end.strftime("%Y-%m-%d"),
        "rate": str(p.rate),
        "timezone": p.timezone,
    }


def _period_from_dict(d: Dict[str, str]) -> FuelPeriod:
    return FuelPeriod(
        start=datetime.strptime(d["start"], "%Y-%m-%d").date(),
        end=datetime.strptime(d["end"], "%Y-%m-%d").date(),
        rate=Decimal(str(d["rate"])),
        carrier=d.get("carrier") or "FedEx",
        service_group=d.get("service_group") or "",
        timezone=d.get("timezone") or DEFAULT_TIMEZONE,
    )


def load_all() -> dict:
    path = store_path()
    if not os.path.exists(path):
        return {"version": STORE_VERSION, "periods": [], "updated_at": None}
    try:
        with open(path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        if not isinstance(data, dict):
            raise ValueError
        data.setdefault("periods", [])
        return data
    except Exception:
        # 记忆文件损坏时不要拖垮主流程，当作空记忆处理
        return {"version": STORE_VERSION, "periods": [], "updated_at": None,
                "warning": "记忆文件无法读取，已按空记忆处理。"}


def load_periods(carrier: Optional[str] = None) -> List[FuelPeriod]:
    """读出已记忆的费率区间（按开始日期排序）。"""
    periods: List[FuelPeriod] = []
    for item in load_all().get("periods", []):
        try:
            p = _period_from_dict(item)
        except Exception:
            continue
        if carrier and p.carrier and p.carrier.strip().lower() != carrier.strip().lower():
            continue
        periods.append(p)
    return sorted(periods, key=lambda x: (x.carrier, x.start))


def _write(data: dict) -> str:
    path = store_path()
    data["version"] = STORE_VERSION
    data["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=2)
    os.replace(tmp, path)   # 原子替换，避免写一半断电导致记忆损坏
    return path


def remember(periods: Sequence[FuelPeriod]) -> str:
    """
    把本次用到的费率区间并入记忆。

    规则：同一承运商下，新区间与旧区间**有重叠就以新的为准**（旧的被移除），
    不重叠的旧区间原样保留。这样按月运行时记忆会自然地越积越全。
    """
    if not periods:
        return store_path()

    data = load_all()
    existing = [
        p for p in (
            _safe_period(item) for item in data.get("periods", [])
        ) if p is not None
    ]

    kept: List[FuelPeriod] = []
    for old in existing:
        overlapped = any(
            (old.carrier or "").strip().lower() == (new.carrier or "").strip().lower()
            and not (old.end < new.start or old.start > new.end)
            for new in periods
        )
        if not overlapped:
            kept.append(old)

    merged = sorted(kept + list(periods), key=lambda x: (x.carrier, x.start))
    data["periods"] = [_period_to_dict(p) for p in merged]
    return _write(data)


def replace_all(periods: Sequence[FuelPeriod]) -> str:
    """整体覆盖记忆（用于「管理已记忆费率」窗口里的保存）。"""
    data = load_all()
    data["periods"] = [_period_to_dict(p) for p in sorted(periods, key=lambda x: (x.carrier, x.start))]
    return _write(data)


def clear() -> None:
    path = store_path()
    if os.path.exists(path):
        os.remove(path)


def _safe_period(item) -> Optional[FuelPeriod]:
    try:
        return _period_from_dict(item)
    except Exception:
        return None


# ---------------------------------------------------------------------------
# 覆盖情况
# ---------------------------------------------------------------------------

def missing_weeks(
    periods: Sequence[FuelPeriod],
    week_ranges: Sequence[Tuple[date, date]],
    carrier: Optional[str] = None,
) -> List[Tuple[date, date]]:
    """返回记忆里还没有费率的周区间（整周未被覆盖才算缺）。"""
    missing = []
    for start, end in week_ranges:
        covered = any(
            (not carrier or not p.carrier or p.carrier.strip().lower() == carrier.strip().lower())
            and p.start <= start and p.end >= end
            for p in periods
        )
        if not covered:
            missing.append((start, end))
    return missing


def describe_memory(periods: Sequence[FuelPeriod]) -> str:
    if not periods:
        return "尚未记忆任何燃油费率"
    first, last = periods[0], periods[-1]
    return (f"已记忆 {len(periods)} 个区间（{first.start:%Y-%m-%d} ~ {last.end:%Y-%m-%d}）")
