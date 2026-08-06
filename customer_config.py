# -*- coding: utf-8 -*-
"""
customer_config.py —— 按客户维护的计费规则配置

★ 本文件的字段名已按五份真实账单（MPAV / HBKE / MUTN / IPQI / GUXE）核对过。
★ 修改客户规则请只改本文件，不要改 billing_core.py。

真实账单结构要点（与原《代码修改建议》文档的描述有出入，以实际表格为准）
--------------------------------------------------------------------
1. 决定账期的「物流类型」实际字段名是【尾程物流】，取值只有三种：
       平台快递 / 自有快递 / 卡派自提
   而「单件代发 / 复合代发 / 一票多件」是另一个字段【订单类型】，与账期无关。
2. 账单里所有"费用"都是**负数**（订单处理费 -2.2、物流费用 -22.04、费用总计 -24.24），
   而附加费明细（住宅地址附加费 / 基础运费 / 燃油附加费 …）是**正数**，作为构成项。
   因此公式为：
       燃油基数 = Σ(附加费明细，正数)
       实际燃油费 = ROUND(燃油基数 × 费率, 2)                     （正数）
       物流费用   = -ROUND(燃油基数 + 实际燃油费 + PlatformFee, 2)  （负数）
       费用总计   = ROUND(Σ处理费类(负数) + 物流费用, 2)            （负数）
3. 各客户的费用列不同（GUXE 连燃油相关列都没有），配置里写的是**全集**，
   程序会自动忽略该客户不存在的列，因此新增客户通常不需要改这里。
"""

from __future__ import annotations

import copy
from typing import Any, Dict, List, Tuple

# ---------------------------------------------------------------------------
# 默认配置（= 五个客户的并集，程序自动忽略不存在的列）
# ---------------------------------------------------------------------------
DEFAULT_CONFIG: Dict[str, Any] = {
    "customer_code": "DEFAULT",
    "customer_name": "默认规则",

    "sheets": {
        "outbound": "出库订单",
        "warehouse_rent": "仓租",
        "inbound_detail": "入库详情",
        "summary": "汇总",
    },

    # ---- 关键字段（全部是表头文字，不是列字母） ----
    "fields": {
        "order_id": "订单编号",
        "sku": "SKU",
        "logistics_type": "尾程物流",     # ★ 决定账期的字段（平台快递/自有快递/卡派自提）
        "order_type": "订单类型",         # 单件代发/复合代发/一票多件，仅作参考
        "created_at": "创建时间",
        "shipped_at": "发货时间",
        "finished_at": "完成时间",
        "status": "订单状态",             # 实际账单里没有这一列，留空即跳过
        "remark": "订单备注",
        # 计算结果写回的列
        "fuel_fee": "实际燃油费",         # 新增列
        "logistics_fee": "物流费用",
        "total": "费用总计",
        # 承运商给的燃油金额（fuel.mode="source" 时直接用它）
        "source_fuel_fee": "燃油附加费",
    },

    # ---- 4. 账期归属：尾程物流 -> 用哪个时间字段 ----
    "period_rules": {
        "卡派自提": "完成时间",
        "平台快递": "创建时间",
        "自有快递": "创建时间",
    },
    "period_default_field": "创建时间",
    "warn_on_period_fallback": True,

    # ---- 燃油 ----
    "fuel": {
        # none = 不计燃油；schedule = 按费率区间表；source = 用源文件燃油金额
        "mode": "schedule",
        "carrier": "FedEx",
        "service_group": "Ground",
        # 燃油费率按哪个日期匹配：period（跟账期字段一致）/ 创建时间 / 发货时间 / 完成时间
        # 实测：本平台账单的燃油费率与【创建时间】对应关系最清晰
        "rate_date_field": "创建时间",
        # 费率周的起始星期：0=周一（FedEx 官网口径）。
        # ★ 实测源账单把费率整体后移了一天（周二才换档），本程序按官网周一口径计算，
        #   因此边界日（如 7/6、7/20、7/27）的重算金额会与源文件不同，属于**有意修正**。
        "week_start_weekday": 0,
        "source_timezone": "Asia/Shanghai",
        "billing_timezone": "Asia/Shanghai",
        # 这些尾程物流逐单强制不计燃油（它们本来就没有运费明细）
        "no_fuel_logistics": ["自有快递", "卡派自提"],
        # 费率未覆盖某订单日期时：error=阻断（推荐）/ warn=按 0 计并记异常
        "missing_rate_action": "error",
    },

    # ---- 燃油基数（正数明细列；不存在的列自动忽略） ----
    "fuel_base_fields": [
        "住宅地址附加费",
        "偏远附加费",
        "基础运费",
        "操作附加费",
        "操作附加费-尺寸",
        "超尺寸附加费",
    ],
    # PlatformFee 即使全月为 0 也保留，避免以后出现平台费再次漏计
    "platform_fee_fields": ["PlatformFee"],

    # ---- 5. 费用总计构成（含手工计费列；物流费用用重算值） ----
    "total_fields": [
        "订单处理费",
        "自有运单计费",
        "复合订单处理费",
        "贴标/换标(手动计费)",
        "拣货费(手动计费)",
        "自提出库费(手动计费)",
        "出库打托费(手动计费)",
        "物流费用",
    ],

    # ---- 1/2. 多文件合并 ----
    "merge": {
        "key": ["订单编号", "SKU"],       # 14. 复合订单不能只按订单编号去重
        "supplement_priority": True,
        "strict_structure": True,
        "allow_extra_fields": False,
    },

    # ---- 3. 剔除规则：只有这些情况才允许删除订单 ----
    "drop_rules": {
        "drop_out_of_period": True,
        "test_order_keywords": [],        # 真实账单里暂无测试单标记，按需填写
        "cancelled_status": ["已取消", "作废", "已作废"],
        "free_order_flags": [],
    },

    # ---- 13. 取整 ----
    "rounding": {
        "digits": 2,
        "mode": "ROUND_HALF_UP",
        "order_level": True,
    },
}


# ---------------------------------------------------------------------------
# 各客户差异（只写与默认不同的部分，程序做深度合并）
# ---------------------------------------------------------------------------
CUSTOMER_OVERRIDES: Dict[str, Dict[str, Any]] = {
    "MPAV": {
        "customer_name": "NBKXKJ",
        # 54 列：订单处理费 / 自有运单计费 / 贴标换标 / 拣货费 + 全套附加费
        # 曾出现补充文件 50 列、主文件 54 列按位置粘贴导致整体错位
        "merge": {"strict_structure": True, "allow_extra_fields": False},
    },
    "HBKE": {
        "customer_name": "Xhoome123",
        # 52 列，含 自提出库费(手动计费) / 出库打托费(手动计费)；有 24 单卡派自提
    },
    "MUTN": {
        "customer_name": "KANGDI-cyl",
        # 48 列，含 复合订单处理费；无自有运单计费、无手工计费列
    },
    "IPQI": {
        "customer_name": "yaoxiang",
        # 47 列，最精简：只有 订单处理费 + 物流费用
        # 历史问题：燃油实际在 AS 列，物流公式却引用空白 AV 列 -> 现已全部按表头名定位
    },
    "GUXE": {
        "customer_name": "senbo",
        # 38 列，完全没有燃油/运费明细列，物流费用恒为 0
        "fuel": {"mode": "none"},
        "total_fields": [
            "订单处理费",
            "自有运单计费",
            "复合订单处理费",
            "物流费用",
        ],
        "fuel_base_fields": [],
        "platform_fee_fields": [],
    },
}


# ---------------------------------------------------------------------------
# 工具
# ---------------------------------------------------------------------------

def _deep_merge(base: Dict[str, Any], override: Dict[str, Any]) -> Dict[str, Any]:
    result = copy.deepcopy(base)
    for key, value in override.items():
        if isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = _deep_merge(result[key], value)
        else:
            result[key] = copy.deepcopy(value)
    return result


def list_customer_codes() -> List[str]:
    return sorted(CUSTOMER_OVERRIDES.keys())


def get_customer_config(customer_code: str) -> Dict[str, Any]:
    """取某客户的完整配置（默认配置 + 该客户差异）。未配置的客户返回默认配置。"""
    code = (customer_code or "").strip().upper()
    config = _deep_merge(DEFAULT_CONFIG, {"customer_code": code or "DEFAULT"})
    if code in CUSTOMER_OVERRIDES:
        config = _deep_merge(config, CUSTOMER_OVERRIDES[code])
        config["customer_code"] = code
    return config


def detect_customer_code(candidate_codes: List[str]) -> str:
    """从账单里读到的客户编码中挑一个已配置的，用于自动选择客户。"""
    for code in candidate_codes:
        code = str(code or "").strip().upper()
        if code in CUSTOMER_OVERRIDES:
            return code
    return ""


def validate_config(config: Dict[str, Any]) -> List[str]:
    """配置自检，返回问题清单。"""
    problems: List[str] = []

    fuel_mode = config.get("fuel", {}).get("mode")
    if fuel_mode not in {"none", "schedule", "source"}:
        problems.append(f"fuel.mode 只能是 none / schedule / source，当前为「{fuel_mode}」。")

    if not config.get("total_fields"):
        problems.append("total_fields 为空，费用总计将恒为 0。")

    logistics_field = config["fields"]["logistics_fee"]
    if logistics_field not in config.get("total_fields", []):
        problems.append(
            f"费用总计构成里缺少【{logistics_field}】，物流费用不会被计入总计，请确认是否有意为之。")

    if not config.get("merge", {}).get("key"):
        problems.append("merge.key 为空，无法确定业务主键，合并与去重会不可靠。")

    digits = config.get("rounding", {}).get("digits")
    if not isinstance(digits, int) or not (0 <= digits <= 6):
        problems.append(f"rounding.digits 应为 0-6 的整数，当前为 {digits}。")

    wsw = config.get("fuel", {}).get("week_start_weekday")
    if wsw not in range(7):
        problems.append(f"fuel.week_start_weekday 应为 0-6（0=周一），当前为 {wsw}。")

    return problems


def period_field_for(config: Dict[str, Any], logistics_type: str) -> Tuple[str, bool]:
    """
    返回 (账期时间字段名, 是否走了兜底)。
    对应建议 4：卡派自提按【完成时间】，其余按【创建时间】，不得全表统一。
    """
    rules = config.get("period_rules", {})
    key = (logistics_type or "").strip()
    if key in rules:
        return rules[key], False
    normalized = key.replace(" ", "").replace("\u3000", "")
    for rule_key, field_name in rules.items():
        if rule_key.replace(" ", "") == normalized:
            return field_name, False
    return config.get("period_default_field", "创建时间"), True
