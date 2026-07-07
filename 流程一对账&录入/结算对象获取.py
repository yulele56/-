# -*- coding: utf-8 -*-
"""从对账台账结算对象查询参考中，按对方账户匹配结算单位。"""

from pathlib import Path

import pandas as pd

_MAPPING_CACHE = {}
DEFAULT_UNKNOWN = "待认"
SHEET_SETTLEMENT_REF = "结算对象查询参考"


def _load_settlement_mapping(ledger_path):
    """读取摘要「/」后字段 → 结算单位 的映射表。"""
    path = Path(ledger_path)
    xl = pd.ExcelFile(path)
    if SHEET_SETTLEMENT_REF not in xl.sheet_names:
        return {}

    df_ref = pd.read_excel(path, sheet_name=SHEET_SETTLEMENT_REF, header=1)
    if "摘要" not in df_ref.columns or "结算单位" not in df_ref.columns:
        return {}

    mapping = {}
    for _, row in df_ref.iterrows():
        summary = row.get("摘要")
        unit = row.get("结算单位")
        if pd.isna(summary) or pd.isna(unit):
            continue
        summary_text = str(summary).strip()
        if "/" not in summary_text:
            continue
        suffix = summary_text.rsplit("/", 1)[-1].strip()
        if suffix:
            mapping[suffix] = str(unit).strip()
    return mapping


def get_settlement_object(ledger_path, counterparty):
    """
    根据对方账户，与结算对象查询参考中摘要「/」后的字段匹配，返回结算单位。

    参数:
        ledger_path: 对账台账 Excel 路径
        counterparty: 银行流水中的对方账户

    返回:
        匹配到的结算单位；无法匹配时返回「待认」
    """
    path_key = str(Path(ledger_path).resolve())
    if path_key not in _MAPPING_CACHE:
        _MAPPING_CACHE[path_key] = _load_settlement_mapping(ledger_path)

    name = str(counterparty or "").strip()
    if not name:
        return DEFAULT_UNKNOWN
    return _MAPPING_CACHE[path_key].get(name, DEFAULT_UNKNOWN)
