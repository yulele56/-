# -*- coding: utf-8 -*-
"""
判断「流水汇总表.xlsx」中「支付宝流水汇总」是否包含指定日期的流水。

RPA 调用:
    has = alipay_has_flow_on_date("2026-06-28")
    has = alipay_has_flow_on_date("2026-06-28", r"...\流水汇总表.xlsx")
"""

from pathlib import Path

import pandas as pd

DEFAULT_EXCEL = Path(__file__).resolve().parent.parent / "流水汇总表.xlsx"
FALLBACK_EXCEL = Path(__file__).resolve().parent / "流水汇总表.xlsx"
SHEET_NAME = "支付宝流水汇总"


def _resolve_excel(excel_path):
    if excel_path:
        return Path(excel_path)
    if DEFAULT_EXCEL.is_file():
        return DEFAULT_EXCEL
    return FALLBACK_EXCEL


def alipay_has_flow_on_date(query_date, excel_path=None):
    """
    查询支付宝汇总流水是否包含某日记录。

    参数:
        query_date: 日期，如 "2026-06-28"
        excel_path: 流水汇总表路径，默认本目录「流水汇总表.xlsx」

    返回:
        True 有该日流水，False 没有
    """
    path = _resolve_excel(excel_path)
    if not path.is_file():
        raise FileNotFoundError("流水汇总表不存在: %s" % path)

    df = pd.read_excel(path, sheet_name=SHEET_NAME)
    if df.empty or "日期" not in df.columns:
        return False

    target = pd.Timestamp(pd.to_datetime(query_date)).normalize()
    dates = pd.to_datetime(df["日期"], errors="coerce").dt.normalize()
    return bool((dates == target).any())


if __name__ == "__main__":
    print("文件:", _resolve_excel(None))
    print("2026-06-28", alipay_has_flow_on_date("2026-06-28"))
    print("2026-05-26", alipay_has_flow_on_date("2026-05-26"))
