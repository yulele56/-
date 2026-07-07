# -*- coding: utf-8 -*-
"""
审计：按收款渠道汇总对账台账收支，并与余额表比对。
"""

from pathlib import Path

import pandas as pd

from 余额表数据处理 import get_balance_account_amounts

__all__ = ["compare_ledger_balance_totals"]

AMT_TOL = 0.02


def _normalize_amount(value):
    if pd.isna(value):
        return 0.0
    text = str(value).strip()
    if text in ("", "-", "__", "nan", "NaN", "None"):
        return 0.0
    text = (
        text.replace(",", "")
        .replace(" ", "")
        .replace("￥", "")
        .replace("¥", "")
        .replace("元", "")
    )
    try:
        return round(float(text), 2)
    except ValueError:
        return 0.0


def _amounts_equal(a, b):
    return abs(float(a) - float(b)) < AMT_TOL


def _calc_channel_totals(ledger_path):
    """按收款渠道聚合银行收入流水汇总的收入/支出总金额。"""
    path = Path(ledger_path)
    xl = pd.ExcelFile(path)
    sheet_names = set(xl.sheet_names)
    sheet_flow = next(
        (s for s in ("银行收入流水", "银行收入流水汇总") if s in sheet_names),
        None,
    )
    if not sheet_flow:
        raise ValueError("未找到银行收入流水工作表，当前工作簿包含: %s" % list(sheet_names))

    df = pd.read_excel(path, sheet_name=sheet_flow)
    col_channel = "收款渠道"
    if col_channel not in df.columns and "收款方式" in df.columns:
        col_channel = "收款方式"
    if col_channel not in df.columns:
        raise ValueError("未找到收款渠道/收款方式列")

    for col in ("收入", "支出"):
        if col not in df.columns:
            df[col] = 0

    df["收入"] = df["收入"].map(_normalize_amount)
    df["支出"] = df["支出"].map(_normalize_amount)
    df[col_channel] = df[col_channel].fillna("").astype(str).str.strip()

    totals = []
    for channel, group in df.groupby(col_channel, sort=False):
        channel_name = str(channel).strip()
        if not channel_name:
            continue
        income_total = round(float(group["收入"].sum()), 2)
        expense_total = round(float(group["支出"].sum()), 2)
        if income_total <= 0 and expense_total <= 0:
            continue
        totals.append(
            {
                "收款渠道": channel_name,
                "台账收入": income_total,
                "台账支出": expense_total,
            }
        )
    return totals


def compare_ledger_balance_totals(ledger_path, balance_path):
    """
    按收款渠道汇总台账收支，并与余额表同账户名称的收入/支出比对。

    参数:
        ledger_path: 对账台账路径
        balance_path: 余额表路径

    返回:
        channel_totals: 各收款渠道台账收支汇总
        compare_results: 比对结果列表，每项结构:
            {
                "收款渠道": str,
                "台账收入": float,
                "台账支出": float,
                "余额表账户": str,
                "余额表收入": float,
                "余额表支出": float,
                "结果": "金额一致" | "金额不一致",
                "差额": {  # 仅金额不一致或未找到账户时有明细
                    "收入差额": float,  # 台账收入 - 余额表收入
                    "支出差额": float,
                },
            }
    """
    channel_totals = _calc_channel_totals(ledger_path)
    compare_results = []

    for item in channel_totals:
        channel = item["收款渠道"]
        ledger_income = item["台账收入"]
        ledger_expense = item["台账支出"]

        balance = get_balance_account_amounts(balance_path, channel)
        balance_income = balance["收入金额"]
        balance_expense = balance["支出金额"]
        balance_account = balance["账户名称"]

        income_ok = _amounts_equal(ledger_income, balance_income)
        expense_ok = _amounts_equal(ledger_expense, balance_expense)
        found = balance["是否找到"]

        if found and income_ok and expense_ok:
            result = {
                "收款渠道": channel,
                "台账收入": ledger_income,
                "台账支出": ledger_expense,
                "余额表账户": balance_account,
                "余额表收入": balance_income,
                "余额表支出": balance_expense,
                "结果": "金额一致",
            }
        else:
            result = {
                "收款渠道": channel,
                "台账收入": ledger_income,
                "台账支出": ledger_expense,
                "余额表账户": balance_account,
                "余额表收入": balance_income,
                "余额表支出": balance_expense,
                "结果": "金额不一致",
                "差额": {
                    "收入差额": round(ledger_income - balance_income, 2),
                    "支出差额": round(ledger_expense - balance_expense, 2),
                },
            }
            if not found:
                result["差额"]["说明"] = "余额表未找到该账户"
        compare_results.append(result)

    return channel_totals, compare_results


if __name__ == "__main__":
    base = Path(__file__).resolve().parent.parent / "流程一对账&录入" / "2026.05"
    balance = base / "1779206400000000余额表数据.xls"
    ledger = base / "对账台账.xlsx"

    if balance.exists():
        info = get_balance_account_amounts(balance, "蓝乾中信5701")
        print("余额表查询:", info)

    if balance.exists() and ledger.exists():
        totals, results = compare_ledger_balance_totals(ledger, balance)
        print("渠道汇总 %d 个，比对 %d 个" % (len(totals), len(results)))
        for row in results:
            print(row)
    elif not ledger.exists():
        print("未找到示例台账:", ledger)
