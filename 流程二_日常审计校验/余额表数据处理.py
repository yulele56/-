# -*- coding: utf-8 -*-
"""
读取余额表 Excel，按账户名称查询收入/支出金额。
"""

from pathlib import Path

import pandas as pd

__all__ = ["get_balance_account_amounts"]


def _normalize_header(value):
    text = str(value).replace("\uf0c9", "").strip()
    return text


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


def _read_balance_sheet(balance_path):
    path = Path(balance_path)
    if not path.exists():
        raise FileNotFoundError("余额表不存在: %s" % path)

    df = pd.read_excel(path, sheet_name=0, header=0)
    rename = {_col: _normalize_header(_col) for _col in df.columns}
    df = df.rename(columns=rename)

    col_account = next((c for c in df.columns if "账户名称" in c), None)
    col_income = next((c for c in df.columns if "收入金额" in c), None)
    col_expense = next((c for c in df.columns if "支出金额" in c), None)
    if not col_account or not col_income or not col_expense:
        raise ValueError("余额表缺少账户名称/收入金额/支出金额列，当前列: %s" % list(df.columns))

    df = df.copy()
    df["_account"] = df[col_account].map(lambda x: "" if pd.isna(x) else str(x).strip())
    df["_income"] = df[col_income].map(_normalize_amount)
    df["_expense"] = df[col_expense].map(_normalize_amount)
    return df[df["_account"] != ""]


def _find_account_row(df, account_name):
    """精确匹配账户名；否则唯一模糊匹配（互含）。"""
    name = str(account_name or "").strip()
    if not name:
        return None

    exact = df[df["_account"] == name]
    if len(exact) == 1:
        return exact.iloc[0]
    if len(exact) > 1:
        return exact.iloc[0]

    fuzzy = []
    for _, row in df.iterrows():
        acc = row["_account"]
        if acc and (acc in name or name in acc):
            fuzzy.append(row)
    if len(fuzzy) == 1:
        return fuzzy[0]
    return None


def get_balance_account_amounts(balance_path, account_name):
    """
    根据账户名称查询余额表中的收入金额、支出金额。

    参数:
        balance_path: 余额表 Excel 路径（.xls / .xlsx）
        account_name: 账户名称（可与收款渠道同名，如「蓝乾中信5701」）

    返回:
        {
            "账户名称": str,
            "收入金额": float,
            "支出金额": float,
            "是否找到": bool,
        }
        未找到时 收入金额/支出金额为 0.0，是否找到为 False。
    """
    df = _read_balance_sheet(balance_path)
    row = _find_account_row(df, account_name)
    if row is None:
        return {
            "账户名称": str(account_name or "").strip(),
            "收入金额": 0.0,
            "支出金额": 0.0,
            "是否找到": False,
        }
    return {
        "账户名称": row["_account"],
        "收入金额": float(row["_income"]),
        "支出金额": float(row["_expense"]),
        "是否找到": True,
    }


if __name__ == "__main__":
    sample = (
        Path(__file__).resolve().parent.parent
        / "流程一对账&录入"
        / "2026.05"
        / "1779206400000000余额表数据.xls"
    )
    if sample.exists():
        for name in ("蓝乾中信5701", "公户中行7475", "不存在账户"):
            info = get_balance_account_amounts(sample, name)
            print(name, "->", info)
    else:
        print("未找到示例余额表:", sample)
