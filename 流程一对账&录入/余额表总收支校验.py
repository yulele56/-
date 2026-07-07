# -*- coding: utf-8 -*-
"""
将流水汇总金额与余额表中的账户收入/支出金额比对校验。

余额表列：账户名称、收入金额、支出金额（首行表头，金额可能带千分位逗号）。
账户匹配：对传入的「收款渠道」去首尾空格后，在余额表「账户名称」中做模糊包含匹配。
"""

import re
from pathlib import Path

import pandas as pd

AMT_TOL = 0.02


def _amounts_equal(a, b):
    return abs(float(a) - float(b)) < AMT_TOL


def verify_balance_sheet_totals(balance_path, check_items):
    """
    按收款渠道汇总传入的流水合计，与余额表同账户的收入/支出金额比对。

    参数:
        balance_path: 余额表 Excel 路径（如 1779206400000000余额表数据.xls）
        check_items: 待校验项列表，每项为 dict，例如:
            [
                {"收款渠道": "支付宝", "收支类型": "收入", "总金额": 79014.3},
                {"收款渠道": "支付宝", "收支类型": "支出", "总金额": 79014.3},
                {"收款渠道": "微信", "收支类型": "收入", "总金额": 4742.0},
            ]
            也支持用「账户名称」键代替「收款渠道」。

    返回:
        compare_results: 比对结果列表，每项结构:
            {
                "收款渠道": str,
                "台账收入": float,      # 待校验项汇总收入
                "台账支出": float,      # 待校验项汇总支出
                "余额表账户": str,
                "余额表收入": float,
                "余额表支出": float,
                "结果": "金额一致" | "金额不一致",
                "差额": {               # 仅金额不一致或未找到账户时
                    "收入差额": float,  # 台账收入 - 余额表收入
                    "支出差额": float,
                },
            }
    """
    path = Path(balance_path)
    if not path.is_file():
        raise ValueError("余额表文件不存在: %s" % balance_path)

    # ---------- 读取余额表 ----------
    raw = pd.read_excel(path, header=None)
    if raw.empty:
        raise ValueError("余额表为空: %s" % balance_path)

    def normalize_header(value):
        """去掉表头里的图标等特殊符号。"""
        text = str(value).strip()
        text = re.sub(r"[\uf000-\uf8ff]", "", text)
        return text.strip()

    def normalize_amount(value):
        """金额转 float，逗号、货币符号等去掉。"""
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

    headers = [normalize_header(c) for c in raw.iloc[0].tolist()]
    df = raw.iloc[1:].copy()
    df.columns = headers
    df = df.dropna(how="all")

    def pick_column(columns, must_contain):
        """在列名中找包含指定关键字的列（子串匹配）。"""
        for col in columns:
            text = normalize_header(col)
            if all(k in text for k in must_contain):
                return col
        return None

    col_account = pick_column(df.columns, ("账户",)) or pick_column(df.columns, ("名称",))
    col_income = pick_column(df.columns, ("收入",))
    col_expense = pick_column(df.columns, ("支出",))
    if not col_account or not col_income or not col_expense:
        raise ValueError(
            "余额表缺少必要列（账户名称/收入金额/支出金额），当前列: %s" % list(df.columns)
        )

    # 余额表账户 -> {收入, 支出}
    balance_map = {}
    account_names = []
    for _, row in df.iterrows():
        name = str(row.get(col_account) or "").strip()
        if not name or name in ("nan", "NaN"):
            continue
        balance_map[name] = {
            "收入": normalize_amount(row.get(col_income)),
            "支出": normalize_amount(row.get(col_expense)),
        }
        account_names.append(name)

    # ---------- 汇总传入的校验项（按收款渠道/账户名称） ----------
    def item_channel(item):
        text = item.get("收款渠道")
        if text is None:
            text = item.get("账户名称")
        return str(text or "").strip()

    def item_flow_type(item):
        return str(item.get("收支类型") or "").strip()

    def item_amount(item):
        return normalize_amount(item.get("总金额", item.get("金额", 0)))

    expected = {}
    for item in check_items or []:
        channel = item_channel(item)
        if not channel:
            continue
        if channel not in expected:
            expected[channel] = {"收入": None, "支出": None}
        flow_type = item_flow_type(item)
        amount = item_amount(item)
        if "收入" in flow_type:
            expected[channel]["收入"] = (
                amount
                if expected[channel]["收入"] is None
                else round(expected[channel]["收入"] + amount, 2)
            )
        elif "支出" in flow_type:
            expected[channel]["支出"] = (
                amount
                if expected[channel]["支出"] is None
                else round(expected[channel]["支出"] + amount, 2)
            )

    def find_balance_account(channel):
        """
        模糊匹配余额表账户：去首尾空格后，账户名称包含传入渠道名即视为命中。
        多个命中时优先完全相等，其次优先以渠道名开头的较短账户名。
        """
        channel = channel.strip()
        if not channel:
            return None
        matches = []
        for name in account_names:
            n = name.strip()
            if channel in n or n in channel:
                matches.append(n)
        if not matches:
            return None
        for name in matches:
            if name == channel:
                return name
        matches.sort(key=lambda x: (not x.startswith(channel), len(x)))
        return matches[0]

    # ---------- 逐渠道比对，输出与审计总额计算一致的结构 ----------
    compare_results = []
    for channel in expected:
        ledger_income = expected[channel]["收入"]
        ledger_expense = expected[channel]["支出"]
        if ledger_income is None:
            ledger_income = 0.0
        if ledger_expense is None:
            ledger_expense = 0.0
        ledger_income = round(float(ledger_income), 2)
        ledger_expense = round(float(ledger_expense), 2)

        matched_name = find_balance_account(channel)
        if not matched_name:
            compare_results.append({
                "收款渠道": channel,
                "台账收入": ledger_income,
                "台账支出": ledger_expense,
                "余额表账户": channel,
                "余额表收入": 0.0,
                "余额表支出": 0.0,
                "结果": "金额不一致",
                "差额": {
                    "收入差额": round(ledger_income - 0.0, 2),
                    "支出差额": round(ledger_expense - 0.0, 2),
                    "说明": "余额表未找到该账户",
                },
            })
            continue

        bal = balance_map[matched_name]
        balance_income = bal["收入"]
        balance_expense = bal["支出"]
        income_ok = _amounts_equal(ledger_income, balance_income)
        expense_ok = _amounts_equal(ledger_expense, balance_expense)

        if income_ok and expense_ok:
            compare_results.append({
                "收款渠道": channel,
                "台账收入": ledger_income,
                "台账支出": ledger_expense,
                "余额表账户": matched_name,
                "余额表收入": balance_income,
                "余额表支出": balance_expense,
                "结果": "金额一致",
            })
        else:
            compare_results.append({
                "收款渠道": channel,
                "台账收入": ledger_income,
                "台账支出": ledger_expense,
                "余额表账户": matched_name,
                "余额表收入": balance_income,
                "余额表支出": balance_expense,
                "结果": "金额不一致",
                "差额": {
                    "收入差额": round(ledger_income - balance_income, 2),
                    "支出差额": round(ledger_expense - balance_expense, 2),
                },
            })

    return compare_results


if __name__ == "__main__":
    base = Path(__file__).resolve().parent
    balance_file = base / "2026.05" / "1779206400000000余额表数据.xls"
    sample_items = [
        {"收款渠道": "支付宝", "收支类型": "收入", "总金额": 79014.3},
        {"收款渠道": "支付宝", "收支类型": "支出", "总金额": 79014.3},
        {"收款渠道": "微信", "收支类型": "收入", "总金额": 4742.0},
    ]
    if balance_file.exists():
        for row in verify_balance_sheet_totals(balance_file, sample_items):
            print(row)
    else:
        print("未找到余额表:", balance_file)
