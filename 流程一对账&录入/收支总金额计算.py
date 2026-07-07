# -*- coding: utf-8 -*-
"""
计算对账台账「银行收入流水汇总」的收支总金额。

输出顺序：先按「日期」聚合，每个日期下再列出各「收款渠道」的收入/支出合计。
"""

from pathlib import Path

import pandas as pd


def calc_bank_flow_totals(ledger_path):
    """
    读取对账台账中的银行流水，按日期汇总各渠道收支总金额。

    参数:
        ledger_path: 对账台账 Excel 文件路径（.xlsx）

    返回:
        result: 列表，每个元素代表「某一天」的汇总，结构如下:
            [
                {
                    "日期": "2026-05-20",
                    "明细": [
                        {"收款渠道": "蓝乾中信5701", "收支类型": "收入", "总金额": 152576.90},
                        {"收款渠道": "蓝乾中信5701", "收支类型": "支出", "总金额": 6429.43},
                        ...
                    ],
                },
                ...
            ]
    """
    path = Path(ledger_path)

    # 打开 Excel，查找「银行收入流水」相关工作表
    xl = pd.ExcelFile(path)
    sheet_names = set(xl.sheet_names)
    sheet_flow = next(
        (s for s in ("银行收入流水", "银行收入流水汇总") if s in sheet_names),
        None,
    )
    if not sheet_flow:
        raise ValueError("未找到银行收入流水工作表，当前工作簿包含: %s" % list(xl.sheet_names))

    df = pd.read_excel(path, sheet_name=sheet_flow)

    # 收款渠道列：不同台账可能叫「收款渠道」或「收款方式」
    col_pay_channel = "收款渠道"
    if col_pay_channel not in df.columns and "收款方式" in df.columns:
        col_pay_channel = "收款方式"
    if col_pay_channel not in df.columns:
        raise ValueError("未找到收款渠道/收款方式列")

    col_date = "日期"
    if col_date not in df.columns:
        raise ValueError("未找到日期列")

    # 没有收入/支出列时补 0，避免后面求和报错
    for col in ("收入", "支出"):
        if col not in df.columns:
            df[col] = 0

    def normalize_amount(value):
        """把单元格里的金额转成数字；空值、横线等当作 0。"""
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

    def normalize_date(value):
        """把日期统一成 YYYY-MM-DD 字符串，方便分组比对。"""
        if pd.isna(value):
            return ""
        try:
            return pd.to_datetime(value).strftime("%Y-%m-%d")
        except Exception:
            text = str(value).strip()
            return text if text not in ("nan", "NaN", "None") else ""

    # 清洗三列，保证分组、求和时类型一致
    df["收入"] = df["收入"].map(normalize_amount)
    df["支出"] = df["支出"].map(normalize_amount)
    df[col_pay_channel] = df[col_pay_channel].fillna("").astype(str)
    df[col_date] = df[col_date].map(normalize_date).fillna("").astype(str)

    result = []

    # 第一层：按「日期」聚合（同一天的数据放在一起）
    for date_val, date_df in df.groupby(col_date, sort=False):
        date_str = str(date_val).strip()
        day_items = []

        # 第二层：在同一天内，按「收款渠道」分别算收入合计、支出合计
        for channel, group in date_df.groupby(col_pay_channel, sort=False):
            channel_name = str(channel).strip()
            income_total = round(float(group["收入"].sum()), 2)
            expense_total = round(float(group["支出"].sum()), 2)

            # 收入大于 0 才输出一条「收入」汇总
            if income_total > 0:
                day_items.append(
                    {
                        "收款渠道": channel_name,
                        "收支类型": "收入",
                        "总金额": income_total,
                    }
                )
            # 支出大于 0 才输出一条「支出」汇总
            if expense_total > 0:
                day_items.append(
                    {
                        "收款渠道": channel_name,
                        "收支类型": "支出",
                        "总金额": expense_total,
                    }
                )

        # 这一天有有效汇总数据时，才加入最终结果
        if day_items:
            result.append(
                {
                    "日期": date_str,
                    "明细": day_items,
                }
            )

    return result


if __name__ == "__main__":
    # 直接运行本文件时，用同目录下示例台账做测试
    default_ledger = Path(__file__).resolve().parent / "2026.05" / "对账台账.xlsx"
    if default_ledger.exists():
        totals = calc_bank_flow_totals(default_ledger)
        print("按日期聚合的收支总金额:")
        for day_group in totals:
            print("  日期 %s:" % day_group["日期"])
            for item in day_group["明细"]:
                print(
                    "    [%s][%s] %.2f"
                    % (item["收款渠道"], item["收支类型"], item["总金额"])
                )
    else:
        print("未找到示例台账:", default_ledger)
