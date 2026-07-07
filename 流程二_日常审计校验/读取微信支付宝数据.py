# -*- coding: utf-8 -*-
"""
从「流水汇总表.xlsx」读取指定时间段内的微信、支付宝流水，按平台与收支类型分组输出。
"""

from pathlib import Path

import pandas as pd

__all__ = ["get_wx_alipay_flows"]

# 默认台账路径（与脚本同目录）
DEFAULT_EXCEL = Path(__file__).resolve().parent / "流水汇总表.xlsx"

# sheet 名 → 收款渠道
SHEET_CHANNEL = {
    "微信流水汇总": "微信5071W",
    "支付宝流水汇总": "支付宝777",
}

DETAIL_COLS = ("日期", "对方账户", "收入", "支出", "收款渠道", "备注", "收支类型")
GROUP_COLS = ("收款渠道", "收支类型")
# 输出顺序：先支出，后收入；先微信，后支付宝
GROUP_ORDER = (
    ("微信5071W", "支出"),
    ("微信5071W", "收入"),
    ("支付宝777", "支出"),
    ("支付宝777", "收入"),
)


def _parse_bound_date(value, label):
    """
    将 RPA 传入的起止日期转为当日 0 点的 Timestamp。
    解析失败时抛出明确错误（避免 NaT.normalize() 报 NaTType 无 normalize）。
    """
    if value is None:
        raise ValueError("%s 为空，请传入如 2025-01-01" % label)
    try:
        if pd.isna(value):
            raise ValueError("%s 为空，请传入如 2025-01-01" % label)
    except (TypeError, ValueError):
        pass

    text = str(value).strip()
    if text in ("", "nan", "NaN", "None", "nat", "NaT"):
        raise ValueError("%s 无效: %r" % (label, value))

    ts = pd.to_datetime(value, errors="coerce")
    if pd.isna(ts):
        raise ValueError("%s 无法解析为日期: %r" % (label, value))
    return pd.Timestamp(ts).normalize()


def get_wx_alipay_flows(start_date, end_date, excel_path=None):
    """
    读取流水汇总表，筛选 [start_date, end_date] 内的微信/支付宝流水并分组。

    参数:
        start_date: 起始日期，如 "2025-01-01"
        end_date:   结束日期，如 "2025-12-31"
        excel_path: Excel 路径，默认脚本目录下「流水汇总表.xlsx」

    返回:
        df_all:  筛选后的全部流水（空值已 fillna("")）
        result:  分组列表，每项为:
            {
                "收款渠道": "微信5071W" 或 "支付宝777",
                "收支类型": "收入" 或 "支出",
                "明细": [ {日期, 对方账户, 收入, 支出, 收款渠道, 备注, 收支类型}, ... ],
            }
    """
    path = Path(excel_path) if excel_path else DEFAULT_EXCEL
    if not path.exists():
        raise FileNotFoundError("流水汇总表不存在: %s" % path)

    start_ts = _parse_bound_date(start_date, "起始日期")
    end_ts = _parse_bound_date(end_date, "结束日期")

    parts = []
    xl = pd.ExcelFile(path)

    # 逐个 sheet 读取，补上收款渠道列
    for sheet_name, channel in SHEET_CHANNEL.items():
        if sheet_name not in xl.sheet_names:
            continue
        df = pd.read_excel(path, sheet_name=sheet_name)
        df["收款渠道"] = channel
        parts.append(df)

    if not parts:
        return pd.DataFrame(), []

    df_all = pd.concat(parts, ignore_index=True)

    # 统一日期格式，并按时间段筛选（闭区间）
    df_all["日期"] = pd.to_datetime(df_all["日期"], errors="coerce")
    df_all = df_all[df_all["日期"].notna()].copy()
    df_all = df_all[(df_all["日期"] >= start_ts) & (df_all["日期"] <= end_ts)].copy()
    df_all["日期"] = df_all["日期"].dt.strftime("%Y-%m-%d")

    if df_all.empty:
        return df_all.fillna(""), []

    # 金额转数字，便于判断收入/支出方向
    df_all["收入"] = pd.to_numeric(df_all["收入"], errors="coerce").fillna(0)
    df_all["支出"] = pd.to_numeric(df_all["支出"], errors="coerce").fillna(0)

    rows = []

    # 一行若只有收入或只有支出，拆成标准记录；收入/支出列互斥一侧留空串
    for _, row in df_all.iterrows():
        date_val = row["日期"]
        cp = row.get("对方账户", "")
        channel = row["收款渠道"]
        remark = row.get("备注", "")
        income_val = round(float(row["收入"]), 2) if row["收入"] > 0 else 0
        expense_val = round(float(row["支出"]), 2) if row["支出"] > 0 else 0

        if income_val <= 0 and expense_val <= 0:
            continue
        if income_val > 0 and expense_val > 0:
            if income_val >= expense_val:
                expense_val = 0
            else:
                income_val = 0

        if income_val > 0:
            rows.append({
                "日期": date_val,
                "对方账户": cp,
                "收入": income_val,
                "支出": "",
                "收款渠道": channel,
                "备注": remark,
                "收支类型": "收入",
            })
        if expense_val > 0:
            rows.append({
                "日期": date_val,
                "对方账户": cp,
                "收入": "",
                "支出": expense_val,
                "收款渠道": channel,
                "备注": remark,
                "收支类型": "支出",
            })

    df_all = pd.DataFrame(rows)
    if df_all.empty:
        return df_all, []

    # 所有空值统一为空字符串，避免 RPA 读到 nan
    df_all = df_all.fillna("")

    # 按收款渠道 + 收支类型分组
    grouped = {}
    for (channel, flow_type), grp in df_all.groupby(list(GROUP_COLS), sort=False):
        grouped[(channel, flow_type)] = grp[list(DETAIL_COLS)].fillna("").to_dict(orient="records")

    result = []
    for key in GROUP_ORDER:
        if key not in grouped:
            continue
        channel, flow_type = key
        result.append({
            "收款渠道": channel,
            "收支类型": flow_type,
            "明细": grouped[key],
        })

    return df_all, result


if __name__ == "__main__":
    df, groups = get_wx_alipay_flows("2025-01-01", "2025-12-31")
    print("共 %d 条，分组 %d 组" % (len(df), len(groups)))
    for g in groups:
        print("  [%s][%s] %d 条" % (g["收款渠道"], g["收支类型"], len(g["明细"])))
    if not df.empty:
        print(df.head(5).to_string(index=False))
