# -*- coding: utf-8 -*-
"""
处理对账台账：读取银行收入流水，按收款渠道与对方账户判断是否算税，比对记账明细后标记录入状态。
"""

import re
import traceback
from pathlib import Path

import pandas as pd

from 结算对象获取 import get_settlement_object


def process_bank_reconciliation_ledger(ledger_path):
    """
    读取对账台账银行收入流水，判断算税/不算税，比对记账明细后标记「是否录入」。

    返回:
        df_new: 待录入流水
        result: 按收款渠道分组的明细列表
        contract_index_list: [(行索引, "已录入")]
        result_flags: [(行索引, 算税/不算税, 录入标记, 金额, 税额, 结算对象)]
    """
    df_new = pd.DataFrame()
    result = []
    contract_index_list = []
    result_flags = []

    try:
        # -------------------------------
        # 1. 常量：算税分类规则
        # -------------------------------
        NO_TAX_CHANNEL_ACCOUNTS = (
            ("微信", None),
            ("支付宝", None),
            ("李希", "中行"),
            ("谢纵生", "华兴"),
            ("蓝乾", "华兴"),
            ("蓝乾", "中信"),
            ("李茂铎", "建行"),
            ("李希", "民生澄海支行"),
        )
        NON_PUBLIC_CHANNELS = (
            "微信5071W",
            "支付宝777",
            "希兄中行3364",
            "生兄华兴0201",
            "蓝乾华兴0219",
            "蓝乾中信5701",
            "茂铎建行8870",
            "茂锋建行8870",
            "希兄民生6603",
        )
        SHAREHOLDER_COUNTERPARTY_NAMES = ("李希", "谢纵生", "蓝乾", "李茂铎")

        def empty_if_na(value):
            if value is None:
                return ""
            try:
                if pd.isna(value):
                    return ""
            except (TypeError, ValueError):
                pass
            return value

        def normalize_amount(value):
            if pd.isna(value):
                return None
            text = str(value).strip()
            if not text:
                return None
            text = (
                text.replace(",", "")
                .replace(" ", "")
                .replace("￥", "")
                .replace("¥", "")
                .replace("元", "")
            )
            return round(float(text), 2)

        def normalize_date(value):
            if pd.isna(value):
                return None
            text = str(value).strip()
            if re.fullmatch(r"\d{8}\.0+", text):
                text = text.split(".")[0]
            if re.fullmatch(r"\d{8}", text):
                return "%s-%s-%s" % (text[0:4], text[4:6], text[6:8])
            return pd.to_datetime(value).strftime("%Y-%m-%d")

        # -------------------------------
        # 2. 读取 Excel，定位工作表
        # -------------------------------
        path = Path(ledger_path)
        xl = pd.ExcelFile(path)
        sheet_names = set(xl.sheet_names)

        sheet_flow = next(
            (s for s in ("银行收入流水", "银行收入流水汇总", "收入流水汇总") if s in sheet_names),
            None,
        )
        sheet_detail = next(
            (s for s in ("记账明细", "对账明细") if s in sheet_names),
            None,
        )
        if not sheet_flow or not sheet_detail:
            raise ValueError("工作表缺失，当前工作簿包含: %s" % list(xl.sheet_names))

        # -------------------------------
        # 3. 读取并清洗银行流水（仅收入>0）
        # -------------------------------
        df_flow = pd.read_excel(path, sheet_name=sheet_flow)
        df_flow = df_flow.dropna(subset=["收入"]).copy()

        if "日期" in df_flow.columns:
            df_flow["日期"] = df_flow["日期"].map(normalize_date)
        if "收入" in df_flow.columns:
            df_flow["收入"] = df_flow["收入"].map(normalize_amount)
            df_flow = df_flow.dropna(subset=["收入"]).copy()
            df_flow = df_flow[df_flow["收入"] > 0].copy()

        col_pay_channel = "收款方式"
        if col_pay_channel not in df_flow.columns and "收款渠道" in df_flow.columns:
            col_pay_channel = "收款渠道"

        for col in ("是否录入", "账户类别", "金额", "税额"):
            if col not in df_flow.columns:
                df_flow[col] = None
            df_flow[col] = df_flow[col].astype(object)

        # -------------------------------
        # 4. 读取并清洗记账明细
        # -------------------------------
        df_detail = pd.read_excel(path, sheet_name=sheet_detail, header=1)
        col_settle = None
        for col in df_detail.columns:
            if "结算账户" in str(col):
                col_settle = col
                break
        if not col_settle:
            for col in df_detail.columns:
                if "银行账户" in str(col):
                    col_settle = col
                    break
        if not col_settle:
            raise ValueError("记账明细缺少结算账户/银行账户列")

        detail_income = df_detail.copy()
        detail_income["_amt"] = detail_income["收入金额(借)"].map(normalize_amount).fillna(0)
        detail_income = detail_income[detail_income["_amt"] > 0].copy()
        detail_income["记账日期"] = detail_income["记账日期"].map(normalize_date)
        detail_income["_date"] = detail_income["记账日期"]
        detail_income["_amt_norm"] = detail_income["_amt"]
        detail_income["_settle"] = detail_income[col_settle].map(
            lambda x: str(x).strip() if pd.notna(x) else ""
        )

        # -------------------------------
        # 5. 算税分类 + 入账比对（内联辅助函数）
        # -------------------------------
        def is_no_tax_channel(channel_text):
            channel = str(channel_text or "").strip()
            if not channel:
                return False
            for account_name, bank_name in NO_TAX_CHANNEL_ACCOUNTS:
                if account_name not in channel:
                    continue
                if not bank_name:
                    return True
                if bank_name in channel:
                    return True
            return False

        def is_non_public_channel(channel_text):
            channel = str(channel_text or "").strip()
            if not channel:
                return False
            if channel in NON_PUBLIC_CHANNELS:
                return True
            matches = [name for name in NON_PUBLIC_CHANNELS if name in channel or channel in name]
            return len(matches) == 1

        def is_shareholder_counterparty(counterparty_text):
            name = str(counterparty_text or "").strip()
            if not name:
                return False
            return any(n in name for n in SHAREHOLDER_COUNTERPARTY_NAMES)

        def classify_tax_label(row):
            channel = str(row[col_pay_channel] if col_pay_channel in row.index else "").strip()
            cp = str(row.get("对方账户", "")).strip() if pd.notna(row.get("对方账户")) else ""
            if is_no_tax_channel(channel):
                return "不算税"
            if is_non_public_channel(channel):
                return "不算税"
            if is_shareholder_counterparty(cp):
                return "不算税"
            if not cp:
                return "不算税"
            return "算税"

        def channel_match(flow_channel, settle_account):
            a = str(flow_channel or "").strip()
            b = str(settle_account or "").strip()
            if not a or not b:
                return False
            return a in b or b in a

        def detail_has_line(flow_date, flow_channel, amount):
            if flow_date is None or amount is None:
                return False
            exp = round(float(amount), 2)
            for _, d in detail_income.iterrows():
                if d["_date"] != flow_date:
                    continue
                if abs(d["_amt_norm"] - exp) >= 0.02:
                    continue
                if not channel_match(flow_channel, d["_settle"]):
                    continue
                return True
            return False

        def is_recorded(flow_row, tax_label, net_amt, tax_amt):
            flow_date = flow_row["日期"]
            flow_channel = str(
                flow_row[col_pay_channel] if col_pay_channel in flow_row.index else ""
            ).strip()
            if tax_label == "算税":
                if net_amt is None or tax_amt is None:
                    return False
                return detail_has_line(flow_date, flow_channel, net_amt) and detail_has_line(
                    flow_date, flow_channel, tax_amt
                )
            if net_amt is None:
                return False
            return detail_has_line(flow_date, flow_channel, net_amt)

        # -------------------------------
        # 6. 逐行处理：算税/比对/标记
        # -------------------------------
        contract_index_list = []
        result_flags = []

        for idx, row in df_flow.iterrows():
            gross_income = row["收入"]
            tax_label = classify_tax_label(row)

            if tax_label == "算税":
                tax_amt = round(gross_income * 0.1, 2)
                net_amount = round(gross_income * 0.9, 2)
            else:
                tax_amt = 0
                net_amount = gross_income

            df_flow.at[idx, "账户类别"] = tax_label
            df_flow.at[idx, "税额"] = tax_amt
            df_flow.at[idx, "金额"] = net_amount
            settle_obj = get_settlement_object(path, row.get("对方账户", ""))

            if is_recorded(row, tax_label, net_amount, tax_amt):
                df_flow.at[idx, "是否录入"] = "已录入"
                contract_index_list.append((idx, "已录入"))
                result_flags.append((
                    idx,
                    empty_if_na(tax_label),
                    "已录入",
                    empty_if_na(net_amount),
                    empty_if_na(tax_amt),
                    empty_if_na(settle_obj),
                ))
            else:
                df_flow.at[idx, "是否录入"] = "待录入"
                result_flags.append((
                    idx,
                    empty_if_na(tax_label),
                    "待录入",
                    empty_if_na(net_amount),
                    empty_if_na(tax_amt),
                    empty_if_na(settle_obj),
                ))

        # -------------------------------
        # 7. 待录入流水 + 结算对象
        # -------------------------------
        df_new = df_flow[df_flow["是否录入"] == "待录入"].copy()

        if not df_new.empty:
            if "结算对象" not in df_new.columns:
                df_new["结算对象"] = None
            df_new["结算对象"] = df_new["对方账户"].apply(
                lambda x: get_settlement_object(path, x)
            )

        # -------------------------------
        # 8. 按收款渠道分组输出
        # -------------------------------
        result = []
        if not df_new.empty:
            if col_pay_channel not in df_new.columns:
                df_new[col_pay_channel] = ""
            df_new[col_pay_channel] = df_new[col_pay_channel].fillna("")

            for group_keys, group_df in df_new.groupby([col_pay_channel], sort=False):
                if not isinstance(group_keys, tuple):
                    group_keys = (group_keys,)
                record = dict(zip([col_pay_channel], [empty_if_na(k) for k in group_keys]))
                records = group_df.to_dict(orient="records")
                record["明细"] = [{k: empty_if_na(v) for k, v in rec.items()} for rec in records]
                result.append(record)

        if df_new is not None and not df_new.empty:
            df_new = df_new.copy().fillna("")

    except Exception as e:
        error_msg = "处理银行台账出错：%s\n%s" % (str(e), traceback.format_exc())
        print(error_msg)

    return df_new, result, contract_index_list, result_flags


if __name__ == "__main__":
    default_ledger = Path(__file__).resolve().parent / "2026.05" / "对账台账.xlsx"
    if default_ledger.exists():
        df_new, result, contract_index_list, result_flags = (
            process_bank_reconciliation_ledger(default_ledger)
        )
        print("待录入银行流水 %d 条" % len(df_new))
    else:
        print("未找到示例台账:", default_ledger)
