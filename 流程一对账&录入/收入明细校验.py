# -*- coding: utf-8 -*-
"""
银行收入流水与记账明细核对。
"""

import re
import traceback
from pathlib import Path

import pandas as pd


def verify_bank_income_flow_details(ledger_path):
    """
    逐条校验银行收入流水是否在记账明细中有对应分录。

    返回:
        df_new: 未完全匹配的流水
        result: 未匹配流水按收款渠道分组
        contract_index_list: [(行索引, "匹配成功")]
        result_flags: [(行索引, 核验结果, 金额, 收款渠道, 对方账户, 结算对象)]
    """
    df_new = pd.DataFrame()
    result = []
    contract_index_list = []
    result_flags = []

    try:
        # -------------------------------
        # 1. 核验状态常量
        # -------------------------------
        COL_VERIFY_STATUS = "核验结果"
        STATUS_OK = "匹配成功"
        STATUS_FAIL = "匹配失败"
        STATUS_NET_FAIL = "净额匹配失败"
        STATUS_TAX_FAIL = "税额匹配失败"
        FAIL_STATUSES = (STATUS_FAIL, STATUS_NET_FAIL, STATUS_TAX_FAIL)

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
            text = text.replace(",", "").replace(" ", "").replace("￥", "").replace("¥", "").replace("元", "")
            try:
                return round(float(text), 2)
            except ValueError:
                return None

        def normalize_date(value):
            if pd.isna(value):
                return None
            text = str(value).strip()
            if re.fullmatch(r"\d{8}\.0+", text):
                text = text.split(".")[0]
            if re.fullmatch(r"\d{8}", text):
                return "%s-%s-%s" % (text[0:4], text[4:6], text[6:8])
            try:
                return pd.to_datetime(value).strftime("%Y-%m-%d")
            except Exception:
                return None

        def normalize_text(value):
            if pd.isna(value):
                return ""
            return str(value).strip()

        def channel_match(flow_channel, detail_channel):
            a, b = normalize_text(flow_channel), normalize_text(detail_channel)
            if not a or not b:
                return False
            return a in b or b in a

        def detail_has_line(detail_df, flow_date, flow_ch, amount):
            if flow_date is None or amount is None:
                return False
            exp = round(float(amount), 2)
            for _, d in detail_df.iterrows():
                if d["_date"] != flow_date:
                    continue
                if abs(d["_amt"] - exp) >= 0.02:
                    continue
                if not channel_match(flow_ch, d["_settle"]):
                    continue
                return True
            return False

        # -------------------------------
        # 2. 读取 Excel，定位工作表
        # -------------------------------
        path = Path(ledger_path)
        xl = pd.ExcelFile(path)
        names = set(xl.sheet_names)

        sheet_flow = next(
            (s for s in ("银行收入流水", "银行收入流水汇总", "收入流水汇总", "收入流水") if s in names),
            None,
        )
        sheet_detail = next((s for s in ("记账明细", "对账明细") if s in names), None)
        if not sheet_flow or not sheet_detail:
            raise ValueError("工作表缺失: %s" % list(xl.sheet_names))

        # -------------------------------
        # 3. 读取并清洗流水
        # -------------------------------
        df_flow = pd.read_excel(path, sheet_name=sheet_flow)
        col_channel = "收款渠道" if "收款渠道" in df_flow.columns else "收款方式"

        for col in (COL_VERIFY_STATUS, "结算对象", "账户类别"):
            if col not in df_flow.columns:
                df_flow[col] = None
            df_flow[col] = df_flow[col].astype(object)

        if "日期" in df_flow.columns:
            df_flow["日期"] = df_flow["日期"].map(normalize_date)
        for col in ("金额", "收入", "税额"):
            if col in df_flow.columns:
                df_flow[col] = df_flow[col].map(
                    lambda x: normalize_amount(x) if col != "税额" else (normalize_amount(x) or 0.0)
                )
        df_flow = df_flow[df_flow["收入"].notna() & (df_flow["收入"] > 0)].copy()

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
        if not col_settle or "收入金额(借)" not in df_detail.columns:
            raise ValueError("记账明细缺少必要列")

        detail = df_detail.copy()
        detail["_date"] = detail["记账日期"].map(normalize_date)
        detail["_amt"] = detail["收入金额(借)"].map(normalize_amount).fillna(0)
        detail = detail[detail["_amt"] > 0]
        detail["_settle"] = detail[col_settle].map(normalize_text)

        # -------------------------------
        # 5. 逐条核验
        # -------------------------------
        contract_index_list = []
        result_flags = []

        for idx, row in df_flow.iterrows():
            gross = normalize_amount(row.get("收入"))
            if not gross or gross <= 0:
                continue

            flow_date = row.get("日期")
            flow_ch = row.get(col_channel, "")
            is_tax = normalize_text(row.get("账户类别")) == "算税"

            if not is_tax:
                amt = normalize_amount(row.get("金额"))
                if amt is None:
                    amt = gross
                if detail_has_line(detail, flow_date, flow_ch, amt):
                    status = STATUS_OK
                else:
                    status = STATUS_FAIL
            else:
                net = normalize_amount(row.get("金额"))
                tax = normalize_amount(row.get("税额"))
                if net is None:
                    net = round(gross * 0.9, 2)
                if tax is None:
                    tax = round(gross * 0.1, 2)

                net_ok = detail_has_line(detail, flow_date, flow_ch, net)
                tax_ok = detail_has_line(detail, flow_date, flow_ch, tax)
                if net_ok and tax_ok:
                    status = STATUS_OK
                elif net_ok:
                    status = STATUS_NET_FAIL
                elif tax_ok:
                    status = STATUS_TAX_FAIL
                else:
                    status = STATUS_FAIL

            df_flow.at[idx, COL_VERIFY_STATUS] = status
            if status == STATUS_OK:
                contract_index_list.append((idx, STATUS_OK))

            result_flags.append((
                idx,
                status,
                empty_if_na(row.get("收入")),
                empty_if_na(row.get(col_channel)),
                empty_if_na(row.get("对方账户")),
                empty_if_na(row.get("结算对象")),
            ))

        # -------------------------------
        # 6. 未匹配流水 + 分组
        # -------------------------------
        df_new = df_flow[df_flow[COL_VERIFY_STATUS].isin(FAIL_STATUSES)].copy()
        if df_new is not None and not df_new.empty:
            df_new = df_new.copy().fillna("")

        result = []
        if not df_new.empty:
            df_new[col_channel] = df_new[col_channel].fillna("")
            for channel, grp in df_new.groupby(col_channel, sort=False):
                records = grp.to_dict(orient="records")
                result.append({
                    col_channel: empty_if_na(channel),
                    "明细": [{k: empty_if_na(v) for k, v in rec.items()} for rec in records],
                })

    except Exception as e:
        error_msg = "收入明细校验出错：%s\n%s" % (str(e), traceback.format_exc())
        print(error_msg)

    return df_new, result, contract_index_list, result_flags


if __name__ == "__main__":
    ledger = Path(__file__).resolve().parent / "2026.05" / "对账台账.xlsx"
    if ledger.exists():
        df_new, result, ok_list, flags = verify_bank_income_flow_details(ledger)
        print("未完全匹配 %d 条，匹配成功 %d 条" % (len(df_new), len(ok_list)))
    else:
        print("未找到:", ledger)
