# -*- coding: utf-8 -*-
"""
处理对账台账：对比支付宝/微信收入流水与记账明细，标记是否录入，并生成分组明细。
"""

from pathlib import Path

import pandas as pd

COL_RECORD_STATUS = "是否录入"
STATUS_RECORDED = "已录入"
STATUS_PENDING = "待录入"


def process_reconciliation_ledger(ledger_path):
    """
    读取对账台账，比对支付宝/微信收入流水与记账明细，标记录入状态并输出待录入明细。

    比对规则（满足任一即视为已入账）:
        1. 记账日期 + 收入金额 与流水完全一致
        2. 收入金额相同且记账明细中仅有一条该金额的收入
        3. 收入金额相同且流水「对方账户」出现在对应记账摘要中

    返回（顺序与 process_invoice_excel 一致）:
        df_new: 处理后的 DataFrame（只包含待录入流水）
        result: 分组明细列表（按收款方式分组，每组含「明细」）
        contract_index_list: [(行索引, "已录入")]
        result_flags: [(行索引, "待录入", 收入, 税额)]（只包含待录入）
    """
    path = Path(ledger_path)
    xl = pd.ExcelFile(path)
    sheet_names = set(xl.sheet_names)

    sheet_flow = next(
        (s for s in ("支付宝微信收入流水", "支付宝&微信收入流水") if s in sheet_names),
        None,
    )
    sheet_detail = next(
        (s for s in ("记账明细", "对账明细") if s in sheet_names),
        None,
    )
    if not sheet_flow or not sheet_detail:
        raise ValueError(f"工作表缺失，当前工作簿包含: {list(xl.sheet_names)}")

    # 读取收入流水，去掉金额为空的占位行
    df_flow = pd.read_excel(path, sheet_name=sheet_flow)
    df_flow = df_flow.dropna(subset=["金额"]).copy()
    if COL_RECORD_STATUS not in df_flow.columns:
        df_flow[COL_RECORD_STATUS] = None
    df_flow[COL_RECORD_STATUS] = df_flow[COL_RECORD_STATUS].astype(object)

    # 读取记账明细并筛出收入分录
    df_detail = pd.read_excel(path, sheet_name=sheet_detail, header=1)
    detail_income = df_detail.copy()
    detail_income["_amt"] = pd.to_numeric(
        detail_income["收入金额(借)"], errors="coerce"
    ).fillna(0)
    detail_income = detail_income[detail_income["_amt"] > 0].copy()
    detail_income["_date"] = detail_income["记账日期"].apply(
        lambda v: None if pd.isna(v) else pd.to_datetime(v).strftime("%Y-%m-%d")
    )
    detail_income["_amt_norm"] = detail_income["_amt"].apply(
        lambda v: round(float(v), 2)
    )

    def is_recorded(flow_row):
        flow_date = (
            None
            if pd.isna(flow_row["日期"])
            else pd.to_datetime(flow_row["日期"]).strftime("%Y-%m-%d")
        )
        flow_amt = round(float(flow_row["金额"]), 2)
        counterparty = ""
        if pd.notna(flow_row.get("对方账户")):
            counterparty = str(flow_row["对方账户"]).strip()

        if flow_date is not None:
            exact = detail_income[
                (detail_income["_date"] == flow_date)
                & (detail_income["_amt_norm"] == flow_amt)
            ]
            if not exact.empty:
                return True

        same_amount = detail_income[detail_income["_amt_norm"] == flow_amt]
        if same_amount.empty:
            return False
        if len(same_amount) == 1:
            return True

        for _, detail_row in same_amount.iterrows():
            summary = str(detail_row["摘要"]).strip()
            if not counterparty or not summary:
                continue
            if counterparty in summary:
                return True
            core = counterparty.replace("*", "").strip()
            if len(core) >= 2 and core in summary:
                return True
        return False

    contract_index_list = []
    result_flags = []

    for idx, row in df_flow.iterrows():
        flow_amt = round(float(row["金额"]), 2)
        if is_recorded(row):
            df_flow.at[idx, COL_RECORD_STATUS] = STATUS_RECORDED
            contract_index_list.append((idx, STATUS_RECORDED))
        else:
            df_flow.at[idx, COL_RECORD_STATUS] = STATUS_PENDING
            result_flags.append((idx, STATUS_PENDING, flow_amt, 0))

    # 只包含待录入流水（对应开票逻辑中的「不行」）
    df_new = df_flow[df_flow[COL_RECORD_STATUS] == STATUS_PENDING].copy()

    result = []
    if not df_new.empty:
        if "收款方式" not in df_new.columns:
            df_new["收款方式"] = ""
        df_new["收款方式"] = df_new["收款方式"].fillna("")

        output_group_cols = ["收款方式"]
        for group_keys, group_df in df_new.groupby(output_group_cols, sort=False):
            if not isinstance(group_keys, tuple):
                group_keys = (group_keys,)
            record = dict(zip(output_group_cols, group_keys))
            record["明细"] = group_df.to_dict(orient="records")
            result.append(record)

    return df_new, result, contract_index_list, result_flags


if __name__ == "__main__":
    default_ledger = Path(__file__).resolve().parent / "2026.05" / "对账台账.xlsx"
    if default_ledger.exists():
        df_new, result, contract_index_list, result_flags = process_reconciliation_ledger(
            default_ledger
        )
        print(f"待录入流水 {len(df_new)} 条")
        print(df_new.to_string(index=False))
        print(f"\n已录入索引: {contract_index_list}")
        print(f"待录入标记: {result_flags}")
        print(f"\n分组明细 ({len(result)} 组):")
        for group in result:
            print(f"  [{group['收款方式']}] {len(group['明细'])} 条")
    else:
        print(f"未找到示例台账: {default_ledger}")
