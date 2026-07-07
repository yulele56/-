# -*- coding: utf-8 -*-
"""
日常审计：银行流水算税预处理 + 收入明细核验。
"""

import re
import traceback
from pathlib import Path

import pandas as pd

__all__ = ["audit_bank_ledger_combined"]


def audit_bank_ledger_combined(ledger_path):
    """
    对账台账审计：算税预处理 + 收入明细核验。

    返回 4 项：
        核验_未匹配流水, 核验_分组明细, 核验_成功匹配索引, 核验_标记

    核验_标记每条为：(行索引, 结果, 金额, 税额, 对方账户, 结算对象)
    """
    df_verify = pd.DataFrame()
    verify_groups = []
    verify_ok_idx = []
    verify_flags = []

    try:
        # -------------------------------
        # 1. 常量与核验状态
        # -------------------------------
        NO_TAX_CHANNELS = (
            ("微信", None), ("支付宝", None),
            ("李希", "中行"), ("谢纵生", "华兴"), ("蓝乾", "华兴"), ("蓝乾", "中信"),
            ("李茂铎", "建行"), ("李希", "民生澄海支行"),
        )
        NON_PUBLIC_CHANNELS = (
            "微信5071W", "支付宝777", "希兄中行3364", "生兄华兴0201",
            "蓝乾华兴0219", "蓝乾中信5701", "茂铎建行8870", "茂锋建行8870", "希兄民生6603",
        )
        SHAREHOLDER_NAMES = ("李希", "谢纵生", "蓝乾", "李茂铎")

        STATUS_OK = "匹配成功"
        STATUS_FAIL = "匹配失败"
        STATUS_NET_FAIL = "净额匹配失败"
        STATUS_TAX_FAIL = "税额匹配失败"
        FAIL_STATUSES = (STATUS_FAIL, STATUS_NET_FAIL, STATUS_TAX_FAIL)

        def empty_val(v):
            if v is None:
                return ""
            try:
                if pd.isna(v):
                    return ""
            except (TypeError, ValueError):
                pass
            return v

        def to_amt(v):
            if pd.isna(v):
                return None
            s = str(v).strip().replace(",", "").replace(" ", "").replace("￥", "").replace("¥", "").replace("元", "")
            if not s:
                return None
            try:
                return round(float(s), 2)
            except ValueError:
                return None

        def to_date(v):
            if pd.isna(v):
                return None
            s = str(v).strip()
            if re.fullmatch(r"\d{8}\.0+", s):
                s = s.split(".")[0]
            if re.fullmatch(r"\d{8}", s):
                return "%s-%s-%s" % (s[0:4], s[4:6], s[6:8])
            try:
                return pd.to_datetime(v).strftime("%Y-%m-%d")
            except Exception:
                return None

        def to_text(v):
            return "" if pd.isna(v) else str(v).strip()

        def classify_tax(channel, cp):
            ch = to_text(channel)
            cp = to_text(cp)
            for name, bank in NO_TAX_CHANNELS:
                if name in ch and (not bank or bank in ch):
                    return "不算税"
            if ch in NON_PUBLIC_CHANNELS:
                return "不算税"
            matches = [n for n in NON_PUBLIC_CHANNELS if n in ch or ch in n]
            if len(matches) == 1:
                return "不算税"
            if cp and any(n in cp for n in SHAREHOLDER_NAMES):
                return "不算税"
            if not cp:
                return "不算税"
            return "算税"

        def channel_match(flow_ch, settle_acct):
            a, b = to_text(flow_ch), to_text(settle_acct)
            return bool(a and b and (a in b or b in a))

        def detail_has_line(flow_date, flow_ch, amount):
            if flow_date is None or amount is None:
                return False
            exp = round(float(amount), 2)
            for _, d in detail.iterrows():
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
            (s for s in ("银行收入流水", "银行收入流水汇总", "收入流水汇总") if s in names),
            None,
        )
        sheet_detail = "记账明细" if "记账明细" in names else ("对账明细" if "对账明细" in names else None)
        if not sheet_flow or not sheet_detail:
            raise ValueError("工作表缺失: %s" % list(names))

        # -------------------------------
        # 3. 结算对象映射
        # -------------------------------
        settle_map = {}
        if "结算对象查询参考" in names:
            ref = pd.read_excel(path, sheet_name="结算对象查询参考", header=1)
            if "摘要" in ref.columns and "结算单位" in ref.columns:
                for _, r in ref.iterrows():
                    if pd.isna(r["摘要"]) or pd.isna(r["结算单位"]):
                        continue
                    s = str(r["摘要"]).strip()
                    if "/" in s:
                        key = s.rsplit("/", 1)[-1].strip()
                        if key:
                            settle_map[key] = str(r["结算单位"]).strip()

        # -------------------------------
        # 4. 读取并清洗银行流水
        # -------------------------------
        df = pd.read_excel(path, sheet_name=sheet_flow).copy()
        ch_col = "收款方式"
        if ch_col not in df.columns and "收款渠道" in df.columns:
            ch_col = "收款渠道"

        if "日期" in df.columns:
            df["日期"] = df["日期"].map(to_date)
        if "收入" not in df.columns:
            df["收入"] = None
        df["收入"] = df["收入"].map(to_amt)
        df = df[df["收入"].notna() & (df["收入"] > 0)].copy()

        for c in ("账户类别", "金额", "税额", "结算对象", "核验结果"):
            if c not in df.columns:
                df[c] = None
            df[c] = df[c].astype(object)
        if "金额" in df.columns:
            df["金额"] = df["金额"].map(to_amt)
        if "税额" in df.columns:
            df["税额"] = df["税额"].map(lambda x: to_amt(x) or 0.0)

        # -------------------------------
        # 5. 读取并清洗记账明细
        # -------------------------------
        detail_raw = pd.read_excel(path, sheet_name=sheet_detail, header=1)
        settle_col = None
        for c in detail_raw.columns:
            if "结算账户" in str(c):
                settle_col = c
                break
        if not settle_col:
            for c in detail_raw.columns:
                if "银行账户" in str(c):
                    settle_col = c
                    break
        if not settle_col or "收入金额(借)" not in detail_raw.columns:
            raise ValueError("记账明细缺少必要列")

        detail = detail_raw.copy()
        detail["_date"] = detail["记账日期"].map(to_date)
        detail["_amt"] = detail["收入金额(借)"].map(to_amt).fillna(0)
        detail = detail[detail["_amt"] > 0]
        detail["_settle"] = detail[settle_col].map(to_text)

        # -------------------------------
        # 6. 算税预处理
        # -------------------------------
        for idx, row in df.iterrows():
            gross = row["收入"]
            channel = to_text(row.get(ch_col, ""))
            cp = to_text(row.get("对方账户", ""))

            tax_label = classify_tax(channel, cp)
            if tax_label == "算税":
                tax_amt = round(gross * 0.1, 2)
                net_amt = round(gross * 0.9, 2)
            else:
                tax_amt = 0
                net_amt = gross

            settle_obj = settle_map.get(cp, "待认") if cp else "待认"
            df.at[idx, "账户类别"] = tax_label
            df.at[idx, "税额"] = tax_amt
            df.at[idx, "金额"] = net_amt
            df.at[idx, "结算对象"] = settle_obj

        # -------------------------------
        # 7. 逐条核验
        # -------------------------------
        verify_ok_idx = []
        verify_flags = []

        for idx, row in df.iterrows():
            gross = to_amt(row.get("收入"))
            if not gross or gross <= 0:
                continue

            flow_date = row.get("日期")
            flow_ch = to_text(row.get(ch_col, ""))
            tax_label = to_text(row.get("账户类别", ""))
            cp = to_text(row.get("对方账户", ""))
            settle_obj = to_text(row.get("结算对象", "")) or cp

            if tax_label != "算税":
                match_amt = to_amt(row.get("金额"))
                if match_amt is None:
                    match_amt = gross
                if detail_has_line(flow_date, flow_ch, match_amt):
                    status = STATUS_OK
                else:
                    status = STATUS_FAIL
            else:
                net = to_amt(row.get("金额"))
                tax = to_amt(row.get("税额"))
                if net is None:
                    net = round(gross * 0.9, 2)
                if tax is None:
                    tax = round(gross * 0.1, 2)

                net_ok = detail_has_line(flow_date, flow_ch, net)
                tax_ok = detail_has_line(flow_date, flow_ch, tax)
                if net_ok and tax_ok:
                    status = STATUS_OK
                elif net_ok:
                    status = STATUS_NET_FAIL
                elif tax_ok:
                    status = STATUS_TAX_FAIL
                else:
                    status = STATUS_FAIL

            df.at[idx, "核验结果"] = status
            if status == STATUS_OK:
                verify_ok_idx.append((idx, STATUS_OK))
            verify_flags.append((
                idx, status, empty_val(gross), empty_val(row.get("税额")),
                empty_val(cp), empty_val(settle_obj),
            ))

        # -------------------------------
        # 8. 未匹配流水 + 分组
        # -------------------------------
        df_verify = df[df["核验结果"].isin(FAIL_STATUSES)].copy()
        verify_groups = []

        if not df_verify.empty:
            df_verify[ch_col] = df_verify[ch_col].fillna("")
            for ch, grp in df_verify.groupby(ch_col, sort=False):
                verify_groups.append({
                    ch_col: empty_val(ch),
                    "明细": [{k: empty_val(v) for k, v in r.items()} for r in grp.to_dict("records")],
                })
            df_verify = df_verify.fillna("")

    except Exception as e:
        error_msg = "审计校验出错：%s\n%s" % (str(e), traceback.format_exc())
        print(error_msg)

    return df_verify, verify_groups, verify_ok_idx, verify_flags
