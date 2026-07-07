# -*- coding: utf-8 -*-
"""
批量将流水（银行 + 支付宝 + 微信）写入对账台账「收入流水汇总」sheet。

RPA 调用:
    written, total = append_flows_to_recon_ledger(台账路径, 流水文件夹路径)

    台账路径: 本目录「对账台账模板.xlsx」或其副本
    clear_sheet=True: 写入前清空流水 sheet 已有数据（保留表头）
"""

from pathlib import Path

from 台账流水获取 import append_flows_to_ledger

__all__ = ["append_flows_to_recon_ledger"]

DEFAULT_TEMPLATE = Path(__file__).resolve().parent / "对账台账模板.xlsx"


def append_flows_to_recon_ledger(ledger_path, flow_folder, clear_sheet=False):
    """
    批量写入对账台账收入流水汇总。

    返回: (本次写入条数, 台账流水 sheet 总行数)
    """
    return append_flows_to_ledger(ledger_path, flow_folder, clear_sheet=clear_sheet)


if __name__ == "__main__":
    base = Path(__file__).resolve().parent.parent
    flow_dir = base / "流水"
    ledger = DEFAULT_TEMPLATE.parent / "_测试写入对账台账.xlsx"

    if not DEFAULT_TEMPLATE.is_file():
        print("未找到模板:", DEFAULT_TEMPLATE)
    elif not flow_dir.is_dir():
        print("未找到流水目录:", flow_dir)
    else:
        import shutil

        shutil.copy(DEFAULT_TEMPLATE, ledger)
        written, total = append_flows_to_recon_ledger(
            ledger, flow_dir, clear_sheet=True
        )
        print("台账:", ledger)
        print("本次写入 %d 条, 共 %d 条" % (written, total))
