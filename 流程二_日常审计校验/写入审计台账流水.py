# -*- coding: utf-8 -*-
"""
批量将银行流水写入审计台账「银行收入流水汇总」sheet。

依赖同项目「流程一对账&录入/银行台账流水获取.py」，实在智能需一并上传。

RPA 调用:
    written, total = append_flows_to_audit_ledger(台账路径, 流水文件夹路径)

    台账路径: 流程二目录下「审计台账模板.xlsx」或其副本
    clear_sheet=True: 写入前清空流水 sheet 已有数据（保留表头）
"""

import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_ROOT / "流程一对账&录入"))

from 银行台账流水获取 import append_flows_to_ledger  # noqa: E402

__all__ = ["append_flows_to_audit_ledger"]

DEFAULT_TEMPLATE = Path(__file__).resolve().parent / "审计台账模板.xlsx"


def append_flows_to_audit_ledger(ledger_path, flow_folder, clear_sheet=False):
    """
    批量写入审计台账银行收入流水。

    返回: (本次写入条数, 台账流水 sheet 总行数)
    """
    return append_flows_to_ledger(ledger_path, flow_folder, clear_sheet=clear_sheet)


if __name__ == "__main__":
    base = _ROOT
    flow_dir = base / "流水"
    ledger = DEFAULT_TEMPLATE.parent / "_测试写入台账.xlsx"

    if not DEFAULT_TEMPLATE.is_file():
        print("未找到模板:", DEFAULT_TEMPLATE)
    elif not flow_dir.is_dir():
        print("未找到流水目录:", flow_dir)
    else:
        import shutil

        shutil.copy(DEFAULT_TEMPLATE, ledger)
        written, total = append_flows_to_audit_ledger(
            ledger, flow_dir, clear_sheet=True
        )
        print("台账:", ledger)
        print("本次写入 %d 条, 共 %d 条" % (written, total))
