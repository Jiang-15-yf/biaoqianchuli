"""tests/test_excel_pricecard.py - 用 samples 验证 Excel → 价格牌核心算法

运行：
    python -m tests.test_excel_pricecard

测试用样本会在 samples/ 下动态生成（_test_input.xlsx）。
"""
from __future__ import annotations

import sys
from pathlib import Path

import openpyxl

HERE = Path(__file__).parent.parent
sys.path.insert(0, str(HERE))

from core.excel_pricecard import parse_table, build_pricecard, export_xlsx


def _ensure_sample(path: Path) -> None:
    """动态生成测试用宽表 xlsx（不存在时创建）"""
    if path.exists():
        return
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["款号", "颜色", "背面", "尺码段", "XS", "S", "M", "L", "XL", "合计", "涉及PO"])
    ws.append(["G1945AX", "黑色", "有价格", "170-95A", 10, 20, 30, 40, 30, 130, "PO001,PO002"])
    ws.append(["G1945AX", "白色", "有价格", "170-95A", 15, 25, 35, 45, 35, 155, "PO001,PO003"])
    ws.append(["G1945AX", "黑色", "空白", "", 5, 10, 15, 20, 15, 65, ""])
    wb.save(path)
    print(f"  已生成测试样本：{path.name}")


def test_parse_xlsx():
    """解析动态生成的 xlsx 测试样本"""
    sample = HERE / "samples" / "_test_input.xlsx"
    _ensure_sample(sample)
    parsed = parse_table(str(sample))
    print(f"  解析：{len(parsed['rows'])} 行，尺码 {parsed['size_headers']}")
    assert len(parsed["rows"]) > 0
    assert len(parsed["size_headers"]) > 0
    print(f"  ✅ 解析通过")


def test_build_and_export():
    """构建竖排数据并导出 xlsx"""
    sample = HERE / "samples" / "_test_input.xlsx"
    _ensure_sample(sample)
    parsed = parse_table(str(sample))
    built = build_pricecard(parsed, rate_pct=4)
    body_n = len(built["body"])
    blank_n = len(built["blank_out"])
    print(f"  主体 {body_n} 行，空白吊牌 {blank_n} 行，合计 {built['total_row'][2]}/{built['total_row'][3]}")
    assert body_n > 0

    out = HERE / "_test_pricecard.xlsx"
    try:
        out.unlink()
    except OSError:
        pass  # WorkBuddy sandbox 可能拦截 unlink，不影响核心测试
    export_xlsx(parsed, built, str(out), tag_w_cm=4.5, tag_h_cm=9)
    assert out.exists()
    print(f"  ✅ 导出通过 → {out.name}")


def test_parse_xls():
    """解析真实样本 samples/output.xls（验证 xlrd==1.2.0 路径）"""
    xls_path = HERE / "samples" / "output.xls"
    if not xls_path.exists():
        print(f"  ⚠️  跳过：未发现 {xls_path.name}")
        return
    parsed = parse_table(str(xls_path))
    print(f"  解析：{len(parsed['rows'])} 行，尺码 {parsed['size_headers']}")
    assert len(parsed["rows"]) > 0
    assert len(parsed["size_headers"]) > 0
    print(f"  ✅ .xls 解析通过")


def main() -> None:
    print("\n== test_parse_xlsx ==")
    test_parse_xlsx()
    print("\n== test_parse_xls ==")
    test_parse_xls()
    print("\n== test_build_and_export ==")
    test_build_and_export()
    print("\n所有测试通过 ✅")


if __name__ == "__main__":
    main()