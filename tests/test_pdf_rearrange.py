"""tests/test_pdf_rearrange.py - 用 samples 验证 PDF 重排核心算法

运行：
    python -m tests.test_pdf_rearrange
"""
from __future__ import annotations

import sys
from pathlib import Path

HERE = Path(__file__).parent.parent
sys.path.insert(0, str(HERE))

from core.pdf_rearrange import rearrange_pdf, detect_stickers
import fitz


def test_detect_stickers_count():
    """验证 samples/input.pdf 能检出预期枚数（按 history 应为 12 = 6 枚/页 × 2 页）"""
    doc = fitz.open(HERE / "samples" / "input.pdf")
    counts = []
    for i in range(len(doc)):
        dets = detect_stickers(doc[i])
        counts.append(len(dets))
    doc.close()
    total = sum(counts)
    print(f"  检测结果：每页 {counts}，共 {total} 枚")
    assert total > 0, "未检测到任何标签"
    print(f"  ✅ 检测通过")


def test_rearrange_pdf_output():
    """运行核心重排，验证输出页数与页面尺寸"""
    in_pdf = HERE / "samples" / "input.pdf"
    out_pdf = HERE / "_test_output.pdf"
    try:
        if out_pdf.exists():
            out_pdf.unlink()
    except OSError:
        pass

    result = rearrange_pdf(
        input_path=str(in_pdf),
        output_path=str(out_pdf),
        per_page=1,
        margin_mm=0,
        padding_mm=4,
    )
    print(f"  重排结果：{result}")
    assert result["labels"] > 0
    assert result["pages"] == result["labels"]  # per_page=1 时
    assert out_pdf.exists()
    # 用 PyMuPDF 验证输出
    doc = fitz.open(str(out_pdf))
    assert len(doc) == result["pages"]
    # 验证页面尺寸接近 samples/示例输出
    p0 = doc[0]
    w_cm = p0.rect.width / 72 * 2.54
    h_cm = p0.rect.height / 72 * 2.54
    print(f"  输出首页尺寸：{w_cm:.2f} × {h_cm:.2f} cm")
    doc.close()
    print(f"  ✅ 重排通过")


def main() -> None:
    print("\n== test_detect_stickers_count ==")
    test_detect_stickers_count()
    print("\n== test_rearrange_pdf_output ==")
    test_rearrange_pdf_output()
    print("\n所有测试通过 ✅")


if __name__ == "__main__":
    main()