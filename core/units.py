"""core.units - 单位换算（与 PDF/Excel 中常用的 pt 单位互转）

PDF 内部坐标系：72 pt = 1 inch = 2.54 cm = 25.4 mm
"""
CM_PER_PT = 2.54 / 72
MM_PER_PT = 25.4 / 72


def cm_to_pt(cm: float) -> float:
    """厘米 → 磅（PDF 点）"""
    return cm / 2.54 * 72


def mm_to_pt(mm: float) -> float:
    """毫米 → 磅"""
    return mm / 25.4 * 72


def pt_to_cm(pt: float) -> float:
    """磅 → 厘米"""
    return pt / 72 * 2.54


def pt_to_mm(pt: float) -> float:
    """磅 → 毫米"""
    return pt / 72 * 25.4