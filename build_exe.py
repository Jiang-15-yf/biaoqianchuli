"""build_exe.py - 用 PyInstaller 打包为单文件 EXE

运行：
    pip install -r requirements.txt
    python build_exe.py

输出：
    dist/价格牌处理工具.exe  （双击直接运行，无需 Python 环境）

参数说明：
    --onefile     : 打包成单个 exe（启动稍慢，但分发简单）
    --windowed    : 不弹黑色控制台窗口
    --add-data    : 把 samples 目录打包进 exe（相对路径解析需要）
"""
from __future__ import annotations

import os
import shutil
import subprocess
import sys
from pathlib import Path

import PyInstaller.__main__


HERE = Path(__file__).parent.resolve()
APP_NAME = "价格牌处理工具"


def build() -> None:
    # 清理旧产物（环境的安全删除 shim 可能拦截删除，忽略失败）
    for d in ["build", "dist"]:
        p = HERE / d
        try:
            if p.exists():
                shutil.rmtree(p, ignore_errors=True)
        except OSError:
            pass
    spec = HERE / f"{APP_NAME}.spec"
    try:
        if spec.exists():
            spec.unlink()
    except OSError:
        pass

    # 打包参数
    args = [
        str(HERE / "main.py"),
        f"--name={APP_NAME}",
        "--onefile",
        "--windowed",
        "--noconfirm",
        "--clean",
        "--noupx",  # 不压缩（部分杀软误报）
        f"--icon={HERE / 'assets' / 'logo.ico'}",
        f"--add-data={HERE / 'assets'}{os.pathsep}assets",
        # ---- 拖拽支持：必须打包 vendor，否则拖入功能静默失效 ----
        # main.py 靠 sys.path.insert(0, "vendor") 运行时引入 tkinterdnd2。
        # 注意要打整个 vendor 目录（目标是 _MEIPASS/vendor），因为 main.py
        # 找的就是 <...>/vendor/tkinterdnd2；且 tkdnd/ 下是平台二进制 DLL，
        # 只有 add-data 才会带上。
        f"--add-data={HERE / 'vendor'}{os.pathsep}vendor",
        # 【不要】加 --paths=vendor / --hidden-import=tkinterdnd2！
        # 那会把 tkinterdnd2 编译进 PYZ 内存归档，冻结后它的 __file__ 变成
        # 虚拟路径(_MEIxxx\tkinterdnd2\TkinterDnD.pyc，目录并不存在)，
        # _require() 拼不出 tkdnd 的 DLL 目录而抛异常 → TkinterDnD.Tk() 失败
        # → main() 回退成普通 Tk() → 既多一个空白窗口、又失去拖放能力。
        # 不打进 PYZ，运行时才会从 _MEIPASS/vendor 的真实文件加载，__file__ 正确。
        # 同理也不要 --hidden-import=tkinterdnd2*（那也会把它收进 PYZ）。
        # 显式包含子模块，确保 PyInstaller 能识别
        "--collect-submodules=core",
        "--hidden-import=fitz",
        "--hidden-import=PIL",
        "--hidden-import=PIL._tkinter_finder",
        "--hidden-import=openpyxl",
        "--hidden-import=xlrd",
        # 排除系统 Python 中损坏/无关的第三方库，避免 PyInstaller 触发其 hook 而崩溃
        "--exclude-module=matplotlib",
        "--exclude-module=numpy",
        "--exclude-module=scipy",
        "--exclude-module=pandas",
        "--exclude-module=IPython",
        "--exclude-module=pytest",
        "--exclude-module=notebook",
        "--exclude-module=jupyter",
    ]

    print("==> 打包参数：", args)
    PyInstaller.__main__.run(args)

    exe_path = HERE / "dist" / f"{APP_NAME}.exe"
    if exe_path.exists():
        size_mb = exe_path.stat().st_size / 1024 / 1024
        print(f"\n✅ 打包完成：{exe_path}  ({size_mb:.1f} MB)")
    else:
        print("\n❌ 打包失败：未找到 dist/*.exe")
        sys.exit(1)


if __name__ == "__main__":
    build()