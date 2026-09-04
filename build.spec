# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller 打包配置（onefile 模式，供 Inno Setup 二次封装）
用法: pyinstaller build.spec
产物: dist/SheepBreeding.exe （单文件，资源经 sys._MEIPASS 访问）
说明:
  - 业务代码未使用 pdfplumber，故不打包；重型 ML 库（torch/pandas）亦排除。
  - 资源：templates/ static/ seed/sheep_farm.db 通过 datas 注入。
"""
import os

APP_DIR = os.getcwd()  # 始终在项目根目录运行（build.bat 已 cd 至此）

added_files = [
    (os.path.join(APP_DIR, "templates"), "templates"),
    (os.path.join(APP_DIR, "static"), "static"),
    (os.path.join(APP_DIR, "seed", "sheep_farm.db"), "seed"),
]

a = Analysis(
    [os.path.join(APP_DIR, "launcher.py")],
    pathex=[APP_DIR],
    binaries=[],
    datas=added_files,
    hiddenimports=[
        "flask", "jinja2", "werkzeug", "click", "itsdangerous", "markupsafe",
        "openpyxl", "reportlab",
        "pywebview", "webview",
        "sqlite3",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "pdfplumber", "pdfminer", "pdfminer.six",
        "torch", "torchvision", "tensorboard",
        "pandas", "numpy", "scipy",
        "tkinter", "unittest", "pydoc", "doctest",
    ],
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    name="SheepBreeding",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    console=False,          # 桌面端隐藏控制台窗口
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    version=os.path.join(APP_DIR, "version_info.txt"),  # EXE 文件属性版本信息(RC 风格)
)
