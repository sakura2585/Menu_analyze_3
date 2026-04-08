# -*- mode: python ; coding: utf-8 -*-
# PyInstaller：於專案目錄執行  python -m PyInstaller order_note.spec
# Windows：會依 app.py 的 _APP_VERSION 寫入 exe「內容→詳細資料」的檔案／產品版本（與程式內顯示一致）。

import os
import re
import sys

block_cipher = None

_spec_dir = os.path.dirname(os.path.abspath(SPECPATH))


def _read_app_version() -> tuple[tuple[int, int, int, int], str]:
    """回傳 (filevers 四元組), 顯示用字串（不含 v 前綴）。"""
    app_py = os.path.join(_spec_dir, "app.py")
    raw = "v0.0.0"
    try:
        with open(app_py, encoding="utf-8") as f:
            m = re.search(
                r'^\s*_APP_VERSION\s*=\s*["\']([^"\']+)["\']',
                f.read(),
                re.M,
            )
            if m:
                raw = (m.group(1) or "").strip()
    except OSError:
        pass
    s = raw.lower().lstrip("v")
    nums = [int(x) for x in re.findall(r"\d+", s)]
    while len(nums) < 4:
        nums.append(0)
    ft = tuple(nums[:4])
    vs_display = ".".join(str(x) for x in ft[:4])
    return ft, vs_display


def _write_win_version_info(path: str, filevers: tuple, vs_display: str) -> None:
    if sys.platform != "win32":
        return
    os.makedirs(os.path.dirname(path), exist_ok=True)
    # PyInstaller 會載入此檔並讀取 VSVersionInfo（語法須與官方範例一致）
    content = f"""# UTF-8
# 由 order_note.spec 依 app.py _APP_VERSION 自動產生
VSVersionInfo(
  ffi=FixedFileInfo(
    filevers={filevers},
    prodvers={filevers},
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
  ),
  kids=[
    StringFileInfo(
      [
      StringTable(
        u'040904B0',
        [StringStruct(u'CompanyName', u''),
        StringStruct(u'FileDescription', u'訂餐備註分析'),
        StringStruct(u'FileVersion', u'{vs_display}'),
        StringStruct(u'InternalName', u'MenuAnalyze'),
        StringStruct(u'LegalCopyright', u''),
        StringStruct(u'OriginalFilename', u'訂餐備註分析.exe'),
        StringStruct(u'ProductName', u'訂餐備註分析'),
        StringStruct(u'ProductVersion', u'{vs_display}')])
      ]),
    VarFileInfo([VarStruct(u'Translation', [1033, 1200])])
  ]
)
"""
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write(content)


_filevers, _vs_display = _read_app_version()
_version_info_path = os.path.join(_spec_dir, "build", "_file_version_info.txt")
_write_win_version_info(_version_info_path, _filevers, _vs_display)

a = Analysis(
    ["app.py"],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        "docx",
        "docx.enum.style",
        "docx.enum.text",
        "docx.oxml",
        "docx.oxml.ns",
        "lxml",
        "lxml._elementpath",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

_exe_options = dict(
    name="訂餐備註分析",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
if sys.platform == "win32" and os.path.isfile(_version_info_path):
    _exe_options["version"] = _version_info_path

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    **_exe_options,
)
