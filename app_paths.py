# -*- coding: utf-8 -*-
"""專案可寫入資料目錄：開發時為程式所在資料夾；PyInstaller 打包後為 .exe 所在資料夾。"""

from __future__ import annotations

import sys
from pathlib import Path


def project_data_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent
