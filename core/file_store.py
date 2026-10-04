from __future__ import annotations

import os
import tempfile
from pathlib import Path
from typing import Callable


def atomic_write(path: Path, writer: Callable[[Path], None], *, rename: bool = False) -> Path:
    """完整寫入同目錄暫存檔後發布；批次輸出遇到重名時保留原檔並更名。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    descriptor, name = tempfile.mkstemp(prefix=".word-engine-", suffix=path.suffix, dir=path.parent)
    os.close(descriptor)
    temporary = Path(name)
    try:
        writer(temporary)
        with temporary.open("r+b") as handle:
            os.fsync(handle.fileno())
        if not rename:
            os.replace(temporary, path)
            return path
        counter = 0
        while True:
            candidate = path if not counter else path.with_name(f"{path.stem}_{counter}{path.suffix}")
            try:
                # 建立硬連結會原子拒絕既有名稱，避免並行輸出互相覆蓋。
                os.link(temporary, candidate)
                return candidate
            except FileExistsError:
                counter += 1
    finally:
        temporary.unlink(missing_ok=True)
