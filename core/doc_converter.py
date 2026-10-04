from __future__ import annotations

import shutil
import subprocess
from dataclasses import dataclass
from pathlib import Path
from tempfile import TemporaryDirectory


@dataclass
class ConversionResult:
    converted_path: Path
    temp_dir: TemporaryDirectory[str] | None = None

    def cleanup(self) -> None:
        if self.temp_dir is not None:
            self.temp_dir.cleanup()


def libreoffice_exists() -> bool:
    return shutil.which("soffice") is not None


def prepare_template(path: str | Path) -> ConversionResult:
    template_path = Path(path)
    if not template_path.is_file():
        raise FileNotFoundError(template_path)
    suffix = template_path.suffix.lower()
    if suffix == ".docx":
        return ConversionResult(converted_path=template_path)
    if suffix != ".doc":
        raise ValueError("僅支援 .docx 或 .doc 版型")
    if not libreoffice_exists():
        raise RuntimeError("未偵測到 LibreOffice，無法轉換 .doc，請先安裝後再試。")

    temp_dir = TemporaryDirectory(prefix="word-merge-tool-")
    output_dir = Path(temp_dir.name)
    command = [
        "soffice",
        f"-env:UserInstallation={(output_dir / 'profile').as_uri()}",
        "--headless",
        "--convert-to",
        "docx",
        "--outdir",
        str(output_dir),
        str(template_path.resolve()),
    ]
    try:
        completed = subprocess.run(command, capture_output=True, text=True, check=False, timeout=60)
        converted_path = output_dir / f"{template_path.stem}.docx"
        if completed.returncode != 0 or not converted_path.is_file():
            stderr = completed.stderr.strip() or completed.stdout.strip() or "未知錯誤"
            raise RuntimeError(f".doc 轉換失敗：{stderr}")
    except subprocess.TimeoutExpired as exc:
        temp_dir.cleanup()
        raise RuntimeError(".doc 轉換超過 60 秒，已取消並清除暫存檔") from exc
    except BaseException:
        temp_dir.cleanup()
        raise
    return ConversionResult(converted_path=converted_path, temp_dir=temp_dir)
