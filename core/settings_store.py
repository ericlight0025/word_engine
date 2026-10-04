from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from core.file_store import atomic_write


@dataclass
class AppSettings:
    data_dir: str = ""
    template_dir: str = ""
    output_dir: str = ""
    theme: str = "Obsidian Violet"
    font_scale: int = 100


class SettingsStore:
    def __init__(self, path: Path) -> None:
        self.path = path

    def load(self) -> AppSettings:
        if not self.path.exists():
            return AppSettings()
        try:
            payload = json.loads(self.path.read_text(encoding="utf-8"))
        except Exception:
            return AppSettings()
        if not isinstance(payload, dict):
            return AppSettings()
        try:
            font_scale = int(payload.get("font_scale", 100))
        except (TypeError, ValueError, OverflowError):
            font_scale = 100
        return AppSettings(
            data_dir=str(payload.get("data_dir") or ""),
            template_dir=str(payload.get("template_dir") or ""),
            output_dir=str(payload.get("output_dir") or ""),
            theme=str(payload.get("theme") or "Obsidian Violet"),
            font_scale=max(90, min(130, font_scale)),
        )

    def save(self, settings: AppSettings) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        text = (
            json.dumps(
                {
                    "data_dir": settings.data_dir,
                    "template_dir": settings.template_dir,
                    "output_dir": settings.output_dir,
                    "theme": settings.theme,
                    "font_scale": settings.font_scale,
                },
                ensure_ascii=False,
                indent=2,
            ) + "\n"
        )
        atomic_write(self.path, lambda temporary: temporary.write_text(text, encoding="utf-8"))
