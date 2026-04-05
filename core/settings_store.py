from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path


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
        return AppSettings(
            data_dir=str(payload.get("data_dir", "")),
            template_dir=str(payload.get("template_dir", "")),
            output_dir=str(payload.get("output_dir", "")),
            theme=str(payload.get("theme", "Obsidian Violet")),
            font_scale=int(payload.get("font_scale", 100)),
        )

    def save(self, settings: AppSettings) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.path.write_text(
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
            ),
            encoding="utf-8",
        )
