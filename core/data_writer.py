from __future__ import annotations

import csv
from pathlib import Path

from openpyxl import Workbook


def write_dataset(path: str | Path, headers: list[str], rows: list[dict[str, str]]) -> None:
    file_path = Path(path)
    suffix = file_path.suffix.lower()
    if suffix == ".csv":
        with file_path.open("w", encoding="utf-8-sig", newline="") as csv_file:
            writer = csv.DictWriter(csv_file, fieldnames=headers)
            writer.writeheader()
            writer.writerows(rows)
        return

    if suffix == ".xlsx":
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(headers)
        for row in rows:
            sheet.append([row.get(header, "") for header in headers])
        workbook.save(file_path)
        return

    raise ValueError("目前僅支援將編輯內容存回 .csv 或 .xlsx")
