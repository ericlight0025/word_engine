from __future__ import annotations

import csv
from pathlib import Path

from core.excel_reader import ExcelDataset


def read_csv(path: str | Path) -> ExcelDataset:
    file_path = Path(path)
    if file_path.suffix.lower() != ".csv":
        raise ValueError("無法讀取，請確認格式為 .csv")

    with file_path.open("r", encoding="utf-8-sig", newline="") as csv_file:
        reader = csv.DictReader(csv_file)
        headers = reader.fieldnames or []
        if not headers:
            raise ValueError("CSV 第一列缺少欄位名稱")

        rows: list[dict[str, str]] = []
        for row in reader:
            normalized = {header: (row.get(header, "") or "").strip() for header in headers}
            if not any(normalized.values()):
                continue
            rows.append(normalized)

    if not rows:
        raise ValueError("CSV 沒有可用資料列")
    return ExcelDataset(headers=headers, rows=rows)
