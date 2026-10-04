from __future__ import annotations

import csv
from pathlib import Path

from core.excel_reader import ExcelDataset, validate_headers


def read_csv(path: str | Path) -> ExcelDataset:
    file_path = Path(path)
    if file_path.suffix.lower() != ".csv":
        raise ValueError("無法讀取，請確認格式為 .csv")

    with file_path.open("r", encoding="utf-8-sig", newline="") as csv_file:
        reader = csv.reader(csv_file)
        headers = [value.strip() for value in next(reader, [])]
        validate_headers(headers, "CSV")

        rows: list[dict[str, str]] = []
        for row in reader:
            if len(row) > len(headers):
                raise ValueError(f"CSV 第 {reader.line_num} 行欄位數超過表頭")
            values = [value.strip() for value in row] + [""] * (len(headers) - len(row))
            normalized = dict(zip(headers, values))
            if not any(normalized.values()):
                continue
            rows.append(normalized)

    if not rows:
        raise ValueError("CSV 沒有可用資料列")
    return ExcelDataset(headers=headers, rows=rows, source_headers=headers.copy(),
                        source_rows=[row.copy() for row in rows])
