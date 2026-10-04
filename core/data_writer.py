from __future__ import annotations

import csv
from pathlib import Path

from openpyxl import Workbook, load_workbook

from core.csv_reader import read_csv
from core.excel_reader import ExcelDataset, read_excel, validate_headers
from core.file_store import atomic_write


def write_dataset(
    path: str | Path, headers: list[str], rows: list[dict[str, str]],
    *, source_dataset: ExcelDataset | None = None,
) -> None:
    """只更新資料表的編輯內容，保留來源工作表結構並以原子方式存回。"""
    file_path = Path(path)
    suffix = file_path.suffix.lower()
    if suffix not in {".csv", ".xlsx"}:
        raise ValueError("目前僅支援將編輯內容存回 .csv 或 .xlsx")
    validate_headers(headers, "資料")
    for row in rows:
        if set(row) - set(headers):
            raise ValueError("資料包含未列入表頭的欄位，已取消存檔")

    current = None
    if file_path.exists():
        current = read_csv(file_path) if suffix == ".csv" else read_excel(file_path)
        if source_dataset and source_dataset.source_headers and (
            current.headers != source_dataset.source_headers or current.rows != source_dataset.source_rows
            or (source_dataset.sheet_name is not None and current.sheet_name != source_dataset.sheet_name)
        ):
            raise ValueError("來源檔已被其他程式修改，請重新匯入後再存檔")

    if suffix == ".csv":
        def save_csv(temporary: Path) -> None:
            with temporary.open("w", encoding="utf-8-sig", newline="") as handle:
                writer = csv.DictWriter(handle, fieldnames=headers)
                writer.writeheader()
                writer.writerows(rows)
        atomic_write(file_path, save_csv)
        return

    workbook = load_workbook(file_path, data_only=False) if current else Workbook()
    try:
        sheet = workbook[current.sheet_name] if current else workbook.active
        if current and (len(headers) != len(current.headers) or len(rows) != len(current.rows)):
            raise ValueError("Excel 存回僅支援編輯既有欄位與資料列，已取消結構變更")
        for column, header in enumerate(headers, start=1):
            cell = sheet.cell(1, column)
            cell.value = header
            cell.data_type = "s"
        numbers = current.row_numbers if current else list(range(2, len(rows) + 2))
        for index, (number, row) in enumerate(zip(numbers, rows)):
            for column, header in enumerate(headers, start=1):
                value = row.get(header, "")
                if current and value == current.rows[index].get(current.headers[column - 1], ""):
                    # 未編輯的公式、日期、數字與格式保留，不能用畫面上的字串覆蓋。
                    continue
                cell = sheet.cell(number, column)
                cell.value = value
                if isinstance(value, str):
                    cell.data_type = "s"
        atomic_write(file_path, workbook.save)
    finally:
        workbook.close()
