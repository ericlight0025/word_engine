from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from openpyxl import load_workbook


SUPPORTED_EXCEL_SUFFIXES = {".xlsx"}


@dataclass
class ExcelDataset:
    headers: list[str]
    rows: list[dict[str, str]]
    row_numbers: list[int] = field(default_factory=list)
    sheet_name: str | None = None
    source_headers: list[str] = field(default_factory=list)
    source_rows: list[dict[str, str]] = field(default_factory=list)


def validate_headers(headers: list[str], source: str) -> None:
    """欄名需非空且唯一，避免字典轉換時默默丟失資料。"""
    if not headers or not any(headers):
        raise ValueError(f"{source} 第一列缺少欄位名稱")
    seen: set[str] = set()
    for header in headers:
        if not header:
            raise ValueError(f"{source} 第一列包含空白欄位名稱")
        if header in seen:
            raise ValueError(f"{source} 欄位名稱重複：{header}")
        seen.add(header)


def _stringify(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


def read_excel(path: str | Path) -> ExcelDataset:
    file_path = Path(path)
    if file_path.suffix.lower() not in SUPPORTED_EXCEL_SUFFIXES:
        raise ValueError("僅支援 .xlsx；舊式 .xls 請先另存為 .xlsx")

    workbook = load_workbook(file_path, data_only=True, read_only=True)
    try:
        sheet = workbook.active
        if sheet is None:
            raise ValueError("Excel 沒有可用工作表")
        rows_iter = sheet.iter_rows(values_only=True)
        headers = [_stringify(cell) for cell in next(rows_iter, ())]
        # 忽略僅套用格式的尾端空欄，但拒絕沒有欄名的實際資料。
        while headers and not headers[-1]:
            headers.pop()
        validate_headers(headers, "Excel")
        data_rows: list[dict[str, str]] = []
        row_numbers: list[int] = []
        for number, row in enumerate(rows_iter, start=2):
            if any(_stringify(cell) for cell in row[len(headers):]):
                raise ValueError(f"Excel 第 {number} 列有資料缺少欄位名稱")
            values = [_stringify(cell) for cell in row[:len(headers)]]
            if not any(values):
                continue
            padded = values + [""] * (len(headers) - len(values))
            data_rows.append(dict(zip(headers, padded)))
            row_numbers.append(number)
        if not data_rows:
            raise ValueError("Excel 沒有可用資料列")
        return ExcelDataset(
            headers=headers, rows=data_rows, row_numbers=row_numbers, sheet_name=sheet.title,
            source_headers=headers.copy(), source_rows=[row.copy() for row in data_rows],
        )
    finally:
        workbook.close()
