from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass
from pathlib import Path

from docxtpl import DocxTemplate


TAG_PATTERN = re.compile(r"{{\s*([^{}]+?)\s*}}")
INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*]+')
WHITESPACE_PATTERN = re.compile(r"\s+")


@dataclass
class TagStatus:
    tag: str
    status: str
    message: str


@dataclass
class MergeWarning:
    row_index: int
    message: str


@dataclass
class MergeFailure:
    row_index: int
    reason: str


@dataclass
class MergeSummary:
    success_count: int
    warning_count: int
    failure_count: int
    output_files: list[Path]
    warnings: list[MergeWarning]
    failures: list[MergeFailure]


def extract_tags(template_path: str | Path) -> list[str]:
    path = Path(template_path)
    if not path.exists():
        raise FileNotFoundError(path)

    with zipfile.ZipFile(path, "r") as archive:
        names = [
            name
            for name in archive.namelist()
            if name.startswith("word/")
            and name.endswith(".xml")
            and (
                name == "word/document.xml"
                or "header" in name
                or "footer" in name
            )
        ]
        text = "".join(
            archive.read(name).decode("utf-8", errors="ignore") for name in names
        )

    tags = sorted({match.strip() for match in TAG_PATTERN.findall(text)})
    return tags


def build_tag_statuses(tags: list[str], headers: list[str], sample_row: dict[str, str] | None) -> list[TagStatus]:
    statuses: list[TagStatus] = []
    for tag in tags:
        if tag in headers:
            preview = ""
            if sample_row is not None:
                preview = sample_row.get(tag, "")
            message = preview if preview else "有對應欄位，首筆資料為空值"
            statuses.append(TagStatus(tag=tag, status="matched", message=message))
        else:
            statuses.append(TagStatus(tag=tag, status="missing", message="找不到對應欄位"))

    tag_set = set(tags)
    for header in headers:
        if header not in tag_set:
            statuses.append(TagStatus(tag=header, status="extra", message="Excel 有欄位，但版型沒有對應 Tag"))
    return statuses


def sanitize_filename(value: str) -> str:
    cleaned = INVALID_FILENAME_CHARS.sub("_", value.strip())
    cleaned = WHITESPACE_PATTERN.sub("_", cleaned)
    cleaned = cleaned.strip("._")
    return cleaned or "output"


def _resolve_filename(
    template_path: Path,
    row: dict[str, str],
    index: int,
    naming_field: str,
    warnings: list[MergeWarning],
) -> str:
    if naming_field:
        field_value = row.get(naming_field, "").strip()
        if field_value:
            return sanitize_filename(field_value) + ".docx"
        warnings.append(MergeWarning(row_index=index, message=f"{naming_field} 為空，改用預設流水號"))
    return f"{template_path.stem}_{index:03d}.docx"


def merge_documents(
    template_path: str | Path,
    rows: list[dict[str, str]],
    output_dir: str | Path,
    naming_field: str,
) -> MergeSummary:
    source = Path(template_path)
    destination = Path(output_dir)
    destination.mkdir(parents=True, exist_ok=True)

    output_files: list[Path] = []
    warnings: list[MergeWarning] = []
    failures: list[MergeFailure] = []

    for index, row in enumerate(rows, start=1):
        filename = _resolve_filename(source, row, index, naming_field, warnings)
        target = destination / filename
        try:
            document = DocxTemplate(str(source))
            document.render(dict(row))
            document.save(str(target))
            output_files.append(target)
        except Exception as exc:
            failures.append(MergeFailure(row_index=index, reason=str(exc)))

    success_count = len(output_files)
    warning_count = len(warnings)
    failure_count = len(failures)
    return MergeSummary(
        success_count=success_count,
        warning_count=warning_count,
        failure_count=failure_count,
        output_files=output_files,
        warnings=warnings,
        failures=failures,
    )
