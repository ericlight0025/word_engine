from __future__ import annotations

import re
import zipfile
from xml.etree import ElementTree
from dataclasses import dataclass
from pathlib import Path

from docxtpl import DocxTemplate
from core.file_store import atomic_write


TAG_PATTERN = re.compile(r"{{\s*([^{}]+?)\s*}}")
INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]+')
WHITESPACE_PATTERN = re.compile(r"\s+")
INVALID_XML_CHAR_PATTERN = re.compile(r"[^\x09\x0a\x0d\x20-\ud7ff\ue000-\ufffd\U00010000-\U0010ffff]")
WINDOWS_RESERVED_NAMES = {"CON", "PRN", "AUX", "NUL", "CONIN$", "CONOUT$"} | {
    f"{prefix}{number}" for prefix in ("COM", "LPT") for number in (*range(1, 10), "¹", "²", "³")
}


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

    # 使用與產出相同的模板解析器，正確處理跨 run、條件與 filter 的變數。
    return sorted(DocxTemplate(str(path)).get_undeclared_template_variables())


def extract_template_preview(template_path: str | Path) -> str:
    """擷取本文、表格、頁首及頁尾的文字預覽，不把 Word XML 當成文字顯示。"""
    namespace = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    paragraphs: list[str] = []
    with zipfile.ZipFile(template_path) as archive:
        names = [name for name in archive.namelist() if name == "word/document.xml"
                 or re.fullmatch(r"word/(?:header|footer)\d+\.xml", name)]
        for name in names:
            root = ElementTree.fromstring(archive.read(name))
            for paragraph in root.iter(f"{namespace}p"):
                text = "".join(node.text or "" for node in paragraph.iter(f"{namespace}t"))
                if text:
                    paragraphs.append(text)
    return "\n".join(paragraphs)


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
    if cleaned.split(".")[0].rstrip(" .").upper() in WINDOWS_RESERVED_NAMES:
        cleaned = "_" + cleaned
    cleaned = cleaned.encode("utf-8")[:180].decode("utf-8", errors="ignore").rstrip("._")
    return cleaned or "output"


def _resolve_filename(
    template_path: Path,
    row: dict[str, str],
    index: int,
    naming_field: str,
    warnings: list[MergeWarning],
) -> str:
    if naming_field:
        field_value = str(row.get(naming_field) or "").strip()
        if field_value:
            return sanitize_filename(field_value) + ".docx"
        warnings.append(MergeWarning(row_index=index, message=f"{naming_field} 為空，改用預設流水號"))
    return f"{sanitize_filename(template_path.stem)}_{index:03d}.docx"


def merge_documents(
    template_path: str | Path,
    rows: list[dict[str, str]],
    output_dir: str | Path,
    naming_field: str,
) -> MergeSummary:
    source = Path(template_path)
    tags = extract_tags(source)
    destination = Path(output_dir)
    destination.mkdir(parents=True, exist_ok=True)

    output_files: list[Path] = []
    warnings: list[MergeWarning] = []
    failures: list[MergeFailure] = []

    for index, row in enumerate(rows, start=1):
        try:
            for header, value in row.items():
                if isinstance(value, str) and INVALID_XML_CHAR_PATTERN.search(value):
                    raise ValueError(f"欄位 {header} 含 Word 不支援的控制字元")
            filename = _resolve_filename(source, row, index, naming_field, warnings)
            target = destination / filename
            missing = [tag for tag in tags if tag not in row]
            if missing:
                warnings.append(MergeWarning(index, f"缺少欄位，已留空：{', '.join(missing)}"))
            document = DocxTemplate(str(source))
            document.render(dict(row), autoescape=True)
            published = atomic_write(target, document.save, rename=True)
            if published != target:
                warnings.append(MergeWarning(index, f"檔名已存在，改存為 {published.name}"))
            output_files.append(published)
        except Exception as exc:
            failures.append(MergeFailure(row_index=index, reason=str(exc)))

    success_count = len(output_files)
    warning_count = len({warning.row_index for warning in warnings})
    failure_count = len(failures)
    return MergeSummary(
        success_count=success_count,
        warning_count=warning_count,
        failure_count=failure_count,
        output_files=output_files,
        warnings=warnings,
        failures=failures,
    )
