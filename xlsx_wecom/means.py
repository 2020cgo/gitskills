"""计算工作表各列均值，并追加到数据最后一行。"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from xlsx_wecom.config import SkillConfig

_NUMBER_TEXT = re.compile(r"^[+-]?(?:\d+\.\d*|\d+|\.\d+)$")
_MARKER = "xlsx-wecom"
_BOLD = Font(bold=True)


class WorkbookError(Exception):
    """这个 xlsx 不能按当前需求处理。"""


@dataclass
class ColumnResult:
    name: str
    letter: str
    mean: Decimal | None
    numeric_count: int
    blank_count: int
    ignored_count: int


@dataclass
class SheetResult:
    name: str
    columns: list[ColumnResult] = field(default_factory=list)
    empty: bool = False
    wrote: bool = False


@dataclass
class Outcome:
    source: Path
    output: Path | None
    sheets: list[SheetResult]
    config: SkillConfig

    @property
    def wrote(self) -> bool:
        return any(sheet.wrote for sheet in self.sheets)


def output_path(source: Path, config: SkillConfig) -> Path:
    if config.output == "inplace":
        return source
    suffix = f"-{config.label}"
    stem = source.stem
    if not stem.endswith(suffix):
        stem = f"{stem}{suffix}"
    return source.with_name(stem + source.suffix)


def process_workbook(source: Path, config: SkillConfig) -> Outcome:
    if source.suffix.lower() != ".xlsx":
        raise WorkbookError(f"只处理 .xlsx 文件，当前是「{source.name}」。请另存为 .xlsx。")
    if not source.is_file():
        raise WorkbookError(f"找不到文件：{source}")

    try:
        workbook = load_workbook(source)
    except Exception as exc:
        raise WorkbookError(f"无法打开 {source.name}：{exc}") from exc

    try:
        targets = _select_sheets(workbook, config)
        sheets = [append_means(worksheet, config) for worksheet in targets]
        outcome = Outcome(source=source, output=None, sheets=sheets, config=config)
        if outcome.wrote:
            destination = output_path(source, config)
            _atomic_save(workbook, destination)
            outcome.output = destination
        return outcome
    finally:
        workbook.close()


def append_means(worksheet: Worksheet, config: SkillConfig) -> SheetResult:
    result = SheetResult(name=worksheet.title)
    last_row, last_col = _used_bounds(worksheet)
    if last_row == 0 or last_col == 0:
        result.empty = True
        return result

    data_end = last_row
    if _is_mean_row(worksheet, last_row, config.label):
        data_end -= 1

    start = 2 if config.header else 1
    if config.header and data_end < 1:
        result.empty = True
        return result
    if data_end < start:
        _write_mean_row(worksheet, start, last_col, config, [])
        result.columns = _column_results(worksheet, config, last_col, [])
        result.wrote = True
        return result

    columns = _collect(worksheet, config, start, data_end, last_col)
    write_row = data_end + 1 if data_end >= start else start
    if _is_mean_row(worksheet, last_row, config.label):
        write_row = last_row
    _write_mean_row(worksheet, write_row, last_col, config, columns)
    result.columns = columns
    result.wrote = True
    return result


def format_means(outcome: Outcome) -> str:
    config = outcome.config
    if config.blank_cells == "as_zero":
        blank_note = "空单元格按 0 计入均值。"
    else:
        blank_note = "空单元格未计入均值。"
    lines = [f"文件：{outcome.source.name}", blank_note]
    for sheet in outcome.sheets:
        lines.append(f"工作表：{sheet.name}")
        if sheet.empty:
            lines.append("没有数据，未追加均值。")
            continue
        for column in sheet.columns:
            lines.append(_format_column(column, config))
        lines.append("均值已写在该表最后一行。")
    return "\n".join(lines)


def _format_column(column: ColumnResult, config: SkillConfig) -> str:
    if column.mean is None:
        return f"{column.name}：不是数值列"
    detail = f"{column.numeric_count} 个数字"
    if column.ignored_count:
        detail += f"，{column.ignored_count} 个非数字未计入"
    if column.blank_count and config.blank_cells == "skip":
        detail += f"，{column.blank_count} 个空单元格未计入"
    if column.blank_count and config.blank_cells == "as_zero":
        detail += f"，{column.blank_count} 个空单元格按 0 计入"
    rendered = f"{column.mean:.{config.decimals}f}"
    return f"{column.name}：{rendered}（{detail}）"


def _select_sheets(workbook, config: SkillConfig) -> list[Worksheet]:
    if config.sheets == "all":
        sheets = [sheet for sheet in workbook.worksheets if isinstance(sheet, Worksheet)]
    else:
        active = workbook.active
        if not isinstance(active, Worksheet):
            raise WorkbookError("当前表不是普通工作表，无法计算均值。")
        sheets = [active]
    if not sheets:
        raise WorkbookError("工作簿里没有可处理的工作表。")
    return sheets


def _collect(
    worksheet: Worksheet,
    config: SkillConfig,
    start: int,
    end: int,
    last_col: int,
) -> list[ColumnResult]:
    seen: set[str] = set()
    columns: list[ColumnResult] = []
    for col in range(1, last_col + 1):
        letter = get_column_letter(col)
        name = _header_name(worksheet, config, col, letter, seen)
        numbers: list[Decimal] = []
        blanks = 0
        ignored = 0
        for row in range(start, end + 1):
            kind, number = classify(worksheet.cell(row, col).value)
            if kind == "number" and number is not None:
                numbers.append(number)
            elif kind == "blank":
                blanks += 1
            else:
                ignored += 1
        if config.blank_cells == "as_zero":
            numbers.extend([Decimal("0")] * blanks)
        mean = _mean(numbers, config.decimals) if numbers else None
        columns.append(
            ColumnResult(
                name=name,
                letter=letter,
                mean=mean,
                numeric_count=len(numbers),
                blank_count=blanks,
                ignored_count=ignored,
            )
        )
    return columns


def _column_results(
    worksheet: Worksheet,
    config: SkillConfig,
    last_col: int,
    columns: list[ColumnResult],
) -> list[ColumnResult]:
    if columns:
        return columns
    seen: set[str] = set()
    built: list[ColumnResult] = []
    for col in range(1, last_col + 1):
        letter = get_column_letter(col)
        built.append(
            ColumnResult(
                name=_header_name(worksheet, config, col, letter, seen),
                letter=letter,
                mean=None,
                numeric_count=0,
                blank_count=0,
                ignored_count=0,
            )
        )
    return built


def _header_name(
    worksheet: Worksheet,
    config: SkillConfig,
    col: int,
    letter: str,
    seen: set[str],
) -> str:
    if config.header:
        value = worksheet.cell(1, col).value
        text = "" if value is None else str(value).strip()
        name = text or letter
    else:
        name = letter
    if name in seen:
        name = f"{name}（{letter}）"
    seen.add(name)
    return name


def _write_mean_row(
    worksheet: Worksheet,
    row: int,
    last_col: int,
    config: SkillConfig,
    columns: list[ColumnResult],
) -> None:
    by_letter = {column.letter: column for column in columns}
    for col in range(1, last_col + 1):
        cell = worksheet.cell(row, col)
        letter = get_column_letter(col)
        column = by_letter.get(letter)
        if col == 1 and (column is None or column.mean is None):
            cell.value = config.label
            cell.number_format = "General"
        elif column is not None and column.mean is not None:
            cell.value = float(column.mean)
            cell.number_format = config.number_format
        else:
            cell.value = None
            cell.number_format = "General"
        cell.font = _BOLD
        if col == 1:
            cell.comment = Comment(_MARKER, "xlsx-wecom")
        elif cell.comment is not None and _MARKER in (cell.comment.text or ""):
            cell.comment = None


def _is_mean_row(worksheet: Worksheet, row: int, label: str) -> bool:
    cell = worksheet.cell(row, 1)
    if cell.comment is not None and _MARKER in (cell.comment.text or ""):
        return True
    return cell.value == label


def _used_bounds(worksheet: Worksheet) -> tuple[int, int]:
    max_row = 0
    max_col = 0
    for row in worksheet.iter_rows():
        for cell in row:
            comment = cell.comment.text if cell.comment is not None else ""
            has_marker = _MARKER in (comment or "")
            if has_marker or _has_content(cell.value):
                max_row = max(max_row, cell.row)
                max_col = max(max_col, cell.column)
    return max_row, max_col


def _has_content(value: object) -> bool:
    if value is None:
        return False
    if isinstance(value, str) and value.strip() == "":
        return False
    return True


def classify(value: object) -> tuple[str, Decimal | None]:
    """把单元格分成 number、blank、ignore。"""
    if value is None:
        return "blank", None
    if isinstance(value, bool):
        return "ignore", None
    if isinstance(value, datetime | date):
        return "ignore", None
    if isinstance(value, int | float):
        number = Decimal(str(value))
        if not number.is_finite():
            return "ignore", None
        return "number", number
    if isinstance(value, str):
        text = value.strip()
        if text == "":
            return "blank", None
        if text.startswith("="):
            return "ignore", None
        if _NUMBER_TEXT.fullmatch(text):
            return "number", Decimal(text)
        return "ignore", None
    return "ignore", None


def _mean(numbers: list[Decimal], decimals: int) -> Decimal:
    total = sum(numbers, Decimal("0"))
    quantum = Decimal("1") if decimals <= 0 else Decimal("1").scaleb(-decimals)
    return (total / Decimal(len(numbers))).quantize(quantum, rounding=ROUND_HALF_UP)


def _atomic_save(workbook, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(path.name + ".tmp")
    try:
        workbook.save(temporary)
        temporary.replace(path)
    except Exception:
        if temporary.exists():
            temporary.unlink()
        raise
