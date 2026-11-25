"""读取 SKILL.md 里的 skill-config 块。"""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

SKILL_NAME = "xlsx-wecom"
CONFIG_FENCE = "skill-config"

_BOOL = {"true": True, "false": False}
_BLANK = {"skip", "as_zero"}
_SHEETS = {"active", "all"}
_OUTPUT = {"copy", "inplace"}


class SkillConfigError(Exception):
    """SKILL.md 里的配置不能用。"""


@dataclass(frozen=True)
class SkillConfig:
    label: str
    header: bool
    decimals: int
    blank_cells: str
    sheets: str
    output: str

    @property
    def number_format(self) -> str:
        if self.decimals <= 0:
            return "0"
        return "0." + ("0" * self.decimals)


def find_skill(explicit: str | None, cwd: Path | None = None) -> Path:
    if explicit:
        path = Path(explicit).expanduser()
        if not path.is_file():
            raise SkillConfigError(f"找不到技能文件：{path}")
        return path

    env = os.environ.get("XLSX_WECOM_SKILL")
    if env:
        path = Path(env).expanduser()
        if not path.is_file():
            raise SkillConfigError(f"XLSX_WECOM_SKILL 指向的文件不存在：{path}")
        return path

    start = (cwd or Path.cwd()).resolve()
    for directory in [start, *start.parents]:
        candidate = directory / ".cursor" / "skills" / SKILL_NAME / "SKILL.md"
        if candidate.is_file():
            return candidate
    raise SkillConfigError(
        "找不到 .cursor/skills/xlsx-wecom/SKILL.md。请在项目根目录运行，或用 --skill 指向它。"
    )


def load_skill(path: Path) -> SkillConfig:
    try:
        text = path.read_text(encoding="utf-8")
    except OSError as exc:
        raise SkillConfigError(f"无法读取技能文件：{path}") from exc
    return parse_skill_text(text, source=str(path))


def parse_skill_text(text: str, source: str = "SKILL.md") -> SkillConfig:
    block = _extract_block(text, source)
    values: dict[str, str] = {}
    for line_no, raw in block:
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        if ":" not in line:
            raise SkillConfigError(f"{source} 的 skill-config 第 {line_no} 行缺少冒号：{raw}")
        key, value = line.split(":", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if not key:
            raise SkillConfigError(f"{source} 的 skill-config 第 {line_no} 行没有配置名。")
        if key in values:
            raise SkillConfigError(f"{source} 里的「{key}」写了两次。")
        values[key] = value

    unknown = sorted(set(values) - _FIELDS)
    if unknown:
        names = "、".join(unknown)
        raise SkillConfigError(
            f"{source} 里出现了程序还不认识的配置：{names}。"
            "如果这是新需求，让 agent 按技能文件实现它，不要沿用旧行为。"
        )
    missing = [name for name in _FIELD_ORDER if name not in values]
    if missing:
        raise SkillConfigError(f"{source} 的 skill-config 缺少：{'、'.join(missing)}。")

    return SkillConfig(
        label=_label(values["label"], source),
        header=_boolean(values["header"], "header", source),
        decimals=_decimals(values["decimals"], source),
        blank_cells=_choice(values["blank_cells"], "blank_cells", _BLANK, source),
        sheets=_choice(values["sheets"], "sheets", _SHEETS, source),
        output=_choice(values["output"], "output", _OUTPUT, source),
    )


_FIELD_ORDER = ("label", "header", "decimals", "blank_cells", "sheets", "output")
_FIELDS = set(_FIELD_ORDER)


def _extract_block(text: str, source: str) -> list[tuple[int, str]]:
    lines = text.splitlines()
    opening = f"```{CONFIG_FENCE}"
    starts = [index for index, line in enumerate(lines) if line.strip() == opening]
    if not starts:
        raise SkillConfigError(f"{source} 里没有 ```skill-config 配置块。")
    if len(starts) > 1:
        raise SkillConfigError(f"{source} 里有多段 skill-config，只保留一段。")
    start = starts[0]
    collected: list[tuple[int, str]] = []
    for index in range(start + 1, len(lines)):
        if lines[index].strip() == "```":
            return collected
        collected.append((index + 1, lines[index]))
    raise SkillConfigError(f"{source} 的 skill-config 没有结束标记。")


def _label(value: str, source: str) -> str:
    if not value:
        raise SkillConfigError(f"{source} 的 label 不能是空的。")
    if len(value) > 20:
        raise SkillConfigError(f"{source} 的 label 不能超过 20 个字符。")
    if any(char in value for char in "\r\n\t\\/:*?\"<>|"):
        raise SkillConfigError(f"{source} 的 label 不能包含换行或文件名里禁用的符号。")
    return value


def _boolean(value: str, key: str, source: str) -> bool:
    if value not in _BOOL:
        raise SkillConfigError(f"{source} 的 {key} 只能写 true 或 false，现在是「{value}」。")
    return _BOOL[value]


def _decimals(value: str, source: str) -> int:
    if not value.isdigit():
        raise SkillConfigError(f"{source} 的 decimals 只能是 0 到 6 的整数，现在是「{value}」。")
    number = int(value)
    if number > 6:
        raise SkillConfigError(f"{source} 的 decimals 只能是 0 到 6 的整数，现在是「{value}」。")
    return number


def _choice(value: str, key: str, allowed: set[str], source: str) -> str:
    if value not in allowed:
        options = "、".join(sorted(allowed))
        raise SkillConfigError(f"{source} 的 {key} 只能是 {options}，现在是「{value}」。")
    return value
