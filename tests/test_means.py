import unittest
from datetime import datetime
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook, load_workbook

from xlsx_wecom.config import SkillConfig, parse_skill_text
from xlsx_wecom.means import classify, format_means, output_path, process_workbook

ROOT = Path(__file__).resolve().parents[1]


def config(**overrides) -> SkillConfig:
    base = parse_skill_text((ROOT / ".cursor" / "skills" / "xlsx-wecom" / "SKILL.md").read_text(encoding="utf-8"))
    data = {
        "label": base.label,
        "header": base.header,
        "decimals": base.decimals,
        "blank_cells": base.blank_cells,
        "sheets": base.sheets,
        "output": base.output,
    }
    data.update(overrides)
    return SkillConfig(**data)


def write_scores(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "成绩"
    sheet.append(["姓名", "语文", "数学"])
    sheet.append(["甲", 90, 70])
    sheet.append(["乙", 80, 100])
    sheet.append(["丙", None, 90])
    workbook.save(path)
    workbook.close()


class ClassifyTests(unittest.TestCase):
    def test_numbers_blanks_and_non_numbers(self):
        self.assertEqual(classify(80)[0], "number")
        self.assertEqual(classify("  -1.5 ")[1], Decimal("-1.5"))
        self.assertEqual(classify("  ")[0], "blank")
        self.assertEqual(classify(None)[0], "blank")
        self.assertEqual(classify("12人")[0], "ignore")
        self.assertEqual(classify("1,234")[0], "ignore")
        self.assertEqual(classify("=SUM(A1)")[0], "ignore")
        self.assertEqual(classify(True)[0], "ignore")
        self.assertEqual(classify(datetime(2024, 1, 1))[0], "ignore")
        self.assertEqual(classify(float("nan"))[0], "ignore")


class MeansTests(unittest.TestCase):
    def test_appends_the_example_from_the_skill(self):
        with self._tmpdir() as directory:
            source = directory / "成绩.xlsx"
            write_scores(source)
            before = source.read_bytes()
            outcome = process_workbook(source, config())
            self.assertEqual(source.read_bytes(), before)
            self.assertEqual(outcome.output, directory / "成绩-均值.xlsx")
            saved = load_workbook(outcome.output)
            sheet = saved.active
            self.assertEqual([cell.value for cell in sheet[5]], ["均值", 85, 86.6667])
            self.assertEqual(sheet["B5"].number_format, "0.0000")
            self.assertIn("xlsx-wecom", sheet["A5"].comment.text)
            self.assertIn("语文：85.0000", format_means(outcome))
            self.assertIn("数学：86.6667", format_means(outcome))
            self.assertIn("1 个空单元格未计入", format_means(outcome))
            saved.close()

    def test_running_twice_replaces_the_mean_row(self):
        with self._tmpdir() as directory:
            source = directory / "成绩.xlsx"
            write_scores(source)
            first = process_workbook(source, config())
            second = process_workbook(first.output, config())
            saved = load_workbook(second.output)
            sheet = saved.active
            self.assertEqual(sheet.max_row, 5)
            self.assertEqual(sheet["B5"].value, 85)
            self.assertEqual(sheet["C5"].value, 86.6667)
            saved.close()

    def test_blank_cells_can_count_as_zero_from_the_skill(self):
        with self._tmpdir() as directory:
            source = directory / "成绩.xlsx"
            write_scores(source)
            outcome = process_workbook(source, config(blank_cells="as_zero", decimals=4))
            saved = load_workbook(outcome.output)
            self.assertEqual(saved.active["B5"].value, 56.6667)
            saved.close()

    def test_header_false_uses_column_letters_and_includes_the_first_row(self):
        with self._tmpdir() as directory:
            source = directory / "数字.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.append([10, 1])
            sheet.append([30, 3])
            workbook.save(source)
            workbook.close()
            outcome = process_workbook(source, config(header=False, decimals=0, label="均值"))
            saved = load_workbook(outcome.output)
            sheet = saved.active
            self.assertEqual(sheet["A3"].value, 20)
            self.assertEqual(sheet["B3"].value, 2)
            self.assertEqual(outcome.sheets[0].columns[0].name, "A")
            saved.close()

    def test_all_sheets_and_inplace(self):
        with self._tmpdir() as directory:
            source = directory / "两张.xlsx"
            workbook = Workbook()
            first = workbook.active
            first.title = "一"
            first.append(["项目", "数量"])
            first.append(["甲", 1])
            first.append(["乙", 3])
            second = workbook.create_sheet("二")
            second.append(["项目", "数量"])
            second.append(["甲", 10])
            workbook.save(source)
            workbook.close()
            outcome = process_workbook(source, config(sheets="all", output="inplace", decimals=0))
            self.assertEqual(outcome.output, source)
            saved = load_workbook(source)
            self.assertEqual(saved["一"]["B4"].value, 2)
            self.assertEqual(saved["二"]["B3"].value, 10)
            saved.close()

    def test_sample_file_in_the_repo(self):
        outcome = process_workbook(ROOT / "examples" / "成绩.xlsx", config())
        try:
            self.assertIn("语文：85.0000", format_means(outcome))
            self.assertIn("数学：86.6667", format_means(outcome))
        finally:
            if outcome.output and outcome.output.exists():
                outcome.output.unlink()

    def test_output_name_uses_the_label(self):
        path = output_path(Path("报表.xlsx"), config(label="平均"))
        self.assertEqual(path.name, "报表-平均.xlsx")

    def _tmpdir(self):
        return _TempDir()


class _TempDir:
    def __enter__(self):
        import tempfile

        self._cm = tempfile.TemporaryDirectory()
        return Path(self._cm.__enter__())

    def __exit__(self, exc_type, exc, tb):
        return self._cm.__exit__(exc_type, exc, tb)


if __name__ == "__main__":
    unittest.main()
