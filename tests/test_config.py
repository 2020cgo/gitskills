import unittest
from pathlib import Path

from xlsx_wecom.config import SkillConfigError, find_skill, parse_skill_text

ROOT = Path(__file__).resolve().parents[1]
SKILL = ROOT / ".cursor" / "skills" / "xlsx-wecom" / "SKILL.md"


def block(body: str) -> str:
    return f"说明\n\n```skill-config\n{body}\n```\n"


class ParseSkillTests(unittest.TestCase):
    def test_shipped_skill_matches_the_basic_requirement(self):
        config = parse_skill_text(SKILL.read_text(encoding="utf-8"), source=str(SKILL))
        self.assertEqual(config.label, "均值")
        self.assertTrue(config.header)
        self.assertEqual(config.decimals, 4)
        self.assertEqual(config.blank_cells, "skip")
        self.assertEqual(config.sheets, "active")
        self.assertEqual(config.output, "copy")
        self.assertEqual(config.number_format, "0.0000")

    def test_decimals_and_label_come_from_the_config_block(self):
        config = parse_skill_text(
            block("label: 平均\nheader: false\ndecimals: 2\nblank_cells: as_zero\nsheets: all\noutput: inplace\n")
        )
        self.assertEqual(config.label, "平均")
        self.assertFalse(config.header)
        self.assertEqual(config.decimals, 2)
        self.assertEqual(config.blank_cells, "as_zero")
        self.assertEqual(config.sheets, "all")
        self.assertEqual(config.output, "inplace")

    def test_unknown_key_stops_instead_of_keeping_the_old_behavior(self):
        text = block(
            "label: 均值\nheader: true\ndecimals: 4\nblank_cells: skip\n"
            "sheets: active\noutput: copy\nstatistic: median\n"
        )
        with self.assertRaises(SkillConfigError) as caught:
            parse_skill_text(text)
        self.assertIn("statistic", str(caught.exception))

    def test_invalid_choice_names_the_field(self):
        text = block(
            "label: 均值\nheader: yes\ndecimals: 4\nblank_cells: skip\nsheets: active\noutput: copy\n"
        )
        with self.assertRaises(SkillConfigError) as caught:
            parse_skill_text(text)
        self.assertIn("header", str(caught.exception))

    def test_missing_block(self):
        with self.assertRaises(SkillConfigError):
            parse_skill_text("没有配置")

    def test_find_skill_walks_up_from_a_subdirectory(self):
        found = find_skill(None, cwd=ROOT / "xlsx_wecom")
        self.assertEqual(found, SKILL)


if __name__ == "__main__":
    unittest.main()
