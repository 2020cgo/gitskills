import os
import subprocess
import sys
import unittest
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]


class CliTests(unittest.TestCase):
    def test_no_send_prints_means_and_writes_a_copy(self):
        proc = subprocess.run(
            [sys.executable, "-m", "xlsx_wecom", "examples/成绩.xlsx", "--no-send"],
            cwd=ROOT,
            capture_output=True,
            text=True,
            check=False,
        )
        self.assertEqual(proc.returncode, 0, proc.stderr)
        self.assertIn("语文：85.0000", proc.stdout)
        self.assertIn("数学：86.6667", proc.stdout)
        self.assertIn("未发送企业微信", proc.stdout)
        output = ROOT / "examples" / "成绩-均值.xlsx"
        self.assertTrue(output.exists())
        saved = load_workbook(output)
        try:
            self.assertEqual(saved.active["A5"].value, "均值")
            self.assertEqual(saved.active["B5"].value, 85)
            self.assertEqual(saved.active["C5"].value, 86.6667)
        finally:
            saved.close()
            output.unlink()

    def test_missing_wecom_credentials_keep_the_file_and_exit_2(self):
        env = os.environ.copy()
        for key in list(env):
            if key.startswith("WECOM_"):
                del env[key]
        env["PYTHONDONTWRITEBYTECODE"] = "1"
        proc = subprocess.run(
            [sys.executable, "-m", "xlsx_wecom", "examples/成绩.xlsx", "--env-file", "/tmp/xlsx-wecom-missing.env"],
            cwd=ROOT,
            capture_output=True,
            text=True,
            check=False,
            env=env,
        )
        self.assertEqual(proc.returncode, 2, proc.stderr)
        self.assertIn("没有发送", proc.stderr)
        output = ROOT / "examples" / "成绩-均值.xlsx"
        self.assertTrue(output.exists())
        output.unlink()


if __name__ == "__main__":
    unittest.main()
