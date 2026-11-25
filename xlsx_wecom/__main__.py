"""命令行：python -m xlsx_wecom 文件.xlsx"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from xlsx_wecom.config import SkillConfigError, find_skill, load_skill
from xlsx_wecom.means import WorkbookError, format_means, process_workbook
from xlsx_wecom.wecom import WeComError, WeComNotConfigured, load_local_env, resolve_credentials, send_outcome


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="按 SKILL.md 给 xlsx 追加各列均值，并用企业微信发给自己。"
    )
    parser.add_argument("workbook", type=Path, help="要处理的 .xlsx 文件")
    parser.add_argument("--skill", help="SKILL.md 的路径。默认向上查找项目里的技能文件。")
    parser.add_argument(
        "--no-send",
        action="store_true",
        help="只生成本地文件，不发送企业微信。",
    )
    parser.add_argument(
        "--env-file",
        type=Path,
        default=Path(".env"),
        help="凭证文件，默认是当前目录的 .env。已有环境变量不会被覆盖。",
    )
    args = parser.parse_args(argv)

    try:
        load_local_env(args.env_file)
        config = load_skill(find_skill(args.skill))
        outcome = process_workbook(args.workbook, config)
    except (SkillConfigError, WorkbookError) as exc:
        print(str(exc), file=sys.stderr)
        return 1

    print(format_means(outcome), flush=True)
    if not outcome.wrote or outcome.output is None:
        print("没有写入文件，也没有发送。", file=sys.stderr)
        return 1

    print(f"已保存：{outcome.output}", flush=True)
    if args.no_send:
        print("未发送企业微信（--no-send）。")
        return 0

    try:
        credentials = resolve_credentials()
        print(send_outcome(outcome, credentials))
    except WeComNotConfigured as exc:
        print(str(exc), file=sys.stderr)
        return 2
    except WeComError as exc:
        print(f"文件已保存在本地，但没有发到企业微信。{exc}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
