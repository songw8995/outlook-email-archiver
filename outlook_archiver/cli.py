from __future__ import annotations

import argparse
import json
import os
import sys
import traceback
from datetime import datetime
from pathlib import Path
from typing import Any

from . import __version__
from .archive import export_archive, load_config, preview, render_archive, verify_archive
from .outlook_reader import OutlookSession, doctor
from .utils import write_json_atomic

if sys.stdout is not None and hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if sys.stderr is not None and hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")


def emit(event: dict[str, Any]) -> None:
    print("EVENT " + json.dumps(event, ensure_ascii=False, default=str), flush=True)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Outlook / PST 邮件转 Markdown")
    parser.add_argument("--version", action="version", version=__version__)
    sub = parser.add_subparsers(dest="command", required=True)
    doctor_parser = sub.add_parser("doctor", help="检测 Python、Tkinter 和经典 Outlook COM")
    doctor_parser.add_argument("--output")
    folders = sub.add_parser("folders", help="列出可选择的邮箱和邮件文件夹")
    folders.add_argument("--source", choices=["outlook", "pst"], default="outlook")
    folders.add_argument("--pst")
    folders.add_argument("--pst-work-root", help="PST 已验证工作副本位置")
    folders.add_argument("--pst-reuse-root", action="append", default=[], help="兼容查找旧版已验证工作副本的位置")
    folders.add_argument("--output-json")
    preview_parser = sub.add_parser("preview", help="预览选定范围")
    preview_parser.add_argument("--config", required=True)
    preview_parser.add_argument("--limit", type=int)
    export_parser = sub.add_parser("export", help="导出选定范围")
    export_parser.add_argument("--config", required=True)
    export_parser.add_argument("--limit", type=int, help="仅用于受控测试；正式模式不设置")
    export_parser.add_argument("--stop-file")
    verify_parser = sub.add_parser("verify", help="核验现有归档")
    verify_parser.add_argument("--root", required=True)
    verify_parser.add_argument("--quick", action="store_true", help="只核对存在和大小，不重算哈希")
    verify_parser.add_argument("--open-msg", action="store_true", help="用经典 Outlook 尝试打开 MSG")
    render_parser = sub.add_parser("render", help="从已保存正文源重建 Markdown")
    render_parser.add_argument("--root", required=True)
    render_parser.add_argument("--hide-addresses", action="store_true")
    sub.add_parser("gui", help="启动中文窗口")
    return parser


def _doctor_markdown(result: dict[str, Any]) -> str:
    ok = result.get("pywin32") and result.get("tkinter") and result.get("outlook_com") and result.get("profile_access")
    lines = [
        "# Outlook 邮件归档器当前环境检测", "",
        f"- 检测时间：{datetime.now().astimezone().isoformat(timespec='seconds')}",
        "- 本次实际执行了当前检测；没有沿用历史环境报告。", "",
        "| 项目 | 结果 |", "| --- | --- |",
        f"| Python | {result.get('python')} |",
        f"| Python 路径 | `{result.get('python_executable')}` |",
        f"| Python 位数 | {result.get('python_bitness')}-bit |",
        f"| Tkinter | {'可用' if result.get('tkinter') else '不可用'} |",
        f"| pywin32 | {'可用' if result.get('pywin32') else '不可用'} |",
        f"| 经典 Outlook COM | {'可用' if result.get('outlook_com') else '不可用'} |",
        f"| Outlook 应用/版本 | {result.get('outlook_name', '未知')} {result.get('outlook_version', '')} |",
        f"| 默认配置/Store | {'可访问，' + str(result.get('stores')) + ' 个 Store' if result.get('profile_access') else '不可访问'} |",
        "", f"## 结论", "", "当前环境满足程序运行条件。" if ok else "当前环境不满足程序运行条件，请按错误提示处理。", "",
    ]
    if result.get("errors"):
        lines.extend(["## 错误", ""] + [f"- {value}" for value in result["errors"]] + [""])
    if result.get("warnings"):
        lines.extend(["## 警告", ""] + [f"- {value}" for value in result["warnings"]] + [""])
    return "\n".join(lines)


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        if args.command == "gui":
            from .gui import run_gui

            run_gui()
            return 0
        if args.command == "doctor":
            result = doctor()
            if args.output:
                Path(args.output).write_text(_doctor_markdown(result), encoding="utf-8-sig")
            emit({"event": "doctor", **result})
            return 0 if not result["errors"] else 2
        if args.command == "folders":
            if args.source == "pst" and not args.pst:
                raise ValueError("PST 来源必须提供 --pst。")
            with OutlookSession(
                args.pst if args.source == "pst" else None,
                args.pst_work_root if args.source == "pst" else None,
                args.pst_reuse_root if args.source == "pst" else None,
            ) as session:
                stores, folders, warnings = session.catalog()
                result = {
                    "source": args.source,
                    "pst_work_copy": str(session.pst_copy.copy) if session.pst_copy else "",
                    "pst_copy_reused": session.pst_copy.reused if session.pst_copy else False,
                    "stores": stores,
                    "folders": [folder.to_dict() for folder in folders],
                    "warnings": warnings,
                }
            if args.output_json:
                warnings.extend(session.cleanup_warnings)
                write_json_atomic(Path(args.output_json), result)
            emit({"event": "folders", "stores": len(stores), "folders": len(folders), "warnings": warnings})
            return 0
        if args.command == "preview":
            preview(load_config(args.config), limit=args.limit, emit=emit)
            return 0
        if args.command == "export":
            export_archive(load_config(args.config), limit=args.limit, stop_file=Path(args.stop_file) if args.stop_file else None, emit=emit)
            return 0
        if args.command == "verify":
            emit({"event": "verify_complete", **verify_archive(args.root, full_hash=not args.quick, open_msg=args.open_msg)})
            return 0
        if args.command == "render":
            emit({"event": "render_complete", **render_archive(args.root, include_addresses=not args.hide_addresses)})
            return 0
    except Exception as exc:
        emit({"event": "fatal", "error_type": type(exc).__name__, "error": str(exc)})
        if os.environ.get("OUTLOOK_ARCHIVER_DEBUG") == "1":
            traceback.print_exc()
        return 1
    return 0
