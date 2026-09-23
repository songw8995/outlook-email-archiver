"""Portable GUI entry; unexpected startup errors remain visible."""
import os
import sys
import traceback
from pathlib import Path

if __name__ == "__main__":
    try:
        from outlook_archiver.gui import run_gui
        run_gui()
    except Exception as exc:
        log = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "OutlookEmailArchiver" / "启动失败日志.txt"
        try:
            log.parent.mkdir(parents=True, exist_ok=True)
            log.write_text(traceback.format_exc(), encoding="utf-8-sig")
        except OSError:
            pass
        import tkinter.messagebox as messagebox
        messagebox.showerror("邮件转 Markdown 启动失败", f"{type(exc).__name__}: {exc}\n\n请完整解压程序包，勿单独移动 runtime 或程序文件。\n日志：{log}")
        sys.exit(1)
