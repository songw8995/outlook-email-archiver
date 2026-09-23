"""Exercise real Tk widgets with simulated worker events; no Outlook access."""
import json
import unittest
from unittest.mock import patch

from outlook_archiver.gui import App, completion_title


class CompletionTests(unittest.TestCase):
    def test_outcomes_are_not_confused_with_success(self):
        for event, expected in [
            ({"selected": 5, "success": 5}, "转换已完成"),
            ({"selected": 5, "skipped": 5}, "转换已完成"),
            ({"selected": 5, "warning": 1}, "有警告"),
            ({"selected": 5, "source_unavailable": 2}, "失败"),
            ({"selected": 5, "partial": 1}, "失败"),
            ({"selected": 5, "unprocessed": 1}, "已停止"),
            ({"selected": 0, "scan_stopped": True}, "已停止"),
            ({"selected": 0}, "没有符合条件"),
        ]:
            with self.subTest(event=event):
                self.assertIn(expected, completion_title(event))

    def test_three_consecutive_results_open_visible_dialogs_after_exit(self):
        with patch.object(App, "_restore_settings"):
            app = App()
        try:
            app.update()
            for index in range(3):
                result = {"event": "run_complete", "selected": 5, "success": 5,
                          "report": f"report-{index}.md"}
                app._handle_line("EVENT " + json.dumps(result))
                self.assertFalse(any(w.winfo_class() == "Toplevel" for w in app.winfo_children()))
                app._finished("export", 0)
                app.update()
                dialogs = [w for w in app.winfo_children() if w.winfo_class() == "Toplevel"]
                self.assertEqual(len(dialogs), 1)
                dialog = dialogs[0]
                self.assertTrue(dialog.winfo_viewable())
                self.assertEqual(dialog.title(), "转换已完成")
                self.assertEqual(app.grab_current(), dialog)
                self.assertIn("转换已完成", app.completion_status.get())
                self.assertEqual(app.last_report, f"report-{index}.md")
                dialog.destroy()
            with patch("outlook_archiver.gui.messagebox.showwarning") as warning:
                app._finished("export", 0)
                warning.assert_called_once()
                self.assertIn("无法确认", app.completion_status.get())
        finally:
            app.destroy()


if __name__ == "__main__":
    unittest.main()
