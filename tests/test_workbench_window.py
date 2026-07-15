import runpy
import tkinter as tk
import unittest
from pathlib import Path


PROJECT_DIR = Path(__file__).resolve().parents[1]
GUI_SCRIPT = next(
    path
    for path in PROJECT_DIR.glob("*.py")
    if path.name != "web_query.py" and path.stat().st_size > 50_000
)
gui_module = runpy.run_path(str(GUI_SCRIPT))


class WorkbenchWindowTests(unittest.TestCase):
    def test_initial_geometry_keeps_title_bar_inside_a_small_logical_screen(self):
        width, height, x, y = gui_module["fit_window_geometry"](1366, 720)

        self.assertEqual(width, 760)
        self.assertEqual(height, 656)
        self.assertEqual(x, 303)
        self.assertEqual(y, 32)
        self.assertGreaterEqual(y, 32)

    def test_workbench_has_normal_window_controls_and_a_close_button(self):
        root = tk.Tk()
        root.withdraw()
        try:
            app = gui_module["PivotTableApp"](root)
            root.update_idletasks()

            self.assertNotEqual(root.overrideredirect(), 1)
            self.assertEqual(root.resizable(), (1, 1))
            self.assertTrue(hasattr(app, "btn_close"))
            self.assertEqual(app.btn_close.cget("text"), "关闭工作台")
            self.assertTrue(hasattr(app, "btn_stop_web"))
            self.assertEqual(app.btn_stop_web.cget("state"), tk.DISABLED)
            self.assertEqual(app.btn_stop_web.cget("text"), "⏹ 停止查询")
            self.assertTrue(hasattr(app, "btn_retry_missing"))
            self.assertEqual(app.btn_retry_missing.cget("text"), "↻ 补查缺失 PO")
            self.assertEqual(app.btn_retry_missing.cget("state"), tk.NORMAL)
            self.assertEqual(app.btn_retry_missing.master, app.btn_web.master)
            self.assertTrue(hasattr(app, "content_canvas"))
        finally:
            root.destroy()

    def test_stop_control_remains_visible_and_legible_at_minimum_width(self):
        root = tk.Tk()
        root.withdraw()
        try:
            app = gui_module["PivotTableApp"](root)
            root.geometry("640x560+20+20")
            root.deiconify()
            root.update_idletasks()
            root.update()

            left = app.btn_stop_web.winfo_rootx() - root.winfo_rootx()
            right = left + app.btn_stop_web.winfo_width()
            self.assertTrue(app.btn_stop_web.winfo_ismapped())
            self.assertGreaterEqual(left, 0)
            self.assertLessEqual(right, root.winfo_width())
            self.assertEqual(app.btn_stop_web.cget("disabledforeground"), "#f8fafc")
        finally:
            root.destroy()

    def test_missing_po_recovery_control_stays_next_to_query_at_minimum_width(self):
        root = tk.Tk()
        root.withdraw()
        try:
            app = gui_module["PivotTableApp"](root)
            root.geometry("640x560+20+20")
            root.deiconify()
            root.update_idletasks()
            root.update()

            self.assertTrue(app.btn_web.winfo_ismapped())
            self.assertTrue(app.btn_retry_missing.winfo_ismapped())
            self.assertGreater(app.btn_web.winfo_width(), 100)
            self.assertGreater(app.btn_retry_missing.winfo_width(), 100)
            self.assertEqual(app.btn_web.master, app.btn_retry_missing.master)
        finally:
            root.destroy()

    def test_completion_dialog_never_captures_workbench_input(self):
        """任务完成提示异常时，主工作台仍必须保持可操作。"""
        root = tk.Tk()
        root.withdraw()
        try:
            app = gui_module["PivotTableApp"](root)
            root.deiconify()
            root.update_idletasks()

            app._show_copyable_dialog("测试完成", "任务已完成")
            root.update_idletasks()

            self.assertIsNone(root.grab_current())
            self.assertEqual(app.btn_web.cget("state"), tk.NORMAL)
        finally:
            root.destroy()

    def test_air_conditioning_price_fill_control_is_visible_only_for_air_conditioning(self):
        root = tk.Tk()
        root.withdraw()
        try:
            app = gui_module["PivotTableApp"](root)
            root.deiconify()
            root.update_idletasks()

            self.assertTrue(hasattr(app, "btn_air_price"))
            self.assertFalse(app.btn_air_price.winfo_ismapped())

            app._switch_category("空调")
            root.update_idletasks()
            self.assertTrue(app.btn_air_price.winfo_ismapped())
            self.assertEqual(app.btn_air_price.cget("text"), "⑥ 填充空调老/新价格")

            app._switch_category("装潢")
            root.update_idletasks()
            self.assertFalse(app.btn_air_price.winfo_ismapped())
        finally:
            root.destroy()


if __name__ == "__main__":
    unittest.main()
