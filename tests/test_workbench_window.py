import runpy
import threading
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
    @staticmethod
    def _destroy_root(root):
        """ttkbootstrap 1.x 的 Style 是进程级单例，测试间须显式清理。"""
        # 先清空主题切换产生的空闲事件，再销毁 Tk 解释器；否则下一项测试
        # 重新创建窗口时，Tcl 可能尝试向已销毁的窗口派发 ThemeChanged。
        root.update_idletasks()
        root.update()
        root.destroy()
        gui_module["ttk"].Style.instance = None

    def test_initial_geometry_keeps_title_bar_inside_a_small_logical_screen(self):
        width, height, x, y = gui_module["fit_window_geometry"](1366, 720)

        self.assertEqual(width, 760)
        self.assertEqual(height, 656)
        self.assertEqual(x, 303)
        self.assertEqual(y, 32)
        self.assertGreaterEqual(y, 32)

    def test_application_root_uses_the_configured_bootstrap_theme(self):
        root = gui_module["create_workbench_root"]()
        try:
            self.assertEqual(
                root.tk.call("ttk::style", "theme", "use"),
                gui_module["WORKBENCH_THEME"],
            )
        finally:
            self._destroy_root(root)

    def test_workbench_has_normal_window_controls_and_a_close_button(self):
        root = gui_module["create_workbench_root"]()
        root.withdraw()
        try:
            app = gui_module["PivotTableApp"](root)
            root.update_idletasks()

            self.assertNotEqual(root.overrideredirect(), 1)
            self.assertEqual(root.resizable(), (1, 1))
            self.assertTrue(hasattr(app, "btn_close"))
            self.assertEqual(app.btn_close.cget("text"), "关闭工作台")
            self.assertTrue(hasattr(app, "btn_stop_web"))
            self.assertEqual(str(app.btn_stop_web.cget("state")), tk.DISABLED)
            self.assertEqual(app.btn_stop_web.cget("text"), "⏹ 停止查询")
            self.assertTrue(hasattr(app, "btn_retry_missing"))
            self.assertEqual(app.btn_retry_missing.cget("text"), "↻ 补查缺失 PO")
            self.assertEqual(str(app.btn_retry_missing.cget("state")), tk.NORMAL)
            self.assertEqual(app.btn_retry_missing.master, app.btn_web.master)
            self.assertTrue(hasattr(app, "content_canvas"))
        finally:
            self._destroy_root(root)

    def test_stop_control_remains_visible_and_legible_at_minimum_width(self):
        root = gui_module["create_workbench_root"]()
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
            self.assertEqual(app.btn_stop_web.winfo_class(), "TButton")
            self.assertEqual(str(app.btn_stop_web.cget("state")), tk.DISABLED)
        finally:
            self._destroy_root(root)

    def test_missing_po_recovery_control_stays_next_to_query_at_minimum_width(self):
        root = gui_module["create_workbench_root"]()
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
            self._destroy_root(root)

    def test_completion_dialog_never_captures_workbench_input(self):
        """任务完成提示异常时，主工作台仍必须保持可操作。"""
        root = gui_module["create_workbench_root"]()
        root.withdraw()
        try:
            app = gui_module["PivotTableApp"](root)
            root.deiconify()
            root.update_idletasks()

            app._show_copyable_dialog("测试完成", "任务已完成")
            root.update_idletasks()

            self.assertIsNone(root.grab_current())
            self.assertEqual(str(app.btn_web.cget("state")), tk.NORMAL)
        finally:
            self._destroy_root(root)

    def test_air_conditioning_price_fill_control_is_visible_only_for_air_conditioning(self):
        root = gui_module["create_workbench_root"]()
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
            self._destroy_root(root)

    def test_modern_workbench_keeps_every_task_action_as_a_themed_button(self):
        """视觉升级不能移除操作入口，所有任务按钮都应使用 ttk 主题控件。"""
        root = gui_module["create_workbench_root"]()
        root.withdraw()
        try:
            app = gui_module["PivotTableApp"](root)
            root.update_idletasks()

            task_buttons = (
                app.btn_factory,
                app.btn,
                app.btn_extra,
                app.btn_web,
                app.btn_retry_missing,
                app.btn_backfill,
                app.btn_tracking,
                app.btn_air_price,
            )
            for button in task_buttons:
                with self.subTest(button=button.cget("text")):
                    self.assertEqual(button.winfo_class(), "TButton")
        finally:
            self._destroy_root(root)

    def test_stop_query_keeps_working_with_themed_controls(self):
        """停止查询仍应更新按钮状态，而不是向 ttk 按钮写入原生颜色参数。"""
        root = gui_module["create_workbench_root"]()
        root.withdraw()
        try:
            app = gui_module["PivotTableApp"](root)
            app._web_stop_event = threading.Event()
            app.btn_stop_web.config(state=tk.NORMAL)

            app._on_stop_web_click()

            self.assertTrue(app._web_stop_event.is_set())
            self.assertEqual(str(app.btn_stop_web.cget("state")), tk.DISABLED)
            self.assertEqual(app.btn_stop_web.cget("text"), "⏹ 正在停止...")
        finally:
            self._destroy_root(root)


if __name__ == "__main__":
    unittest.main()
