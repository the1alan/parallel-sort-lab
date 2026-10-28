"""Дымовой тест GUI: пропускается, если нет Tkinter или графического дисплея."""
import os
import sys
import time
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

try:
    import tkinter as tk
    _root = tk.Tk()
    _root.destroy()
    HAS_TK = True
except Exception:  # нет модуля или дисплея
    HAS_TK = False


@unittest.skipUnless(HAS_TK, "Tkinter или дисплей недоступны")
class UiSmokeTests(unittest.TestCase):
    def test_run_experiment_in_gui(self):
        from ui.app import ParallelSortApp
        root = tk.Tk()
        try:
            app = ParallelSortApp(root)
            app.sizes_var.set("500")
            app.workers_var.set("1, 2")
            app.repeats_var.set("2")
            app.threshold_var.set("0")
            app.start_run()
            deadline = time.time() + 60
            while (app.is_running or not app.queue.empty()) and time.time() < deadline:
                root.update()
                time.sleep(0.02)
            for _ in range(5):
                root.update()
                time.sleep(0.12)
            self.assertEqual(len(app.rows), 4)  # merge, builtin, parallel p=1, p=2
            self.assertEqual(len(app.tree.get_children()), 4)
            self.assertTrue(all(r.correct for r in app.rows))
            self.assertTrue(app.chart_type_var.get())
        finally:
            app._on_close()

    def test_invalid_input_shows_error_without_running(self):
        from unittest import mock
        from ui.app import ParallelSortApp
        root = tk.Tk()
        try:
            app = ParallelSortApp(root)
            app.sizes_var.set("abc")
            with mock.patch("ui.app.messagebox.showerror") as err:
                app.start_run()
            err.assert_called_once()
            self.assertFalse(app.is_running)
        finally:
            app._on_close()


if __name__ == "__main__":
    unittest.main()
