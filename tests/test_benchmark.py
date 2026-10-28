import csv
import os
import sys
import tempfile
import threading
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from benchmark.data_generator import generate_data  # noqa: E402
from benchmark.metrics import efficiency, median_time, speedup  # noqa: E402
from benchmark.runner import (  # noqa: E402
    CSV_FIELDS, ExperimentConfig, InputError, ResultRow, build_config,
    parse_int_list, run_experiment, save_csv,
)


class DataGeneratorTests(unittest.TestCase):
    def test_kinds(self):
        r = generate_data("random", 500, seed=1)
        self.assertEqual(len(r), 500)
        self.assertEqual(r, generate_data("random", 500, seed=1))  # воспроизводимость
        self.assertEqual(generate_data("sorted", 500, seed=1), sorted(r))
        self.assertEqual(generate_data("reversed", 500, seed=1), sorted(r, reverse=True))

    def test_invalid(self):
        with self.assertRaises(ValueError):
            generate_data("zigzag", 10)
        with self.assertRaises(ValueError):
            generate_data("random", -1)


class MetricsTests(unittest.TestCase):
    def test_median(self):
        self.assertEqual(median_time([3.0, 1.0, 2.0]), 2.0)
        self.assertEqual(median_time([1.0, 2.0, 3.0, 4.0]), 2.5)
        with self.assertRaises(ValueError):
            median_time([])

    def test_speedup_efficiency(self):
        self.assertAlmostEqual(speedup(4.0, 2.0), 2.0)
        self.assertAlmostEqual(efficiency(2.0, 4), 0.5)

    def test_no_division_by_zero(self):
        self.assertIsNone(speedup(1.0, 0.0))
        self.assertIsNone(speedup(0.0, 1.0))
        self.assertIsNone(speedup(None, 1.0))
        self.assertIsNone(efficiency(None, 2))
        self.assertIsNone(efficiency(2.0, 0))


class InputParsingTests(unittest.TestCase):
    def test_parse_list(self):
        self.assertEqual(parse_int_list("10000, 100000; 300000 10000", "x", 1, 10**7), (10000, 100000, 300000))

    def test_parse_errors(self):
        for text in ("", "abc", "1.5", "-5", "0"):
            with self.subTest(text):
                with self.assertRaises(InputError):
                    parse_int_list(text, "Поле", 1, 1000)
        with self.assertRaises(InputError):
            parse_int_list("2000", "Поле", 1, 1000)

    def test_build_config(self):
        cfg = build_config("1000", ["random"], ["merge", "parallel"], "1 2", "3")
        self.assertEqual(cfg.sizes, (1000,))
        self.assertEqual(cfg.workers, (1, 2))
        self.assertEqual(cfg.repeats, 3)

    def test_build_config_errors(self):
        with self.assertRaises(InputError):
            build_config("abc", ["random"], ["merge"], "1", "3")
        with self.assertRaises(InputError):
            build_config("1000", [], ["merge"], "1", "3")
        with self.assertRaises(InputError):
            build_config("1000", ["random"], [], "1", "3")
        with self.assertRaises(InputError):
            build_config("1000", ["random"], ["parallel"], "0", "3")
        with self.assertRaises(InputError):
            build_config("1000", ["random"], ["merge"], "1", "0")
        with self.assertRaises(InputError):
            build_config("1000", ["random"], ["merge"], "1", "1 2")

    def test_workers_ignored_without_parallel(self):
        cfg = build_config("1000", ["random"], ["merge"], "мусор", "1")
        self.assertEqual(cfg.workers, (1,))


class RunnerTests(unittest.TestCase):
    def setUp(self):
        self.cfg = ExperimentConfig(
            sizes=(300,), data_types=("random", "reversed"), repeats=2,
            workers=(1, 2), algorithms=("merge", "builtin", "parallel"), threshold=0,
        )

    def test_run(self):
        progress = []
        rows = run_experiment(self.cfg, on_progress=lambda d, t, m: progress.append((d, t)))
        # 2 типа данных × (merge + builtin + 2 варианта parallel)
        self.assertEqual(len(rows), 2 * 4)
        self.assertTrue(all(r.correct for r in rows))
        self.assertTrue(all(r.median_time > 0 for r in rows))
        self.assertEqual(progress[-1][0], progress[-1][1])
        for r in rows:
            if r.algorithm == "parallel":
                self.assertIsNotNone(r.speedup)
                self.assertAlmostEqual(r.efficiency, r.speedup / r.workers)
            else:
                self.assertIsNone(r.speedup)
                self.assertIsNone(r.efficiency)

    def test_baseline_added_when_only_parallel(self):
        cfg = ExperimentConfig((200,), ("random",), 1, (2,), ("parallel",), threshold=0)
        rows = run_experiment(cfg)
        self.assertEqual({r.algorithm for r in rows}, {"merge", "parallel"})

    def test_stop(self):
        stop = threading.Event()
        stop.set()
        self.assertEqual(run_experiment(self.cfg, stop_event=stop), [])

    def test_wrong_result_is_detected(self):
        # Проверка корректности действительно сравнивает с sorted(): подменим алгоритм.
        import benchmark.runner as runner
        original = runner.merge_sort
        runner.merge_sort = lambda d: list(d)  # «сортировка», не сортирующая
        try:
            cfg = ExperimentConfig((50,), ("random",), 1, (1,), ("merge",))
            rows = run_experiment(cfg)
        finally:
            runner.merge_sort = original
        self.assertFalse(rows[0].correct)


class CsvTests(unittest.TestCase):
    def test_save_csv(self):
        rows = [
            ResultRow("merge", 1000, 1, "random", 3, 0.0123456, None, None, True),
            ResultRow("parallel", 1000, 2, "random", 3, 0.02, 0.61728, 0.30864, True),
        ]
        with tempfile.TemporaryDirectory() as tmp:
            path = save_csv(rows, os.path.join(tmp, "sub", "out.csv"))
            self.assertTrue(path.exists())
            with open(path, encoding="utf-8-sig", newline="") as f:
                data = list(csv.DictReader(f))
        self.assertEqual(list(data[0].keys()), CSV_FIELDS)
        self.assertEqual(len(data), 2)
        self.assertEqual(data[0]["speedup"], "")          # нет фиктивных значений
        self.assertEqual(data[1]["speedup"], "0.6173")
        self.assertEqual(data[1]["workers"], "2")


if __name__ == "__main__":
    unittest.main()
