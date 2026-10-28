import os
import random
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from algorithms import MAX_WORKERS, merge, merge_sort, parallel_merge_sort  # noqa: E402
from algorithms.parallel_merge_sort import merge_sorted_chunks, split_into_chunks  # noqa: E402


def sample_cases():
    rng = random.Random(1)
    return {
        "empty": [],
        "single": [7],
        "few": [5, 2, 9, 1, 5, 6],
        "negative": [-3, 5, -10, 0, 7, -1],
        "duplicates": [3, 1, 3, 1, 2, 2, 3, 1],
        "all_equal": [4] * 50,
        "sorted": list(range(200)),
        "reversed": list(range(200, 0, -1)),
        "random": [rng.randint(-1000, 1000) for _ in range(5000)],
    }


class MergeSortTests(unittest.TestCase):
    def test_matches_sorted(self):
        for name, data in sample_cases().items():
            with self.subTest(name):
                self.assertEqual(merge_sort(data), sorted(data))

    def test_input_not_modified(self):
        data = [5, 3, 1, 4]
        original = list(data)
        result = merge_sort(data)
        self.assertEqual(data, original)
        self.assertIsNot(result, data)

    def test_accepts_tuple(self):
        self.assertEqual(merge_sort((3, 1, 2)), [1, 2, 3])

    def test_merge_function(self):
        self.assertEqual(merge([1, 4, 6], [2, 3, 7, 8]), [1, 2, 3, 4, 6, 7, 8])
        self.assertEqual(merge([], [1]), [1])
        self.assertEqual(merge([], []), [])

    def test_stable_merge(self):
        # равные элементы из левого списка идут раньше (проверка по идентичности)
        a, b = [1], [1]
        self.assertIs(merge(a, b)[0], a[0])


class ParallelMergeSortTests(unittest.TestCase):
    def test_matches_sorted_with_processes(self):
        # Один реальный запуск пула проверяет интеграцию multiprocessing.
        # Остальные граничные случаи ниже проверяются без запуска процессов,
        # чтобы весь набор тестов не создавал десятки дорогих spawn-пулов.
        data = sample_cases()["random"][:500]
        original = list(data)
        self.assertEqual(parallel_merge_sort(data, workers=2, threshold=0), sorted(data))
        self.assertEqual(data, original)

    def test_parallel_small_cases_use_sequential_fallback(self):
        for data in ([], [1], [3, 1, 2], [-2, 0, -2, 9], [4, 4, 4]):
            with self.subTest(data=data):
                self.assertEqual(parallel_merge_sort(data, workers=4, threshold=100), sorted(data))

    def test_edge_cases(self):
        for name in ("empty", "single"):
            data = sample_cases()[name]
            self.assertEqual(parallel_merge_sort(data, workers=4, threshold=0), sorted(data))

    def test_workers_one_is_sequential(self):
        data = sample_cases()["random"]
        self.assertEqual(parallel_merge_sort(data, workers=1), sorted(data))

    def test_below_threshold_is_sequential(self):
        data = [3, 2, 1]
        self.assertEqual(parallel_merge_sort(data, workers=4, threshold=10_000), [1, 2, 3])

    def test_more_workers_than_elements(self):
        self.assertEqual(parallel_merge_sort([2, 1], workers=8, threshold=0), [1, 2])

    def test_input_not_modified(self):
        data = sample_cases()["random"]
        original = list(data)
        parallel_merge_sort(data, workers=2, threshold=0)
        self.assertEqual(data, original)

    def test_invalid_workers(self):
        for bad in (0, -1, MAX_WORKERS + 1):
            with self.assertRaises(ValueError):
                parallel_merge_sort([1, 2, 3], workers=bad)
        for bad in (2.5, "2", None, True):
            with self.assertRaises(TypeError):
                parallel_merge_sort([1, 2, 3], workers=bad)

    def test_invalid_threshold(self):
        with self.assertRaises(ValueError):
            parallel_merge_sort([1, 2, 3], workers=2, threshold=-5)

    def test_helpers(self):
        chunks = split_into_chunks(list(range(10)), 3)
        self.assertEqual([len(c) for c in chunks], [4, 3, 3])
        self.assertEqual(sum(chunks, []), list(range(10)))
        self.assertEqual(merge_sorted_chunks([[1, 5], [2, 6], [0, 9]]), [0, 1, 2, 5, 6, 9])
        self.assertEqual(merge_sorted_chunks([]), [])


if __name__ == "__main__":
    unittest.main()
