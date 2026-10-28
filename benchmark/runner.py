"""Запуск экспериментов: честное сравнение трёх вариантов сортировки.

Условия честности:
  * массив генерируется один раз для (размер, тип) и НЕ входит в измеряемое время;
  * всем алгоритмам передаются одни и те же данные;
  * измеряется полное время вызова (для параллельной версии — включая запуск пула,
    передачу данных процессам и финальное слияние);
  * проверка корректности выполняется ПОСЛЕ замера и в время не входит;
  * по нескольким повторам берётся медиана.
"""
from __future__ import annotations

import csv
import re
import threading
import time
from dataclasses import dataclass
from functools import partial
from pathlib import Path
from typing import Callable, Optional

from algorithms import DEFAULT_THRESHOLD, MAX_WORKERS, merge_sort, parallel_merge_sort

from .data_generator import DATA_TYPES, generate_data
from .metrics import efficiency, median_time, speedup

ALGORITHM_NAMES = {
    "merge": "Merge Sort (последовательный)",
    "parallel": "Parallel Merge Sort",
    "builtin": "sorted() (встроенная)",
}
ALGORITHM_ORDER = ("merge", "builtin", "parallel")

MAX_SIZE = 5_000_000
MAX_REPEATS = 50

CSV_FIELDS = [
    "algorithm", "size", "workers", "data_type", "repeats",
    "median_time_s", "speedup", "efficiency", "correct",
]


class InputError(ValueError):
    """Некорректные параметры эксперимента (сообщение — на русском, для пользователя)."""


@dataclass(frozen=True)
class ExperimentConfig:
    sizes: tuple[int, ...]
    data_types: tuple[str, ...]
    repeats: int
    workers: tuple[int, ...]
    algorithms: tuple[str, ...]
    seed: int = 42
    threshold: int = DEFAULT_THRESHOLD

    def validate(self) -> None:
        if not self.sizes or any(s < 1 or s > MAX_SIZE for s in self.sizes):
            raise InputError(f"Размер массива должен быть от 1 до {MAX_SIZE:,}.".replace(",", " "))
        if not self.data_types or any(t not in DATA_TYPES for t in self.data_types):
            raise InputError("Выберите допустимые типы данных.")
        if not 1 <= self.repeats <= MAX_REPEATS:
            raise InputError(f"Число повторов должно быть от 1 до {MAX_REPEATS}.")
        if not self.algorithms or any(a not in ALGORITHM_NAMES for a in self.algorithms):
            raise InputError("Выберите хотя бы один алгоритм.")
        if "parallel" in self.algorithms and (
            not self.workers or any(w < 1 or w > MAX_WORKERS for w in self.workers)
        ):
            raise InputError(f"Число процессов должно быть от 1 до {MAX_WORKERS}.")
        if self.threshold < 0:
            raise InputError("Порог параллельности не может быть отрицательным.")


@dataclass(frozen=True)
class ResultRow:
    algorithm: str          # ключ: merge / parallel / builtin
    size: int
    workers: int            # для последовательных алгоритмов всегда 1
    data_type: str
    repeats: int
    median_time: float      # секунды
    speedup: Optional[float]
    efficiency: Optional[float]
    correct: bool

    @property
    def algorithm_name(self) -> str:
        return ALGORITHM_NAMES[self.algorithm]


# ----------------------------------------------------------------- разбор ввода

def parse_int_list(text: str, field: str, min_value: int, max_value: int) -> tuple[int, ...]:
    """Разбирает «10000, 100000; 300000» в кортеж без повторов с проверкой диапазона."""
    tokens = [t for t in re.split(r"[,;\s]+", (text or "").strip()) if t]
    if not tokens:
        raise InputError(f"Поле «{field}»: укажите хотя бы одно значение.")
    values: list[int] = []
    for token in tokens:
        try:
            value = int(token.replace("_", ""))
        except ValueError:
            raise InputError(
                f"Поле «{field}»: «{token}» — не целое число. Пример: 10000, 100000."
            ) from None
        if not min_value <= value <= max_value:
            raise InputError(f"Поле «{field}»: значение {value} вне диапазона {min_value}…{max_value}.")
        if value not in values:
            values.append(value)
    return tuple(values)


def parse_single_int(text: str, field: str, min_value: int, max_value: int) -> int:
    values = parse_int_list(text, field, min_value, max_value)
    if len(values) != 1:
        raise InputError(f"Поле «{field}»: нужно ровно одно число.")
    return values[0]


def build_config(
    sizes_text: str,
    data_types,
    algorithms,
    workers_text: str,
    repeats_text: str,
    threshold_text: str = str(DEFAULT_THRESHOLD),
    seed: int = 42,
) -> ExperimentConfig:
    """Собирает и проверяет конфигурацию из строк, введённых пользователем."""
    sizes = parse_int_list(sizes_text, "Размер массива", 1, MAX_SIZE)
    data_types = tuple(t for t in DATA_TYPES if t in set(data_types))
    if not data_types:
        raise InputError("Выберите хотя бы один тип данных.")
    algorithms = tuple(a for a in ALGORITHM_ORDER if a in set(algorithms))
    if not algorithms:
        raise InputError("Выберите хотя бы один алгоритм.")
    workers: tuple[int, ...] = (1,)
    if "parallel" in algorithms:
        workers = parse_int_list(workers_text, "Число процессов", 1, MAX_WORKERS)
    repeats = parse_single_int(repeats_text, "Число повторов", 1, MAX_REPEATS)
    threshold = parse_single_int(threshold_text, "Порог параллельности", 0, MAX_SIZE)
    config = ExperimentConfig(sizes, data_types, repeats, workers, algorithms, seed, threshold)
    config.validate()
    return config


# ------------------------------------------------------------------ эксперимент

ProgressCallback = Callable[[int, int, str], None]
RowCallback = Callable[[ResultRow], None]


def _measure(func, data, expected, snapshot, repeats, stop_event, report):
    """Возвращает (времена, корректность) или None, если эксперимент остановлен."""
    times: list[float] = []
    correct = True
    for i in range(1, repeats + 1):
        if stop_event is not None and stop_event.is_set():
            return None
        start = time.perf_counter()
        result = func(data)
        elapsed = time.perf_counter() - start
        # Проверка — вне измеряемого времени.
        correct = correct and result == expected and data == snapshot
        times.append(elapsed)
        report(i)
    return times, correct


def run_experiment(
    config: ExperimentConfig,
    on_progress: Optional[ProgressCallback] = None,
    on_row: Optional[RowCallback] = None,
    stop_event: Optional[threading.Event] = None,
) -> list[ResultRow]:
    """Выполняет эксперимент. При остановке возвращает уже полученные строки.

    Остановка проверяется между отдельными замерами: уже начатая сортировка
    доводится до конца (прерывать пул процессов посреди работы небезопасно).
    Если выбран parallel, последовательный Merge Sort измеряется всегда —
    он служит базой для ускорения S(p).
    """
    config.validate()
    algos = [a for a in ALGORITHM_ORDER if a in config.algorithms]
    if "parallel" in algos and "merge" not in algos:
        algos.insert(0, "merge")

    per_group = sum(config.repeats * (len(config.workers) if a == "parallel" else 1) for a in algos)
    total = per_group * len(config.sizes) * len(config.data_types)
    done = 0
    rows: list[ResultRow] = []

    for size in config.sizes:
        for kind in config.data_types:
            data = generate_data(kind, size, config.seed)   # вне замера
            expected = sorted(data)                          # эталон для проверки
            snapshot = list(data)                            # контроль неизменности входа
            baseline: Optional[float] = None

            for algo in algos:
                variants = config.workers if algo == "parallel" else (1,)
                for p in variants:
                    if algo == "merge":
                        func = merge_sort
                    elif algo == "builtin":
                        func = sorted
                    else:
                        func = partial(parallel_merge_sort, workers=p, threshold=config.threshold)

                    def report(i, algo=algo, p=p, size=size, kind=kind, base=done):
                        if on_progress:
                            label = ALGORITHM_NAMES[algo] + (f", p={p}" if algo == "parallel" else "")
                            on_progress(base + i, total, f"{label}: n={size}, {kind}, повтор {i}/{config.repeats}")

                    outcome = _measure(func, data, expected, snapshot, config.repeats, stop_event, report)
                    if outcome is None:
                        return rows
                    times, correct = outcome
                    done += config.repeats

                    med = median_time(times)
                    if algo == "merge":
                        baseline = med
                    s = speedup(baseline, med) if algo == "parallel" else None
                    e = efficiency(s, p) if s is not None else None
                    row = ResultRow(algo, size, p, kind, config.repeats, med, s, e, correct)
                    rows.append(row)
                    if on_row:
                        on_row(row)
    return rows


# ------------------------------------------------------------------------- CSV

def save_csv(rows: list[ResultRow], path) -> Path:
    """Сохраняет результаты в CSV (UTF-8 с BOM — корректно открывается в Excel/Numbers)."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_FIELDS)
        writer.writeheader()
        for r in rows:
            writer.writerow({
                "algorithm": r.algorithm_name,
                "size": r.size,
                "workers": r.workers,
                "data_type": r.data_type,
                "repeats": r.repeats,
                "median_time_s": f"{r.median_time:.6f}",
                "speedup": "" if r.speedup is None else f"{r.speedup:.4f}",
                "efficiency": "" if r.efficiency is None else f"{r.efficiency:.4f}",
                "correct": r.correct,
            })
    return path
