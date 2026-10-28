"""Параллельная сортировка слиянием на ProcessPoolExecutor.

Схема:
  1. Массив делится на `workers` почти равных частей.
  2. Части сортируются ПАРАЛЛЕЛЬНО в отдельных процессах (merge_sort).
  3. Отсортированные части сливаются в главном процессе парами.

Пул процессов создаётся ровно один раз за вызов (не рекурсивно).
Ускорение не гарантировано: на малых массивах расходы на запуск процессов,
сериализацию (pickle) данных и финальное слияние превышают выигрыш.
Заключительное слияние выполняется последовательно (закон Амдала).
"""
from __future__ import annotations

import multiprocessing
from concurrent.futures import ProcessPoolExecutor
from typing import Iterable

from .merge_sort import merge, merge_sort

MAX_WORKERS = 64
# Массивы короче порога сортируются последовательно: процессы не создаются.
DEFAULT_THRESHOLD = 1000


def validate_workers(workers: int) -> int:
    if isinstance(workers, bool) or not isinstance(workers, int):
        raise TypeError("Число процессов должно быть целым числом.")
    if workers < 1 or workers > MAX_WORKERS:
        raise ValueError(f"Число процессов должно быть от 1 до {MAX_WORKERS}.")
    return workers


def split_into_chunks(items: list[int], parts: int) -> list[list[int]]:
    """Делит список на `parts` непрерывных частей, размеры различаются не более чем на 1."""
    if parts < 1:
        raise ValueError("Число частей должно быть положительным.")
    base, extra = divmod(len(items), parts)
    chunks: list[list[int]] = []
    start = 0
    for k in range(parts):
        size = base + (1 if k < extra else 0)
        chunks.append(items[start:start + size])
        start += size
    return chunks


def merge_sorted_chunks(chunks: list[list[int]]) -> list[int]:
    """Попарно сливает отсортированные части (дерево слияния глубиной log2(k))."""
    if not chunks:
        return []
    chunks = list(chunks)
    while len(chunks) > 1:
        merged = [merge(chunks[i], chunks[i + 1]) for i in range(0, len(chunks) - 1, 2)]
        if len(chunks) % 2:
            merged.append(chunks[-1])
        chunks = merged
    return chunks[0]


def parallel_merge_sort(
    data: Iterable[int],
    workers: int = 2,
    threshold: int = DEFAULT_THRESHOLD,
) -> list[int]:
    """Возвращает новый отсортированный список; исходные данные не изменяются.

    workers   — число процессов (1..MAX_WORKERS). При workers=1 пул не создаётся,
                используется последовательный merge_sort.
    threshold — минимальная длина массива для параллельной обработки.
    """
    validate_workers(workers)
    if isinstance(threshold, bool) or not isinstance(threshold, int) or threshold < 0:
        raise ValueError("Порог должен быть неотрицательным целым числом.")

    items = list(data)
    n = len(items)
    if workers == 1 or n < 2 or n < threshold:
        return merge_sort(items)

    chunks = split_into_chunks(items, min(workers, n))
    # «spawn» — единое поведение на macOS/Windows/Linux, безопасно при работе из потока GUI.
    context = multiprocessing.get_context("spawn")
    with ProcessPoolExecutor(max_workers=len(chunks), mp_context=context) as executor:
        sorted_chunks = list(executor.map(merge_sort, chunks))
    return merge_sorted_chunks(sorted_chunks)
