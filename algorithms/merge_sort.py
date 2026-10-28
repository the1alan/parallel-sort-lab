"""Последовательная сортировка слиянием (Merge Sort).

Сложность: время O(n log n) в худшем, среднем и лучшем случае;
память O(n) (временные списки при слиянии). Сортировка устойчивая.
"""
from __future__ import annotations

from typing import Iterable


def merge(left: list[int], right: list[int]) -> list[int]:
    """Сливает два отсортированных списка в один отсортированный за O(len(left)+len(right)).

    Знак «<=» при сравнении сохраняет порядок равных элементов (устойчивость).
    """
    result: list[int] = []
    append = result.append
    i = j = 0
    n_left, n_right = len(left), len(right)
    while i < n_left and j < n_right:
        if left[i] <= right[j]:
            append(left[i])
            i += 1
        else:
            append(right[j])
            j += 1
    # Один из списков исчерпан — остаток другого уже отсортирован.
    result.extend(left[i:])
    result.extend(right[j:])
    return result


def _sort(items: list[int]) -> list[int]:
    n = len(items)
    if n <= 1:
        return items
    mid = n // 2
    return merge(_sort(items[:mid]), _sort(items[mid:]))


def merge_sort(data: Iterable[int]) -> list[int]:
    """Возвращает НОВЫЙ отсортированный список; исходные данные не изменяются.

    Глубина рекурсии — log2(n), поэтому переполнения стека для больших массивов нет.
    """
    return _sort(list(data))
