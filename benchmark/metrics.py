"""Метрики параллельной обработки.

Speedup:    S(p) = T_sequential / T_parallel(p)
Efficiency: E(p) = S(p) / p
Если метрика неприменима (нет времени, нулевой знаменатель) — возвращается None,
а не фиктивное значение.
"""
from __future__ import annotations

import statistics
from typing import Optional, Sequence


def median_time(times: Sequence[float]) -> float:
    if not times:
        raise ValueError("Нет измерений для вычисления медианы.")
    return float(statistics.median(times))


def speedup(t_sequential: Optional[float], t_parallel: Optional[float]) -> Optional[float]:
    if t_sequential is None or t_parallel is None:
        return None
    if t_sequential <= 0 or t_parallel <= 0:
        return None
    return t_sequential / t_parallel


def efficiency(speedup_value: Optional[float], processes: int) -> Optional[float]:
    if speedup_value is None or processes is None or processes < 1:
        return None
    return speedup_value / processes
