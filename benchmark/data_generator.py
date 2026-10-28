"""Генерация входных данных. Время генерации в измерения не входит."""
from __future__ import annotations

import random

DATA_TYPES = ("random", "sorted", "reversed")
DATA_TYPE_NAMES = {
    "random": "случайные",
    "sorted": "отсортированные",
    "reversed": "обратный порядок",
}


def generate_data(kind: str, size: int, seed: int = 42) -> list[int]:
    """Возвращает список из `size` целых чисел.

    random   — случайные значения (с фиксированным seed → эксперимент воспроизводим);
    sorted   — те же значения по возрастанию;
    reversed — те же значения по убыванию.
    """
    if kind not in DATA_TYPES:
        raise ValueError(f"Неизвестный тип данных: {kind!r}. Допустимо: {', '.join(DATA_TYPES)}.")
    if isinstance(size, bool) or not isinstance(size, int) or size < 0:
        raise ValueError("Размер массива должен быть неотрицательным целым числом.")
    rng = random.Random(seed)
    data = [rng.randint(-10**6, 10**6) for _ in range(size)]
    if kind == "sorted":
        data.sort()
    elif kind == "reversed":
        data.sort(reverse=True)
    return data
