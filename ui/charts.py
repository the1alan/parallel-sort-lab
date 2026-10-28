"""Графики по фактическим измерениям (Matplotlib). Выдуманных данных здесь нет:
если измерений нет — на графике выводится сообщение «Нет данных»."""
from __future__ import annotations

from pathlib import Path
from typing import Iterable, Optional

from matplotlib.backends.backend_agg import FigureCanvasAgg
from matplotlib.figure import Figure

from benchmark.data_generator import DATA_TYPE_NAMES
from benchmark.runner import ALGORITHM_NAMES, ResultRow

NO_DATA = "Нет данных для построения графика"


def available_data_types(rows: Iterable[ResultRow]) -> list[str]:
    present = {r.data_type for r in rows}
    return [t for t in DATA_TYPE_NAMES if t in present]


def _fmt_size(n: int) -> str:
    return f"{n:,}".replace(",", " ")


def _new_axes(fig: Figure, title: str, data_type: Optional[str]):
    fig.clear()
    ax = fig.add_subplot(111)
    suffix = f" ({DATA_TYPE_NAMES.get(data_type, data_type)})" if data_type else ""
    ax.set_title(title + suffix, fontsize=11)
    ax.grid(True, alpha=0.3)
    return ax


def _no_data(fig: Figure, ax, message: str = NO_DATA) -> None:
    ax.text(0.5, 0.5, message, ha="center", va="center", transform=ax.transAxes, color="#777777")
    ax.set_xticks([])
    ax.set_yticks([])
    _finish(fig)


def _finish(fig: Figure) -> None:
    try:
        fig.tight_layout()
    except Exception:  # слишком маленькое окно — не критично
        pass


def _annotate(ax, xs, ys, fmt):
    for x, y in zip(xs, ys):
        ax.annotate(fmt(y), (x, y), textcoords="offset points", xytext=(0, 6),
                    ha="center", fontsize=7)


def draw_time_chart(fig: Figure, rows: list[ResultRow], data_type: Optional[str]) -> None:
    """Медианное время (с) в зависимости от размера массива, лог. шкала по Y."""
    ax = _new_axes(fig, "Время выполнения алгоритмов", data_type)
    rows = [r for r in rows if r.data_type == data_type]
    if not rows:
        return _no_data(fig, ax)
    sizes = sorted({r.size for r in rows})
    index = {s: i for i, s in enumerate(sizes)}

    series: dict[str, list[ResultRow]] = {}
    for r in rows:
        if r.algorithm == "parallel":
            key = f"{ALGORITHM_NAMES['parallel']} (p={r.workers})"
        else:
            key = ALGORITHM_NAMES[r.algorithm]
        series.setdefault(key, []).append(r)

    n_points = sum(len(v) for v in series.values())
    for name, items in series.items():
        items.sort(key=lambda r: r.size)
        xs = [index[r.size] for r in items]
        ys = [r.median_time for r in items]
        # Не подписываем все точки времени: близкие линии (Merge Sort и p=1)
        # иначе перекрывают друг друга. Точные значения остаются в таблице/CSV.
        ax.plot(xs, ys, marker="o", label=name)
    ax.set_yscale("log")
    ax.set_xticks(range(len(sizes)))
    ax.set_xticklabels([_fmt_size(s) for s in sizes])
    ax.set_xlabel("Размер массива, элементов")
    ax.set_ylabel("Медианное время, с (лог. шкала)")
    ax.legend(fontsize=8)
    _finish(fig)


def _parallel_rows(rows, data_type, attr):
    return [r for r in rows
            if r.algorithm == "parallel" and r.data_type == data_type and getattr(r, attr) is not None]


def _draw_by_processes(fig, rows, data_type, attr, title, ylabel, ideal_label, ideal_fn):
    ax = _new_axes(fig, title, data_type)
    par = _parallel_rows(rows, data_type, attr)
    if not par:
        return _no_data(fig, ax, NO_DATA + "\n(нужен запуск Parallel Merge Sort вместе с последовательным)")
    workers = sorted({r.workers for r in par})
    for size in sorted({r.size for r in par}):
        pts = sorted((r for r in par if r.size == size), key=lambda r: r.workers)
        xs = [r.workers for r in pts]
        ys = [getattr(r, attr) for r in pts]
        ax.plot(xs, ys, marker="o", label=f"n = {_fmt_size(size)}")
        _annotate(ax, xs, ys, lambda v: f"{v:.2f}")
    ideal_x = [min(workers), max(workers)] if len(workers) > 1 else [workers[0] - 0.5, workers[0] + 0.5]
    ax.plot(ideal_x, [ideal_fn(x) for x in ideal_x], "k--", linewidth=1, label=ideal_label)
    ax.set_xticks(workers)
    ax.set_xlabel("Число процессов p")
    ax.set_ylabel(ylabel)
    ax.set_ylim(bottom=0)
    ax.legend(fontsize=8)
    _finish(fig)


def draw_speedup_chart(fig: Figure, rows: list[ResultRow], data_type: Optional[str]) -> None:
    _draw_by_processes(fig, rows, data_type, "speedup",
                       "Ускорение параллельной сортировки", "Ускорение S(p) = T_посл / T_парал(p)",
                       "Идеальное S(p) = p", lambda x: x)


def draw_efficiency_chart(fig: Figure, rows: list[ResultRow], data_type: Optional[str]) -> None:
    _draw_by_processes(fig, rows, data_type, "efficiency",
                       "Эффективность параллельной обработки", "Эффективность E(p) = S(p) / p",
                       "Идеальная E(p) = 1", lambda x: 1.0)


CHARTS = {
    "time": draw_time_chart,
    "speedup": draw_speedup_chart,
    "efficiency": draw_efficiency_chart,
}


def save_charts(rows: list[ResultRow], out_dir) -> list[Path]:
    """Сохраняет PNG для каждого типа данных и каждого вида графика."""
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    saved: list[Path] = []
    for data_type in available_data_types(rows):
        for name, draw in CHARTS.items():
            fig = Figure(figsize=(8, 5))
            FigureCanvasAgg(fig)
            draw(fig, rows, data_type)
            path = out_dir / f"{name}_{data_type}.png"
            fig.savefig(path, dpi=150)
            saved.append(path)
    return saved
