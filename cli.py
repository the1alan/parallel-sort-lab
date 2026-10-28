"""Консольный запуск эксперимента без GUI (запасной вариант и быстрая проверка).

Пример:
    python3 cli.py --sizes 10000 100000 --types random --workers 1 2 4 --repeats 3 \
        --csv results/run.csv --charts results/charts
"""
import argparse
import multiprocessing
import sys

from benchmark.runner import ALGORITHM_NAMES, InputError, build_config, run_experiment, save_csv


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Parallel Sort Lab — консольный запуск")
    ap.add_argument("--sizes", nargs="+", default=["10000"])
    ap.add_argument("--types", nargs="+", default=["random"], choices=["random", "sorted", "reversed"])
    ap.add_argument("--algorithms", nargs="+", default=list(ALGORITHM_NAMES), choices=list(ALGORITHM_NAMES))
    ap.add_argument("--workers", nargs="+", default=["1", "2", "4"])
    ap.add_argument("--repeats", default="3")
    ap.add_argument("--threshold", default="1000")
    ap.add_argument("--csv", help="путь для сохранения CSV")
    ap.add_argument("--charts", help="папка для сохранения графиков PNG")
    args = ap.parse_args(argv)

    try:
        config = build_config(" ".join(args.sizes), args.types, args.algorithms,
                              " ".join(args.workers), args.repeats, args.threshold)
    except InputError as exc:
        print(f"Ошибка ввода: {exc}", file=sys.stderr)
        return 2

    def progress(done, total, msg):
        print(f"[{done}/{total}] {msg}", flush=True)

    rows = run_experiment(config, on_progress=progress)
    print(f"\n{'Алгоритм':32} {'n':>9} {'p':>2} {'данные':9} {'время, с':>10} {'S':>7} {'E':>7} ок")
    for r in rows:
        s = "—" if r.speedup is None else f"{r.speedup:.2f}"
        e = "—" if r.efficiency is None else f"{r.efficiency:.2f}"
        print(f"{r.algorithm_name:32} {r.size:>9} {r.workers:>2} {r.data_type:9} "
              f"{r.median_time:>10.5f} {s:>7} {e:>7} {'да' if r.correct else 'НЕТ'}")
    if args.csv:
        print("CSV:", save_csv(rows, args.csv))
    if args.charts:
        from ui.charts import save_charts
        for p in save_charts(rows, args.charts):
            print("График:", p)
    return 0 if all(r.correct for r in rows) else 1


if __name__ == "__main__":
    multiprocessing.freeze_support()
    sys.exit(main())
