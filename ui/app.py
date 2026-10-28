"""Графический интерфейс Parallel Sort Lab (Tkinter/ttk, русский язык).

Измерения выполняются в фоновом потоке. Поток НЕ обращается к виджетам:
он кладёт сообщения в queue.Queue, а главный поток раз в 100 мс забирает их
через root.after() и обновляет интерфейс.
"""
from __future__ import annotations

import queue
import sys
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

from algorithms import DEFAULT_THRESHOLD
from benchmark.data_generator import DATA_TYPE_NAMES, DATA_TYPES
from benchmark.runner import (
    ALGORITHM_NAMES, ALGORITHM_ORDER, ExperimentConfig, InputError, ResultRow,
    build_config, run_experiment, save_csv,
)

from .charts import CHARTS, available_data_types, save_charts

PROJECT_ROOT = Path(__file__).resolve().parent.parent
RESULTS_DIR = PROJECT_ROOT / "results"

ALGORITHM_INFO = """\
Merge Sort (последовательный)  — algorithms/merge_sort.py
  Делит массив пополам, рекурсивно сортирует половины и сливает их.
  Время: O(n log n) во всех случаях. Память: O(n). Устойчивая.

Parallel Merge Sort  — algorithms/parallel_merge_sort.py
  Массив делится на p частей; части сортируются параллельно в p процессах
  (ProcessPoolExecutor), затем сливаются в главном процессе.
  Время: O((n/p)·log(n/p)) на этапе сортировки + O(n·log p) на слияние
  + накладные расходы (запуск процессов, pickle-передача данных).
  Память: O(n). Ускорение НЕ гарантировано: на малых n накладные расходы
  больше выигрыша; финальное слияние последовательное (закон Амдала).
  При p = 1 пул не создаётся — используется последовательный Merge Sort.

sorted() (встроенная)  — эталон производительности, НЕ собственная реализация
  Timsort на языке C: высокооптимизированный, O(n log n), на уже
  отсортированных данных — O(n). Служит практическим ориентиром.

Метрики
  Ускорение     S(p) = T_последовательного Merge Sort / T_параллельного(p)
  Эффективность E(p) = S(p) / p
  Время — медиана по повторам; корректность — сравнение с sorted(входа).
"""

COLUMNS = (
    ("algorithm", "Алгоритм", 210, "w"),
    ("size", "Размер", 60, "e"),
    ("workers", "Проц.", 60, "center"),
    ("data_type", "Данные", 85, "w"),
    ("median", "Медиана, с", 125, "e"),
    ("speedup", "S(p)", 60, "e"),
    ("efficiency", "E(p)", 60, "e"),
    ("correct", "Верно?", 70, "center"),
)


class ParallelSortApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        root.title("Parallel Sort Lab — параллельные алгоритмы сортировки")
        root.geometry("1280x780")
        root.minsize(1100, 680)
        style = ttk.Style(root)
        if sys.platform != "darwin" and "clam" in style.theme_names():
            style.theme_use("clam")
        style.configure("Title.TLabel", font=("TkDefaultFont", 11, "bold"))
        style.configure("Hint.TLabel", foreground="#666666")
        style.configure("Treeview", rowheight=24)

        self.queue: queue.Queue = queue.Queue()
        self.rows: list[ResultRow] = []
        self.worker: threading.Thread | None = None
        self.stop_event = threading.Event()
        self.last_config: ExperimentConfig | None = None

        self._build_settings()
        self._build_main()
        self._build_status()
        root.columnconfigure(1, weight=1)
        root.rowconfigure(0, weight=1)
        root.protocol("WM_DELETE_WINDOW", self._on_close)
        self._set_running(False)
        self._poll_queue()

    # ------------------------------------------------------------------ виджеты
    def _build_settings(self) -> None:
        frame = ttk.LabelFrame(self.root, text=" Параметры эксперимента ", padding=12)
        frame.grid(row=0, column=0, sticky="ns", padx=(12, 6), pady=12)

        self.sizes_var = tk.StringVar(value="10000, 100000")
        self.workers_var = tk.StringVar(value="1, 2, 4")
        self.repeats_var = tk.StringVar(value="3")
        self.threshold_var = tk.StringVar(value=str(DEFAULT_THRESHOLD))
        self.type_vars = {t: tk.BooleanVar(value=(t == "random")) for t in DATA_TYPES}
        self.algo_vars = {a: tk.BooleanVar(value=True) for a in ALGORITHM_ORDER}

        r = 0
        ttk.Label(frame, text="Размер массива", style="Title.TLabel").grid(row=r, sticky="w"); r += 1
        ttk.Entry(frame, textvariable=self.sizes_var, width=24).grid(row=r, sticky="ew"); r += 1
        ttk.Label(frame, text="через запятую; напр.: 10000, 100000, 300000", style="Hint.TLabel",
                  wraplength=220).grid(row=r, sticky="w", pady=(0, 10)); r += 1

        ttk.Label(frame, text="Тип данных", style="Title.TLabel").grid(row=r, sticky="w"); r += 1
        for t in DATA_TYPES:
            ttk.Checkbutton(frame, text=DATA_TYPE_NAMES[t], variable=self.type_vars[t]).grid(row=r, sticky="w"); r += 1

        ttk.Label(frame, text="Алгоритмы", style="Title.TLabel").grid(row=r, sticky="w", pady=(10, 0)); r += 1
        for a in ALGORITHM_ORDER:
            ttk.Checkbutton(frame, text=ALGORITHM_NAMES[a], variable=self.algo_vars[a]).grid(row=r, sticky="w"); r += 1

        ttk.Label(frame, text="Число процессов", style="Title.TLabel").grid(row=r, sticky="w", pady=(10, 0)); r += 1
        ttk.Entry(frame, textvariable=self.workers_var, width=24).grid(row=r, sticky="ew"); r += 1
        ttk.Label(frame, text="для Parallel Merge Sort; напр.: 1, 2, 4", style="Hint.TLabel").grid(
            row=r, sticky="w", pady=(0, 10)); r += 1

        ttk.Label(frame, text="Число повторов", style="Title.TLabel").grid(row=r, sticky="w"); r += 1
        ttk.Spinbox(frame, from_=1, to=50, textvariable=self.repeats_var, width=8).grid(
            row=r, sticky="w", pady=(0, 10)); r += 1

        ttk.Label(frame, text="Порог параллельности", style="Title.TLabel").grid(row=r, sticky="w"); r += 1
        ttk.Entry(frame, textvariable=self.threshold_var, width=12).grid(row=r, sticky="w"); r += 1
        ttk.Label(frame, text="массивы короче — сортируются последовательно", style="Hint.TLabel",
                  wraplength=220).grid(row=r, sticky="w", pady=(0, 12)); r += 1

        self.run_btn = ttk.Button(frame, text="▶  Запустить", command=self.start_run)
        self.run_btn.grid(row=r, sticky="ew"); r += 1
        self.stop_btn = ttk.Button(frame, text="■  Остановить", command=self.stop_run)
        self.stop_btn.grid(row=r, sticky="ew", pady=4); r += 1
        self.repeat_btn = ttk.Button(frame, text="↻  Повторить эксперимент", command=self.repeat_run)
        self.repeat_btn.grid(row=r, sticky="ew"); r += 1

    def _build_main(self) -> None:
        self.tabs = ttk.Notebook(self.root)
        self.tabs.grid(row=0, column=1, sticky="nsew", padx=(6, 12), pady=12)

        # --- таблица результатов
        table_tab = ttk.Frame(self.tabs, padding=6)
        self.tabs.add(table_tab, text="Результаты")
        self.tree = ttk.Treeview(table_tab, columns=[c[0] for c in COLUMNS], show="headings")
        for key, title, width, anchor in COLUMNS:
            self.tree.heading(key, text=title)
            self.tree.column(key, width=width, anchor=anchor, stretch=True)
        self.tree.tag_configure("bad", foreground="#c0392b")
        scroll = ttk.Scrollbar(table_tab, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")
        table_tab.rowconfigure(0, weight=1)
        table_tab.columnconfigure(0, weight=1)

        # --- графики
        chart_tab = ttk.Frame(self.tabs, padding=6)
        self.tabs.add(chart_tab, text="Графики")
        top = ttk.Frame(chart_tab)
        top.pack(fill="x")
        ttk.Label(top, text="Тип данных на графиках:").pack(side="left")
        self.chart_type_var = tk.StringVar()
        self.chart_type_box = ttk.Combobox(top, textvariable=self.chart_type_var, state="readonly", width=22)
        self.chart_type_box.pack(side="left", padx=8)
        self.chart_type_box.bind("<<ComboboxSelected>>", lambda e: self.redraw_charts())
        self.chart_tabs = ttk.Notebook(chart_tab)
        self.chart_tabs.pack(fill="both", expand=True, pady=(6, 0))
        self.figures: dict[str, tuple[Figure, FigureCanvasTkAgg]] = {}
        for key, title in (("time", "Время"), ("speedup", "Ускорение"), ("efficiency", "Эффективность")):
            tab = ttk.Frame(self.chart_tabs)
            self.chart_tabs.add(tab, text=title)
            fig = Figure(figsize=(7, 4.5), dpi=100)
            canvas = FigureCanvasTkAgg(fig, master=tab)
            canvas.get_tk_widget().pack(fill="both", expand=True)
            self.figures[key] = (fig, canvas)
        # Перерисовка при показе вкладки — чтобы график занимал всю доступную область.
        for nb in (self.tabs, self.chart_tabs):
            nb.bind("<<NotebookTabChanged>>", lambda e: self.root.after(80, self.redraw_charts))
        self.redraw_charts()

        # --- справка
        info_tab = ttk.Frame(self.tabs, padding=10)
        self.tabs.add(info_tab, text="Об алгоритмах")
        text = tk.Text(info_tab, wrap="word", relief="flat", padx=6, pady=6)
        text.insert("1.0", ALGORITHM_INFO)
        text.configure(state="disabled")
        text.pack(fill="both", expand=True)

    def _build_status(self) -> None:
        bar = ttk.Frame(self.root, padding=(12, 0, 12, 10))
        bar.grid(row=1, column=0, columnspan=2, sticky="ew")
        bar.columnconfigure(1, weight=1)
        self.csv_btn = ttk.Button(bar, text="Экспорт в CSV…", command=self.export_csv)
        self.csv_btn.grid(row=0, column=2, padx=(10, 4))
        self.png_btn = ttk.Button(bar, text="Сохранить графики (PNG)…", command=self.export_charts)
        self.png_btn.grid(row=0, column=3)
        self.progress = ttk.Progressbar(bar, length=260, mode="determinate")
        self.progress.grid(row=0, column=0, sticky="w")
        self.status_var = tk.StringVar(value="Готов к запуску. Задайте параметры и нажмите «Запустить».")
        ttk.Label(bar, textvariable=self.status_var).grid(row=0, column=1, sticky="w", padx=10)

    # ------------------------------------------------------------ состояние/запуск
    @property
    def is_running(self) -> bool:
        return self.worker is not None and self.worker.is_alive()

    def _set_running(self, running: bool) -> None:
        self.run_btn.state(["disabled"] if running else ["!disabled"])
        self.repeat_btn.state(["disabled"] if running or self.last_config is None else ["!disabled"])
        self.stop_btn.state(["!disabled"] if running else ["disabled"])
        for b in (self.csv_btn, self.png_btn):
            b.state(["disabled"] if running or not self.rows else ["!disabled"])

    def _collect_config(self) -> ExperimentConfig:
        return build_config(
            self.sizes_var.get(),
            [t for t, v in self.type_vars.items() if v.get()],
            [a for a, v in self.algo_vars.items() if v.get()],
            self.workers_var.get(),
            self.repeats_var.get(),
            self.threshold_var.get(),
        )

    def start_run(self, config: ExperimentConfig | None = None) -> None:
        if self.is_running:
            return
        try:
            config = config or self._collect_config()
        except InputError as exc:
            messagebox.showerror("Некорректные параметры", str(exc), parent=self.root)
            return
        if max(config.sizes) > 1_000_000 and not messagebox.askokcancel(
            "Большой массив",
            "Размер больше 1 000 000 элементов: эксперимент может занять много времени "
            "и памяти. Продолжить?", parent=self.root):
            return
        self.last_config = config
        self.rows.clear()
        self.tree.delete(*self.tree.get_children())
        self.stop_event.clear()
        self.progress.configure(value=0, maximum=1)
        self.status_var.set("Эксперимент выполняется…")
        self.tabs.select(0)
        self.worker = threading.Thread(target=self._worker_main, args=(config,), daemon=True)
        self._set_running(True)
        self.worker.start()

    def repeat_run(self) -> None:
        if self.last_config is not None:
            self.start_run(self.last_config)

    def stop_run(self) -> None:
        if self.is_running:
            self.stop_event.set()
            self.status_var.set("Остановка: дожидаемся завершения текущего замера…")
            self.stop_btn.state(["disabled"])

    def _worker_main(self, config: ExperimentConfig) -> None:
        """Фоновый поток: только вычисления и queue.put — никаких обращений к Tk."""
        try:
            run_experiment(
                config,
                on_progress=lambda d, t, m: self.queue.put(("progress", d, t, m)),
                on_row=lambda row: self.queue.put(("row", row)),
                stop_event=self.stop_event,
            )
            self.queue.put(("done", self.stop_event.is_set()))
        except Exception as exc:  # noqa: BLE001 — показываем пользователю любую ошибку
            self.queue.put(("error", f"{type(exc).__name__}: {exc}"))

    def _poll_queue(self) -> None:
        try:
            while True:
                msg = self.queue.get_nowait()
                kind = msg[0]
                if kind == "progress":
                    _, done, total, text = msg
                    self.progress.configure(maximum=max(total, 1), value=done)
                    if not self.stop_event.is_set():
                        self.status_var.set(f"[{done}/{total}] {text}")
                elif kind == "row":
                    self._add_row(msg[1])
                elif kind == "done":
                    self._finish(stopped=msg[1])
                elif kind == "error":
                    self._finish(error=msg[1])
        except queue.Empty:
            pass
        self._after_id = self.root.after(100, self._poll_queue)

    def _add_row(self, row: ResultRow) -> None:
        self.rows.append(row)
        values = (
            row.algorithm_name,
            f"{row.size:,}".replace(",", " "),
            row.workers,
            DATA_TYPE_NAMES[row.data_type],
            f"{row.median_time:.6f}",
            "—" if row.speedup is None else f"{row.speedup:.2f}",
            "—" if row.efficiency is None else f"{row.efficiency:.2f}",
            "да" if row.correct else "НЕТ",
        )
        item = self.tree.insert("", "end", values=values, tags=(() if row.correct else ("bad",)))
        self.tree.see(item)

    def _finish(self, stopped: bool = False, error: str | None = None) -> None:
        self._refresh_chart_choices()
        self.redraw_charts()
        self._set_running(False)
        if error:
            self.status_var.set("Ошибка при выполнении эксперимента.")
            messagebox.showerror("Ошибка", error, parent=self.root)
        elif stopped:
            self.status_var.set(f"Остановлено пользователем. Получено строк: {len(self.rows)} (результаты неполные).")
        else:
            bad = sum(1 for r in self.rows if not r.correct)
            note = "все результаты корректны" if not bad else f"ВНИМАНИЕ: некорректных результатов: {bad}"
            self.status_var.set(f"Готово. Строк в таблице: {len(self.rows)}; {note}.")

    # -------------------------------------------------------------------- графики
    def _refresh_chart_choices(self) -> None:
        types = available_data_types(self.rows)
        self.chart_type_box.configure(values=[DATA_TYPE_NAMES[t] for t in types])
        current = self.chart_type_var.get()
        names = [DATA_TYPE_NAMES[t] for t in types]
        self.chart_type_var.set(current if current in names else (names[0] if names else ""))

    def _selected_type(self) -> str | None:
        for key, name in DATA_TYPE_NAMES.items():
            if name == self.chart_type_var.get():
                return key
        return None

    def redraw_charts(self) -> None:
        data_type = self._selected_type()
        for key, (fig, canvas) in self.figures.items():
            CHARTS[key](fig, self.rows, data_type)
            canvas.draw()

    # --------------------------------------------------------------------- экспорт
    def export_csv(self) -> None:
        if not self.rows:
            messagebox.showinfo("Нет данных", "Сначала выполните эксперимент.", parent=self.root)
            return
        RESULTS_DIR.mkdir(exist_ok=True)
        path = filedialog.asksaveasfilename(
            parent=self.root, title="Сохранить результаты", defaultextension=".csv",
            initialdir=str(RESULTS_DIR),
            initialfile=f"results_{datetime.now():%Y%m%d_%H%M%S}.csv",
            filetypes=[("CSV", "*.csv")])
        if not path:
            return
        try:
            saved = save_csv(self.rows, path)
        except OSError as exc:
            messagebox.showerror("Ошибка сохранения", f"Не удалось сохранить файл: {exc}", parent=self.root)
            return
        self.status_var.set(f"CSV сохранён: {saved}")

    def export_charts(self) -> None:
        if not self.rows:
            messagebox.showinfo("Нет данных", "Сначала выполните эксперимент.", parent=self.root)
            return
        RESULTS_DIR.mkdir(exist_ok=True)
        folder = filedialog.askdirectory(parent=self.root, title="Папка для графиков",
                                         initialdir=str(RESULTS_DIR))
        if not folder:
            return
        try:
            files = save_charts(self.rows, folder)
        except OSError as exc:
            messagebox.showerror("Ошибка сохранения", f"Не удалось сохранить графики: {exc}", parent=self.root)
            return
        self.status_var.set(f"Сохранено графиков: {len(files)} → {folder}")

    def _on_close(self) -> None:
        if self.is_running:
            if not messagebox.askokcancel("Выход", "Эксперимент ещё выполняется. Остановить и выйти?",
                                          parent=self.root):
                return
            self.stop_event.set()
            self.worker.join(timeout=5)
        if getattr(self, "_after_id", None):
            self.root.after_cancel(self._after_id)
        self.root.destroy()


def run() -> None:
    root = tk.Tk()
    ParallelSortApp(root)
    root.mainloop()
