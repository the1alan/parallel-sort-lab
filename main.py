"""Точка входа Parallel Sort Lab.

Запуск: python3 main.py

Защита `if __name__ == "__main__"` обязательна: на macOS multiprocessing
использует метод запуска «spawn», при котором дочерние процессы заново
импортируют главный модуль. Без защиты приложение запускалось бы рекурсивно.
Тяжёлые импорты (Tkinter, Matplotlib) выполняются только внутри main().
"""
import multiprocessing


def main() -> None:
    from ui.app import run

    run()


if __name__ == "__main__":
    multiprocessing.freeze_support()
    main()
