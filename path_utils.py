"""
Утиліти для визначення шляхів у звичайному та bundled (PyInstaller) режимі.

- get_base_path() — шлях до bundled ресурсів (templates/, emblem.png)
- get_app_dir()   — шлях до робочих файлів (output/, app.db, xlsx-файли)
"""
import sys
import os


def get_base_path() -> str:
    """Повертає шлях до bundled ресурсів.
    У PyInstaller — sys._MEIPASS, інакше — директорія цього файлу."""
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))


def get_app_dir() -> str:
    """Повертає робочу директорію додатку (де лежить exe або скрипт).
    Сюди зберігаються app.db, output/, xlsx-файли тощо."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))
