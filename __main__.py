"""
Точка входа при запуске пакета через `python -m verifier`. Делегирует выполнение CLI
"""
from .cli import main

if __name__ == "__main__":
    raise SystemExit(main())
