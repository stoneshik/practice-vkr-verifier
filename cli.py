"""
Обрабатывает аргументы командной строки, запускает анализ DOCX и выводит результат в JSON
"""
import argparse
import json
import time
from pathlib import Path

from .report import analyze_docx


def main() -> int:
    """Обрабатывает аргументы CLI, запускает анализ и печатает JSON-отчёт"""
    parser = argparse.ArgumentParser(description="DOCX verifier")
    parser.add_argument("file", help="Path to .docx file")
    parser.add_argument("--pretty", action="store_true", help="Pretty print JSON")
    args = parser.parse_args()
    start = time.time()
    try:
        path = Path(args.file).expanduser().resolve()
        report = analyze_docx(path)
        report["meta"] = {"analysis_time_ms": int((time.time() - start) * 1000)}
        print(json.dumps(report, ensure_ascii=False, indent=2 if args.pretty else None))
        return 0
    except Exception as exc:
        error_report = {
            "status": "error",
            "message": str(exc),
            "meta": {"analysis_time_ms": int((time.time() - start) * 1000)},
        }
        print(json.dumps(error_report, ensure_ascii=False, indent=2 if args.pretty else None))
        return 1
