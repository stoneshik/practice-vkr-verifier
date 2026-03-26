"""
Проверки файла: валидность DOCX, структура архива, MIME-тип и базовые требования к входному файлу
"""
import mimetypes
import zipfile
from pathlib import Path
from typing import Any, Dict, List

from ..constants import ALLOWED_EXTENSIONS, ALLOWED_MIME_TYPES
from ..model import make_finding


def check_file_format(path: Path) -> List[Dict[str, Any]]:
    """Проверяет формат файла, MIME и базовую целостность DOCX-архива"""
    findings: List[Dict[str, Any]] = []
    guessed_mime, _ = mimetypes.guess_type(str(path))
    if path.suffix.lower() not in ALLOWED_EXTENSIONS:
        findings.append(make_finding(
            "F01/F02", "error",
            "Недопустимое расширение файла",
            path.suffix.lower(), ".docx",
            recommendation="Разрешить загрузку только файлов .docx"
        ))
    if guessed_mime not in ALLOWED_MIME_TYPES:
        findings.append(make_finding(
            "F02", "error",
            "Недопустимый MIME-тип файла",
            guessed_mime, next(iter(ALLOWED_MIME_TYPES)),
            recommendation="Проверять MIME-тип на стороне загрузки"
        ))
    if not zipfile.is_zipfile(path):
        findings.append(make_finding(
            "F01/F02", "error",
            "Файл не является корректным DOCX архивом",
            "not_zip", "docx_zip",
            recommendation="Отклонить повреждённый или неверный файл"
        ))
        return findings
    try:
        with zipfile.ZipFile(path) as zf:
            names = set(zf.namelist())
            needed = {"[Content_Types].xml", "word/document.xml"}
            if not needed.issubset(names):
                findings.append(make_finding(
                    "F01/F02", "error",
                    "DOCX структура архива неполная",
                    sorted(list(names))[:10], sorted(list(needed)),
                    recommendation="Проверять целостность DOCX перед анализом"
                ))
    except Exception as exc:
        findings.append(make_finding(
            "F01/F02", "error",
            "Не удалось прочитать DOCX архив",
            str(exc), "Корректный DOCX",
            recommendation="Отклонить файл и сообщить пользователю о повреждении документа"
        ))
    return findings
