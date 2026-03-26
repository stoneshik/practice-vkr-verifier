"""
Вспомогательные функции: нормализация текста, работа со строками, определение типов элементов и общие утилиты
"""
import re
from typing import Any, Dict, Optional, Sequence

from docx.enum.text import WD_ALIGN_PARAGRAPH

from .constants import (
    STRUCTURAL_SECTIONS,
    HEADING_NUMBER_RE,
    LIST_MARKER_RE
)


def normalize_text(text: Optional[str]) -> str:
    """Нормализует текст: схлопывает пробелы и обрезает края"""
    if not text:
        return ""
    return re.sub(r"\s+", " ", text).strip()


def compact_upper(text: Optional[str]) -> str:
    """Возвращает нормализованный текст в верхнем регистре"""
    return normalize_text(text).upper()


def is_uppercase(text: str) -> bool:
    """Проверяет, что текст состоит из прописных букв"""
    t = normalize_text(text)
    return bool(t) and any(ch.isalpha() for ch in t) and t == t.upper()


def is_heading_like(text: str) -> bool:
    """Определяет, похож ли текст на заголовок"""
    t = normalize_text(text)
    if not t:
        return False
    if compact_upper(t) in STRUCTURAL_SECTIONS:
        return True
    return bool(
        HEADING_NUMBER_RE.match(t) or t.startswith("ГЛАВА ") or t.startswith("РАЗДЕЛ ") or t.startswith("ПРИЛОЖЕНИЕ ")
    )


def is_list_like(text: str) -> bool:
    """Определяет, похож ли абзац на элемент списка"""
    t = normalize_text(text)
    if not t:
        return False
    if re.match(r"^\d+\.\s+[А-ЯA-ZА-ЯЁ]", t):
        return False
    return bool(LIST_MARKER_RE.match(t))


def list_marker_type(text: str) -> Optional[str]:
    """Определяет тип маркера списка: bullet, number, alpha или unknown"""
    t = normalize_text(text)
    if not is_list_like(t):
        return None
    m = LIST_MARKER_RE.match(t)
    if not m:
        return None
    marker = m.group(2)
    if marker in {"-", "•", "*"}:
        return "bullet"
    if re.fullmatch(r"\d+\)", marker):
        return "number"
    if re.fullmatch(r"[a-zа-я]\)", marker, re.IGNORECASE):
        return "alpha"
    return "unknown"


def length_to_mm(length: Any) -> Optional[float]:
    """Преобразует объект длины python-docx в миллиметры"""
    if length is None:
        return None
    try:
        return round(float(length.mm), 2)
    except Exception:
        return None


def get_alignment_name(value: Any) -> Optional[str]:
    """Преобразует значение выравнивания в строковое имя"""
    if value is None:
        return None
    try:
        if value == WD_ALIGN_PARAGRAPH.LEFT:
            return "LEFT"
        if value == WD_ALIGN_PARAGRAPH.CENTER:
            return "CENTER"
        if value == WD_ALIGN_PARAGRAPH.RIGHT:
            return "RIGHT"
        if value == WD_ALIGN_PARAGRAPH.JUSTIFY:
            return "JUSTIFY"
    except Exception:
        pass
    return str(value)


def block_text(block: Dict[str, Any]) -> str:
    """Возвращает нормализованный текст блока"""
    return normalize_text(block.get("text", ""))


def previous_nonempty_paragraph(blocks: Sequence[Dict[str, Any]], idx: int) -> Optional[Dict[str, Any]]:
    """Возвращает предыдущий непустой абзац перед указанным блоком"""
    for i in range(idx - 1, -1, -1):
        b = blocks[i]
        if b["kind"] == "paragraph" and normalize_text(b["text"]):
            return b
    return None


def next_nonempty_paragraph(blocks: Sequence[Dict[str, Any]], idx: int) -> Optional[Dict[str, Any]]:
    """Возвращает следующий непустой абзац после указанного блока"""
    for i in range(idx + 1, len(blocks)):
        b = blocks[i]
        if b["kind"] == "paragraph" and normalize_text(b["text"]):
            return b
    return None


def next_nonempty_block(blocks: Sequence[Dict[str, Any]], idx: int) -> Optional[Dict[str, Any]]:
    """Возвращает следующий непустой блок после указанного индекса"""
    for i in range(idx + 1, len(blocks)):
        b = blocks[i]
        if b["kind"] == "table" or normalize_text(b.get("text", "")):
            return b
    return None
