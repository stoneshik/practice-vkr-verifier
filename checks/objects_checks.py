"""
Проверки объектов документа: рисунки, таблицы, подписи и корректность их оформления
"""
import re
from typing import Any, Dict, List, Sequence

from ..constants import FIGURE_CAPTION_RE, TABLE_CAPTION_RE, FORMULA_NUMBER_RE, COMMON_UNITS
from ..model import make_finding
from ..utils import next_nonempty_paragraph, previous_nonempty_paragraph, normalize_text, block_text, next_nonempty_block


def check_figures(blocks: Sequence[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Проверяет наличие и выравнивание подписей рисунков"""
    findings: List[Dict[str, Any]] = []
    for i, b in enumerate(blocks):
        if b["kind"] != "paragraph" or not b["has_drawing"]:
            continue
        next_p = next_nonempty_paragraph(blocks, i)
        if not next_p or not FIGURE_CAPTION_RE.match(next_p["text"]):
            findings.append(make_finding(
                "F11", "error",
                "Не найдена корректная подпись рисунка под изображением",
                next_p["text"] if next_p else None,
                "Рисунок N — Название рисунка",
                paragraph=b["paragraph_index"] + 1,
                recommendation="Добавить подпись рисунка под изображением по центру"
            ))
        elif next_p.get("alignment") not in {None, "CENTER"}:
            findings.append(make_finding(
                "F11", "warning",
                "Подпись рисунка должна быть выровнена по центру",
                next_p.get("alignment"), "CENTER",
                paragraph=next_p["paragraph_index"] + 1,
                recommendation="Выровнять подпись рисунка по центру"
            ))
    return findings


def check_tables(blocks: Sequence[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Проверяет наличие подписей таблиц и факт использования объектов table"""
    findings: List[Dict[str, Any]] = []
    for i, b in enumerate(blocks):
        if b["kind"] != "table":
            continue
        prev_p = previous_nonempty_paragraph(blocks, i)
        if not prev_p:
            findings.append(make_finding(
                "F12", "error",
                "Не найдена подпись таблицы над таблицей", None,
                "Таблица N — Название таблицы",
                block=b["block_index"] + 1,
                recommendation="Добавить подпись таблицы над объектом таблицы"
            ))
            continue
        prev_text = normalize_text(prev_p["text"])
        if prev_text.lower().startswith("продолжение таблицы") or prev_text.lower().startswith("листинг"):
            continue
        if not TABLE_CAPTION_RE.match(prev_text):
            findings.append(make_finding(
                "F12", "error",
                "Не найдена корректная подпись таблицы над таблицей",
                prev_p["text"],
                "Таблица N — Название таблицы",
                block=b["block_index"] + 1,
                recommendation="Добавить подпись таблицы над объектом таблицы"
            ))
    for i, b in enumerate(blocks):
        if b["kind"] != "paragraph":
            continue
        txt = block_text(b)
        if not TABLE_CAPTION_RE.match(txt):
            continue
        nxt = next_nonempty_block(blocks, i)
        if nxt is None:
            findings.append(make_finding(
                "F12", "warning",
                "После подписи таблицы не найден объект таблицы Word",
                "eof", "table",
                paragraph=b["paragraph_index"] + 1,
                recommendation="Убедиться, что таблица вставлена именно как объект таблицы"
            ))
        elif nxt["kind"] == "paragraph" and nxt.get("has_drawing"):
            findings.append(make_finding(
                "F12", "warning",
                "Подпись таблицы обнаружена рядом с изображением, а не объектом таблицы",
                nxt["text"], "Нативная таблица Word",
                paragraph=b["paragraph_index"] + 1,
                recommendation="Заменить растровое изображение на объект таблицы Word"
            ))
        elif nxt["kind"] != "table":
            findings.append(make_finding(
                "F12", "warning",
                "После подписи таблицы не найден объект таблицы Word",
                nxt["kind"], "table",
                paragraph=b["paragraph_index"] + 1,
                recommendation="Убедиться, что таблица вставлена именно как объект таблицы"
            ))
    return findings


def check_formulae(blocks: Sequence[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Проверяет базовое оформление формул и связанных с ними пояснений."""
    findings: List[Dict[str, Any]] = []
    for i, b in enumerate(blocks):
        if b["kind"] != "paragraph" or not b["has_omml"]:
            continue
        text = block_text(b)
        prev_p = previous_nonempty_paragraph(blocks, i)
        next_p = next_nonempty_paragraph(blocks, i)
        is_standalone = len(text) <= 24 or re.fullmatch(r"[\s,.;:()\-–—\dIVXLivxlc]+", text) is not None
        if not is_standalone:
            continue
        if next_p is not None:
            next_text = normalize_text(next_p["text"])
            if next_text and not next_text.lower().startswith("где") and not next_p.get("is_heading_like"):
                findings.append(make_finding(
                    "F17", "warning",
                    "После формулы ожидается строка с пояснениями, начинающаяся со слова «где»", next_p["text"],
                    "Пояснения после формулы, если они есть, начинаются со слова «где»",
                    paragraph=b["paragraph_index"] + 1,
                    recommendation="Проверить оформление пояснений после формулы")
                )
        if not FORMULA_NUMBER_RE.search(b["raw_text"]):
            findings.append(make_finding(
                "F17", "warning",
                "Не удалось подтвердить нумерацию формулы в конце строки", b["raw_text"],
                "Номер формулы в круглых скобках справа",
                paragraph=b["paragraph_index"] + 1,
                recommendation="Добавить номер формулы в круглых скобках по правому краю"
            ))
        plain = b["raw_text"]
        for unit in COMMON_UNITS:
            if plain and re.search(rf"\d+ {re.escape(unit)}\b", plain):
                findings.append(make_finding(
                    "F17", "warning",
                    "Между числом и единицей измерения должен быть неразрывный пробел",
                    f"обычный пробел перед {unit}",
                    f"неразрывный пробел перед {unit}",
                    paragraph=b["paragraph_index"] + 1,
                    recommendation="Заменить обычный пробел на неразрывный между числом и единицей измерения"
                ))
                break
    return findings
