"""
Специальные проверки: формулы, списки, термины, сокращения и список источников
"""
import re
from typing import Any, Dict, List, Sequence

from ..constants import APPENDIX_RE, SOURCE_ENTRY_RE
from ..model import make_finding
from ..utils import (
    block_text,
    compact_upper,
    is_uppercase,
    previous_nonempty_paragraph,
    is_heading_like
)


def check_appendices(
    blocks: Sequence[Dict[str, Any]], toc_lookup: Dict[str, int], page_lookup: Dict[str, Dict[int, int]]
) -> List[Dict[str, Any]]:
    """Проверяет оформление и размещение приложений"""
    findings: List[Dict[str, Any]] = []
    paragraphs = [b for b in blocks if b["kind"] == "paragraph"]
    found = []
    for b in paragraphs:
        m = APPENDIX_RE.match(block_text(b))
        if m:
            found.append((b, m.group(1).upper()))
    if not found:
        findings.append(make_finding(
            "F13", "info",
            "Приложения не обнаружены", "none",
            "Приложения, если они предусмотрены работой",
            recommendation="Проверить, нужны ли приложения по структуре конкретной ВКР"
        ))
        return findings
    for p, letter in found:
        page = toc_lookup.get(compact_upper(p["text"])) or page_lookup["paragraphs"].get(p["paragraph_index"])
        if not is_uppercase(p["text"]):
            findings.append(make_finding(
                "F13", "warning",
                "Заголовок приложения должен быть набран прописными буквами",
                p["text"], f"ПРИЛОЖЕНИЕ {letter}",
                paragraph=p["paragraph_index"] + 1, page=page,
                recommendation="Привести заголовок приложения к прописным буквам"
            ))
        if p.get("alignment") not in {None, "CENTER"}:
            findings.append(make_finding(
                "F13",
                "warning",
                "Заголовок приложения должен располагаться по центру",
                p.get("alignment"),
                "CENTER",
                paragraph=p["paragraph_index"] + 1, page=page,
                recommendation="Выровнять заголовок приложения по центру"
            ))
        prev_p = previous_nonempty_paragraph(blocks, p["block_index"])
        prev_page = page_lookup["paragraphs"].get(prev_p["paragraph_index"]) if prev_p else None
        if not (p.get("page_break_before") or p.get("has_page_break") or p.get("has_rendered_page_break")) and not (page and prev_page and page > prev_page):
            findings.append(make_finding(
                "F13", "warning",
                "Приложение не выглядит начатым с новой страницы", p["text"],
                "Новая страница перед приложением",
                paragraph=p["paragraph_index"] + 1, page=page,
                recommendation="Добавить разрыв страницы перед приложением"
            ))
    return findings


def check_abbreviations_list(blocks: Sequence[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Проверяет оформление списка сокращений."""
    findings: List[Dict[str, Any]] = []
    paragraphs = [b for b in blocks if b["kind"] == "paragraph"]
    start = next((i for i, b in enumerate(paragraphs) if compact_upper(b["text"]) == "СПИСОК СОКРАЩЕНИЙ И УСЛОВНЫХ ОБОЗНАЧЕНИЙ"), None)
    if start is None:
        return findings
    items: List[Dict[str, Any]] = []
    for i in range(start + 1, len(paragraphs)):
        b = paragraphs[i]
        txt = block_text(b)
        if not txt:
            continue
        if is_heading_like(txt) and i > start + 1:
            break
        items.append(b)
    if not items:
        findings.append(make_finding(
            "F15", "warning",
            "В списке сокращений не найдены элементы", "none",
            "Сокращение — расшифровка",
            paragraph=paragraphs[start]["paragraph_index"] + 1,
            recommendation="Оформить сокращения столбцом"
        ))
        return findings
    if not any(re.search(r"[–—-]", b["text"]) for b in items):
        findings.append(make_finding(
            "F15", "warning",
            "В списке сокращений не найдены элементы с расшифровкой через тире",
            "none","Сокращение — расшифровка",
            paragraph=paragraphs[start]["paragraph_index"] + 1,
            recommendation="Оформить сокращения столбцом: сокращение слева, расшифровка через тире"
        ))
    for b in items:
        txt = block_text(b)
        if not txt:
            continue
        if txt.endswith((".", ",", ";", ":")):
            findings.append(make_finding(
                "F15", "warning",
                "В строке списка сокращений не должно быть знаков препинания в конце", txt,
                "Без точки, запятой и т.п. в конце",
                paragraph=b["paragraph_index"] + 1,
                recommendation="Убрать завершающие знаки препинания"
            ))
    return findings


def check_terms_list(blocks: Sequence[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Проверяет оформление раздела терминов и определений."""
    findings: List[Dict[str, Any]] = []
    paragraphs = [b for b in blocks if b["kind"] == "paragraph"]
    heading = next((b for b in paragraphs if compact_upper(b["text"]) == "ТЕРМИНЫ И ОПРЕДЕЛЕНИЯ"), None)
    if heading is None:
        return findings

    start_block = heading["block_index"] + 1
    items: List[Dict[str, Any]] = []
    tables: List[Dict[str, Any]] = []
    for i in range(start_block, len(blocks)):
        blk = blocks[i]
        if blk["kind"] == "paragraph":
            txt = block_text(blk)
            if not txt:
                continue
            if is_heading_like(txt):
                break
            items.append(blk)
        else:
            tables.append(blk)

    if any(
        t["cols"] >= 2
        and "ТЕРМИН" in compact_upper(" ".join(" ".join(r) for r in t.get("text_rows", [])))
        and "ОПРЕДЕЛЕНИ" in compact_upper(" ".join(" ".join(r) for r in t.get("text_rows", [])))
        for t in tables
    ):
        return findings

    term_like = [b for b in items if re.search(r"[–—-]", b["text"])]
    if not term_like:
        findings.append(make_finding(
            "F16", "warning",
            "В разделе «Термины и определения» не найдено табличное или колонковое оформление",
            "none", "Двухколоночная таблица или строки «Термин — определение»",
            paragraph=heading["paragraph_index"] + 1,
            recommendation="Оформить раздел таблицей либо колонкой с тире"
        ))
        return findings
    for b in term_like:
        if block_text(b).endswith((".", ",", ";", ":")):
            findings.append(make_finding(
                "F16", "warning",
                "Строка термина не должна заканчиваться знаками препинания",
                b["text"], "Без точки и запятой в конце",
                paragraph=b["paragraph_index"] + 1,
                recommendation="Убрать завершающие знаки препинания"
            ))
    return findings


def check_sources(blocks: Sequence[Dict[str, Any]], toc_lookup: Dict[str, int]) -> List[Dict[str, Any]]:
    """Проверяет наличие и нумерацию списка использованных источников."""
    findings: List[Dict[str, Any]] = []
    paragraphs = [b for b in blocks if b["kind"] == "paragraph"]
    start = next((i for i, b in enumerate(paragraphs) if compact_upper(b["text"]) == "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"), None)
    if start is None:
        findings.append(make_finding(
            "F18", "error",
            "Не найден заголовок списка использованных источников", "missing",
            "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ",
            page=toc_lookup.get("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"),
            recommendation="Добавить раздел со списком использованных источников"
        ))
        return findings
    entries: List[Dict[str, Any]] = []
    for i in range(start + 1, len(paragraphs)):
        blk = paragraphs[i]
        txt = block_text(blk)
        if not txt:
            continue
        if is_heading_like(txt):
            break
        entries.append(blk)
    if not entries:
        findings.append(make_finding(
            "F18", "error",
            "В списке источников не найдены элементы списка", "none",
            "Нумерованный список источников",
            paragraph=paragraphs[start]["paragraph_index"] + 1,
            page=toc_lookup.get("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"),
            recommendation="Оформить источники как нумерованный список Word или с явной нумерацией"
        ))
        return findings

    explicit_numbers: List[int] = []
    number_like_count = 0
    for blk in entries:
        txt = block_text(blk)
        m = SOURCE_ENTRY_RE.match(txt)
        if m:
            explicit_numbers.append(int(m.group(1)))
            number_like_count += 1
        elif blk.get("is_numbered"):
            number_like_count += 1

    if number_like_count == 0:
        findings.append(make_finding(
            "F18", "error",
            "В списке источников не найдены нумерованные элементы", "none",
            "1. Источник ...",
            paragraph=paragraphs[start]["paragraph_index"] + 1,
            page=toc_lookup.get("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"),
            recommendation="Оформить источники арабскими цифрами с точкой"
        ))
    elif explicit_numbers and explicit_numbers != list(range(1, len(explicit_numbers) + 1)):
        findings.append(make_finding(
            "F18", "warning",
            "Нумерация источников выглядит непоследовательной",
            explicit_numbers, list(range(1, len(explicit_numbers) + 1)),
            paragraph=paragraphs[start]["paragraph_index"] + 1,
            page=toc_lookup.get("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"),
            recommendation="Привести нумерацию источников к последовательной"
        ))
    return findings
