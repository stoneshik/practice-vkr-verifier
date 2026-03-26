"""
Проверки структуры документа: разделы, заголовки, нумерация и соответствие оглавлению
"""
from typing import Any, Dict, List, Sequence, Tuple

from ..constants import (
    REQUIRED_SECTIONS,
    HEADING_NUMBER_RE,
    HEADING_NUMBER_BAD_TRAILING_DOT_RE,
    STRUCTURAL_SECTIONS,
    OPTIONAL_SECTIONS
)
from ..model import make_finding
from ..utils import (
    compact_upper,
    is_uppercase,
    block_text,
    previous_nonempty_paragraph,
    is_heading_like
)


def check_required_sections(blocks: Sequence[Dict[str, Any]], raw_paragraph_texts: Sequence[str]) -> List[Dict[str, Any]]:
    """Проверяет наличие обязательных разделов и их базовое оформление"""
    findings: List[Dict[str, Any]] = []
    paragraphs = [b for b in blocks if b["kind"] == "paragraph"]
    text_upper = [compact_upper(b["text"]) for b in paragraphs]
    raw_upper = [compact_upper(t) for t in raw_paragraph_texts if t]
    for section_name in REQUIRED_SECTIONS:
        exact = [i for i, t in enumerate(text_upper) if t == section_name]
        if exact:
            idx = exact[0]
            if not is_uppercase(paragraphs[idx]["text"]):
                findings.append(make_finding(
                    "F04", "warning",
                    f"Заголовок раздела должен быть написан прописными буквами: {section_name}",
                    paragraphs[idx]["text"], section_name,
                    paragraph=paragraphs[idx]["paragraph_index"] + 1,
                    recommendation="Привести заголовок к прописным буквам"
                ))
            continue
        if any(t == section_name for t in raw_upper):
            continue
        contained = [i for i, t in enumerate(text_upper) if section_name in t]
        if contained:
            idx = contained[0]
            findings.append(make_finding(
                "F04", "warning",
                f"Раздел найден не в точном виде: {section_name}",
                paragraphs[idx]["text"], section_name,
                paragraph=paragraphs[idx]["paragraph_index"] + 1,
                recommendation=f"Оформить заголовок как «{section_name}»"
            ))
        else:
            findings.append(make_finding(
                "F04", "error",
                f"Отсутствует обязательный раздел: {section_name}",
                "missing", section_name,
                recommendation=f"Добавить раздел «{section_name}»"
            ))
    return findings


def check_heading_numbers_and_start_pages(
    blocks: Sequence[Dict[str, Any]], toc_lookup: Dict[str, int], page_lookup: Dict[str, Dict[int, int]]
) -> List[Dict[str, Any]]:
    """Проверяет нумерацию заголовков и старт структурных элементов с новой страницы."""
    findings: List[Dict[str, Any]] = []
    paragraphs = [b for b in blocks if b["kind"] == "paragraph"]
    top_level: List[Tuple[int, Dict[str, Any]]] = []
    top_headings: List[Dict[str, Any]] = []
    checked_page_start = 0
    unverifiable_page_start = 0

    for b in paragraphs:
        text = block_text(b)
        if not text or not b["is_heading_like"]:
            continue
        m = HEADING_NUMBER_RE.match(text)
        if m:
            number = m.group(1)
            dot = m.group(2)
            level = number.count(".") + 1
            if level == 1:
                try:
                    top_level.append((int(number), b))
                    top_headings.append(b)
                except Exception:
                    pass
            if dot == "." and level == 1:
                findings.append(make_finding(
                    "F05", "error",
                    "Точка в конце номера заголовка глав не допускается", text,
                    "Например: 1 Заголовок, без точки после номера",
                    paragraph=b["paragraph_index"] + 1,
                    page=toc_lookup.get(compact_upper(text)),
                    recommendation="Удалить точку после номера заголовка"
                ))
            elif HEADING_NUMBER_BAD_TRAILING_DOT_RE.match(text):
                findings.append(make_finding(
                    "F05", "error",
                    "В номере заголовка не допускается точка в конце номера", text,
                    "Без точки после номера",
                    paragraph=b["paragraph_index"] + 1,
                    page=toc_lookup.get(compact_upper(text)),
                    recommendation="Убрать точку после номера заголовка"
                ))

        is_required_top = (
            compact_upper(text) in REQUIRED_SECTIONS
            or text.startswith("ПРИЛОЖЕНИЕ ")
            or (m is not None and m.group(1).count(".") == 0)
        )
        if not is_required_top:
            continue

        checked_page_start += 1

        if b.get("page_break_before") or b.get("has_page_break") or b.get("has_rendered_page_break"):
            continue

        prev_p = previous_nonempty_paragraph(blocks, b["block_index"])
        cur_page = page_lookup["paragraphs"].get(b["paragraph_index"])
        prev_page = page_lookup["paragraphs"].get(prev_p["paragraph_index"]) if prev_p else None
        if cur_page is not None and prev_page is not None and cur_page > prev_page:
            continue

        if cur_page is None:
            unverifiable_page_start += 1
            findings.append(make_finding(
                "F05",
                "info",
                "Не удалось надёжно подтвердить начало структурного заголовка с новой страницы",
                text,
                "Новая страница перед разделом/главой/приложением",
                paragraph=b["paragraph_index"] + 1,
                recommendation="Для точной проверки использовать рендер в PDF или явный разрыв страницы в DOCX"
            ))
            continue

        findings.append(make_finding(
            "F05", "warning",
            "Структурный заголовок не выглядит начатым с новой страницы",
            text,
            "Новая страница перед разделом/главой/приложением",
            paragraph=b["paragraph_index"] + 1,
            page=toc_lookup.get(compact_upper(text)) or cur_page,
            recommendation="Добавить разрыв страницы перед заголовком структурного элемента"
        ))

    if top_level:
        nums = [n for n, _ in top_level]
        if nums != list(range(1, len(nums) + 1)):
            findings.append(make_finding(
                "F05", "warning",
                "Нумерация глав выглядит не последовательной", nums, list(range(1, len(nums) + 1)),
                paragraph=top_headings[0]["paragraph_index"] + 1,
                page=page_lookup["paragraphs"].get(top_headings[0]["paragraph_index"]),
                recommendation="Проверить последовательность нумерации глав"
            ))

    if checked_page_start and checked_page_start == unverifiable_page_start:
        findings.append(make_finding(
            "F05", "info",
            "Проверка начала структурных заголовков с новой страницы выполнена только эвристически",
            actual="heuristic_only",
            expected="Явный разрыв страницы или достоверная пагинация",
            recommendation="Для строгой проверки пагинации использовать PDF-рендер"
        ))
    return findings


def check_toc(
    blocks: Sequence[Dict[str, Any]], raw_toc_entries: Sequence[Dict[str, Any]], page_lookup: Dict[str, Dict[int, int]]
) -> List[Dict[str, Any]]:
    """Проверяет наличие, структуру и согласованность оглавления."""
    findings: List[Dict[str, Any]] = []
    paragraphs = [b for b in blocks if b["kind"] == "paragraph"]
    toc_heading = next((b for b in paragraphs if compact_upper(b["text"]) == "СОДЕРЖАНИЕ"), None)

    if toc_heading is None and not raw_toc_entries:
        findings.append(make_finding(
            "F06", "error",
            "Не найдено оглавление («СОДЕРЖАНИЕ»)",
            "missing","Наличие оглавления",
            recommendation="Добавить оглавление"
        ))
        return findings

    if not raw_toc_entries:
        findings.append(make_finding(
            "F06", "warning",
            "Оглавление обнаружено не полностью или не удалось извлечь его строки",
            "missing_entries", "Строки оглавления с названиями и номерами страниц",
            paragraph=(toc_heading["paragraph_index"] + 1 if toc_heading else None),
            recommendation="Проверить, что оглавление создано как стандартное поле Word"
        ))
        return findings

    heading_by_title = {compact_upper(b["text"]): b for b in paragraphs if is_heading_like(b["text"])}

    missing_separator = [e["title"] for e in raw_toc_entries if not e.get("has_separator")]
    if missing_separator:
        findings.append(make_finding(
            "F06", "warning",
            "В оглавлении не для всех строк удалось подтвердить разделитель между названием и номером страницы",
            actual=missing_separator[:5],
            expected="Табуляция или иной разделитель между заголовком и номером страницы",
            paragraph=(toc_heading["paragraph_index"] + 1 if toc_heading else None),
            recommendation="Обновить автособираемое оглавление Word",
        ))

    missing_leader = [e["title"] for e in raw_toc_entries if not e.get("has_leader")]
    if missing_leader:
        findings.append(make_finding(
            "F06", "info",
            "В оглавлении не для всех строк удалось подтвердить точечный заполнитель",
            actual=missing_leader[:5],
            expected="Лидер-заполнитель между заголовком и номером страницы",
            paragraph=(toc_heading["paragraph_index"] + 1 if toc_heading else None),
            recommendation="Проверить настройки табуляции и лидеров в оглавлении",
        ))

    for e in raw_toc_entries:
        title_u = compact_upper(e.get("title"))
        page = str(e.get("page", "")).strip()
        heading = heading_by_title.get(title_u)

        if heading is None and title_u not in STRUCTURAL_SECTIONS and not title_u.startswith("ПРИЛОЖЕНИЕ "):
            findings.append(make_finding(
                "F06", "warning",
                "Пункт оглавления не найден среди заголовков документа",
                e.get("title"), "Совпадающий заголовок в теле документа",
                paragraph=(toc_heading["paragraph_index"] + 1 if toc_heading else None),
                recommendation="Сверить оглавление с текстом документа",
            ))

        if not page.isdigit() or int(page) <= 0:
            findings.append(make_finding(
                "F06", "warning",
                "Некорректный номер страницы в оглавлении",
                page, "Положительное целое число",
                paragraph=(toc_heading["paragraph_index"] + 1 if toc_heading else None),
                recommendation="Проверить формат номера страницы в оглавлении",
            ))
            continue

        if heading is not None:
            actual_page = page_lookup["paragraphs"].get(heading["paragraph_index"])
            if actual_page is not None and actual_page != int(page):
                findings.append(make_finding(
                    "F06", "warning",
                    "Номер страницы в оглавлении не совпадает с вычисленным положением заголовка",
                    actual=f"{e.get('title')} — TOC {page}, doc {actual_page}",
                    expected=f"{e.get('title')} — {page}",
                    paragraph=heading["paragraph_index"] + 1,
                    page=actual_page,
                    recommendation="Обновить поле оглавления Word"
                ))
    return findings

def check_optional_sections(blocks: Sequence[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Проверяет наличие необязательных разделов структуры ВКР"""
    findings: List[Dict[str, Any]] = []
    upper_texts = [compact_upper(b["text"]) for b in blocks if b["kind"] == "paragraph"]
    for section_name in OPTIONAL_SECTIONS:
        if section_name not in upper_texts:
            findings.append(make_finding(
                "F14", "info",
                f"Необязательный раздел не обнаружен: {section_name}",
                "missing", section_name,
                recommendation="Ничего не делать, если раздел не предусмотрен работой"
            ))
    return findings
