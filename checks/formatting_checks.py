"""
Проверки форматирования: шрифт, интервалы, выравнивание, отступы и параметры страницы
"""
from typing import Any, Dict, List, Sequence

from ..constants import SOURCE_ENTRY_RE, HEADING_NUMBER_RE
from ..model import make_finding
from ..utils import block_text, compact_upper, is_uppercase


def check_body_formatting(
    blocks: Sequence[Dict[str, Any]], sections: Sequence[Dict[str, Any]]
) -> List[Dict[str, Any]]:
    """Проверяет форматирование основного текста и параметры полей страницы"""
    findings: List[Dict[str, Any]] = []
    for b in blocks:
        if b["kind"] != "paragraph":
            continue
        text = block_text(b)
        if not text:
            continue
        low = text.lower()
        if b["is_heading_like"] or b["is_caption"] or b["is_list_like"] or b["has_omml"]:
            continue
        if (compact_upper(text) == "СОДЕРЖАНИЕ" or low.startswith("где")
                or low.startswith("продолжение таблицы") or low.startswith("листинг")):
            continue
        if SOURCE_ENTRY_RE.match(text) or b.get("is_numbered"):
            continue

        if b.get("font_name") and "times new roman" not in b["font_name"].lower():
            findings.append(make_finding(
                "F07", "error",
                "Основной текст должен быть набран шрифтом Times New Roman",
                b["font_name"], "Times New Roman",
                paragraph=b["paragraph_index"] + 1,
                recommendation="Применить шрифт Times New Roman ко всему основному тексту"
            ))
        if b.get("font_size_pt") is not None and b["font_size_pt"] < 12:
            findings.append(make_finding(
                "F07", "error",
                "Размер основного текста меньше минимально допустимого",
                f"{b['font_size_pt']} pt", "Не менее 12 pt",
                paragraph=b["paragraph_index"] + 1,
                recommendation="Увеличить кегль основного текста до 12–14 pt"
            ))
        elif b.get("font_size_pt") is not None and 12 <= b["font_size_pt"] < 14:
            findings.append(make_finding(
                "F07", "warning",
                "Кегль основного текста меньше 14 pt, но не нарушает нижнюю границу",
                f"{b['font_size_pt']} pt", "14 pt, допускается не менее 12 pt",
                paragraph=b["paragraph_index"] + 1,
                recommendation="По возможности привести кегль к 14 pt"
            ))
        if b.get("color") and b["color"].upper() not in {"000000", "AUTO"}:
            findings.append(make_finding(
                "F07", "warning",
                "Цвет основного текста отличается от чёрного", b["color"],
                "000000 / black",
                paragraph=b["paragraph_index"] + 1,
                recommendation="Установить цвет текста чёрным"
            ))
        if b.get("line_spacing") is not None:
            try:
                if abs(float(b["line_spacing"]) - 1.5) > 0.15:
                    findings.append(make_finding(
                        "F07", "error",
                        "Межстрочный интервал не соответствует полуторному",
                        b["line_spacing"], 1.5,
                        paragraph=b["paragraph_index"] + 1,
                        recommendation="Установить межстрочный интервал 1.5"
                    ))
            except Exception:
                pass
        if b.get("first_line_indent_mm") is not None and b["first_line_indent_mm"] < 10.0:
            findings.append(make_finding(
                "F07", "error",
                "Абзацный отступ меньше 1,25 см",
                f"{b['first_line_indent_mm']} mm", "12.5 mm",
                paragraph=b["paragraph_index"] + 1,
                recommendation="Установить абзацный отступ 1,25 см"
            ))
        if b.get("alignment") and b["alignment"] != "JUSTIFY":
            findings.append(make_finding(
                "F07", "warning",
                "Основной текст должен быть выровнен по ширине",
                b["alignment"], "JUSTIFY",
                paragraph=b["paragraph_index"] + 1,
                recommendation="Выравнивать основной текст по ширине"
            ))

    for sec in sections:
        idx = sec["section_index"] + 1
        if sec["left_mm"] is not None and not (29.0 <= sec["left_mm"] <= 31.5):
            findings.append(make_finding(
                "F08", "error",
                "Левое поле не соответствует требованию",
                f"{sec['left_mm']} mm", "30 mm",
                section=idx,
                recommendation="Установить левое поле 30 мм"
            ))
        if sec["right_mm"] is not None and not (10.0 <= sec["right_mm"] <= 15.5):
            findings.append(make_finding(
                "F08", "error",
                "Правое поле не соответствует требованию",
                f"{sec['right_mm']} mm", "10–15 mm",
                section=idx,
                recommendation="Установить правое поле в диапазоне 10–15 мм"
            ))
        if sec["top_mm"] is not None and not (19.0 <= sec["top_mm"] <= 21.5):
            findings.append(make_finding(
                "F08", "error",
                "Верхнее поле не соответствует требованию",
                f"{sec['top_mm']} mm", "20 mm",
                section=idx,
                recommendation="Установить верхнее поле 20 мм"
            ))
        if sec["bottom_mm"] is not None and not (19.0 <= sec["bottom_mm"] <= 21.5):
            findings.append(make_finding(
                "F08", "error",
                "Нижнее поле не соответствует требованию",
                f"{sec['bottom_mm']} mm", "20 mm",
                section=idx,
                recommendation="Установить нижнее поле 20 мм"
            ))
    return findings


def check_page_numbering(sections: Sequence[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Проверяет наличие и базовые признаки оформления нумерации страниц"""
    findings: List[Dict[str, Any]] = []
    found_any = any(sec.get("footer_has_page_field") or sec.get("header_has_page_field") for sec in sections)
    if not found_any:
        findings.append(make_finding(
            "F10", "warning",
            "Не удалось подтвердить наличие номера страницы в колонтитуле",
            "not_detected", "Номера страниц в центре нижнего поля",
            recommendation="Проверить вставку номера страницы в нижний колонтитул"
        ))
    for sec in sections:
        if sec.get("footer_has_page_field") and sec.get("footer_alignment") not in {None, "CENTER"}:
            findings.append(make_finding(
                "F10", "warning",
                "Номер страницы должен располагаться по центру нижнего поля",
                sec.get("footer_alignment"), "CENTER",
                section=sec["section_index"] + 1,
                recommendation="Выровнять номер страницы по центру"
            ))
    return findings


def check_heading_specific_rules(blocks: Sequence[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Проверяет дополнительные требования к оформлению заголовков."""
    findings: List[Dict[str, Any]] = []
    for b in blocks:
        if b["kind"] != "paragraph":
            continue
        text = block_text(b)
        if not text or not b["is_heading_like"]:
            continue

        m = HEADING_NUMBER_RE.match(text)
        level = (m.group(1).count(".") + 1) if m else 1

        if "\n" in b["raw_text"]:
            findings.append(make_finding(
                "F09", "warning",
                "Заголовок не должен содержать переносы строк", text,
                "Одна строка без переноса слов",
                paragraph=b["paragraph_index"] + 1,
                recommendation="Убрать ручные переносы строк из заголовка"
            ))

        if b.get("contains_abbreviation"):
            findings.append(make_finding(
                "F09", "warning",
                "В заголовке обнаружена аббревиатура", text,
                "Заголовок без аббревиатур",
                paragraph=b["paragraph_index"] + 1,
                recommendation="Переформулировать заголовок без сокращений"
            ))

        if text.endswith("."):
            findings.append(make_finding(
                "F09", "warning",
                "Точка в конце заголовка не допускается", text,
                "Без точки в конце заголовка",
                paragraph=b["paragraph_index"] + 1,
                recommendation="Убрать точку в конце заголовка")
            )

        if b.get("first_line_indent_mm") is not None and b["first_line_indent_mm"] < 10.0:
            findings.append(make_finding(
                "F09", "warning",
                "Заголовок должен начинаться с абзацного отступа",
                f"{b['first_line_indent_mm']} mm",
                "12.5 mm",
                paragraph=b["paragraph_index"] + 1,
                recommendation="Установить абзацный отступ заголовка"
            ))

        if level == 1:
            is_sentence_case = bool(text[:1]) and text[:1].isupper() and not is_uppercase(text)
            is_upper = is_uppercase(text)
            bold = b.get("bold")

            if not (is_upper or is_sentence_case):
                findings.append(make_finding(
                    "F09", "warning",
                    "Заголовок первого уровня не соответствует допустимому регистру", text,
                    "Прописные буквы либо стандартный регистр с заглавной первой буквой",
                    paragraph=b["paragraph_index"] + 1,
                    recommendation="Привести заголовок первого уровня к одному из допустимых шаблонов"
                ))

            if is_upper and bold is False:
                findings.append(make_finding(
                    "F09", "warning",
                    "Заголовок первого уровня в прописных буквах должен быть полужирным", text,
                    "Полностью прописные буквы с полужирным начертанием",
                    paragraph=b["paragraph_index"] + 1,
                    recommendation="Сделать заголовок полужирным либо использовать стандартный регистр"
                ))
        else:
            if b.get("font_name") and "times new roman" not in b["font_name"].lower():
                findings.append(make_finding(
                    "F09", "warning",
                    "Заголовок оформлен шрифтом, отличным от Times New Roman",
                    b["font_name"], "Times New Roman",
                    paragraph=b["paragraph_index"] + 1,
                    recommendation="Привести шрифт заголовка к Times New Roman"
                ))
            if b.get("font_size_pt") is not None and b["font_size_pt"] < 12:
                findings.append(make_finding(
                    "F09", "warning",
                    "Подчинённый заголовок имеет слишком маленький кегль",
                    f"{b['font_size_pt']} pt", "Не меньше 12 pt",
                    paragraph=b["paragraph_index"] + 1,
                    recommendation="Увеличить кегль заголовка"
                ))
    return findings

def check_lists(blocks: Sequence[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Проверяет единообразие и пунктуацию в перечислениях"""
    findings: List[Dict[str, Any]] = []
    i = 0
    while i < len(blocks):
        b = blocks[i]
        if b["kind"] != "paragraph" or not b["is_list_like"]:
            i += 1
            continue
        group = [b]
        j = i + 1
        while j < len(blocks) and blocks[j]["kind"] == "paragraph" and blocks[j]["is_list_like"]:
            group.append(blocks[j])
            j += 1
        marker_types = {x.get("list_marker_type") for x in group if x.get("list_marker_type")}
        if len(marker_types) > 1:
            findings.append(make_finding(
                "F19", "warning",
                "В пределах одного списка обнаружены разные типы маркеров/нумерации",
                sorted(marker_types), "Единообразный тип маркера или нумерации в списке",
                paragraph=group[0]["paragraph_index"] + 1,
                recommendation="Привести маркеры списка к одному стилю"
            ))
        for idx, item in enumerate(group):
            txt = block_text(item)
            if not txt:
                continue
            if idx == len(group) - 1:
                if not (txt.endswith(".") or txt.endswith(")") or txt.endswith("]")):
                    findings.append(make_finding(
                        "F19", "info",
                        "Последний элемент списка обычно завершается точкой", txt,
                        "Точка или согласованный завершающий знак",
                        paragraph=item["paragraph_index"] + 1,
                        recommendation="Проверить оформление знаков препинания в списке"
                    ))
            else:
                if not (txt.endswith(";") or txt.endswith(",") or txt.endswith(".")):
                    findings.append(make_finding(
                        "F19", "info",
                        "Элемент списка обычно завершается точкой с запятой", txt,
                        "; или другой согласованный знак препинания",
                        paragraph=item["paragraph_index"] + 1,
                        recommendation="Проверить знаки препинания в элементах списка"
                    ))
        i = j
    return findings
